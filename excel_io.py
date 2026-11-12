"""Windows Excel COM writer; never controls KakaoTalk or the desktop UI."""
from __future__ import annotations

import shutil
from datetime import datetime, time
from pathlib import Path

from reservations import Reservation, validate

XL_UP = -4162
HEADERS = ("입실일", "예약번호", "예약경로", "상호/이름")


def _existing_excel(excel, path: Path):
    target = str(path.resolve()).casefold()
    for wb in excel.Workbooks:
        try:
            if str(Path(wb.FullName).resolve()).casefold() == target:
                return wb
        except Exception:
            continue
    return None


def _backup(path: Path) -> None:
    out = path.parent / "향그린_자동백업"
    out.mkdir(parents=True, exist_ok=True)
    target = out / f"{path.stem}_{datetime.now():%Y%m%d_%H%M%S_%f}{path.suffix}"
    shutil.copy2(path, target)  # If backup fails, abort rather than risk the only copy.
    files = sorted(out.glob(f"{path.stem}_*{path.suffix}"), key=lambda p: p.stat().st_mtime, reverse=True)
    for old in files[30:]:
        try: old.unlink()
        except OSError: pass


def _find_row(ws, last_row: int, booking: str) -> int | None:
    if last_row < 2: return None
    values = ws.Range(f"B2:B{last_row}").Value2
    if last_row == 2:
        values = ((values,) if not isinstance(values, tuple) else values)
    for offset, entry in enumerate(values):
        cell = entry[0] if isinstance(entry, tuple) else entry
        if cell is None: continue
        value = str(int(cell)) if isinstance(cell, (int, float)) and float(cell).is_integer() else str(cell).strip()
        if value == booking: return offset + 2
    return None


def apply_reservation(excel_path: str, res: Reservation, sheet_name: str = "데이터") -> str:
    """Write one booking. Called on a worker thread. Excel must be installed on Windows."""
    issue = validate(res)
    if issue:
        raise ValueError(f"{res.number}: {issue}")
    path = Path(excel_path).expanduser().resolve()
    if not path.is_file() or path.suffix.lower() not in (".xlsx", ".xlsm"):
        raise FileNotFoundError("올바른 엑셀 파일(.xlsx/.xlsm)을 선택하세요.")

    # Import at runtime so parser tests run without Windows or Excel.
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    excel = None
    wb = None
    created_excel = False
    opened_workbook = False
    try:
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            excel = win32com.client.DispatchEx("Excel.Application")
            created_excel = True
            excel.Visible = False
            excel.DisplayAlerts = False
        wb = _existing_excel(excel, path)
        if wb is None:
            wb = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)
            opened_workbook = True
        if wb.ReadOnly:
            raise PermissionError("엑셀이 읽기 전용입니다. 다른 PC에서 사용 중인지 확인하세요.")
        if not wb.Saved:
            raise RuntimeError("엑셀에 저장되지 않은 변경사항이 있습니다. 먼저 저장한 후 다시 실행하세요.")
        ws = wb.Worksheets(sheet_name)
        for i, header in enumerate(HEADERS, start=1):
            if str(ws.Cells(1, i).Value).strip() != header:
                raise ValueError(f"'{sheet_name}' 시트의 {i}열 제목이 예상과 다릅니다. 입력 중지.")
        last_row = max(1, int(ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row))
        found = _find_row(ws, last_row, res.number)
        if found is not None:
            note = str(ws.Cells(found, 12).Value or "")
            if not res.cancelled:
                return f"중복 예약 건너뜀: {res.number} (행 {found})"
            if "예약취소" in note:
                return f"이미 취소됨: {res.number} (행 {found})"
            _backup(path)
            ws.Cells(found, 5).ClearContents()  # E vehicle
            ws.Cells(found, 8).ClearContents()  # H room (excluded from counts)
            ws.Cells(found, 9).ClearContents()  # I amount
            ws.Cells(found, 12).Value = "예약취소"  # L note
            ws.Range(f"A{found}:M{found}").Font.Strikethrough = True
            ws.Range(f"A{found}:M{found}").Interior.Color = 13421823
            wb.Save()
            return f"예약취소 처리: {res.number} (행 {found})"
        if res.cancelled:
            return f"취소 대상 예약번호 없음: {res.number} (수동 확인 필요)"
        _backup(path)
        row = last_row + 1
        # Critical: real Excel dates, NOT ISO strings; the user's 합계 sheet uses COUNTIFS.
        ws.Cells(row, 1).Value = datetime.combine(res.check_in, time.min)
        ws.Cells(row, 1).NumberFormat = "yyyy-mm-dd"
        ws.Cells(row, 2).NumberFormat = "@"
        ws.Cells(row, 2).Value = res.number
        ws.Cells(row, 3).Value = res.channel
        ws.Cells(row, 4).Value = res.guest
        if res.vehicle is not None: ws.Cells(row, 5).Value = res.vehicle
        ws.Cells(row, 8).Value = res.room
        ws.Cells(row, 9).Value = res.amount
        ws.Cells(row, 9).NumberFormat = "#,##0"
        ws.Cells(row, 13).NumberFormat = "@"
        if res.phone: ws.Cells(row, 13).Value = res.phone
        wb.Save()
        return f"입력 완료: {res.source} / {res.number} / {res.amount:,}원 (행 {row})"
    finally:
        try:
            if opened_workbook and wb is not None:
                # Never save on error in finally: only explicit success paths save.
                wb.Close(SaveChanges=False)
        finally:
            if created_excel and excel is not None:
                excel.Quit()
            pythoncom.CoUninitialize()