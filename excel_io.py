"""Windows Excel COM writer; never controls KakaoTalk or the desktop UI."""
from __future__ import annotations

import shutil
from datetime import date, datetime
from zipfile import ZipFile
import posixpath
import xml.etree.ElementTree as ET
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



# Excel stores dates as serial day numbers. Sending a Python datetime through
# COM can shift midnight across time zones; Value2 accepts a timezone-free number.
def excel_date_serial(day: date, date1904: bool = False) -> int:
    epoch = date(1904, 1, 1) if date1904 else date(1899, 12, 30)
    return (day - epoch).days


def _empty(value) -> bool:
    return value is None or value == ""


def _cancel_complete(ws, row: int) -> bool:
    return (all(_empty(ws.Range(f"{col}{row}").Value2) for col in ("E", "H", "I"))
            and "예약취소" in str(ws.Range(f"L{row}").Value2 or ""))


def _mark_cancelled(ws, row: int) -> None:
    note = str(ws.Range(f"L{row}").Value2 or "")
    for col in ("E", "H", "I"):
        cell = ws.Range(f"{col}{row}")
        cell.ClearContents()
        if not _empty(cell.Value2):
            cell.Value2 = ""
    if "예약취소" not in note:
        ws.Range(f"L{row}").Value2 = (note + " | " if note else "") + "예약취소"
    band = ws.Range(f"A{row}:M{row}")
    band.Font.Strikethrough = True
    band.Interior.Color = 13421823
    if not _cancel_complete(ws, row):
        raise RuntimeError(f"취소 셀 변경 검증 실패 (행 {row}) · E/H/I/L열 확인 필요")


# Verify the saved workbook, not just the Excel COM object. Reading OOXML here
# does not rewrite the workbook, and preserves its formulas, macros and layout.
_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _saved_row(path: Path, sheet_name: str, row: int) -> dict:
    with ZipFile(path) as z:
        book = ET.fromstring(z.read("xl/workbook.xml"))
        sheet = next((x for x in book.findall(f"{_NS}sheets/{_NS}sheet") if x.get("name") == sheet_name), None)
        if sheet is None:
            raise RuntimeError("저장된 파일에서 대상 시트를 찾지 못했습니다.")
        rid = sheet.get(_REL + "id")
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        target = next(x.get("Target") for x in rels if x.get("Id") == rid)
        target = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
        root = ET.fromstring(z.read(target))
        strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            ss = ET.fromstring(z.read("xl/sharedStrings.xml"))
            strings = ["".join(t.text or "" for t in si.iter(_NS + "t")) for si in ss.findall(_NS + "si")]
        result = {}
        for record in root.findall(f"{_NS}sheetData/{_NS}row"):
            if record.get("r") != str(row):
                continue
            for cell in record.findall(_NS + "c"):
                ref = cell.get("r", "")
                col = ref.rstrip("0123456789")
                kind = cell.get("t")
                raw = cell.findtext(_NS + "v")
                if kind == "s" and raw is not None:
                    result[col] = strings[int(raw)]
                elif kind == "inlineStr":
                    result[col] = "".join(t.text or "" for t in cell.iter(_NS + "t"))
                else:
                    result[col] = raw
            break
        return result


def _verify_saved(path: Path, sheet_name: str, row: int, res: Reservation, serial: int | None = None) -> None:
    try:
        values = _saved_row(path, sheet_name, row)
    except Exception as exc:
        raise RuntimeError(f"엑셀 저장 후 파일 검증 실패 (행 {row}) · 파일을 다시 열어 확인하세요: {type(exc).__name__}") from exc
    if values.get("B") != res.number:
        raise RuntimeError(f"저장 파일 예약번호 검증 실패 (행 {row})")
    if res.cancelled:
        if not all(_empty(values.get(col)) for col in ("E", "H", "I")) or "예약취소" not in str(values.get("L") or ""):
            raise RuntimeError(f"저장 파일에 취소가 반영되지 않았습니다 (행 {row}) · E/H/I/L열 확인 필요")
    else:
        try:
            same = float(values.get("A", "nan")) == serial and float(values.get("I", "nan")) == res.amount
        except (ValueError, TypeError):
            same = False
        if not same or values.get("D") != res.guest or values.get("H") != res.room or values.get("C") != res.channel:
            raise RuntimeError(f"저장 파일 신규 예약 검증 실패 (행 {row}) · 날짜/이름/객실/금액/경로 확인 필요")
        if res.phone and values.get("M") != res.phone:
            raise RuntimeError(f"저장 파일 연락처 검증 실패 (행 {row})")
        if res.vehicle is not None and values.get("E") != str(res.vehicle):
            raise RuntimeError(f"저장 파일 차량 여부 검증 실패 (행 {row})")

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
            if "예약취소" in note and _cancel_complete(ws, found):
                _verify_saved(path, sheet_name, found, res)
                return f"이미 취소됨 (저장 검증 완료): {res.number} (행 {found})"
            _backup(path)
            _mark_cancelled(ws, found)
            wb.Save()
            if not wb.Saved or not _cancel_complete(ws, found):
                raise RuntimeError(f"취소 저장 후 셀 검증 실패 (행 {found})")
            _verify_saved(path, sheet_name, found, res)
            return f"예약취소 처리 (저장 검증 완료): {res.number} (행 {found})"
        if res.cancelled:
            return f"취소 대상 예약번호 없음: {res.number} (수동 확인 필요)"
        _backup(path)
        row = last_row + 1
        # Critical: real Excel dates, NOT ISO strings; the user's 합계 sheet uses COUNTIFS.
        serial = excel_date_serial(res.check_in, bool(wb.Date1904))
        ws.Range(f"A{row}").Value2 = serial
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
        if not wb.Saved:
            raise RuntimeError(f"신규 예약 저장 확인 실패 (행 {row})")
        _verify_saved(path, sheet_name, row, res, serial)
        return f"입력 완료 (저장 검증 완료): {res.source} / {res.number} / {res.amount:,}원 (행 {row})"
    finally:
        try:
            if opened_workbook and wb is not None:
                # Never save on error in finally: only explicit success paths save.
                wb.Close(SaveChanges=False)
        finally:
            if created_excel and excel is not None:
                excel.Quit()
            pythoncom.CoUninitialize()
