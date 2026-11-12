"""Safe reservation import UI. No synthetic keys, screen scraping or global hooks."""
from __future__ import annotations

import hashlib
import json
import os
import time
import tkinter as tk
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from excel_io import apply_reservation
from reservations import parse_many, validate

APP_DIR = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "HygreenAuto"
CONFIG = APP_DIR / "settings.json"
HISTORY = APP_DIR / "imported_files.json"


def read_json(path: Path) -> dict:
    try: return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError): return {}


def save_json(path: Path, obj: dict) -> None:
    APP_DIR.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(".tmp")
    temp.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")
    temp.replace(path)


def execute(path: str, content: str) -> list[str]:
    events = parse_many(content)
    if not events:
        raise ValueError("예약문자를 찾지 못했습니다. 전체 알림 내용을 복사해 주세요.")
    failed = [(e.number, validate(e)) for e in events if validate(e)]
    if failed:
        raise ValueError("정보가 불완전한 예약이 있습니다: " + ", ".join(f"{n}({why})" for n, why in failed))
    return [apply_reservation(path, e) for e in events]


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("향그린 예약 관리 · 안전 모드")
        self.root.geometry("840x690")
        self.root.minsize(680, 560)
        self.settings = read_json(CONFIG)
        self.history = read_json(HISTORY)
        self.excel = tk.StringVar(value=self.settings.get("excel", ""))
        self.folder = tk.StringVar(value=self.settings.get("folder", ""))
        self.watch = tk.BooleanVar(value=False)  # Never silently enable live writes at launch.
        self.pool = ThreadPoolExecutor(max_workers=1, thread_name_prefix="ExcelWriter")
        self.future = None
        self.work_file = None
        self.work_hash = None
        self._ui()
        self.root.after(400, self._poll)
        self.root.after(10000, self._watch)
        self.root.protocol("WM_DELETE_WINDOW", self._close)
        self._log("준비 완료. 카카오톡 화면에는 접근하지 않습니다.")

    def _ui(self):
        outer = ttk.Frame(self.root, padding=16)
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="향그린 예약 자동입력", font=("맑은 고딕", 19, "bold")).pack(anchor="w")
        ttk.Label(outer, text="NOL · 아고다 · 여기어때  |  중복 방지 · 취소 처리 · 엑셀 원본 백업").pack(anchor="w", pady=(2, 14))
        row = ttk.Frame(outer); row.pack(fill="x", pady=5)
        ttk.Label(row, text="예약자 엑셀", width=13).pack(side="left")
        ttk.Entry(row, textvariable=self.excel).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="파일 선택", command=self._select_excel).pack(side="left", padx=(7, 0))
        ttk.Label(outer, text="1. 카카오톡 알림을 복사해 아래에 붙여넣으세요. 여러 건을 한꺼번에 넣어도 됩니다.").pack(anchor="w", pady=(15, 5))
        self.text = scrolledtext.ScrolledText(outer, height=12, font=("맑은 고딕", 10), wrap="word")
        self.text.pack(fill="both", expand=True)
        buttons = ttk.Frame(outer); buttons.pack(fill="x", pady=8)
        ttk.Button(buttons, text="클립보드 붙여넣기", command=self._paste).pack(side="left")
        ttk.Button(buttons, text="예약 분석", command=self._preview).pack(side="left", padx=7)
        ttk.Button(buttons, text="엑셀 기록", command=self._manual).pack(side="left")
        ttk.Button(buttons, text="텍스트 파일 열기", command=self._open_text).pack(side="left", padx=7)
        ttk.Separator(outer, orient="horizontal").pack(fill="x", pady=9)
        ttk.Label(outer, text="2. 선택 사항: 별도 알림 전달 도구가 .txt 파일을 만드는 폴더를 자동 감시").pack(anchor="w")
        row = ttk.Frame(outer); row.pack(fill="x", pady=6)
        ttk.Entry(row, textvariable=self.folder).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="폴더 선택", command=self._select_folder).pack(side="left", padx=7)
        ttk.Checkbutton(row, text="자동 입력 켜기", variable=self.watch).pack(side="left")
        self.state = ttk.Label(outer, text="대기 중")
        self.state.pack(anchor="w", pady=5)
        self.log = scrolledtext.ScrolledText(outer, height=9, state="disabled", font=("맑은 고딕", 9))
        self.log.pack(fill="both", expand=True)

    def _log(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", time.strftime("[%H:%M:%S] ") + msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _select_excel(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if p:
            self.excel.set(p)
            self._save_settings()

    def _select_folder(self):
        p = filedialog.askdirectory()
        if p:
            self.folder.set(p)
            self._save_settings()

    def _save_settings(self):
        save_json(CONFIG, {"excel": self.excel.get(), "folder": self.folder.get()})

    def _paste(self):
        try:
            self.text.delete("1.0", "end")
            self.text.insert("1.0", self.root.clipboard_get())
            self._preview()
        except tk.TclError:
            messagebox.showwarning("클립보드", "복사된 텍스트가 없습니다.")

    def _open_text(self):
        p = filedialog.askopenfilename(filetypes=[("텍스트", "*.txt")])
        if p:
            try: content = Path(p).read_text(encoding="utf-8-sig")
            except UnicodeError: content = Path(p).read_text(encoding="cp949")
            self.text.delete("1.0", "end")
            self.text.insert("1.0", content)
            self._preview()

    def _preview(self):
        events = parse_many(self.text.get("1.0", "end"))
        if not events:
            self.state.configure(text="예약이 인식되지 않았습니다.")
            return
        bad = [x for x in events if validate(x)]
        self.state.configure(text=f"인식 {len(events)}건 / 확인 필요 {len(bad)}건")
        for r in events[:8]:
            self._log(f"분석: {r.source} {r.number} / {r.room or '객실 미인식'} / {'취소' if r.cancelled else f'{r.amount:,}원'}" + (f" | {validate(r)}" if validate(r) else ""))

    def _manual(self):
        if not self._valid_excel(): return
        content = self.text.get("1.0", "end")
        events = parse_many(content)
        if not events:
            messagebox.showwarning("예약 없음", "붙여넣은 내용에서 예약을 인식하지 못했습니다.")
            return
        if not messagebox.askyesno("엑셀 입력 확인", f"{len(events)}건을 엑셀에 반영할까요?\n작업 전 자동 백업합니다."):
            return
        self._submit(content)

    def _valid_excel(self) -> bool:
        if not Path(self.excel.get()).is_file():
            self._log("엑셀 파일을 먼저 선택하세요.")
            return False
        return True

    def _submit(self, content: str, path: Path | None = None, digest: str | None = None):
        if self.future is not None: return
        self.work_file, self.work_hash = path, digest
        self.future = self.pool.submit(execute, self.excel.get(), content)
        self.state.configure(text="엑셀 처리 중…")

    def _poll(self):
        if self.future is not None and self.future.done():
            try:
                messages = self.future.result()
                for m in messages: self._log(m)
                self.state.configure(text=f"처리 완료 ({len(messages)}건)")
                if self.work_file and self.work_hash:
                    self.history[str(self.work_file)] = self.work_hash
                    save_json(HISTORY, self.history)
            except Exception as e:
                self._log("처리 중지: " + str(e))
                self.state.configure(text="오류 발생 · 로그 확인")
            finally:
                self.future = None
                self.work_file = None
                self.work_hash = None
        self.root.after(400, self._poll)

    def _watch(self):
        if self.watch.get() and self.future is None and self._valid_excel():
            folder = Path(self.folder.get())
            if folder.is_dir():
                for p in sorted(folder.glob("*.txt")):
                    try:
                        if time.time() - p.stat().st_mtime < 3: continue  # Do not read partial writes.
                        payload = p.read_bytes()
                        digest = hashlib.sha256(payload).hexdigest()
                        if self.history.get(str(p)) == digest: continue
                        try: content = payload.decode("utf-8-sig")
                        except UnicodeError: content = payload.decode("cp949")
                        if not parse_many(content): continue
                        self._log("수신 파일 처리: " + p.name)
                        self._submit(content, p, digest)
                        break
                    except (OSError, UnicodeError) as e:
                        self._log(f"수신 파일 읽기 실패: {p.name}: {e}")
            else:
                self._log("자동감시 폴더가 존재하지 않습니다.")
                self.watch.set(False)
        self.root.after(10000, self._watch)

    def _close(self):
        self.watch.set(False)
        self._save_settings()
        if self.future and not self.future.done():
            messagebox.showinfo("처리 중", "엑셀 작업이 진행 중입니다. 완료된 후 종료하세요.")
            return
        self.pool.shutdown(wait=False, cancel_futures=True)
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()