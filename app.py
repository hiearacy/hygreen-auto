"""Hygreen desktop app. PC screen OCR is opt-in and read-only; no input injection."""
from __future__ import annotations

import hashlib
import json
import os
import time
import tkinter as tk
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk
from lan_receiver import NotificationReceiver, PORT, PATH, get_local_ip, new_token
from pc_kakao_ocr import scan_chats, signature, enumerate_window_candidates, looks_like_kakao, CHAT_TITLES

from excel_io import apply_reservation
from reservations import parse_many, validate

APP_DIR = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "HygreenAuto"
CONFIG = APP_DIR / "settings.json"
HISTORY = APP_DIR / "imported_files.json"
LAN_INBOX = APP_DIR / "lan_inbox"
REVIEW = APP_DIR / "review_needed"


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
    results = []
    for event in events:
        result = apply_reservation(path, event)
        if "취소 대상 예약번호 없음" in result:
            raise ValueError(result + " · 알림을 확인 폴더에 보관합니다.")
        results.append(result)
    return results


def execute_one(path: str, reservation) -> list[str]:
    problem = validate(reservation)
    if problem:
        raise ValueError(f"OCR 데이터 확인 필요: {problem}")
    return [apply_reservation(path, reservation)]


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("향그린 예약 관리 · PC 카카오톡 자동감시 v3.2")
        self.root.geometry("950x850")
        self.root.minsize(780, 720)
        self.settings = read_json(CONFIG)
        self.history = read_json(HISTORY)
        self.excel = tk.StringVar(value=self.settings.get("excel", ""))
        self.folder = tk.StringVar(value=self.settings.get("folder", ""))
        self.watch = tk.BooleanVar(value=False)  # Never silently enable live writes at launch.
        self.bridge_write = tk.BooleanVar(value=False)  # Legacy LAN mode, not used in PC OCR.
        self.pc_watch = tk.BooleanVar(value=False)  # Explicit opt-in per run.
        self.pc_write = tk.BooleanVar(value=False)  # Explicit opt-in per run.
        self.pc_scan_future = None
        self.pc_requested_manual_scan = False
        self.pc_streak: dict[tuple, int] = {}
        self.pc_finished: set[tuple] = set()
        self.pc_failed: set[tuple] = set()
        self.pc_queue = []
        self.pc_processing = None
        self.pc_latest_count = 0
        self.pc_verified_preview = False
        self.pc_selected_windows: dict[str, int] = {}
        self.pc_scan_error = ""
        self.bridge = None
        self.bridge_seen = set()
        self.bridge_token = self.settings.get("bridge_token") or new_token()
        if not self.settings.get("bridge_token"):
            self.settings["bridge_token"] = self.bridge_token
            save_json(CONFIG, self.settings)
        self.pool = ThreadPoolExecutor(max_workers=1, thread_name_prefix="ExcelWriter")
        self.pc_pool = ThreadPoolExecutor(max_workers=1, thread_name_prefix="KakaoOCR")
        self.future = None
        self.work_file = None
        self.work_hash = None
        self._ui()
        self.root.after(400, self._poll)
        self.root.after(10000, self._watch)
        self.root.after(1000, self._pc_poll)
        self.root.protocol("WM_DELETE_WINDOW", self._close)
        self._log("준비 완료. PC 카카오톡 화면 읽기는 사용자가 켜야 작동합니다. 창을 가리지 마세요.")

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
        self.text = scrolledtext.ScrolledText(outer, height=9, font=("맑은 고딕", 10), wrap="word")
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
        ttk.Separator(outer, orient="horizontal").pack(fill="x", pady=(8, 7))
        ttk.Label(outer, text="3. PC 카카오톡 화면 자동감시 (NOL 알림 / 여기어때 독립 채팅창, 화면에 보여야 함)").pack(anchor="w")
        pc_row = ttk.Frame(outer); pc_row.pack(fill="x", pady=5)
        ttk.Button(pc_row, text="화면 읽기 테스트", command=self._pc_read_once).pack(side="left")
        ttk.Button(pc_row, text="채팅창 직접 선택", command=self._pc_choose_windows).pack(side="left", padx=(6, 0))
        ttk.Checkbutton(pc_row, text="7초마다 감시", variable=self.pc_watch).pack(side="left", padx=(15, 6))
        ttk.Checkbutton(pc_row, text="자동 엑셀 기록 (검증 후 켜기)", variable=self.pc_write,
                        command=self._pc_toggle_write).pack(side="left", padx=6)
        self.pc_label = ttk.Label(outer, text="기본 OFF · 캡처 이미지 저장 안 함 · 카카오톡 클릭/키보드 조작 없음")
        self.pc_label.pack(anchor="w", pady=(0, 5))
        ttk.Label(outer, text="주의: 다른 창이 카톡을 가리거나 화면이 잠기면 감지가 불가능합니다. OCR 오인식 방지를 위해 2회 재확인 후 입력합니다.", foreground="#914500").pack(anchor="w")
        self.state = ttk.Label(outer, text="대기 중")
        self.state.pack(anchor="w", pady=5)
        self.log = scrolledtext.ScrolledText(outer, height=7, state="disabled", font=("맑은 고딕", 9))
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
        save_json(CONFIG, {"excel": self.excel.get(), "folder": self.folder.get(), "bridge_token": self.bridge_token})

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
                if self.pc_processing is not None:
                    self.pc_finished.add(self.pc_processing)
                for m in messages: self._log(m)
                self.state.configure(text=f"처리 완료 ({len(messages)}건)")
                if self.work_file and self.work_file.parent == LAN_INBOX:
                    self.work_file.unlink(missing_ok=True)  # Delete processed private notification.
                    self.bridge_seen.discard(str(self.work_file))
                elif self.work_file and self.work_hash:
                    self.history[str(self.work_file)] = self.work_hash
                    save_json(HISTORY, self.history)
            except Exception as e:
                if self.pc_processing is not None:
                    self.pc_failed.add(self.pc_processing)
                    self._log("PC OCR 기록 실패 · 안전을 위해 자동 재시도하지 않습니다.")
                if self.work_file and self.work_file.parent == LAN_INBOX:
                    self._needs_review(self.work_file, "엑셀 처리 실패")
                self._log("처리 중지: " + str(e))
                self.state.configure(text="오류 발생 · 로그 확인")
            finally:
                self.future = None
                self.work_file = None
                self.work_hash = None
                self.pc_processing = None
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

    def _pc_choose_windows(self):
        """Select visible chat windows explicitly when auto-discovery fails."""
        if self.pc_scan_future is not None:
            messagebox.showinfo("채팅창 선택", "현재 화면 읽기를 마친 다음 창을 선택해 주세요.")
            return
        try:
            candidates = enumerate_window_candidates()
        except Exception as exc:
            messagebox.showerror("창 목록 오류", f"Windows 창 목록을 확인하지 못했습니다.\\n{exc}")
            return
        if not candidates:
            messagebox.showwarning("창 없음", "보이는 창이 없습니다. 카카오톡 채팅창을 먼저 열어 주세요.")
            return
        candidates = sorted(candidates, key=lambda w: (not looks_like_kakao(w), w.left, w.top))
        descriptions = [
            f"{i + 1:02d} | {w.process[:25]} | 위치 {w.left},{w.top} | {w.width}x{w.height} | {(w.title or '제목 없음')[:35]}"
            for i, w in enumerate(candidates)
        ]
        dialog = tk.Toplevel(self.root)
        dialog.title("카카오톡 채팅창 지정 · v3.2")
        dialog.geometry("930x245")
        dialog.transient(self.root)
        dialog.grab_set()
        frame = ttk.Frame(dialog, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="열려 있는 창에서 카카오톡의 NOL 알림과 여기어때 창을 각각 골라 주세요.").pack(anchor="w", pady=(0, 8))
        ttk.Label(frame, text="카톡 창을 왼쪽·오른쪽에 두고, '위치'와 프로세스 이름으로 구분하세요. 다른 창은 선택하지 마세요.").pack(anchor="w", pady=(0, 12))
        chosen = {}
        for title in CHAT_TITLES:
            row = ttk.Frame(frame)
            row.pack(fill="x", pady=4)
            ttk.Label(row, text=title, width=14).pack(side="left")
            combo = ttk.Combobox(row, state="readonly", values=descriptions)
            combo.pack(side="left", fill="x", expand=True)
            for i, w in enumerate(candidates):
                if self.pc_selected_windows.get(title) == w.hwnd:
                    combo.current(i)
                    break
            chosen[title] = combo

        def save_selection():
            indexes = [c.current() for c in chosen.values()]
            if any(i < 0 for i in indexes) or indexes[0] == indexes[1]:
                messagebox.showwarning("창 선택", "NOL 알림과 여기어때를 서로 다른 창으로 지정해야 합니다.", parent=dialog)
                return
            self.pc_selected_windows = {title: candidates[chosen[title].current()].hwnd for title in CHAT_TITLES}
            self.pc_verified_preview = False
            self.pc_streak.clear()
            self.pc_queue.clear()
            self.pc_label.configure(text="창 직접 선택 완료 · 화면 읽기 테스트로 채팅방 이름을 검증하세요 · 자동기록 OFF")
            self.pc_write.set(False)
            self._log("카카오톡 채팅창 2개를 직접 지정했습니다. 화면 읽기 테스트로 확인하세요.")
            dialog.destroy()

        bottom = ttk.Frame(frame)
        bottom.pack(fill="x", pady=(12, 0))
        ttk.Button(bottom, text="선택 완료", command=save_selection).pack(side="left")
        ttk.Button(bottom, text="자동 검색으로 되돌리기", command=lambda: self._pc_clear_selection(dialog)).pack(side="left", padx=10)
        ttk.Button(bottom, text="닫기", command=dialog.destroy).pack(side="right")

    def _pc_clear_selection(self, dialog):
        self.pc_selected_windows.clear()
        self.pc_verified_preview = False
        self.pc_streak.clear()
        self.pc_queue.clear()
        self.pc_write.set(False)
        self.pc_label.configure(text="자동 창 검색 모드 · 화면 읽기 테스트로 확인하세요")
        self._log("카카오톡 창 직접 선택을 해제했습니다.")
        dialog.destroy()

    def _pc_read_once(self):
        if self.pc_scan_future is not None:
            self._log("이미 OCR 화면 분석 중입니다.")
            return
        self.pc_requested_manual_scan = True
        self.pc_label.configure(text="화면 읽기 진행 중 · 프로그램 화면을 잠깐 숨깁니다.")
        self.root.withdraw()  # Hide *our own* window to avoid obscuring the chat screenshots.
        self.root.after(700, self._pc_launch_manual_scan)

    def _pc_launch_manual_scan(self):
        self.pc_scan_future = self.pc_pool.submit(scan_chats, dict(self.pc_selected_windows) or None)

    def _pc_toggle_write(self):
        if not self.pc_write.get():
            return
        if not self.pc_verified_preview:
            messagebox.showwarning("화면 읽기 테스트 먼저", "먼저 화면 읽기 테스트에서 실제 예약이 정확하게 인식되는지 확인하세요.")
            self.pc_write.set(False)
            return
        if not messagebox.askyesno("실제 엑셀 자동 기록", "정말 자동 기록을 켤까요?\nOCR은 잘못 인식할 수 있으므로 테스트용 엑셀로 여러 건 검증한 후 켜세요."):
            self.pc_write.set(False)

    def _pc_poll(self):
        if self.pc_scan_future is not None and self.pc_scan_future.done():
            manual = self.pc_requested_manual_scan
            self.pc_requested_manual_scan = False
            try:
                observations = self.pc_scan_future.result()
                issues = [o for o in observations if o.warning]
                seen = {}
                for item in observations:
                    for res in item.reservations:
                        seen[signature(res)] = res
                self.pc_streak = {k: min(self.pc_streak.get(k, 0) + 1, 1000) for k in seen}
                self.pc_latest_count = len(seen)
                names = ", ".join(o.title for o in observations)
                self.pc_label.configure(text=f"연결: {names} · 검증된 예약 {len(seen)}건 · 자동 기록 {'ON' if self.pc_write.get() else 'OFF'}")
                if manual:
                    self.pc_verified_preview = bool(seen)
                    blocks = [f"[{o.title}] OCR 결과 (직접 검토 필요)\n{o.visible_text}" for o in observations]
                    self.text.delete("1.0", "end")
                    self.text.insert("1.0", "\n\n--------\n\n".join(blocks))
                    self._log(f"화면 읽기 테스트: 카카오톡 {len(observations)}개 창 / 예약 {len(seen)}건")
                    for o in issues:
                        self._log(f"{o.title}: {o.warning}")
                    for res in seen.values():
                        self._log(f"OCR: {res.source} / {res.number} / {res.check_in} / {res.amount:,}원")
                if self.pc_watch.get() and self.pc_write.get():
                    for key, res in seen.items():
                        needed = 3 if res.cancelled else 2
                        if self.pc_streak[key] < needed:
                            continue
                        if key in self.pc_finished or key in self.pc_failed:
                            continue
                        if key == self.pc_processing or any(signature(x) == key for x in self.pc_queue):
                            continue
                        self.pc_queue.append(res)
                self.pc_scan_error = ""
            except Exception as exc:
                msg = str(exc)
                if manual or msg != self.pc_scan_error:
                    self._log("PC 화면 감시: " + msg)
                self.pc_scan_error = msg
                self.pc_label.configure(text="화면 읽기 실패 · 로그와 OCR 설치 상태 확인")
                # Incomplete OCR dependencies should not spam background processes.
                if "설치" in msg or "언어팩" in msg or "Windows" in msg:
                    self.pc_watch.set(False)
            finally:
                self.pc_scan_future = None
                if manual:
                    self.root.deiconify()  # Show only after local screenshot has finished.
        if self.pc_watch.get() and self.pc_scan_future is None:
            self.pc_scan_future = self.pc_pool.submit(scan_chats, dict(self.pc_selected_windows) or None)
            self.pc_requested_manual_scan = False
        if self.pc_write.get() and self.pc_watch.get() and self.pc_queue and self.future is None:
            if self._valid_excel():
                res = self.pc_queue.pop(0)
                self.pc_processing = signature(res)
                self.future = self.pool.submit(execute_one, self.excel.get(), res)
                self.state.configure(text="PC 예약 자동 엑셀 기록 중…")
            else:
                self.pc_write.set(False)
                self._log("엑셀 선택 오류로 PC 자동 기록이 중단됐습니다.")
        self.root.after(7000, self._pc_poll)

    def _toggle_bridge(self):
        if self.bridge is not None:
            self.bridge.stop()
            self.bridge = None
            self.bridge_write.set(False)
            self.bridge_button.configure(text="LAN 수신 시작")
            self.bridge_label.configure(text="수신 서버 꺼짐 · 자동 기록 꺼짐")
            self._log("휴대폰 수신 서버를 종료했습니다.")
            return
        try:
            bridge = NotificationReceiver(LAN_INBOX, self.bridge_token)
            bridge.start()
            self.bridge = bridge
            self.bridge_button.configure(text="LAN 수신 중지")
            self.bridge_label.configure(text=f"수신 대기: {get_local_ip()}:{PORT} · 자동 기록은 별도 선택")
            self._log("휴대폰 알림 수신 준비 완료. 개인 Wi-Fi에서만 사용하세요.")
        except OSError as e:
            messagebox.showerror("수신 서버 시작 실패", f"포트 {PORT}을 열 수 없습니다.\n{e}")

    def _copy_bridge(self):
        if self.bridge is None:
            messagebox.showwarning("수신 서버", "LAN 수신 시작을 먼저 누르세요.")
            return
        instructions = (f"URL: http://{get_local_ip()}:{PORT}{PATH}\n"
                        "POST / Content-Type: text/plain; charset=utf-8\n"
                        f"헤더 이름: X-Hygreen-Token\n헤더 값: {self.bridge_token}\n"
                        "본문: {not_title} (줄바꿈) {notification} (줄바꿈) {not_text_big}")
        self.root.clipboard_clear()
        self.root.clipboard_append(instructions)
        self._log("휴대폰 연동 정보(비밀키 포함)를 클립보드에 복사했습니다. 외부에 공유하지 마세요.")
        messagebox.showinfo("연동 정보 복사 완료", "연동 정보가 클립보드에 복사됐습니다.\n메모장에 붙여넣어 MacroDroid에 입력하세요.\n인증키는 외부에 공유하지 마세요.")

    def _needs_review(self, path: Path, reason: str):
        REVIEW.mkdir(parents=True, exist_ok=True)
        try:
            target = REVIEW / path.name
            path.replace(target)
            self.bridge_seen.discard(str(path))
            self._log(f"수동 확인 필요: {reason}. 보관 위치: {REVIEW}")
        except OSError as e:
            self._log(f"확인 필요 파일 이동 실패: {e}")
            self.bridge_seen.add(str(path))

    def _bridge_poll(self):
        if self.bridge is not None:
            for p in sorted(LAN_INBOX.glob("notice_*.txt")):
                try:
                    if time.time() - p.stat().st_mtime < 0.5: continue
                    if not self.bridge_write.get():
                        if str(p) not in self.bridge_seen:
                            self.bridge_seen.add(str(p))
                            self._log("카카오 알림 수신됨 · 자동 기록 OFF. 확인 후 체크박스를 켜세요.")
                        continue
                    if self.future is not None: break
                    if not self._valid_excel(): break
                    content = p.read_text(encoding="utf-8")
                    events = parse_many(content)
                    invalid = [e for e in events if validate(e)]
                    if not events or invalid:
                        why = ", ".join(validate(e) for e in invalid) if invalid else "예약 본문 인식 실패"
                        self._needs_review(p, why)
                        continue
                    self._log(f"자동 수신 분석 성공: {len(events)}건 · 엑셀 기록 시작")
                    self._submit(content, p, "queued")
                    break
                except (OSError, UnicodeError) as e:
                    self._needs_review(p, f"읽기 실패: {e}")
        self.root.after(2000, self._bridge_poll)

    def _close(self):
        self.watch.set(False)
        self.pc_watch.set(False)
        self.pc_write.set(False)
        self._save_settings()
        if self.future and not self.future.done():
            messagebox.showinfo("처리 중", "엑셀 작업이 진행 중입니다. 완료된 후 종료하세요.")
            return
        if self.bridge is not None:
            self.bridge.stop()
        self.pool.shutdown(wait=False, cancel_futures=True)
        self.pc_pool.shutdown(wait=False, cancel_futures=True)
        self.root.destroy()


if __name__ == "__main__":
    if os.name == "nt":
        # Make GetClientRect/ClientToScreen coordinates match mss physical pixels.
        import ctypes
        try: ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except (AttributeError, OSError): pass
    root = tk.Tk()
    App(root)
    root.mainloop()