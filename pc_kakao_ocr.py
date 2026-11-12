"""Read-only OCR of *visible* KakaoTalk desktop chat windows.

No keyboard/mouse injection, focus changes, process memory access, or persistent screenshots.
This is a best-effort screen reader: minimized, occluded, locked or off-screen windows
cannot be read reliably. OCR output is untrusted until strict parsing succeeds.
"""
from __future__ import annotations

import os
import re
import shutil
import sys
from functools import lru_cache
from dataclasses import dataclass
from pathlib import Path
from typing import Sequence

from reservations import Reservation, parse_one, validate

CHAT_TITLES = ("여기어때", "NOL 알림")
MIN_WINDOW_WIDTH = 230
MIN_WINDOW_HEIGHT = 200


@dataclass(frozen=True)
class ChatObservation:
    title: str
    visible_text: str
    reservations: tuple[Reservation, ...]
    warning: str = ""


def _check_win32() -> None:
    if sys.platform != "win32":
        raise RuntimeError("PC 카카오톡 화면 감시는 Windows에서만 지원합니다.")


def find_kakao_windows() -> list[tuple[int, str]]:
    """Only accept chat windows owned by KakaoTalk.exe, never similarly titled apps."""
    _check_win32()
    import psutil
    import win32gui
    import win32process

    windows: list[tuple[int, str]] = []

    def each(hwnd, unused):
        if not win32gui.IsWindowVisible(hwnd) or win32gui.IsIconic(hwnd):
            return
        raw = win32gui.GetWindowText(hwnd).strip()
        title = next((n for n in CHAT_TITLES if raw == n), None)
        if title is None:
            return
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if psutil.Process(pid).name().lower() != "kakaotalk.exe":
                return
        except (psutil.Error, OSError):
            return
        windows.append((hwnd, title))

    win32gui.EnumWindows(each, None)
    return windows


def tesseract_executable() -> str:
    """Find a locally installed OCR engine; never download binaries silently."""
    found = shutil.which("tesseract")
    if found:
        return found
    common = [
        Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Tesseract-OCR" / "tesseract.exe",
        Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "Tesseract-OCR" / "tesseract.exe",
    ]
    return next((str(path) for path in common if path.is_file()), "")


@lru_cache(maxsize=1)
def check_ocr() -> str:
    """Return a Korean actionable error or an empty string when ready."""
    try:
        import pytesseract
    except ImportError:
        return "pytesseract가 없습니다. install_windows.bat을 다시 실행하세요."
    exe = tesseract_executable()
    if not exe:
        return "Tesseract OCR이 설치되지 않았습니다. 한국어(kor) 언어팩과 함께 설치하세요."
    pytesseract.pytesseract.tesseract_cmd = exe
    try:
        langs = set(pytesseract.get_languages(config=""))
    except Exception as exc:
        return f"OCR 프로그램 실행 실패: {exc}"
    if not {"kor", "eng"}.issubset(langs):
        return "OCR 언어팩 kor 및 eng가 모두 필요합니다. Tesseract 설치 옵션을 확인하세요."
    return ""


def image_to_text(img) -> str:
    """Use only local Tesseract; transient image exists in RAM, not on disk."""
    import pytesseract
    from PIL import ImageOps, ImageEnhance

    w, h = img.size
    if w < MIN_WINDOW_WIDTH or h < MIN_WINDOW_HEIGHT:
        return ""
    # Read Korean text in narrow chat bubbles at an enlarged resolution.
    img = img.convert("RGB")
    enlarged = img.resize((w * 3, h * 3), resample=3)
    gray = ImageOps.autocontrast(ImageOps.grayscale(enlarged), cutoff=1)
    gray = ImageEnhance.Sharpness(gray).enhance(1.25)
    return pytesseract.image_to_string(gray, lang="kor+eng", config="--psm 11").strip()


def _split_visible_notice(title: str, text: str) -> list[str]:
    """Conservative segmenting of OCR content. No cross-message concatenation.

    If multiple booking identifiers occur without clear safe boundaries, skip them
    rather than mix one guest with another guest's money/check-in date.
    """
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    if title == "여기어때":
        # Each booking block uses 예약번호; include announcement line before it.
        hits = list(re.finditer(r"예약\s*번호\s*[:：]", normalized))
        if not hits:
            return []
        if len(hits) == 1:
            return ["여기어때\n" + normalized]
        # Start at each heading (if recognized), not at the previous booking body.
        blocks = []
        for i, m in enumerate(hits):
            next_start = hits[i + 1].start() if i + 1 < len(hits) else len(normalized)
            prefix = normalized[max(0, m.start() - 140):m.start()]
            headings = list(re.finditer(r"(?:여기어때|예약\s*취소|당일\s*예약|숙박\s*취소)", prefix))
            if not headings:
                continue  # Unsafe to decide whether this is a cancellation.
            prefix = prefix[headings[-1].start():]
            blocks.append("여기어때\n" + prefix + normalized[m.start():next_start])
        return blocks

    # NOL: booking number occupies its own line. Date/phone never meet 14 digits.
    ids = list(re.finditer(r"(?m)^\s*(?:[A-Za-z0-9_-]{14,32})\s*$", normalized))
    ids = [m for m in ids if re.fullmatch(r"[0-9]{14,20}", m.group(0).strip())]
    if not ids:
        return []
    if len(ids) == 1:
        return ["NOL 알림\n" + normalized]
    blocks = []
    for i, m in enumerate(ids):
        next_start = ids[i + 1].start() if i + 1 < len(ids) else len(normalized)
        before = normalized[max(0, m.start() - 220):m.start()]
        markers = list(re.finditer(r"(?:NOL\s*[^\n]*예약|[<〈]\s*숙박\s*(?:취소)?\s*[>〉])", before))
        if not markers:
            continue
        # Use the nearest preceding NOL heading or 숙박 marker in this message.
        before = before[markers[-1].start():]
        blocks.append("NOL 알림\n" + before + normalized[m.start():next_start])
    return blocks


def parse_ocr(title: str, text: str) -> tuple[Reservation, ...]:
    if title not in CHAT_TITLES:
        return ()
    result = []
    seen = set()
    for block in _split_visible_notice(title, text):
        parsed = parse_one(block)
        if parsed is None or validate(parsed):
            continue
        if title == "여기어때" and parsed.channel != "여기":
            continue
        if title == "NOL 알림" and parsed.channel != "야놀자":
            continue
        k = (parsed.number, parsed.cancelled)
        if k not in seen:
            result.append(parsed)
            seen.add(k)
    return tuple(result)


def _foreground_obscures(hwnd: int, left: int, top: int, width: int, height: int) -> bool:
    """Refuse OCR when another titled top-level window covers chat message area.

    This catches accidental reading of unrelated text from a browser in front
    of KakaoTalk. Small tooltips/transparent overlays are ignored.
    """
    import win32con
    import win32gui

    ahead = []
    # EnumWindows returns top-most first. Iterate until the target.
    found = False
    def collect_up_to(candidate, unused):
        nonlocal found
        if candidate == hwnd:
            found = True
            return
        if not found:
            ahead.append(candidate)
    win32gui.EnumWindows(collect_up_to, None)
    if not found:
        return True
    for other in ahead:
        if not win32gui.IsWindowVisible(other) or win32gui.IsIconic(other):
            continue
        if not win32gui.GetWindowText(other).strip():
            continue
        if win32gui.GetWindowLong(other, win32con.GWL_EXSTYLE) & win32con.WS_EX_TOOLWINDOW:
            continue
        try:
            a, b, c, d = win32gui.GetWindowRect(other)
        except win32gui.error:
            continue
        overlap_width = max(0, min(left + width, c) - max(left, a))
        overlap_height = max(0, min(top + height, d) - max(top, b))
        if overlap_width * overlap_height > width * height * 0.12:
            return True
    return False


def capture_one(hwnd: int, title: str) -> ChatObservation:
    import mss
    import win32gui
    from PIL import Image

    if not win32gui.IsWindowVisible(hwnd) or win32gui.IsIconic(hwnd):
        return ChatObservation(title, "", (), "창이 최소화되었거나 보이지 않습니다.")
    left, top = win32gui.ClientToScreen(hwnd, (0, 0))
    x0, y0, x1, y1 = win32gui.GetClientRect(hwnd)
    width, height = x1 - x0, y1 - y0
    if width < MIN_WINDOW_WIDTH or height < MIN_WINDOW_HEIGHT:
        return ChatObservation(title, "", (), "채팅창이 너무 작습니다. 창을 넓혀 주세요.")
    if _foreground_obscures(hwnd, left, top, width, height):
        return ChatObservation(title, "", (), "다른 창이 채팅창을 가리고 있습니다. 창을 앞으로 꺼내 주세요.")
    # mss reads visible pixels. Do not use PrintWindow, hooks, simulated keys.
    with mss.mss() as ctx:
        monitor = ctx.monitors[0]
        if not (monitor["left"] <= left and monitor["top"] <= top and
                left + width <= monitor["left"] + monitor["width"] and
                top + height <= monitor["top"] + monitor["height"]):
            return ChatObservation(title, "", (), "화면 밖으로 나간 채팅창이 있습니다.")
        shot = ctx.grab({"left": left, "top": top, "width": width, "height": height})
        img = Image.frombytes("RGB", shot.size, shot.rgb)
    text = image_to_text(img)
    reservations = parse_ocr(title, text)
    return ChatObservation(title, text, reservations,
                           "읽을 수 있는 예약이 없습니다. 글씨 크기·창 가림 상태를 확인하세요." if not reservations else "")


def scan_chats() -> list[ChatObservation]:
    issue = check_ocr()
    if issue:
        raise RuntimeError(issue)
    windows = find_kakao_windows()
    if not windows:
        raise RuntimeError("열려 있는 카카오톡 '여기어때' 또는 'NOL 알림' 독립 채팅창을 찾지 못했습니다.")
    return [capture_one(hwnd, title) for hwnd, title in windows]


def signature(r: Reservation) -> tuple:
    return (r.channel, r.number, r.cancelled, r.guest, r.check_in, r.room, r.amount, r.vehicle)