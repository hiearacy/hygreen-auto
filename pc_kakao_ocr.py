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


@dataclass(frozen=True)
class WindowCandidate:
    hwnd: int
    title: str
    process: str
    class_name: str
    left: int
    top: int
    width: int
    height: int


def looks_like_kakao(window: WindowCandidate) -> bool:
    """Accept known Kakao-like processes/classes; user may also manually select."""
    return (
        "kakao" in window.process.casefold()
        or window.title.strip() in CHAT_TITLES
        or window.class_name.casefold().startswith("eva_window_")
    )


def enumerate_window_candidates() -> list[WindowCandidate]:
    """Only enumerate visible window metadata; does not capture their contents."""
    _check_win32()
    import psutil
    import win32gui
    import win32process

    result = []

    def each(hwnd: int, _unused):
        if not win32gui.IsWindowVisible(hwnd) or win32gui.IsIconic(hwnd):
            return
        try:
            bounds = win32gui.GetClientRect(hwnd)
            w, h = bounds[2] - bounds[0], bounds[3] - bounds[1]
            if w < MIN_WINDOW_WIDTH or h < MIN_WINDOW_HEIGHT:
                return
            left, top, _, _ = win32gui.GetWindowRect(hwnd)
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            try:
                process = psutil.Process(pid).name()
            except (psutil.Error, OSError, ValueError):
                process = "(확인 불가)"
            try:
                classname = win32gui.GetClassName(hwnd)
            except Exception:
                classname = ""
            result.append(WindowCandidate(
                hwnd, win32gui.GetWindowText(hwnd).strip(),
                process, classname, left, top, w, h,
            ))
        except Exception:
            return

    win32gui.EnumWindows(each, None)
    return result


def _check_win32() -> None:
    if sys.platform != "win32":
        raise RuntimeError("PC 카카오톡 화면 감시는 Windows에서만 지원합니다.")


def find_kakao_windows() -> list[tuple[int, str]]:
    """Find candidate windows without requiring exact KakaoTalk.exe naming."""
    return [(w.hwnd, w.title) for w in enumerate_window_candidates() if looks_like_kakao(w)]



def classify_chat_header(text: str) -> str | None:
    """Identify only allowlisted chat rooms from the header's first two OCR lines."""
    lines = [re.sub(r"\s+", "", line).lower() for line in text.splitlines() if line.strip()]
    lines = lines[:2]
    yeogi = any("여기어때" in line for line in lines)
    nol = any("nol알림" in line for line in lines)
    if yeogi == nol:
        return None
    return "여기어때" if yeogi else "NOL 알림"


def _detect_chat_name(hwnd: int) -> str | None:
    """Read KakaoTalk's small title region rather than the mixed header + notices."""
    import mss
    import win32gui
    from PIL import Image, ImageEnhance, ImageOps
    import pytesseract

    left, top = win32gui.ClientToScreen(hwnd, (0, 0))
    rect = win32gui.GetClientRect(hwnd)
    width, height = rect[2] - rect[0], rect[3] - rect[1]
    if width < MIN_WINDOW_WIDTH or height < MIN_WINDOW_HEIGHT:
        return None
    if _foreground_obscures(hwnd, left, top, width, height):
        return None
    # Header text begins roughly 45 px from the client's left edge. Keep the
    # screenshot strictly within this window and exclude the first chat bubble.
    x = min(40, max(0, width - 150))
    h = min(45, height)
    cw = min(width - x, 260)
    with mss.mss() as ctx:
        monitor = ctx.monitors[0]
        if not _on_monitor(monitor, left + x, top, cw, h):
            return None
        shot = ctx.grab({"left": left + x, "top": top,
                         "width": cw, "height": h})
        header = Image.frombytes("RGB", shot.size, shot.rgb)
    enlarged = header.resize((cw * 4, h * 4), Image.Resampling.LANCZOS)
    gray = ImageOps.autocontrast(ImageOps.grayscale(enlarged))
    gray = ImageEnhance.Contrast(gray).enhance(1.6)
    header_text = pytesseract.image_to_string(gray, lang="kor+eng", config="--psm 6")
    return classify_chat_header(header_text)


def _on_monitor(monitor: dict, left: int, top: int, width: int, height: int) -> bool:
    return (monitor["left"] <= left and monitor["top"] <= top and
            left + width <= monitor["left"] + monitor["width"] and
            top + height <= monitor["top"] + monitor["height"])


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
    """OCR one isolated message bubble (not the full chat desktop)."""
    import pytesseract
    from PIL import Image, ImageOps

    w, h = img.size
    if w < 75 or h < 25:
        return ""
    image = ImageOps.autocontrast(ImageOps.grayscale(img.convert("RGB")))
    image = image.resize((w * 4, h * 4), Image.Resampling.LANCZOS)
    # Uniform text lines are much more reliable with PSM 6 than PSM 11 when
    # Korean, amounts, dates and phone numbers share one narrow bubble.
    return pytesseract.image_to_string(image, lang="kor+eng", config="--psm 6").strip()


def _bubble_row_active(pixels, width: int, y: int) -> bool:
    """Identify a left-aligned white/yellow KakaoTalk notice against blue chat UI.

    Never scan the white footer: it must have the KakaoTalk blue background to
    the left of the candidate message rectangle.
    """
    bg_x = max(12, min(width - 1, round(width * 0.07)))
    br, bg, bb = pixels[bg_x, y]
    if not (bb - br >= 14 and bg - br >= 7):
        return False
    x0, x1 = max(45, round(width * .14)), min(width-35, round(width * .71))
    if x1 <= x0:
        return False
    hits = 0
    scanned = 0
    for x in range(x0, x1, 3):
        r,g,b = pixels[x,y]
        white = min(r,g,b) >= 235 and max(r,g,b)-min(r,g,b) < 24
        yellow = r >= 215 and g >= 160 and b <= 115
        hits += int(white or yellow)
        scanned += 1
    return scanned > 0 and hits >= scanned * .52


def _light_run(pixels, width: int, y: int) -> tuple[int,int] | None:
    """Locate the horizontal extent of the mostly uniform bubble background."""
    runs=[]
    start=None
    for x in range(25,width-12):
        r,g,b=pixels[x,y]
        good=(min(r,g,b)>=235 and max(r,g,b)-min(r,g,b)<24) or (r>=215 and g>=160 and b<=115)
        if good and start is None:
            start=x
        elif not good and start is not None:
            if x-start>=65:
                runs.append((start,x))
            start=None
    if start is not None and width-12-start>=65:
        runs.append((start,width-12))
    runs=[run for run in runs if run[0] <= width * .4]
    return max(runs,key=lambda ab:ab[1]-ab[0]) if runs else None


def message_bubble_boxes(img) -> list[tuple[int,int,int,int]]:
    """Find separate visible white/yellow notification bubbles.

    Uses only visible pixels from the user-chosen chat, never private KakaoTalk
    files or hooks. A missing/ambiguous bubble is omitted rather than joined to
    adjacent bookings; this prevents mixing names, amounts and booking numbers.
    """
    image=img.convert("RGB")
    width,height=image.size
    if width < MIN_WINDOW_WIDTH or height < MIN_WINDOW_HEIGHT:
        return []
    pix=image.load()
    spans=[]
    active_start=None
    last_seen=None
    # Browser bars + message input area contain unrelated text.
    for y in range(40, max(40,height-100)):
        active=_bubble_row_active(pix,width,y)
        if active:
            if active_start is None:
                active_start=y
            last_seen=y
        elif active_start is not None and y-last_seen>8:
            if last_seen-active_start>=34:
                spans.append((active_start,last_seen+1))
            active_start=None
            last_seen=None
    if active_start is not None and last_seen-active_start>=34:
        spans.append((active_start,last_seen+1))
    boxes=[]
    for y0,y1 in spans:
        # At least one empty/light text row should expose the bubble's extent.
        probes=[y0+3,min(y1-1,y0+10),(y0+y1)//2]
        runs=[_light_run(pix,width,yy) for yy in probes]
        runs=[r for r in runs if r is not None]
        if not runs:
            continue
        # The top of a notice may be yellow while the body is white; they share
        # the same x-extent. Prefer longest clean horizontal run.
        left,right=max(runs,key=lambda ab:ab[1]-ab[0])
        # Leave a margin for clipped bubble edges, but not adjacent chats.
        box=(max(0,left-5),max(0,y0-5),min(width-8,max(right+10,int(width*.77))),min(height,y1+5))
        if box[2]-box[0] >= 90:
            boxes.append(box)
    return boxes[-5:]


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
    x0,y0,x1,y1 = win32gui.GetClientRect(hwnd)
    width,height=x1-x0,y1-y0
    if width < MIN_WINDOW_WIDTH or height < MIN_WINDOW_HEIGHT:
        return ChatObservation(title, "", (), "채팅창이 너무 작습니다. 창을 넓혀 주세요.")
    if _foreground_obscures(hwnd, left, top, width, height):
        return ChatObservation(title, "", (), "다른 창이 채팅창을 가리고 있습니다. 창을 앞으로 꺼내 주세요.")
    with mss.mss() as ctx:
        if not _on_monitor(ctx.monitors[0],left,top,width,height):
            return ChatObservation(title,"",(),"화면 밖으로 나간 채팅창이 있습니다.")
        shot=ctx.grab({"left":left,"top":top,"width":width,"height":height})
        image=Image.frombytes("RGB",shot.size,shot.rgb)
    boxes=message_bubble_boxes(image)
    chunks=[]
    matches=[]
    for box in boxes:
        text=image_to_text(image.crop(box))
        if not text:
            continue
        chunks.append(text)
        # Never combine OCR from two different bubbles: one message must
        # independently supply an intact booking ID/name/date/price/room.
        matches.extend(parse_ocr(title,text))
    unique={}
    for reservation in matches:
        unique[(reservation.number,reservation.cancelled)]=reservation
    visible="\n\n----- 메시지 구분 -----\n\n".join(chunks)
    if not boxes:
        return ChatObservation(title,"",(),"흰색 예약 말풍선을 찾지 못했습니다. 채팅창/배율을 확인하세요.")
    if not unique:
        return ChatObservation(title,visible,(),
            f"말풍선 {len(boxes)}개는 찾았지만 완전한 예약을 인식하지 못했습니다. OCR 글자를 확인하세요.")
    return ChatObservation(title,visible,tuple(unique.values()))


def scan_chats(selected_windows: dict[str, int] | None = None) -> list[ChatObservation]:
    """OCR only recognized chat rooms, or user-selected windows after OCR header verification."""
    issue = check_ocr()
    if issue:
        raise RuntimeError(issue)
    candidates = enumerate_window_candidates()
    if not candidates:
        raise RuntimeError("Windows에서 크기 조건을 충족하는 보이는 창이 없습니다. 카카오톡을 화면에 띄워 주세요.")
    by_hwnd = {w.hwnd: w for w in candidates}
    observations = []
    if selected_windows:
        for title in CHAT_TITLES:
            hwnd = selected_windows.get(title)
            if hwnd is None:
                continue
            if hwnd not in by_hwnd:
                observations.append(ChatObservation(title, "", (), "선택했던 창이 닫혔거나 최소화되었습니다. 다시 선택하세요."))
                continue
            try:
                header = _detect_chat_name(hwnd)
                if header != title:
                    preview=capture_one(hwnd,title)
                    observations.append(ChatObservation(title, preview.visible_text, (),
                        "채팅방 상단 이름을 OCR로 확인하지 못했습니다. 아래 내용은 진단 전용이며 자동 기록하지 않습니다."))
                else:
                    observations.append(capture_one(hwnd, title))
            except Exception as exc:
                observations.append(ChatObservation(title, "", (), f"창 읽기 실패: {type(exc).__name__}: {exc}"))
        if not observations:
            raise RuntimeError("선택한 채팅창이 없습니다. '채팅창 직접 선택'에서 두 채팅방을 지정하세요.")
        return observations

    possible = [w for w in candidates if looks_like_kakao(w)]
    if not possible:
        raise RuntimeError(
            f"보이는 창 {len(candidates)}개를 확인했지만 자동 검색에서 카카오톡 후보가 0개입니다. "
            "'채팅창 직접 선택' 버튼으로 NOL 알림과 여기어때를 지정하세요."
        )
    found = set()
    for window in possible:
        try:
            name = _detect_chat_name(window.hwnd)
            if name is None and window.title in CHAT_TITLES:
                name = window.title
            if name is None or name in found:
                continue
            observations.append(capture_one(window.hwnd, name))
            found.add(name)
        except Exception:
            continue
    if not observations:
        raise RuntimeError(
            f"카카오톡 후보 창 {len(possible)}개를 발견했지만 채팅방 이름을 OCR로 확인하지 못했습니다. "
            "'채팅창 직접 선택'으로 두 채팅방을 지정하거나 화면 상단이 가려졌는지 확인하세요."
        )
    return observations



def signature(r: Reservation) -> tuple:
    return (r.channel, r.number, r.cancelled, r.guest, r.check_in, r.room, r.amount, r.vehicle)