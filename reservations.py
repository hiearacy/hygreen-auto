"""Parse NOL and 여기어때 reservation notifications. No UI automation or network."""
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date

ROOM_TYPES = ("작은방", "도보특가", "침대두개", "큰침대", "욕조방", "푹신한방", "따뜻한방")
HEADERS = re.compile(r"(?m)^\s*\[(?:NOL 알림|여기어때)\]\s*\[[^\]\n]+\]\s*")


@dataclass(frozen=True)
class Reservation:
    channel: str                 # Match the pre-existing Excel channel labels.
    source: str                  # Actual provider, when visible in the notice.
    number: str
    guest: str
    phone: str
    check_in: date | None
    room: str
    amount: int
    vehicle: int | None
    cancelled: bool


def split_notifications(text: str) -> list[str]:
    """Handle either one pasted alert or a multi-message KakaoTalk export."""
    text = text.replace("\r\n", "\n").replace("\r", "\n").strip()
    matches = list(HEADERS.finditer(text))
    if not matches:
        return [text] if text else []
    chunks = []
    for i, hit in enumerate(matches):
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        chunks.append(text[hit.start():end].strip())
    return chunks


def parse_many(text: str) -> list[Reservation]:
    seen: set[tuple[str, bool]] = set()
    result: list[Reservation] = []
    for block in split_notifications(text):
        found = parse_one(block)
        if found is None:
            continue
        key = (found.number, found.cancelled)
        if key not in seen:
            seen.add(key)
            result.append(found)
    return result


def parse_one(text: str) -> Reservation | None:
    text = "\n".join(line.strip() for line in text.replace("\r", "").split("\n") if line.strip())
    if "예약번호" in text and "여기어때" in text:
        return _parse_yeogi(text)
    if ("<숙박>" in text or "<숙박 취소>" in text) and ("NOL" in text or "향그린" in text):
        return _parse_nol(text)
    return None


def _room(s: str) -> str:
    s = s.replace(" ", "")
    if "작은방" in s: return "작은방"
    if "도보특가" in s: return "도보특가"
    if "침대두개" in s or "트윈" in s or "싱글침대2" in s: return "침대두개"
    if "큰침대" in s: return "큰침대"
    if "욕조" in s: return "욕조방"
    if "푹신" in s: return "푹신한방"
    if "따뜻" in s or "따듯" in s: return "따뜻한방"
    return ""  # Unrecognized room must be reviewed, not silently misclassified.


def _vehicle(s: str) -> int | None:
    if re.search(r"차량\s*방문|방문방법\s*:\s*차량", s): return 1
    if re.search(r"도보\s*방문|도보\s*전용|주차\s*불가|방문방법\s*:\s*도보", s): return 0
    return None


def _phone(text: str) -> str:
    # Protect leading zeros; store every phone number as text in Excel.
    m = re.search(r"(?:안심번호|연락처)\s*:\s*(\+?[\d\s-]{9,20})", text)
    if not m:
        m = re.search(r"\b(050\d{8,9}|01[016789]\d{7,8})\b", re.sub(r"[ -]", "", text))
    if not m:
        m = re.search(r"/\s*(\+?\d[\d\s-]{7,19}\d)\s*(?:\n|$)", text)
    return re.sub(r"[^\d+]", "", m.group(1 if m.lastindex else 0)) if m else ""


def _date(year: int, month: int, day: int) -> date | None:
    try:
        if not 2020 <= year <= 2100: return None
        return date(year, month, day)
    except ValueError:
        return None


def _parse_nol(text: str) -> Reservation | None:
    # Avoid inadvertently parsing a contact number as the booking identifier.
    m = re.search(r"(?m)^([A-Za-z0-9_-]{14,32})$", text)
    if not m or not any(ch.isdigit() for ch in m.group(1)):
        return None
    number = m.group(1)
    rest = text[m.end():]
    price = re.search(r"(?m)^([\d,]+)\s*원$", rest)
    amount = int(price.group(1).replace(",", "")) if price else 0
    before_price = rest[:price.start()] if price else rest
    after_price = rest[price.end():] if price else rest
    room = _room(before_price)
    guest = ""
    if price:
        for line in after_price.splitlines():
            if line.strip():
                guest = line.split("/")[0].strip()
                break
    d = re.search(r"(20\d{2})[-./](\d{1,2})[-./](\d{1,2})\s*\([^)]*\)\s*\d{1,2}:\d{2}", after_price)
    if not d:
        d = re.search(r"(20\d{2})[-./](\d{1,2})[-./](\d{1,2})", after_price)
    check_in = _date(*map(int, d.groups())) if d else None
    first_line = text.splitlines()[0:4]
    label = " ".join(first_line)
    provider = "아고다" if "아고다" in label else "씨트립" if "씨트립" in label else "카카오예약" if "카카오" in label else "NOL"
    return Reservation("야놀자", provider, number, guest, _phone(text), check_in, room, amount,
                       _vehicle(text), "<숙박 취소>" in text or "<숙박취소>" in text)


def _field(text: str, name: str) -> str:
    m = re.search(rf"(?m)^{re.escape(name)}\s*:\s*(.+)$", text)
    return m.group(1).strip() if m else ""


def _parse_yeogi(text: str) -> Reservation | None:
    number = _field(text, "예약번호")
    if not re.fullmatch(r"[A-Za-z0-9_-]{8,32}", number):
        return None
    input_date = _field(text, "입실일시")
    d = re.search(r"(?:(20\d{2})[-./])?(\d{1,2})[-./](\d{1,2})", input_date)
    year = int(d.group(1)) if d and d.group(1) else (2000 + int(number[:2]) if re.match(r"^\d{2}", number) else date.today().year)
    check_in = _date(year, int(d.group(2)), int(d.group(3))) if d else None
    amount_text = _field(text, "판매금액")
    price = re.search(r"([\d,]+)\s*원", amount_text)
    amount = int(price.group(1).replace(",", "")) if price else 0
    first = " ".join(text.splitlines()[:6])
    cancelled = bool(re.search(r"예약\s*취소|숙박\s*취소|취소\s*완료", first))
    return Reservation("여기", "여기어때", number, _field(text, "예약자명"), _phone(text),
                       check_in, _room(_field(text, "객실정보")), amount, _vehicle(text), cancelled)


def validate(res: Reservation) -> str | None:
    if not res.number: return "예약번호 없음"
    if res.cancelled: return None  # A cancellation only requires its reservation number.
    if not res.check_in: return "입실일 인식 실패"
    if not res.guest: return "예약자명 인식 실패"
    if res.room not in ROOM_TYPES: return "객실타입 인식 실패 (수동 확인 필요)"
    if res.amount <= 0: return "예약금액 인식 실패"
    return None