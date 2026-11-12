import unittest
from datetime import date
from reservations import parse_many, parse_one, validate

NOL = '''NOL 아고다 당일예약
<숙박>
이천 향그린
2699990000000001
이천 푹신한방 [푹신한방-도보전용, 주차불가]
45,000원
테스트고객 / 82216441346
2026-11-11(수) 17:00
~2026-11-12(목) 12:00 (1박)
도보방문'''
CANCEL = '''NOL 미리예약
<숙박 취소>
이천 향그린
2699990000000002
이천 작은방 [작은방-도보전용, 주차불가]
40,000원
가명고객 / 050700000000
2026-11-17(화) 17:00
~2026-11-18(수) 12:00 (1박)
도보방문'''
YEOGI = '''[여기어때] [오후 9:48] [여기어때] 당일예약 - 숙박
[여기어때] 당일예약 - 숙박
예약번호: 269999001356BYE1
객실정보: 이천 향그린 / 도보특가 랜덤배정
판매금액: 40,000원
입실일시: 11/10 (화) 17:00
퇴실일시: 11/11 (수) 12:00 (1박)
예약자명: 가상예약자
안심번호: 050440000000
방문방법: 도보'''

class TestParsing(unittest.TestCase):
    def test_agoda(self):
        r = parse_one(NOL)
        self.assertEqual((r.channel, r.source, r.room, r.vehicle, r.amount), ("야놀자", "아고다", "푹신한방", 0, 45000))
        self.assertEqual(r.check_in, date(2026, 11, 11))
        self.assertIsNone(validate(r))
    def test_cancel(self):
        r = parse_one(CANCEL)
        self.assertTrue(r.cancelled)
        self.assertEqual(r.number, "2699990000000002")
        self.assertIsNone(validate(r))
    def test_yeogi(self):
        r = parse_one(YEOGI)
        self.assertEqual((r.channel,r.room,r.amount,r.vehicle), ("여기","도보특가",40000,0))
        self.assertEqual(r.check_in, date(2026,11,10))
        self.assertEqual(r.phone, "050440000000")
    def test_multi_and_repeat(self):
        pasted = '[NOL 알림] [오전 9:47] ' + NOL + '\n[NOL 알림] [오전 9:48] ' + CANCEL + '\n[NOL 알림] [오전 9:49] ' + NOL
        events = parse_many(pasted)
        self.assertEqual(len(events), 2)
    def test_unknown_room_blocks_write(self):
        r = parse_one(NOL.replace("푹신한방", "미지의객실"))
        self.assertIn("객실타입", validate(r))
    def test_unrelated_does_not_parse(self):
        self.assertIsNone(parse_one("안녕하세요. 내일 체크인 예약입니다."))

if __name__ == '__main__':
    unittest.main()