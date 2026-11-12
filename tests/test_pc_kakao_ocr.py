import unittest
from pc_kakao_ocr import parse_ocr, signature
from test_reservations import NOL, CANCEL, YEOGI

class TestPCOCRParsing(unittest.TestCase):
    def test_nol_single(self):
        reservations = parse_ocr("NOL 알림", NOL)
        self.assertEqual(len(reservations), 1)
        self.assertEqual(reservations[0].amount, 45000)
    def test_yeogi_single(self):
        reservations = parse_ocr("여기어때", YEOGI)
        self.assertEqual(len(reservations), 1)
        self.assertEqual(reservations[0].amount, 40000)
    def test_nol_cancel(self):
        reservations = parse_ocr("NOL 알림", CANCEL)
        self.assertEqual(len(reservations), 1)
        self.assertTrue(reservations[0].cancelled)
    def test_wrong_sender_disallowed(self):
        self.assertEqual(parse_ocr("친구 단톡", NOL), ())
        self.assertEqual(parse_ocr("여기어때", NOL), ())
    def test_incomplete_ocr_blocks(self):
        self.assertEqual(parse_ocr("NOL 알림", NOL.replace("45,000원", "45,OOO원")), ())
    def test_stable_signature(self):
        a = parse_ocr("NOL 알림", NOL)[0]
        b = parse_ocr("NOL 알림", NOL)[0]
        self.assertEqual(signature(a), signature(b))

if __name__ == "__main__":
    unittest.main()
