import unittest
from unittest.mock import patch
from pc_kakao_ocr import (parse_ocr, signature, classify_chat_header,
                          WindowCandidate, looks_like_kakao, scan_chats, ChatObservation)
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
    def test_room_name_from_header_ocr(self):
        self.assertEqual(classify_chat_header("여기어때 #\n1670-6250"), "여기어때")
        self.assertEqual(classify_chat_header("NOL 알림 @\n1644-1346"), "NOL 알림")

    def test_untrusted_header_rejected(self):
        self.assertIsNone(classify_chat_header("친구 채팅\nNOL 아고다 예약"))
        self.assertIsNone(classify_chat_header("여기어때\nNOL 알림"))
        self.assertIsNone(classify_chat_header("카카오톡\n기타 채팅"))

    def test_kakao_process_names_and_manual_fallback(self):
        self.assertTrue(looks_like_kakao(WindowCandidate(1, "", "KakaoTalkBeta.exe", "", 0, 0, 300, 500)))
        self.assertTrue(looks_like_kakao(WindowCandidate(2, "", "unknown.exe", "EVA_Window_Dblclk", 0, 0, 300, 500)))
        self.assertFalse(looks_like_kakao(WindowCandidate(3, "", "msedge.exe", "Chrome_WidgetWin_1", 0, 0, 300, 500)))

    def test_manual_selection_requires_correct_chat_header(self):
        wins = [WindowCandidate(9, "", "UnknownHosted.exe", "", 100, 300, 310, 500),
                WindowCandidate(11, "", "UnknownHosted.exe", "", 420, 300, 310, 500)]
        with patch("pc_kakao_ocr.check_ocr", return_value=""), \
             patch("pc_kakao_ocr.enumerate_window_candidates", return_value=wins), \
             patch("pc_kakao_ocr._detect_chat_name", side_effect=lambda hwnd: "여기어때" if hwnd == 9 else "NOL 알림"), \
             patch("pc_kakao_ocr.capture_one", side_effect=lambda hwnd,title: ChatObservation(title,"",())):
            obs = scan_chats({"여기어때": 9, "NOL 알림": 11})
            self.assertEqual([x.title for x in obs], ["여기어때", "NOL 알림"])
            mismatch = scan_chats({"여기어때": 11, "NOL 알림": 9})
            self.assertEqual(len(mismatch), 2)
            self.assertTrue(all(not x.reservations and "확인하지 못했습니다" in x.warning for x in mismatch))

    def test_stable_signature(self):
        a = parse_ocr("NOL 알림", NOL)[0]
        b = parse_ocr("NOL 알림", NOL)[0]
        self.assertEqual(signature(a), signature(b))

if __name__ == "__main__":
    unittest.main()
