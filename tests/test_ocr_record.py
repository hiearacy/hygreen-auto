import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

from app import App, execute_ocr_batch
from pc_kakao_ocr import parse_ocr
from reservations import parse_many
from test_reservations import NOL, CANCEL, YEOGI


class OCRRecordTests(unittest.TestCase):
    def setUp(self):
        self.events = parse_ocr('여기어때', YEOGI) + parse_ocr('NOL 알림', NOL)
        self.raw = '[여기어때] OCR 결과 (직접 검토 필요)\n' + YEOGI + '\n\n--------\n\n[NOL 알림] OCR 결과 (직접 검토 필요)\n' + NOL + '\n'
        self.app = App.__new__(App)
        self.app.text = Mock()
        self.app.text.get.return_value = self.raw
        self.app.pc_preview_text = self.raw
        self.app.pc_preview_events = self.events
        self.app.state = Mock()
        self.app._log = Mock()
        self.app._valid_excel = Mock(return_value=True)
        self.app.future = None
        self.app.pc_requested_manual_scan = False
        self.app.excel = Mock()
        self.app.excel.get.return_value = 'test.xlsx'
        self.app.pool = Mock()

    def test_original_loss_reproduced_and_fixed(self):
        self.assertEqual(len(parse_many(self.raw)), 1)
        self.assertEqual(self.app._displayed_events(), self.events)
        with patch('app.messagebox.askyesno', return_value=True) as confirm:
            self.app._manual()
        self.assertIn('2건', confirm.call_args.args[1])
        self.app.pool.submit.assert_called_once_with(execute_ocr_batch, 'test.xlsx', self.events)

    def test_batch_writes_both_exact_objects(self):
        with patch('app.apply_reservation', side_effect=['첫 예약 입력', '두번째 예약 입력']) as write:
            self.assertEqual(len(execute_ocr_batch('test.xlsx', self.events)), 2)
        self.assertEqual([c.args[1] for c in write.call_args_list], list(self.events))

    def test_edited_ocr_text_cannot_write_stale_data(self):
        self.app.text.get.return_value += '수정'
        with self.assertRaises(ValueError):
            self.app._displayed_events()
        with patch('app.messagebox.showwarning'):
            self.app._manual()
        self.app.pool.submit.assert_not_called()

    def test_empty_ocr_never_reparsed(self):
        self.app.pc_preview_events = ()
        self.assertEqual(self.app._displayed_events(), ())
        with patch('app.messagebox.showwarning'):
            self.app._manual()
        self.app.pool.submit.assert_not_called()

    def test_diagnostic_text_not_allowed_after_reset(self):
        self.app.pc_preview_events = None
        with self.assertRaises(ValueError):
            self.app._displayed_events()

    def test_pasted_text_mode_still_works(self):
        self.app.pc_preview_events = None
        self.app.text.get.return_value = NOL
        self.assertEqual(tuple(self.app._displayed_events()), parse_ocr('NOL 알림', NOL))

    def test_cancel_passed_as_cancel_and_missing_target_reported(self):
        cancel = parse_ocr('NOL 알림', CANCEL.replace('<숙박 취소>', '<숙박 취소'))[0]
        self.assertTrue(cancel.cancelled)
        with patch('app.apply_reservation', side_effect=['취소 대상 예약번호 없음: TEST (수동 확인 필요)', '입력 완료']) as write:
            result = execute_ocr_batch('test.xlsx', (cancel, self.events[0]))
        self.assertEqual(len(result), 2)
        self.assertIn('취소 대상 예약번호 없음', result[0])
        self.assertTrue(write.call_args_list[0].args[1].cancelled)

    def test_poll_missing_cancel_not_marked_finished(self):
        self.app.future = Mock()
        self.app.future.done.return_value = True
        self.app.future.result.return_value = ['취소 대상 예약번호 없음: TEST (수동 확인 필요)']
        self.app.pc_processing = ('test',)
        self.app.pc_finished = set()
        self.app.pc_failed = set()
        self.app.work_file = self.app.work_hash = None
        self.app.root = Mock()
        self.app._poll()
        self.assertNotIn(('test',), self.app.pc_finished)
        self.assertIn(('test',), self.app.pc_failed)

if __name__ == '__main__':
    unittest.main()
