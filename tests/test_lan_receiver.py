import tempfile
import unittest
from pathlib import Path
from urllib.error import HTTPError
from urllib.request import Request, urlopen

from lan_receiver import NotificationReceiver, PATH, safe_sender, permitted_address


class BridgeTest(unittest.TestCase):
    def test_sender_whitelist(self):
        self.assertTrue(safe_sender('NOL 알림\n<숙박>\n이천 향그린'))
        self.assertTrue(safe_sender('여기어때\n예약번호: 260101001TEST'))
        self.assertFalse(safe_sender('친구이름\n개인 대화 내용'))
        self.assertFalse(safe_sender(''))

    def test_private_network_only(self):
        self.assertTrue(permitted_address('127.0.0.1'))
        self.assertTrue(permitted_address('192.168.1.53'))
        self.assertFalse(permitted_address('8.8.8.8'))
        self.assertFalse(permitted_address('not-an-ip'))

    def test_server_auth_and_local_queue(self):
        with tempfile.TemporaryDirectory() as temp:
            receiver = NotificationReceiver(Path(temp), 'x' * 40, host='127.0.0.1', port=0)
            port = receiver.start()
            url = f'http://127.0.0.1:{port}{PATH}'
            try:
                def post(message, key='x'*40):
                    data=message.encode('utf-8')
                    req=Request(url, data=data, method='POST', headers={
                        'X-Hygreen-Token':key,
                        'Content-Type':'text/plain; charset=utf-8'})
                    try:
                        with urlopen(req, timeout=3) as resp:
                            return resp.status
                    except HTTPError as exc:
                        return exc.code
                self.assertEqual(post('NOL 알림\n<숙박>\n이천 향그린',key='bad'), 401)
                self.assertEqual(post('친구이름\n안녕하세요'), 422)
                self.assertEqual(post('NOL 알림\n<숙박>\n이천 향그린'), 202)
                incoming=list(Path(temp).glob('notice_*.txt'))
                self.assertEqual(len(incoming),1)
                self.assertIn('이천 향그린',incoming[0].read_text(encoding='utf-8'))
                self.assertEqual(len(list(Path(temp).glob('*.part'))),0)
            finally:
                receiver.stop()

if __name__ == '__main__':
    unittest.main()