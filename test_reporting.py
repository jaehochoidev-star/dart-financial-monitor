import json
import logging
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("DART_API_KEY", "test")
os.environ.setdefault("TELEGRAM_BOT_TOKEN", "test")
os.environ.setdefault("TELEGRAM_CHAT_ID", "test")

import pandas as pd
import dart_monitor as monitor
from report_archive import archive_report, build_site


class ReportingTests(unittest.TestCase):
    def test_logs_redact_credentials_and_recipient_in_traceback(self):
        with patch.multiple(monitor, DART_API_KEY="sample-api-key",
                            TELEGRAM_TOKEN="sample-bot-token", TELEGRAM_CHAT_ID="123456789",
                            SMTP_PASSWORD="sample-password", EMAIL_TO="private@example.com"):
            try:
                raise RuntimeError("sample-api-key sample-bot-token 123456789 sample-password private@example.com")
            except RuntimeError:
                import sys
                record = logging.LogRecord("test", logging.ERROR, "test", 1,
                                           "Failed for %s", ("private@example.com",), sys.exc_info())
            output = monitor.PrivateLogFormatter().format(record)
            for value in ("sample-api-key", "sample-bot-token", "123456789", "sample-password", "private@example.com"):
                self.assertNotIn(value, output)
            self.assertIn("[REDACTED]", output)

    def test_history_order_escape_attachment_and_rerun(self):
        with tempfile.TemporaryDirectory() as temp:
            root, output = Path(temp) / "reports", Path(temp) / "site"
            attachment = Path(temp) / "test.xlsx"
            attachment.write_bytes(b"test workbook")
            archive_report("<b>older</b>", "20260908", 4, 0,
                           root=root, run_date="2026-09-09")
            archive_report("<b>newer</b> &lt;script&gt;", "20260909", 5, 1,
                           attachment, root=root, run_date="2026-09-10")
            build_site(root, output)
            html = (output / "index.html").read_text(encoding="utf-8")
            self.assertLess(html.index('datetime="2026-09-10"'), html.index('datetime="2026-09-09"'))
            self.assertIn("&lt;script&gt;", html)
            self.assertNotIn("<script>", html)
            self.assertEqual((output / "downloads/2026-09-10/report.xlsx").read_bytes(), attachment.read_bytes())
            archive_report("없음", "20260909", 5, 0, root=root, run_date="2026-09-10")
            self.assertEqual(len(list(root.glob("*/report.json"))), 2)
            record = json.loads((root / "2026-09-10/report.json").read_text(encoding="utf-8"))
            self.assertIsNone(record["attachment"])

    def test_email_same_content_and_attachment_both_tls_modes(self):
        with tempfile.TemporaryDirectory() as temp:
            attachment = Path(temp) / "report.xlsx"
            attachment.write_bytes(b"workbook")
            for port in (587, 465):
                with self.subTest(port=port), patch.multiple(
                    monitor, EMAIL_TO="one@example.com,two@example.com",
                    EMAIL_FROM="sender@example.com", SMTP_USER="sender@example.com",
                    SMTP_PASSWORD="fake", SMTP_PORT=port,
                ), patch.object(monitor.smtplib, "SMTP") as smtp, patch.object(monitor.smtplib, "SMTP_SSL") as ssl:
                    monitor.send_email("결과", "<b>결과</b> &amp; 기업", str(attachment))
                    client = (ssl if port == 465 else smtp).return_value.__enter__.return_value
                    sent = client.send_message.call_args.args[0]
                    self.assertIn("결과 & 기업", sent.get_body(preferencelist=("plain",)).get_content())
                    self.assertIn("<b>결과</b> &amp; 기업", sent.get_body(preferencelist=("html",)).get_content())
                    self.assertEqual(next(sent.iter_attachments()).get_payload(decode=True), b"workbook")
                    self.assertEqual(len(sent["To"].addresses), 2)
                    self.assertEqual(client.starttls.call_count, int(port == 587))

    def test_telegram_failure_still_attempts_email_and_archive(self):
        with patch.object(monitor, "collect_data", return_value=(pd.DataFrame(), "20260909", 5)), \
             patch.object(monitor, "send_message", side_effect=RuntimeError("offline")), \
             patch.object(monitor, "send_email") as email, \
             patch.object(monitor, "archive_report") as archive:
            with self.assertRaises(SystemExit) as result:
                monitor.main()
            self.assertEqual(result.exception.code, 1)
            email.assert_called_once()
            archive.assert_called_once()
            self.assertEqual(email.call_args.args[1], archive.call_args.args[0])

    def test_collection_failure_does_not_record_empty_result(self):
        with patch.object(monitor, "collect_data", side_effect=RuntimeError("failed")), \
             patch.object(monitor, "send_message"), patch.object(monitor, "archive_report") as archive:
            with self.assertRaises(SystemExit):
                monitor.main()
            archive.assert_not_called()

    def test_nonempty_result_uses_same_message_and_workbook(self):
        frame = pd.DataFrame([{"기업명": "테스트 기업"}])
        with patch.object(monitor, "collect_data", return_value=(frame, "20260909", 5)), \
             patch.object(monitor, "save_excel") as workbook, \
             patch.object(monitor, "build_message", return_value="<b>동일 결과</b>"), \
             patch.object(monitor, "send_message") as telegram, \
             patch.object(monitor, "send_file") as document, \
             patch.object(monitor, "send_email") as email, \
             patch.object(monitor, "archive_report") as archive:
            monitor.main()
            path = workbook.call_args.args[1]
            self.assertEqual(telegram.call_args.args[0], email.call_args.args[1])
            self.assertEqual(email.call_args.args[1], archive.call_args.args[0])
            self.assertEqual(path, document.call_args.args[0])
            self.assertEqual(path, email.call_args.args[2])
            self.assertEqual(path, archive.call_args.args[4])


if __name__ == "__main__":
    unittest.main()
