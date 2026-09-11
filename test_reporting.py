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
from report_archive import archive_report, build_site, render_financial_table
from openpyxl import Workbook


def make_workbook(path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "증가 기업 목록"
    sheet.append(["제목"])
    sheet.append(["설명"])
    sheet.append(["종목코드", "기업명", "보고서유형", "기준연도", "재무제표", "비교기준",
                  "매출액(당기)", "매출액(전년동기)", "매출액증감률(%)",
                  "영업이익(당기)", "영업이익(전년동기)", "영업이익증감률(%)",
                  "당기순이익(당기)", "당기순이익(전년동기)", "당기순이익증감률(%)"])
    sheet.append(["005930", "기업 <테스트>", "사업보고서", "2026", "연결", "전기(연간)",
                  23594361108, 16512932058, 42.88, 22523390467, 14317961314, 57.31,
                  19646303318, 11953832363, 64.35])
    sheet.append(["350520", "두 번째 기업", "사업보고서", "2026", "연결", "전기(연간)",
                  0, -100, None, -50, -100, 50, 0, 0, None])
    workbook.save(path)
    workbook.close()


class ReportingTests(unittest.TestCase):
    def test_financial_table_exact_amounts_percent_codes_and_escape(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "report.xlsx"
            make_workbook(path)
            html = render_financial_table(path, "2026-09-10")
            for value in ("005930", "23,594,361,108", "42.88%", "64.35%", "두 번째 기업", "—", "-100"):
                self.assertIn(value, html)
            self.assertIn("기업 &lt;테스트&gt;", html)
            self.assertNotIn("기업 <테스트>", html)
            self.assertEqual(html.count('<th scope="col">'), 15)
            self.assertEqual(html.count('<td'), 30)
            self.assertIn('tabindex="0"', html)

    def test_month_pages_isolate_records_and_preserve_downloads(self):
        from html.parser import HTMLParser

        class Links(HTMLParser):
            def __init__(self):
                super().__init__()
                self.hrefs = []

            def handle_starttag(self, tag, attrs):
                if tag == "a":
                    self.hrefs.append(dict(attrs)["href"])

        with tempfile.TemporaryDirectory() as temp:
            root, output = Path(temp) / "reports", Path(temp) / "site"
            workbook = Path(temp) / "sample.xlsx"
            make_workbook(workbook)
            for day in ("2025-12-31", "2026-01-01", "2026-01-15"):
                archive_report("result-" + day, day.replace("-", ""), 1, 1,
                               workbook, root=root, run_date=day)
            build_site(root, output)
            january = (output / "2026-01.html").read_text(encoding="utf-8")
            december = (output / "2025-12.html").read_text(encoding="utf-8")
            self.assertEqual((output / "index.html").read_text(encoding="utf-8"), january)
            self.assertIn("result-2025-12-31", december)
            self.assertNotIn("result-2025-12-31", january)
            self.assertNotIn("result-2026-01-01", december)
            self.assertLess(january.index("result-2026-01-15"), january.index("result-2026-01-01"))
            self.assertIn('href="2025-12.html" aria-current="page"', december)
            self.assertIn('href="2026-01.html" aria-current="page"', january)
            for page in output.glob("*.html"):
                links = Links()
                links.feed(page.read_text(encoding="utf-8"))
                for href in links.hrefs:
                    self.assertTrue((output / href).is_file(), href)
            archive_report("new-month", "20260201", 0, 0, root=root, run_date="2026-02-01")
            build_site(root, output)
            newest = (output / "index.html").read_text(encoding="utf-8")
            self.assertIn("new-month", newest)
            self.assertNotIn("result-2026-01-01", newest)
            self.assertIn('href="2026-02.html"', (output / "2025-12.html").read_text(encoding="utf-8"))

    def test_empty_archive_renders_without_month_tabs(self):
        with tempfile.TemporaryDirectory() as temp:
            output = Path(temp) / "site"
            build_site(Path(temp) / "reports", output)
            html = (output / "index.html").read_text(encoding="utf-8")
            self.assertIn("아직 저장된 결과가 없습니다", html)
            self.assertNotIn('<nav', html)

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
            make_workbook(attachment)
            archive_report("<b>older</b>", "20260908", 4, 0,
                           root=root, run_date="2026-09-09")
            archive_report("<b>newer</b> &lt;script&gt;", "20260909", 5, 1,
                           attachment, root=root, run_date="2026-09-10")
            build_site(root, output)
            html = (output / "index.html").read_text(encoding="utf-8")
            self.assertLess(html.index('datetime="2026-09-10"'), html.index('datetime="2026-09-09"'))
            self.assertIn("&lt;script&gt;", html)
            self.assertNotIn("<script>", html)
            self.assertIn("23,594,361,108", html)
            self.assertIn("기업별 상세 내역", html)
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
