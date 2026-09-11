"""Daily report storage and a JavaScript-free GitHub Pages archive."""

import json
import shutil
from datetime import datetime, timedelta, timezone
from html import escape, unescape
from pathlib import Path

KST = timezone(timedelta(hours=9))


def plain_message(message):
    # The shared message only uses these Telegram formatting tags.
    for tag in ("<b>", "</b>", "<code>", "</code>"):
        message = message.replace(tag, "")
    return unescape(message)


def archive_report(message, base_date, total_count, matched_count,
                   attachment_path=None, root="reports", run_date=None):
    now = datetime.now(KST)
    day = run_date or now.strftime("%Y-%m-%d")
    datetime.strptime(day, "%Y-%m-%d")
    directory = Path(root) / day
    directory.mkdir(parents=True, exist_ok=True)
    attachment = None
    if attachment_path:
        attachment = "report.xlsx"
        shutil.copyfile(attachment_path, directory / attachment)
    record = {
        "date": day, "updated_at": now.isoformat(), "base_date": base_date,
        "total_count": total_count, "matched_count": matched_count,
        "message": plain_message(message), "attachment": attachment,
    }
    temporary = directory / "report.json.tmp"
    temporary.write_text(json.dumps(record, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temporary.replace(directory / "report.json")


def build_site(root="reports", output="site"):
    output = Path(output)
    output.mkdir(parents=True, exist_ok=True)
    cards_by_month = {}
    for source in sorted(Path(root).glob("????-??-??/report.json"), reverse=True):
        record = json.loads(source.read_text(encoding="utf-8"))
        day = source.parent.name
        datetime.strptime(day, "%Y-%m-%d")
        month = day[:7]
        link = ""
        if record["attachment"]:
            target = output / "downloads" / day
            target.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(source.parent / "report.xlsx", target / "report.xlsx")
            link = f'<a href="downloads/{day}/report.xlsx" download>상세 엑셀 다운로드 ↗</a>'
        base = record["base_date"]
        base = f"{base[:4]}-{base[4:6]}-{base[6:]}" if len(base) == 8 else "해당 공시 없음"
        cards_by_month.setdefault(month, []).append(f'''<article>
<div class="heading"><h2><time datetime="{day}">{day}</time></h2>
<span class="badge">조건 충족 {int(record['matched_count'])}개</span></div>
<p class="meta">공시 기준일 {escape(base)} · 조회 기업 {int(record['total_count'])}개</p>
<pre>{escape(record['message'])}</pre>{link}</article>''')
    template = '''<!doctype html>
<html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="description" content="DART 재무성장 기업 모니터링의 날짜별 결과">
<title>DART | __MONTH_TITLE__</title>
<style>
:root{color-scheme:light;font-family:system-ui,-apple-system,"Malgun Gothic",sans-serif;color:#172b3a;background:#f3f5f7}
*{box-sizing:border-box}body{margin:0;border-top:5px solid #166b60}main{max-width:920px;margin:auto;padding:48px 24px}
header{margin-bottom:32px}.eyebrow{color:#166b60;font-size:13px;font-weight:750;letter-spacing:.12em}
h1{font-size:32px;letter-spacing:-.06em;margin:12px 0}header p,.meta,footer{color:#63717d;line-height:1.7}
article{background:white;border:1px solid #dce3e7;border-radius:14px;padding:28px;margin:0 0 20px}
.heading{display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap}h2{font-size:23px;margin:0}
.badge{background:#e4f3ed;color:#176348;padding:7px 12px;border-radius:30px;font-size:13px;font-weight:650}
.meta{font-size:13px;margin:10px 0 20px}pre{white-space:pre-wrap;overflow-wrap:anywhere;font:inherit;line-height:1.85;font-size:14px}
a{display:inline-block;color:#126458;font-weight:650;text-underline-offset:4px;margin-top:12px}a:focus-visible{outline:3px solid #166b60;outline-offset:5px}
.months{display:flex;gap:8px;overflow-x:auto;padding:6px 3px 14px;margin-bottom:16px}
.months a{flex-shrink:0;margin:0;padding:10px 16px;border:1px solid #cbd8d5;border-radius:8px;text-decoration:none;background:white}
.months a[aria-current="page"]{background:#166b60;color:white;border-color:#166b60}
.month-title{font-size:19px;margin:0 0 20px}.month-title span{font-size:13px;font-weight:400;color:#63717d;margin-left:10px}
footer{font-size:12px;margin-top:28px}@media(max-width:520px){main{padding:30px 16px}article{padding:20px}h1{font-size:27px}pre{font-size:13px}}
</style></head><body><main><header><div class="eyebrow">DART / DAILY REPORT</div>
<h1>재무성장 기업 모니터링</h1><p>월을 선택하면 해당 월의 결과를 최신 날짜부터 볼 수 있습니다.<br>매출액 · 영업이익 · 당기순이익이 모두 증가한 기업을 기록합니다.</p>
</header>__MONTH_CONTENT__<footer>날짜는 한국시간 실행일 기준입니다. 휴일에는 이전 공시일의 결과가 반복될 수 있습니다.<br>같은 날 다시 실행하면 해당 날짜의 결과가 갱신됩니다.</footer></main></body></html>'''
    months = sorted(cards_by_month, reverse=True)
    for month in months or [None]:
        title = f"{month[:4]}년 {int(month[5:])}월" if month else "일별 모니터링 기록"
        tabs = []
        for target_month in months:
            current = ' aria-current="page"' if target_month == month else ''
            label = f"{target_month[:4]}년 {int(target_month[5:])}월"
            tabs.append(f'<a href="{target_month}.html"{current}>{label}</a>')
        if month:
            content = ('<nav class="months" aria-label="월별 결과">' + ''.join(tabs) + '</nav>'
                       + f'<h2 class="month-title">{title}<span>{len(cards_by_month[month])}일의 기록</span></h2>'
                       + '\n'.join(cards_by_month[month]))
        else:
            content = '<article><h2>아직 저장된 결과가 없습니다.</h2><p>첫 모니터링 실행 후 날짜별 결과가 여기에 쌓입니다.</p></article>'
        page = template.replace('__MONTH_TITLE__', title).replace('__MONTH_CONTENT__', content)
        if month:
            (output / f"{month}.html").write_text(page, encoding="utf-8")
        if not month or month == months[0]:
            (output / "index.html").write_text(page, encoding="utf-8")
    (output / ".nojekyll").touch()


if __name__ == "__main__":
    build_site()
