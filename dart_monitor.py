"""
DART 재무제표 모니터링 스크립트
- 사업보고서, 분기보고서, 반기보고서에서 매출액/영업이익/당기순이익 증가 기업 필터링
- 결과를 엑셀로 저장 후 텔레그램 전송
"""

import os
import sys
import time
import logging
import requests
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── 로깅 설정 ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)
log = logging.getLogger(__name__)

# ── 환경변수 ───────────────────────────────────────────────────────────────────
DART_API_KEY      = os.environ["DART_API_KEY"]
TELEGRAM_TOKEN    = os.environ["TELEGRAM_BOT_TOKEN"]
TELEGRAM_CHAT_ID  = os.environ["TELEGRAM_CHAT_ID"]

BASE_URL = "https://opendart.fss.or.kr/api"

# 조회할 보고서 유형
REPORT_TYPES = {
    "11011": "사업보고서",
    "11012": "반기보고서",
    "11013": "분기보고서(1분기)",
    "11014": "분기보고서(3분기)",
}

# 재무 지표 계정코드
ACCOUNT_CODES = {
    "매출액":     "ifrs-full_Revenue",
    "영업이익":   "dart_OperatingIncomeLoss",
    "당기순이익": "ifrs-full_ProfitLoss",
}


# ── DART API 호출 유틸 ─────────────────────────────────────────────────────────

def dart_get(endpoint: str, params: dict, retries: int = 3) -> dict | None:
    """DART Open API GET 요청 (재시도 포함)"""
    params["crtfc_key"] = DART_API_KEY
    url = f"{BASE_URL}/{endpoint}.json"
    for attempt in range(1, retries + 1):
        try:
            res = requests.get(url, params=params, timeout=30)
            res.raise_for_status()
            data = res.json()
            if data.get("status") == "000":
                return data
            log.warning("DART 오류 [%s] %s (시도 %d)", data.get("status"), data.get("message"), attempt)
            if data.get("status") in ("020",):  # API 한도 초과
                time.sleep(5 * attempt)
        except requests.RequestException as e:
            log.warning("요청 실패: %s (시도 %d)", e, attempt)
            time.sleep(2 * attempt)
    return None


def get_corp_list() -> pd.DataFrame:
    """상장법인 목록 조회 (KOSPI + KOSDAQ)"""
    all_corps = []
    for market in ("Y", "K"):   # Y=KOSPI, K=KOSDAQ
        data = dart_get("corpCode", {"market": market})
        # corpCode 엔드포인트는 XML zip 반환 → 별도 처리
    # fallback: 전체 corpCode.xml 다운로드
    return _download_corp_codes()


def _download_corp_codes() -> pd.DataFrame:
    """DART 전체 기업 코드 XML zip 다운로드 & 파싱"""
    import zipfile, xml.etree.ElementTree as ET
    url = f"{BASE_URL}/corpCode.xml"
    res = requests.get(url, params={"crtfc_key": DART_API_KEY}, timeout=60)
    res.raise_for_status()
    with zipfile.ZipFile(BytesIO(res.content)) as z:
        with z.open("CORPCODE.xml") as f:
            tree = ET.parse(f)
    rows = []
    for item in tree.getroot().findall("list"):
        stock_code = item.findtext("stock_code", "").strip()
        if stock_code:   # 상장사만
            rows.append({
                "corp_code": item.findtext("corp_code", "").strip(),
                "corp_name": item.findtext("corp_name", "").strip(),
                "stock_code": stock_code,
            })
    df = pd.DataFrame(rows)
    log.info("상장법인 %d개 로드 완료", len(df))
    return df


def get_recent_report_year_period() -> list[tuple[str, str]]:
    """
    오늘 날짜 기준으로 조회할 (연도, 보고서코드) 목록 반환
    - 최근 1년 분량의 보고서를 대상으로 함
    """
    today = datetime.today()
    year  = today.year
    month = today.month

    targets = []
    # 사업보고서: 전년도 12월 결산
    targets.append((str(year - 1), "11011"))
    # 반기보고서: 당해 6월
    if month >= 8:
        targets.append((str(year), "11012"))
    # 분기보고서
    if month >= 5:
        targets.append((str(year), "11013"))   # 1분기
    if month >= 11:
        targets.append((str(year), "11014"))   # 3분기
    return targets


def fetch_financial(corp_code: str, year: str, reprt_code: str) -> dict | None:
    """단일 기업의 재무제표 주요 계정 조회"""
    data = dart_get("fnlttSinglAcntAll", {
        "corp_code": corp_code,
        "bsns_year": year,
        "reprt_code": reprt_code,
        "fs_div": "CFS",    # 연결재무제표 우선
    })
    if not data:
        # 연결 없으면 별도재무제표
        data = dart_get("fnlttSinglAcntAll", {
            "corp_code": corp_code,
            "bsns_year": year,
            "reprt_code": reprt_code,
            "fs_div": "OFS",
        })
    return data


def parse_financials(data: dict) -> dict:
    """API 응답에서 매출액/영업이익/당기순이익 추출"""
    result = {}
    if not data or "list" not in data:
        return result

    for item in data["list"]:
        acnt_id  = item.get("account_id", "")
        acnt_nm  = item.get("account_nm", "")
        thstrm   = item.get("thstrm_amount", "").replace(",", "").strip()   # 당기
        frmtrm   = item.get("frmtrm_amount", "").replace(",", "").strip()   # 전기

        for label, code in ACCOUNT_CODES.items():
            if code in acnt_id or label in acnt_nm:
                try:
                    result[f"{label}_당기"] = int(thstrm) if thstrm else None
                    result[f"{label}_전기"] = int(frmtrm) if frmtrm else None
                except ValueError:
                    pass
    return result


def is_all_increasing(fin: dict) -> bool:
    """매출액·영업이익·당기순이익 모두 전기 대비 증가 여부"""
    for label in ACCOUNT_CODES:
        cur  = fin.get(f"{label}_당기")
        prev = fin.get(f"{label}_전기")
        if cur is None or prev is None:
            return False
        if cur <= prev:
            return False
    return True


# ── 메인 수집 로직 ─────────────────────────────────────────────────────────────

def collect_data() -> pd.DataFrame:
    log.info("기업 코드 다운로드 중...")
    corps = _download_corp_codes()

    targets = get_recent_report_year_period()
    log.info("조회 대상 보고서: %s", targets)

    records = []
    total   = len(corps)

    for idx, row in corps.iterrows():
        corp_code = row["corp_code"]
        corp_name = row["corp_name"]
        stock_code = row["stock_code"]

        if idx % 100 == 0:
            log.info("진행 중: %d / %d", idx, total)

        for year, reprt_code in targets:
            data = fetch_financial(corp_code, year, reprt_code)
            fin  = parse_financials(data)
            if not fin:
                continue
            if is_all_increasing(fin):
                records.append({
                    "종목코드":     stock_code,
                    "기업명":       corp_name,
                    "보고서유형":   REPORT_TYPES.get(reprt_code, reprt_code),
                    "기준연도":     year,
                    "매출액_당기":  fin.get("매출액_당기"),
                    "매출액_전기":  fin.get("매출액_전기"),
                    "매출액_증감률": _growth(fin.get("매출액_당기"), fin.get("매출액_전기")),
                    "영업이익_당기": fin.get("영업이익_당기"),
                    "영업이익_전기": fin.get("영업이익_전기"),
                    "영업이익_증감률": _growth(fin.get("영업이익_당기"), fin.get("영업이익_전기")),
                    "당기순이익_당기": fin.get("당기순이익_당기"),
                    "당기순이익_전기": fin.get("당기순이익_전기"),
                    "당기순이익_증감률": _growth(fin.get("당기순이익_당기"), fin.get("당기순이익_전기")),
                })
            time.sleep(0.05)   # API 호출 간격 준수

    df = pd.DataFrame(records)
    log.info("조건 충족 기업: %d개", len(df))
    return df


def _growth(cur, prev) -> float | None:
    if cur is None or prev is None or prev == 0:
        return None
    return round((cur - prev) / abs(prev) * 100, 2)


# ── 엑셀 저장 ──────────────────────────────────────────────────────────────────

HEADER_FILL  = PatternFill("solid", start_color="1F4E79")
SUB_FILL     = PatternFill("solid", start_color="2E75B6")
ALT_FILL     = PatternFill("solid", start_color="EBF3FB")
WHITE_FILL   = PatternFill("solid", start_color="FFFFFF")
WHITE_FONT   = Font(bold=True, color="FFFFFF", name="Arial", size=10)
DARK_FONT    = Font(bold=True, color="1F4E79", name="Arial", size=10)
NORMAL_FONT  = Font(name="Arial", size=9)
GREEN_FONT   = Font(name="Arial", size=9, color="375623")
CENTER       = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT         = Alignment(horizontal="left",   vertical="center")
RIGHT        = Alignment(horizontal="right",  vertical="center")

def _border():
    s = Side(style="thin", color="BDD7EE")
    return Border(left=s, right=s, top=s, bottom=s)


def save_excel(df: pd.DataFrame, filepath: str):
    wb = Workbook()

    # ── 시트 1: 요약 ────────────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "증가 기업 목록"

    today_str = datetime.today().strftime("%Y년 %m월 %d일")
    ws.merge_cells("A1:N1")
    title_cell = ws["A1"]
    title_cell.value = f"DART 재무성장 기업 모니터링  |  {today_str}"
    title_cell.font  = Font(bold=True, color="FFFFFF", name="Arial", size=13)
    title_cell.fill  = HEADER_FILL
    title_cell.alignment = CENTER
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:N2")
    sub = ws["A2"]
    sub.value = "매출액 · 영업이익 · 당기순이익 전기 대비 전부 증가한 기업"
    sub.font  = Font(color="FFFFFF", name="Arial", size=10)
    sub.fill  = SUB_FILL
    sub.alignment = CENTER
    ws.row_dimensions[2].height = 20

    headers = [
        "종목코드", "기업명", "보고서유형", "기준연도",
        "매출액(당기)", "매출액(전기)", "매출액증감률(%)",
        "영업이익(당기)", "영업이익(전기)", "영업이익증감률(%)",
        "당기순이익(당기)", "당기순이익(전기)", "당기순이익증감률(%)",
    ]
    col_widths = [10, 22, 14, 10, 16, 16, 14, 16, 16, 14, 16, 16, 14]

    for col_idx, (header, width) in enumerate(zip(headers, col_widths), start=1):
        cell = ws.cell(row=3, column=col_idx, value=header)
        cell.font      = WHITE_FONT
        cell.fill      = HEADER_FILL
        cell.alignment = CENTER
        cell.border    = _border()
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[3].height = 36

    money_cols  = [5, 6, 8, 9, 11, 12]
    pct_cols    = [7, 10, 13]

    df_sorted = df.sort_values("매출액_증감률", ascending=False).reset_index(drop=True)

    for r_idx, row_data in df_sorted.iterrows():
        excel_row = r_idx + 4
        fill = ALT_FILL if r_idx % 2 == 0 else WHITE_FILL

        vals = [
            row_data["종목코드"], row_data["기업명"], row_data["보고서유형"], row_data["기준연도"],
            row_data["매출액_당기"], row_data["매출액_전기"], row_data["매출액_증감률"],
            row_data["영업이익_당기"], row_data["영업이익_전기"], row_data["영업이익_증감률"],
            row_data["당기순이익_당기"], row_data["당기순이익_전기"], row_data["당기순이익_증감률"],
        ]
        for col_idx, val in enumerate(vals, start=1):
            cell = ws.cell(row=excel_row, column=col_idx, value=val)
            cell.fill   = fill
            cell.border = _border()
            if col_idx in money_cols:
                cell.number_format = '#,##0'
                cell.alignment = RIGHT
                cell.font = NORMAL_FONT
            elif col_idx in pct_cols:
                cell.number_format = '0.00"%"'
                cell.alignment = RIGHT
                cell.font = GREEN_FONT
            elif col_idx == 2:
                cell.alignment = LEFT
                cell.font = NORMAL_FONT
            else:
                cell.alignment = CENTER
                cell.font = NORMAL_FONT

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:{get_column_letter(len(headers))}{len(df_sorted)+3}"

    # ── 시트 2: 보고서 유형별 요약 ─────────────────────────────────────────────
    ws2 = wb.create_sheet("보고서유형별 요약")
    ws2.merge_cells("A1:E1")
    t = ws2["A1"]
    t.value = "보고서 유형별 증가 기업 수"
    t.font  = Font(bold=True, color="FFFFFF", name="Arial", size=12)
    t.fill  = HEADER_FILL
    t.alignment = CENTER
    ws2.row_dimensions[1].height = 28

    summary_headers = ["보고서유형", "기업수", "매출액증감률평균(%)", "영업이익증감률평균(%)", "당기순이익증감률평균(%)"]
    for ci, h in enumerate(summary_headers, 1):
        c = ws2.cell(row=2, column=ci, value=h)
        c.font = WHITE_FONT; c.fill = SUB_FILL; c.alignment = CENTER; c.border = _border()
        ws2.column_dimensions[get_column_letter(ci)].width = 22

    for ri, (rtype, grp) in enumerate(df.groupby("보고서유형"), start=3):
        ws2.cell(row=ri, column=1, value=rtype).alignment = CENTER
        ws2.cell(row=ri, column=2, value=len(grp)).alignment = CENTER
        ws2.cell(row=ri, column=3, value=round(grp["매출액_증감률"].mean(), 2)).number_format = '0.00"%"'
        ws2.cell(row=ri, column=4, value=round(grp["영업이익_증감률"].mean(), 2)).number_format = '0.00"%"'
        ws2.cell(row=ri, column=5, value=round(grp["당기순이익_증감률"].mean(), 2)).number_format = '0.00"%"'
        for ci in range(1, 6):
            ws2.cell(row=ri, column=ci).border = _border()

    wb.save(filepath)
    log.info("엑셀 저장 완료: %s", filepath)


# ── 텔레그램 전송 ──────────────────────────────────────────────────────────────

def send_telegram_message(text: str):
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": text,
        "parse_mode": "HTML",
        "disable_web_page_preview": True,
    }
    res = requests.post(url, json=payload, timeout=30)
    res.raise_for_status()
    log.info("텔레그램 메시지 전송 완료")


def send_telegram_file(filepath: str, caption: str):
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendDocument"
    with open(filepath, "rb") as f:
        res = requests.post(
            url,
            data={"chat_id": TELEGRAM_CHAT_ID, "caption": caption, "parse_mode": "HTML"},
            files={"document": (os.path.basename(filepath), f,
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
            timeout=60,
        )
    res.raise_for_status()
    log.info("텔레그램 파일 전송 완료: %s", filepath)


def build_summary_message(df: pd.DataFrame) -> str:
    today_str = datetime.today().strftime("%Y-%m-%d")
    lines = [
        f"📊 <b>DART 재무성장 기업 모니터링</b>",
        f"📅 기준일: {today_str}",
        f"",
        f"✅ <b>매출액·영업이익·당기순이익 모두 증가한 기업</b>: {len(df)}개",
        f"",
    ]

    for rtype, grp in df.groupby("보고서유형"):
        lines.append(f"📁 <b>{rtype}</b> ({len(grp)}개)")
        top5 = grp.nlargest(5, "매출액_증감률")
        for _, r in top5.iterrows():
            lines.append(
                f"  • {r['기업명']} ({r['종목코드']}) "
                f"매출 +{r['매출액_증감률']:.1f}% / 영업이익 +{r['영업이익_증감률']:.1f}%"
            )
        lines.append("")

    lines.append("📎 상세 내역은 첨부 엑셀 파일을 확인하세요.")
    return "\n".join(lines)


# ── 엔트리포인트 ───────────────────────────────────────────────────────────────

def main():
    today_str = datetime.today().strftime("%Y%m%d")
    excel_path = f"dart_report_{today_str}.xlsx"

    try:
        df = collect_data()

        if df.empty:
            send_telegram_message(
                f"📊 <b>DART 모니터링 결과</b>\n"
                f"📅 {datetime.today().strftime('%Y-%m-%d')}\n\n"
                f"오늘은 조건을 충족하는 기업이 없습니다."
            )
            return

        save_excel(df, excel_path)

        summary = build_summary_message(df)
        send_telegram_message(summary)
        send_telegram_file(excel_path, caption=f"DART 재무성장 기업 목록 ({today_str})")

    except Exception as e:
        log.error("실행 중 오류 발생: %s", e, exc_info=True)
        try:
            send_telegram_message(f"⚠️ DART 모니터링 오류 발생\n<code>{e}</code>")
        except Exception:
            pass
        sys.exit(1)


if __name__ == "__main__":
    main()
