"""
DART 재무제표 모니터링 스크립트
- 전날 공시된 사업/반기/분기 보고서 기업만 대상
- 전년도 동일 분기 대비 매출액·영업이익·당기순이익 모두 증가한 기업 필터링
- 연결재무제표(CFS) 우선, 없을 때만 별도재무제표(OFS) 사용
- 결과를 엑셀로 저장 후 텔레그램 전송

[비교 기준]
- 사업보고서  : thstrm_amount      vs frmtrm_amount      (당기 연간 vs 전년도 연간)
- 분기/반기   : thstrm_add_amount  vs frmtrm_add_amount  (당기 누적 vs 전년 동기 누적)
  * 단일 분기 비교가 필요하면 thstrm_amount vs frmtrm_q_amount 로 변경 가능
- 연결재무제표(CFS) 우선 → 없으면 별도재무제표(OFS)
"""

import os, sys, time, logging, smtplib, requests
from datetime import datetime, timedelta
from email.message import EmailMessage
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── 로깅 ───────────────────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s [%(levelname)s] %(message)s",
                    handlers=[logging.StreamHandler(sys.stdout)])
log = logging.getLogger(__name__)

# ── 환경변수 ───────────────────────────────────────────────────────────────────
DART_API_KEY     = os.environ["DART_API_KEY"]
TELEGRAM_TOKEN   = os.environ["TELEGRAM_BOT_TOKEN"]
TELEGRAM_CHAT_ID = os.environ["TELEGRAM_CHAT_ID"]
SMTP_HOST        = os.getenv("SMTP_HOST") or "smtp.gmail.com"
SMTP_PORT        = int(os.getenv("SMTP_PORT") or "587")
SMTP_USER        = os.getenv("SMTP_USER", "")
SMTP_PASSWORD    = os.getenv("SMTP_PASSWORD", "")
EMAIL_FROM       = os.getenv("EMAIL_FROM", SMTP_USER)
EMAIL_TO         = os.getenv("EMAIL_TO", "")


BASE_URL = "https://opendart.fss.or.kr/api"

# 공시 상세 유형코드 → (reprt_code, 보고서명, 연간여부)
TARGET_REPORT_CODES = {
    "A001": ("11011", "사업보고서",       True),   # 연간: 당기 vs 전기 연간
    "A002": ("11012", "반기보고서",       False),  # 누적: 상반기 vs 전년 상반기
    "A003": ("11013", "분기보고서(1분기)", False),  # 누적: 1Q vs 전년 1Q
    "A004": ("11014", "분기보고서(3분기)", False),  # 누적: 9M vs 전년 9M
}

# 손익계산서(IS) 또는 포괄손익계산서(CIS) 계정 ID
ACCOUNT_MAP = {
    "매출액":     ["ifrs-full_Revenue",
                   "ifrs-full_GrossProfit",       # 일부 기업 대체 표기
                   "dart_Revenues"],
    "영업이익":   ["dart_OperatingIncomeLoss",
                   "ifrs-full_ProfitLossFromOperatingActivities"],
    "당기순이익": ["ifrs-full_ProfitLoss",
                   "ifrs-full_ProfitLossAttributableToOwnersOfParent"],
}

# 손익 관련 재무제표 구분만 사용 (재무상태표 BS 제외)
INCOME_SJ_DIV = {"IS", "CIS"}


# ── DART API ───────────────────────────────────────────────────────────────────

def dart_get(endpoint: str, params: dict, retries: int = 3) -> dict | None:
    params = {**params, "crtfc_key": DART_API_KEY}
    url = f"{BASE_URL}/{endpoint}.json"
    for attempt in range(1, retries + 1):
        try:
            res = requests.get(url, params=params, timeout=30)
            res.raise_for_status()
            data = res.json()
            if data.get("status") == "000":
                return data
            log.warning("DART 오류 [%s] %s", data.get("status"), data.get("message"))
            if data.get("status") == "020":   # API 한도 초과
                time.sleep(5 * attempt)
        except requests.RequestException as e:
            log.warning("요청 실패: %s (시도 %d)", e, attempt)
            time.sleep(2 * attempt)
    return None


# ── 공시 목록 조회 ─────────────────────────────────────────────────────────────

def get_recent_disclosures() -> tuple[list[dict], str]:
    """
    가장 최근 영업일(최대 7일 이전)에 공시된 보고서 목록 반환
    주말·공휴일 자동 대응
    """
    today = datetime.today()
    for days_back in range(1, 8):
        target_date = today - timedelta(days=days_back)
        date_str = target_date.strftime("%Y%m%d")
        disclosures = []

        for pblntf_detail_ty, (reprt_code, reprt_name, is_annual) in TARGET_REPORT_CODES.items():
            page = 1
            while True:
                data = dart_get("list", {
                    "bgn_de": date_str,
                    "end_de": date_str,
                    "pblntf_detail_ty": pblntf_detail_ty,
                    "page_no": str(page),
                    "page_count": "100",
                })
                if not data or "list" not in data:
                    break
                for item in data["list"]:
                    stock_code = item.get("stock_code", "").strip()
                    if not stock_code:   # 비상장사 제외
                        continue
                    disclosures.append({
                        "corp_code":  item["corp_code"],
                        "corp_name":  item["corp_name"],
                        "stock_code": stock_code,
                        "reprt_code": reprt_code,
                        "reprt_name": reprt_name,
                        "is_annual":  is_annual,
                        "rcept_dt":   item.get("rcept_dt", date_str),
                        "bsns_year":  item.get("rcept_dt", date_str)[:4],
                    })
                total = int(data.get("total_count", 0))
                if page * 100 >= total:
                    break
                page += 1

        if disclosures:
            log.info("%s 공시 %d건 발견 (상장사)", date_str, len(disclosures))
            return disclosures, date_str

        log.info("%s 공시 없음, 이전 날 탐색 중...", date_str)

    log.warning("최근 7일간 공시 없음")
    return [], ""


# ── 재무 데이터 수집 ───────────────────────────────────────────────────────────

def fetch_financials(corp_code: str, bsns_year: str, reprt_code: str) -> tuple[dict | None, str]:
    """
    연결재무제표(CFS) 우선 조회.
    연결이 없는 기업(자회사 없음)이면 별도재무제표(OFS) 사용.
    사용된 재무제표 구분(fs_div)도 함께 반환.
    """
    for fs_div in ("CFS", "OFS"):
        data = dart_get("fnlttSinglAcntAll", {
            "corp_code": corp_code,
            "bsns_year": bsns_year,
            "reprt_code": reprt_code,
            "fs_div": fs_div,
        })
        if data and data.get("list"):
            return data, fs_div
    return None, ""


def _to_int(value: str) -> int | None:
    """콤마 제거 후 정수 변환. 빈 값·비숫자는 None"""
    v = value.replace(",", "").strip() if value else ""
    if not v or v in ("-", ""):
        return None
    try:
        return int(v)
    except ValueError:
        return None


def parse_financials(data: dict, is_annual: bool) -> dict:
    """
    API 응답에서 매출액·영업이익·당기순이익의 당기/전년동기 금액 추출.

    [비교 로직]
    - 사업보고서(is_annual=True):
        당기  = thstrm_amount   (당기 연간 합계)
        전년  = frmtrm_amount   (전기 연간 합계)

    - 분기/반기(is_annual=False):
        당기  = thstrm_add_amount   (당기 누적, e.g. 1Q·반기·9M)
        전년  = frmtrm_add_amount   (전년 동기 누적)
        * frmtrm_add_amount가 없으면 frmtrm_q_amount(단일분기 전년값) 로 대체
    """
    result = {}
    if not data:
        return result

    for item in data.get("list", []):
        # 손익계산서·포괄손익계산서만 사용
        if item.get("sj_div", "") not in INCOME_SJ_DIV:
            continue

        acnt_id = item.get("account_id", "")
        acnt_nm = item.get("account_nm", "")

        for label, id_list in ACCOUNT_MAP.items():
            if label in result:   # 이미 찾은 계정은 skip
                continue
            matched = any(code in acnt_id for code in id_list) or label in acnt_nm
            if not matched:
                continue

            if is_annual:
                cur  = _to_int(item.get("thstrm_amount", ""))
                prev = _to_int(item.get("frmtrm_amount", ""))
            else:
                cur  = _to_int(item.get("thstrm_add_amount", ""))
                prev = _to_int(item.get("frmtrm_add_amount", ""))
                # frmtrm_add_amount 없으면 frmtrm_q_amount(전년 동일 분기 단독) 대체
                if prev is None:
                    prev = _to_int(item.get("frmtrm_q_amount", ""))

            if cur is not None:   # 당기값이 있어야 유효
                result[f"{label}_당기"] = cur
                result[f"{label}_전년동기"] = prev

    return result


def is_all_increasing(fin: dict) -> bool:
    """세 지표 모두 전년 동기 대비 증가 여부"""
    for label in ACCOUNT_MAP:
        cur  = fin.get(f"{label}_당기")
        prev = fin.get(f"{label}_전년동기")
        if cur is None or prev is None:
            return False
        if cur <= prev:
            return False
    return True


def growth(cur, prev) -> float | None:
    if cur is None or prev is None or prev == 0:
        return None
    return round((cur - prev) / abs(prev) * 100, 2)


# ── 메인 수집 ──────────────────────────────────────────────────────────────────

def collect_data() -> tuple[pd.DataFrame, str, int]:
    disclosures, base_date = get_recent_disclosures()
    if not disclosures:
        return pd.DataFrame(), "", 0

    # 중복 제거 (재공시 등)
    seen, unique = set(), []
    for d in disclosures:
        key = (d["corp_code"], d["reprt_code"])
        if key not in seen:
            seen.add(key)
            unique.append(d)

    total_count = len(unique)
    log.info("중복 제거 후 조회 대상: %d개", total_count)

    records = []
    for i, disc in enumerate(unique):
        corp_name  = disc["corp_name"]
        stock_code = disc["stock_code"]
        reprt_name = disc["reprt_name"]
        is_annual  = disc["is_annual"]

        log.info("[%d/%d] %s (%s) - %s", i + 1, total_count, corp_name, stock_code, reprt_name)

        data, fs_div = fetch_financials(disc["corp_code"], disc["bsns_year"], disc["reprt_code"])
        fin = parse_financials(data, is_annual)

        if not fin:
            log.info("  → 재무 데이터 없음")
            continue

        found = {label: f"{label}_당기" in fin for label in ACCOUNT_MAP}
        log.info("  → 추출: %s | 재무제표: %s",
                 " / ".join(f"{k}={'✅' if v else '❌'}" for k, v in found.items()), fs_div)

        if is_all_increasing(fin):
            records.append({
                "종목코드":          stock_code,
                "기업명":            corp_name,
                "보고서유형":        reprt_name,
                "기준연도":          disc["bsns_year"],
                "재무제표구분":      "연결" if fs_div == "CFS" else "별도",
                "비교기준":          "전년동기(누적)" if not is_annual else "전기(연간)",
                "매출액_당기":       fin.get("매출액_당기"),
                "매출액_전년동기":   fin.get("매출액_전년동기"),
                "매출액_증감률":     growth(fin.get("매출액_당기"), fin.get("매출액_전년동기")),
                "영업이익_당기":     fin.get("영업이익_당기"),
                "영업이익_전년동기": fin.get("영업이익_전년동기"),
                "영업이익_증감률":   growth(fin.get("영업이익_당기"), fin.get("영업이익_전년동기")),
                "당기순이익_당기":   fin.get("당기순이익_당기"),
                "당기순이익_전년동기": fin.get("당기순이익_전년동기"),
                "당기순이익_증감률": growth(fin.get("당기순이익_당기"), fin.get("당기순이익_전년동기")),
            })
            log.info("  → ✅ 조건 충족!")
        else:
            log.info("  → 조건 미충족")

        time.sleep(0.3)

    df = pd.DataFrame(records)
    log.info("최종 조건 충족: %d개 / %d개", len(df), total_count)
    return df, base_date, total_count


# ── 엑셀 저장 ──────────────────────────────────────────────────────────────────

def save_excel(df: pd.DataFrame, filepath: str, base_date: str):
    wb  = Workbook()
    ws  = wb.active
    ws.title = "증가 기업 목록"

    H  = PatternFill("solid", start_color="1F4E79")
    S  = PatternFill("solid", start_color="2E75B6")
    A  = PatternFill("solid", start_color="EBF3FB")
    W  = PatternFill("solid", start_color="FFFFFF")
    bs = Side(style="thin", color="BDD7EE")
    BD = Border(left=bs, right=bs, top=bs, bottom=bs)
    CC = Alignment(horizontal="center", vertical="center", wrap_text=True)
    RT = Alignment(horizontal="right",  vertical="center")
    LT = Alignment(horizontal="left",   vertical="center")

    date_fmt = (f"{base_date[:4]}년 {base_date[4:6]}월 {base_date[6:]}일"
                if len(base_date) == 8 else base_date)

    ws.merge_cells("A1:O1")
    c = ws["A1"]
    c.value = f"DART 재무성장 기업 모니터링  |  공시일: {date_fmt}  (전년 동기 대비 증가)"
    c.font  = Font(bold=True, color="FFFFFF", name="Arial", size=13)
    c.fill  = H; c.alignment = CC
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:O2")
    c = ws["A2"]
    c.value = ("매출액 · 영업이익 · 당기순이익 전년 동기(누적) 대비 모두 증가한 기업  "
               "| 연결재무제표 우선 적용  (단위: 백만원)")
    c.font  = Font(color="FFFFFF", name="Arial", size=10)
    c.fill  = S; c.alignment = CC
    ws.row_dimensions[2].height = 20

    headers = [
        "종목코드", "기업명", "보고서유형", "기준연도", "재무제표", "비교기준",
        "매출액(당기)", "매출액(전년동기)", "매출액증감률(%)",
        "영업이익(당기)", "영업이익(전년동기)", "영업이익증감률(%)",
        "당기순이익(당기)", "당기순이익(전년동기)", "당기순이익증감률(%)",
    ]
    widths = [10, 20, 16, 10, 10, 14, 16, 16, 13, 16, 16, 13, 16, 16, 13]
    money_c = {7, 8, 10, 11, 13, 14}
    pct_c   = {9, 12, 15}

    for ci, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(row=3, column=ci, value=h)
        c.font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        c.fill = H; c.alignment = CC; c.border = BD
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[3].height = 36

    df_s = df.sort_values("매출액_증감률", ascending=False).reset_index(drop=True)
    for ri, row in df_s.iterrows():
        er   = ri + 4
        fill = A if ri % 2 == 0 else W
        vals = [
            row["종목코드"], row["기업명"], row["보고서유형"], row["기준연도"],
            row["재무제표구분"], row["비교기준"],
            row["매출액_당기"], row["매출액_전년동기"], row["매출액_증감률"],
            row["영업이익_당기"], row["영업이익_전년동기"], row["영업이익_증감률"],
            row["당기순이익_당기"], row["당기순이익_전년동기"], row["당기순이익_증감률"],
        ]
        for ci, val in enumerate(vals, 1):
            c = ws.cell(row=er, column=ci, value=val)
            c.fill = fill; c.border = BD
            c.font = Font(name="Arial", size=9)
            if ci in money_c:
                c.number_format = '#,##0'; c.alignment = RT
            elif ci in pct_c:
                c.number_format = '0.00"%"'; c.alignment = RT
                c.font = Font(name="Arial", size=9, color="375623")
            elif ci == 2:
                c.alignment = LT
            else:
                c.alignment = CC

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:{get_column_letter(len(headers))}{len(df_s)+3}"

    # 시트2: 요약
    ws2 = wb.create_sheet("보고서유형별 요약")
    ws2.merge_cells("A1:F1")
    c = ws2["A1"]
    c.value = "보고서 유형별 전년동기 대비 증가 기업 요약"
    c.font  = Font(bold=True, color="FFFFFF", name="Arial", size=12)
    c.fill  = H; c.alignment = CC
    ws2.row_dimensions[1].height = 28

    sh = ["보고서유형", "기업수", "연결/별도비율",
          "매출액증감률평균(%)", "영업이익증감률평균(%)", "당기순이익증감률평균(%)"]
    for ci, h in enumerate(sh, 1):
        c = ws2.cell(row=2, column=ci, value=h)
        c.font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        c.fill = S; c.alignment = CC; c.border = BD
        ws2.column_dimensions[get_column_letter(ci)].width = 22

    for ri, (rtype, grp) in enumerate(df.groupby("보고서유형"), 3):
        cfs_cnt = (grp["재무제표구분"] == "연결").sum()
        ratio   = f"연결 {cfs_cnt}/{len(grp)}"
        ws2.cell(row=ri, column=1, value=rtype).alignment = CC
        ws2.cell(row=ri, column=2, value=len(grp)).alignment = CC
        ws2.cell(row=ri, column=3, value=ratio).alignment = CC
        for ci, col in enumerate(["매출액_증감률", "영업이익_증감률", "당기순이익_증감률"], 4):
            c = ws2.cell(row=ri, column=ci, value=round(grp[col].dropna().mean(), 2))
            c.number_format = '0.00"%"'; c.alignment = RT
        for ci in range(1, 7):
            ws2.cell(row=ri, column=ci).border = BD

    wb.save(filepath)
    log.info("엑셀 저장 완료: %s", filepath)


# ── 텔레그램 ───────────────────────────────────────────────────────────────────

def send_message(text: str):
    res = requests.post(
        f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage",
        json={"chat_id": TELEGRAM_CHAT_ID, "text": text, "parse_mode": "HTML"},
        timeout=30
    )
    res.raise_for_status()
    log.info("텔레그램 메시지 전송 완료")


def send_file(filepath: str, caption: str):
    with open(filepath, "rb") as f:
        res = requests.post(
            f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendDocument",
            data={"chat_id": TELEGRAM_CHAT_ID, "caption": caption, "parse_mode": "HTML"},
            files={"document": (os.path.basename(filepath), f,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
            timeout=60
        )
    res.raise_for_status()
    log.info("텔레그램 파일 전송 완료")


def send_email(subject: str, text: str, attachment_path: str | None = None):
    if not EMAIL_TO:
        return
    if not SMTP_USER or not SMTP_PASSWORD or not EMAIL_FROM:
        raise ValueError("메일 전송 설정이 부족합니다: SMTP_USER, SMTP_PASSWORD, EMAIL_FROM 필요")

    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = EMAIL_FROM
    message["To"] = [address.strip() for address in EMAIL_TO.split(",") if address.strip()]
    message.set_content(text)

    if attachment_path:
        with open(attachment_path, "rb") as f:
            message.add_attachment(
                f.read(),
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename=os.path.basename(attachment_path),
            )

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as smtp:
        smtp.starttls()
        smtp.login(SMTP_USER, SMTP_PASSWORD)
        smtp.send_message(message)
    log.info("메일 전송 완료: %s", EMAIL_TO)


def build_message(df: pd.DataFrame, base_date: str, total_count: int) -> str:
    date_fmt = (f"{base_date[:4]}-{base_date[4:6]}-{base_date[6:]}"
                if len(base_date) == 8 else base_date)
    lines = [
        "📊 <b>DART 재무성장 기업 모니터링</b>",
        f"📅 공시일: {date_fmt}",
        f"🔍 공시 기업: {total_count}개 → 조건 충족: <b>{len(df)}개</b>",
        f"📌 기준: 전년 동기(누적) 대비 매출·영업이익·순이익 모두 증가",
        f"📋 재무제표: 연결 우선, 없으면 별도",
        "",
    ]
    for rtype, grp in df.groupby("보고서유형"):
        cfs = (grp["재무제표구분"] == "연결").sum()
        lines.append(f"📁 <b>{rtype}</b> ({len(grp)}개 | 연결 {cfs}개)")
        for _, r in grp.nlargest(5, "매출액_증감률").iterrows():
            fs_tag = "연결" if r["재무제표구분"] == "연결" else "별도"
            lines.append(
                f"  • {r['기업명']} ({r['종목코드']}) [{fs_tag}]  "
                f"매출 <b>+{r['매출액_증감률']:.1f}%</b> / "
                f"영업이익 <b>+{r['영업이익_증감률']:.1f}%</b>"
            )
        lines.append("")
    lines.append("📎 상세 내역은 첨부 엑셀 파일을 확인하세요.")
    return "\n".join(lines)


# ── 엔트리포인트 ───────────────────────────────────────────────────────────────

def main():
    today_str  = datetime.today().strftime("%Y%m%d")
    excel_path = f"dart_report_{today_str}.xlsx"

    try:
        df, base_date, total_count = collect_data()

        if df.empty:
            message = (
                f"📊 <b>DART 모니터링 결과</b>\n"
                f"📅 {datetime.today().strftime('%Y-%m-%d')}\n\n"
                f"공시된 보고서 중 조건을 충족하는 기업이 없습니다."
            )
            send_message(message)
            send_email("DART 모니터링 결과", message)
            return

        save_excel(df, excel_path, base_date)
        message = build_message(df, base_date, total_count)
        send_message(message)
        send_file(excel_path, caption=f"DART 재무성장 기업 목록 ({base_date})")
        send_email(
            f"DART 재무성장 기업 목록 ({base_date})",
            message,
            attachment_path=excel_path,
        )

    except Exception as e:
        log.error("오류 발생: %s", e, exc_info=True)
        try:
            send_message(f"⚠️ DART 모니터링 오류\n<code>{e}</code>")
        except Exception:
            pass
        sys.exit(1)


if __name__ == "__main__":
    main()
