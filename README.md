# DART 재무성장 기업 모니터링

매일 자동으로 DART(전자공시시스템)에서 **매출액·영업이익·당기순이익이 모두 전기 대비 증가한 기업**을 추출해 엑셀 파일로 저장하고 텔레그램으로 발송합니다.

---

## 📁 파일 구조

```
dart-financial-monitor/
├── .github/
│   └── workflows/
│       └── daily_report.yml   # GitHub Actions 스케줄러
├── dart_monitor.py            # 메인 파이썬 스크립트
├── requirements.txt           # 의존성 패키지
└── README.md
```

---

## ⚙️ 초기 설정 (딱 한 번만)

### 1. DART API 키 발급

1. [DART Open API](https://opendart.fss.or.kr) 접속
2. 회원가입 → 로그인 → **API 신청** → 인증키 발급
3. 발급된 API Key를 복사해 두기

---

### 2. 텔레그램 봇 생성

1. 텔레그램에서 **@BotFather** 검색 → 채팅 시작
2. `/newbot` 입력 → 봇 이름과 username 설정
3. 발급된 **Bot Token** (예: `123456:ABC-DEF...`) 복사
4. 본인 채팅 ID 확인: **@userinfobot** 채팅 → `/start` → `Your chat ID: 123456789`
5. 만든 봇에게 먼저 `/start` 메시지를 보내야 봇이 메시지를 전송할 수 있습니다

---

### 3. GitHub 저장소 생성 & 코드 업로드

```bash
git init dart-financial-monitor
cd dart-financial-monitor

# 파일 복사 후
git add .
git commit -m "초기 커밋"
git remote add origin https://github.com/YOUR_USERNAME/dart-financial-monitor.git
git push -u origin main
```

---

### 4. GitHub Secrets 등록

저장소 → **Settings** → **Secrets and variables** → **Actions** → **New repository secret**

| Secret 이름          | 값                          |
|---------------------|-----------------------------|
| `DART_API_KEY`      | DART API 인증키              |
| `TELEGRAM_BOT_TOKEN`| 텔레그램 봇 토큰              |
| `TELEGRAM_CHAT_ID`  | 본인 텔레그램 채팅 ID         |

---

## 🚀 실행 방법

### 자동 실행 (매일 오전 8시 KST)
설정 후 아무것도 하지 않아도 됩니다. GitHub Actions가 자동으로 실행합니다.

### 수동 실행
GitHub 저장소 → **Actions** 탭 → **DART 재무 모니터링** → **Run workflow** 클릭

### 로컬 테스트
```bash
pip install -r requirements.txt

export DART_API_KEY="your_key"
export TELEGRAM_BOT_TOKEN="your_token"
export TELEGRAM_CHAT_ID="your_chat_id"

python dart_monitor.py
```

---

## 📊 결과물

### 텔레그램 메시지
```
📊 DART 재무성장 기업 모니터링
📅 기준일: 2025-05-01

✅ 매출액·영업이익·당기순이익 모두 증가한 기업: 87개

📁 사업보고서 (42개)
  • 삼성전자 (005930) 매출 +12.3% / 영업이익 +45.2%
  • ...

📎 상세 내역은 첨부 엑셀 파일을 확인하세요.
```

### 엑셀 파일 (`dart_report_YYYYMMDD.xlsx`)
- **시트 1 - 증가 기업 목록**: 매출액·영업이익·당기순이익 당기/전기 금액 및 증감률
- **시트 2 - 보고서유형별 요약**: 유형별 기업 수 및 평균 증감률

엑셀 파일은 GitHub Actions **Artifacts**에도 90일간 보관됩니다.

---

## ⚠️ 주의사항

- DART API는 **1일 10,000건** 요청 제한이 있습니다. 상장사 약 2,500개 × 보고서 유형 수를 감안하면 충분합니다.
- 조회 대상: **KOSPI + KOSDAQ 상장법인** (비상장 제외)
- 재무제표는 **연결재무제표** 우선, 없으면 별도재무제표 사용
- GitHub Actions 무료 플랜: 월 2,000분 무료 (1회 실행 약 30~60분 예상)

---

## 🔧 커스터마이징

### 실행 시간 변경 (`daily_report.yml`)
```yaml
- cron: "0 23 * * *"   # UTC 23:00 = KST 08:00
# → UTC 22:00 = KST 07:00 으로 바꾸려면:
- cron: "0 22 * * *"
```

### 특정 시장만 조회 (`dart_monitor.py`)
`_download_corp_codes()` 함수에서 KOSPI/KOSDAQ 필터를 추가할 수 있습니다.
