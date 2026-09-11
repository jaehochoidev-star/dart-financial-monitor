# DART 재무성장 기업 모니터링

매일 자동으로 DART(전자공시시스템)에서 **매출액·영업이익·당기순이익이 모두 전기 대비 증가한 기업**을 추출해 엑셀 파일로 저장하고 텔레그램과 이메일로 발송합니다. 날짜별 결과는 GitHub Pages에서 최신순으로 확인할 수 있습니다.

---

## 📁 파일 구조

```
dart-financial-monitor/
├── .github/
│   └── workflows/
│       └── daily_report.yml   # GitHub Actions 스케줄러
├── dart_monitor.py            # 메인 파이썬 스크립트
├── report_archive.py          # 날짜별 저장 및 정적 웹페이지 생성
├── test_reporting.py          # 외부 전송 없이 실행하는 검증
├── reports/                   # 운영 실행 후 누적되는 결과와 엑셀
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
| `SMTP_HOST`         | SMTP 서버 (예: `smtp.gmail.com`) |
| `SMTP_PORT`         | SMTP 포트 (일반적으로 `587`)    |
| `SMTP_USER`         | SMTP 로그인 계정                 |
| `SMTP_PASSWORD`     | SMTP 비밀번호 또는 앱 비밀번호    |
| `EMAIL_FROM`        | 발신 이메일 주소                 |
| `EMAIL_TO`          | 수신 이메일 주소 (여러 개는 쉼표로 구분) |

`EMAIL_TO`를 설정하면 텔레그램 메시지와 함께 결과 메일이 전송됩니다. Gmail은 일반 계정 비밀번호 대신 앱 비밀번호를 사용해야 합니다. 메일을 사용하지 않으면 메일 관련 secret은 설정하지 않아도 됩니다.

메일은 텔레그램과 동일한 요약을 HTML/일반 텍스트로 보내고, 결과가 있으면 동일한 엑셀을 첨부합니다. `EMAIL_FROM`은 비워 두면 `SMTP_USER`를 사용합니다. `SMTP_PORT`는 STARTTLS `587`(기본값) 또는 SSL `465`를 지원합니다. 텔레그램이나 이메일 중 하나가 실패해도 다른 전송과 기록 저장을 시도하며, 실패는 Actions에 표시됩니다.

### 5. 날짜별 웹페이지 (GitHub Pages)

1. 변경된 코드를 저장소의 기본 브랜치에 업로드합니다.
2. 저장소 **Settings → Pages → Build and deployment → Source**를 **GitHub Actions**로 설정합니다.
3. **Settings → Actions → General → Workflow permissions**에서 저장소 쓰기가 허용되어 있는지 확인합니다. 조직 정책이나 브랜치 보호 규칙이 자동 기록 커밋을 차단하면 예외 설정이 필요합니다.
4. **Actions → DART 재무 모니터링 → Run workflow**로 한 번 실행합니다.
5. 배포가 끝나면 Pages 설정에 표시된 주소에서 결과를 확인합니다. 현재 저장소의 기본 주소는 `https://jaehochoidev-star.github.io/dart-financial-monitor/`입니다.

GitHub Pages의 Actions 배포 방식은 [GitHub 공식 안내](https://docs.github.com/en/pages/getting-started-with-github-pages/using-custom-workflows-with-github-pages)를 따릅니다. 웹페이지에 알림 요약과 결과 엑셀이 게시됩니다. GitHub Pages의 공개 범위를 확인해 사용하세요.

- 결과를 한국시간 **실행 날짜의 월별 페이지**로 나누고, 각 월 안에서는 최신 날짜부터 표시합니다. 상단 월 탭으로 이동하며 검색이나 JavaScript는 사용하지 않습니다.
- 기본 주소는 기록이 있는 가장 최근 월을 표시합니다. `2026-09.html`처럼 월별 주소로 직접 접속할 수 있고, 새 달의 기록이 생기면 탭이 자동으로 추가됩니다. 모바일에서는 월 탭을 가로로 스크롤할 수 있습니다.
- 각 날짜 아래에는 엑셀 첫 시트의 기업별 상세 내역 15개 열을 표로 표시합니다. 기존에 저장된 엑셀도 반영하며, 표를 좌우로 스크롤하여 금액(원)과 증감률을 확인할 수 있습니다. 엑셀 다운로드도 계속 제공합니다.
- 공시 기준일은 실행일과 별도로 표시합니다. 기존의 최근 7일 공시 탐색 방식을 유지하므로 휴일에는 같은 공시 결과가 다시 기록될 수 있습니다.
- 조건 충족 기업이 없는 날도 기록합니다. 같은 날 재실행하면 해당 날짜의 결과를 갱신합니다.
- 원본 기록과 엑셀은 `reports/YYYY-MM-DD/`에 커밋하여 계속 보관합니다. 아티팩트 90일 보관 기간과 무관합니다.
- 텔레그램/메일 전송이 실패하더라도 생성된 기록은 저장하고 웹페이지 배포를 시도합니다. 수집 자체가 실패하면 새 정상 결과로 기록하지 않습니다.
- 변경 이전의 텔레그램 메시지는 자동으로 가져오지 않습니다. 누적 기록은 적용 이후부터 시작합니다.
- 자동/수동 실행은 모두 기본 브랜치의 코드를 사용합니다. 웹페이지 생성 결과인 `site/`는 커밋하지 않습니다.

로컬에서 `python report_archive.py`를 실행한 뒤 `site/index.html`을 열면 저장된 기록을 볼 수 있습니다. 기록이 없으면 안내 화면이 표시됩니다.

Pages 배포 파일과 엑셀 아티팩트 이름에는 실행 ID와 재실행 회차가 포함됩니다. 배포 작업은 업로드 작업이 전달한 이름을 사용하므로, 전체 재실행과 배포만 재실행하는 경우 모두 같은 이름의 아티팩트 충돌을 방지합니다. 이 수정 이전에 발생한 중복 오류는 기존 실행의 재시도 대신 **Actions → Run workflow**로 새 실행을 시작하세요. 새 실행에서는 데이터 수집과 알림도 다시 수행합니다.

---

## 🚀 실행 방법

### 자격증명 관리

- API 키, 봇 토큰, 채팅 ID, 메일 계정과 비밀번호는 GitHub Actions Secrets 또는 로컬 환경변수로만 설정합니다.
- `__bot_test.py`는 `TELEGRAM_BOT_TOKEN`과 `TELEGRAM_CHAT_ID` 환경변수를 사용합니다. 직접 실행할 때만 테스트 메시지를 보냅니다.
- `__run_darrt_monitor.py`는 설정된 환경변수로 모니터링을 실행하는 보조 진입점입니다.
- `.env` 파일은 커밋에서 제외되지만 자동으로 읽지는 않습니다. 실행 전에 환경변수를 설정해야 합니다.
- 토큰을 재발급하면 **Settings → Secrets and variables → Actions → TELEGRAM_BOT_TOKEN** 값을 교체합니다. 채팅 ID는 기존 Secret 값을 그대로 사용합니다.
- 로그에는 메일 수신 주소를 출력하지 않고, 오류 로그의 설정된 자격증명은 가립니다. 과거 실행 로그는 소스 수정으로 변경되지 않습니다.

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

export SMTP_HOST="smtp.gmail.com"
export SMTP_PORT="587"
export SMTP_USER="your_email@gmail.com"
export SMTP_PASSWORD="your_app_password"
export EMAIL_FROM="your_email@gmail.com"
export EMAIL_TO="recipient@example.com"

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
