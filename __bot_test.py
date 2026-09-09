"""Send a Telegram test using environment variables."""
import os
import requests


def main():
    token = os.environ["TELEGRAM_BOT_TOKEN"]
    chat_id = os.environ["TELEGRAM_CHAT_ID"]
    try:
        response = requests.post(
            f"https://api.telegram.org/bot{token}/sendMessage",
            json={"chat_id": chat_id, "text": "DART 모니터링 테스트: 텔레그램 연동 성공!"},
            timeout=10,
        )
        response.raise_for_status()
        if not response.json().get("ok"):
            raise RuntimeError("Telegram API returned an unsuccessful response")
    except (requests.RequestException, RuntimeError):
        # Exception URLs and response bodies can include tokens or personal data.
        raise SystemExit("텔레그램 테스트 실패: 환경변수와 연결 상태를 확인하세요.") from None
    print("텔레그램 테스트 전송 완료")


if __name__ == "__main__":
    main()
