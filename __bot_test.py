import requests

token = '[REMOVED]'
chat_id = '[REMOVED]'

url = f'https://api.telegram.org/bot{token}/sendMessage'
payload = {
    'chat_id': chat_id,
    'text': '✅ <b>DART 모니터링 테스트</b>\n\n텔레그램 연동 성공! 🎉',
    'parse_mode': 'HTML'
}
res = requests.post(url, json=payload, timeout=10)
print(res.status_code, res.json())