@'
import requests

base = "http://211.104.10.171:7071"
headers = {"X-API-Key": "Hw-_k-QPRgzolGqkLFIGYLzwqDnep53-wprci845GWw"}
tests = [
    ("GET", "/"),
    ("POST", "/api/v1/company/search"),
    ("GET", "/swagger/index.html"),
    ("GET", "/swagger"),
    ("GET", "/api/swagger"),
    ("GET", "/accounting/api/v1/company/search")
]

for method, path in tests:
    try:
        r = requests.request(method, base + path, headers=headers, timeout=3)
        print(f"{method:4s} {path:35s} -> 상태코드: {r.status_code}")
    except Exception as e:
        print(f"{method:4s} {path:35s} -> 연결 실패: {e}")
'@ | Out-File -Encoding utf8 test_api.py; python test_api.py