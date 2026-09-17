import requests
base = 'http://211.104.10.171:7071'
headers = {'X-API-Key': 'Hw-_k-QPRgzolGqkLFIGYLzwqDnep53-wprci845GWw'}
tests = [('GET', '/'), ('POST', '/api/v1/company/search'), ('GET', '/swagger/index.html'), ('GET', '/swagger'), ('GET', '/api/swagger'), ('GET', '/accounting/api/v1/company/search')]
for m, p in tests:
    try:
        r = requests.request(m, base + p, headers=headers, timeout=3)
        print(f'{m:4s} {p:35s} -> STATUS: {r.status_code}')
    except Exception as e:
        print(f'{m:4s} {p:35s} -> FAIL')