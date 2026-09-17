import requests
base = 'http://211.104.10.171:7071'
headers = {'X-API-Key': 'Hw-_k-QPRgzolGqkLFIGYLzwqDnep53-wprci845GWw', 'Content-Type': 'application/json'}
prefixes = ['', '/accounting', '/erp', '/api', '/gateway', '/ntstl', '/service', '/account']
endpoint = '/api/v1/ntstl/list'
for p in prefixes:
    url = base + p + endpoint
    try:
        r = requests.post(url, headers=headers, json={'company_code':'','target_date':'20260916'}, timeout=3)
        print(f'{url:60s} -> STATUS: {r.status_code}')
    except Exception:
        print(f'{url:60s} -> FAIL')