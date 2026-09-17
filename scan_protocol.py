import requests
from requests.packages.urllib3.exceptions import InsecureRequestWarning
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

headers = {'X-API-Key': 'Hw-_k-QPRgzolGqkLFIGYLzwqDnep53-wprci845GWw', 'Content-Type': 'application/json'}
payload = {'company_code':'', 'target_date':'20260916'}

targets = [
    'https://211.104.10.171:7071/api/v1/ntstl/list',
    'http://211.104.10.171/api/v1/ntstl/list',
    'https://211.104.10.171/api/v1/ntstl/list',
    'http://211.104.10.171:8080/api/v1/ntstl/list',
    'http://211.104.10.171:7071/api/v1/ntstl/list'
]

for url in targets:
    try:
        r = requests.post(url, headers=headers, json=payload, timeout=3, verify=False)
        print(f'{url:50s} -> STATUS: {r.status_code}')
    except Exception as e:
        print(f'{url:50s} -> FAIL ({type(e).__name__})')