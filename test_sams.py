import requests
urls=['http://211.104.10.171:7071/sams/api/v1/ntstl/list', 'http://211.104.10.171:7071/samsapi/api/v1/ntstl/list', 'http://211.104.10.171:7071/SAMS/api/v1/ntstl/list']
for u in urls:
    try:
        r = requests.post(u, headers={'X-API-Key': 'Hw-_k-QPRgzolGqkLFIGYLzwqDnep53-wprci845GWw', 'Content-Type': 'application/json'}, json={'company_code':'', 'target_date':'20260916'}, timeout=3)
        print(f'{u} -> STATUS: {r.status_code}')
    except Exception as e:
        print(f'{u} -> FAIL')
