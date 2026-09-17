import requests
url = 'http://211.104.10.171:7071/api/v1/ntstl/list'
headers = {'X-API-Key': 'Hw-_k-QPRgzolGqkLFIGYLzwqDnep53-wprci845GWw', 'Content-Type': 'application/json'}
params = {'page': 1, 'pageSize': 1000}
payload = {'company_code': '', 'target_date': '20260916', 'type_account_code': [], 'type_customer_code': []}
r = requests.post(url, headers=headers, params=params, json=payload, timeout=5)
print('Status:', r.status_code)
print('Response:', r.text[:300])