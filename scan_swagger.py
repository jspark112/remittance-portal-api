import requests
base = 'http://211.104.10.171:7071'
headers = {'X-API-Key': 'Hw-_k-QPRgzolGqkLFIGYLzwqDnep53-wprci845GWw'}
swagger_paths = [
    '/v1/swagger.json',
    '/swagger/v1/swagger.json',
    '/openapi.json',
    '/api/openapi.json',
    '/api/swagger.json'
]
for p in swagger_paths:
    try:
        r = requests.get(base + p, headers=headers, timeout=3)
        print(f'{p:30s} -> STATUS: {r.status_code}')
    except Exception:
        print(f'{p:30s} -> FAIL')