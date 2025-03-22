import requests

url = 'https://api.upstox.com/v2/login/authorization/token'
headers = {
    'accept': 'application/json',
    'Content-Type': 'application/x-www-form-urlencoded',
}

data = {
    'code': 'FNmDbC',
    'client_id': '2e32809b-adaf-40d2-b957-1e22483fbf43',
    'client_secret': 'rlnvf3s1jb',
    'redirect_uri': 'https://google.com/',
    'grant_type': 'authorization_code',
}

response = requests.post(url, headers=headers, data=data)

print(response.status_code)
print(response.json())
