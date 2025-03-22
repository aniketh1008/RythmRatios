
api_key = "2e32809b-adaf-40d2-b957-1e22483fbf43"
redirect_uri = "https://google.com/"

login_url = f"https://api.upstox.com/v2/login/authorization/dialog?client_id={api_key}&redirect_uri={redirect_uri}&response_type=code"
print("Login URL:", login_url)


#nehaapikey:"fe20e987-0367-4bf8-857a-27d582598a2b"
#nehapisecret:"oxrso6ozq2"

# import requests
#
# url = "https://api.upstox.com/v2/login/authorization/dialog"
#
# payload={}
# headers = {
#     'client_id':'f8f73882-2788-446c-a78a-491be6f7cca5',
#     'redirect_uri':'https://google.com/'
#     'response_type':'code'
# }
#
# response = requests.request("GET", url, headers=headers, data=payload)
#
# print(response.text)