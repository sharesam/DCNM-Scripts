import json
import requests
from requests.auth import HTTPBasicAuth

dcnm_ip = "10.59.10.80"
fabric_name = "SoAZ-PHX"
url_login = f"https://{dcnm_ip}/rest/logon"
url_logout = f"https://{dcnm_ip}/rest/logout"
requests.packages.urllib3.disable_warnings()
dcnm_creds = HTTPBasicAuth('kharesamudram', 'Kp2daySbP!')
headers = {'Content-Type': 'application/json'}
request_body = {"expirationTime": "999999"}

response = requests.post(url_login, headers=headers, auth=dcnm_creds, verify=False, data=json.dumps(request_body))
dcnm_token = json.loads(response.text)['Dcnm-Token']  # the token returned from dcnm is stored in this variable
headers_token = {'Content-Type': 'application/json', 'dcnm-token': str(dcnm_token)}

if __name__ == '__main__':
    print(headers_token)  # prints the token
