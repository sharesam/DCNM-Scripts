import json
import requests
from requests.auth import HTTPBasicAuth

dcnm_ip = "10.122.104.50"
url = f"https://{dcnm_ip}/rest/logon"
requests.packages.urllib3.disable_warnings()
dcnm_creds = HTTPBasicAuth('admin', 'C!sc0123')
headers = {'Content-Type': 'application/json'}
request_body = {"expirationTime": "999999"}

response = requests.post(url, headers=headers, auth=dcnm_creds, verify=False, data=json.dumps(request_body))

print(json.loads(response.text)['Dcnm-Token'])