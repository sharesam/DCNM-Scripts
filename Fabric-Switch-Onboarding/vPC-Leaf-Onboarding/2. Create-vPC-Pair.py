import json
import requests
from requests.auth import HTTPBasicAuth
from openpyxl import Workbook, load_workbook

##### this section gathers the token from DCNM which will be used for the subsequent Get or Post requests####
dcnm_ip = "10.122.104.50"
fabric_name = "AZ-Phoenix"
url_login = f"https://{dcnm_ip}/rest/logon"
url_logout = f"https://{dcnm_ip}/rest/logout"
requests.packages.urllib3.disable_warnings()
dcnm_creds = HTTPBasicAuth('admin', 'C!sc0123')
headers = {'Content-Type': 'application/json'}
request_body = {"expirationTime": "999999"}

response = requests.post(url_login, headers=headers, auth=dcnm_creds, verify=False, data=json.dumps(request_body))
dcnm_token = json.loads(response.text)['Dcnm-Token']  ##the token returned from dcnm is stored in this variable
headers_token = {'Content-Type': 'application/json', 'dcnm-token': str(dcnm_token)}

#### this section loads the Spine-Details.xlsx into the program

wb = load_workbook(
    '/Users/krishna/python-projects/My_venvs/venv1-enterprise/dcnm-python-Testing/Fabric-Switch-Onboarding/vPC-Leaf-Onboarding/vPC-Leaf-Details.xlsx')
vPC_leaf_onboarding = wb["vPC_Leaf_Onboarding"]
url_vpc_pairing = f"https://{dcnm_ip}/rest/vpcpair"
# url_role = f"https://{dcnm_ip}/rest/control/switches/roles" ####default role is Leaf

####vPC - Pairing

vPC_Peering = {"peerOneId": vPC_leaf_onboarding['A2'].value,
               "peerTwoId": vPC_leaf_onboarding['A3'].value}  # "useVirtualPeerlink":'true'}

requests.post(url_vpc_pairing, headers=headers_token, verify=False,
              data=json.dumps(vPC_Peering))  ####send a request to vPC the leaf pairs
requests.post(url_logout, headers=headers_token, verify=False)
