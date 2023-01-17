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

#### this section loads the vPC-Leaf-Details.xlsx into the program

wb = load_workbook(
    '/Users/krishna/python-projects/My_venvs/venv1-enterprise/dcnm-python-Testing/Fabric-Switch-Onboarding/vPC-Leaf-Onboarding/vPC-Leaf-Details.xlsx')
vPC_leaf_onboarding = wb["vPC_Leaf_Onboarding"]
url_poap = f"https://{dcnm_ip}/rest/control/fabrics/{fabric_name}/inventory/poap"
url_vpc_pairing = f"https://{dcnm_ip}/rest/vpcpair"
# url_role = f"https://{dcnm_ip}/rest/control/switches/roles" ####default role is Leaf

for i in range(2, len(vPC_leaf_onboarding['A']) + 1):
    leaf_details = [{
        "serialNumber": vPC_leaf_onboarding[f'A{i}'].value,
        "model": vPC_leaf_onboarding[f'B{i}'].value,
        "version": vPC_leaf_onboarding[f'D{i}'].value,
        "hostname": vPC_leaf_onboarding[f'E{i}'].value,
        "ipAddress": vPC_leaf_onboarding[f'F{i}'].value,
        "password": vPC_leaf_onboarding[f'G{i}'].value,
        "discoveryUsername": "admin",
        "discoveryPassword": vPC_leaf_onboarding[f'G{i}'].value,
        "discoveryAuthProtocol": 0,
        "data": "{\"gateway\": \"192.168.201.1/24\", \"modulesModel\": [\"N9K-X9364v\", \"N9K-vSUP\"]}"}]
    requests.post(url_poap, headers=headers_token, verify=False, data=json.dumps(leaf_details))

requests.post(url_logout, headers=headers_token, verify=False)
