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

#### this section loads the Interface-Details.xlsx into the program

wb = load_workbook(
    '/Users/krishna/python-projects/My_venvs/venv1-enterprise/dcnm-python-Testing/Create Interfaces/Interface-Details.xlsx')
access_interface_add = wb["Acess-Interface-Details"]
url_access_int_add = f"https://{dcnm_ip}/rest/interface"
url_access_int_deploy = f"https://{dcnm_ip}/rest/interface/deploy"

###### Create vPC-Interface #####
for i in range(2, len(vPC_add['A']) + 1):
    vPC_details = {
        "policy": vPC_add[f'A{i}'].value,
        "interfaceType": "INTERFACE_VPC",
        "interfaces": [
            {
                "serialNumber": vPC_add[f'B{i}'].value + "~" + vPC_add[f'C{i}'].value,
                "interfaceType": "INTERFACE_VPC",
                "ifName": "vPC" + vPC_add[f'D{i}'].value,
                "fabricName": "AZ-Phoenix",
                "nvPairs": {
                    "PEER1_PCID": vPC_add[f'D{i}'].value,
                    "PEER2_PCID": vPC_add[f'D{i}'].value,
                    "ENABLE_MIRROR_CONFIG": 'false',
                    "PEER1_MEMBER_INTERFACES": vPC_add[f'E{i}'].value,
                    "PEER2_MEMBER_INTERFACES": vPC_add[f'F{i}'].value,
                    "PC_MODE": vPC_add[f'L{i}'].value,
                    "BPDUGUARD_ENABLED": vPC_add[f'J{i}'].value,
                    "PORTTYPE_FAST_ENABLED": vPC_add[f'K{i}'].value,
                    "MTU": "jumbo",
                    "SPEED": "Auto",
                    "PEER1_ALLOWED_VLANS": "none",
                    "PEER2_ALLOWED_VLANS": "none",
                    "PEER1_PO_DESC": vPC_add[f'G{i}'].value,
                    "PEER2_PO_DESC": vPC_add[f'H{i}'].value,
                    "PEER1_PO_CONF": "",
                    "PEER2_PO_CONF": "",
                    "ADMIN_STATE": 'true',
                    "INTF_NAME": "vPC" + vPC_add[f'D{i}'].value
                }
            }
        ],
        "skipResourceCheck": 'false'
    }
    deploy_details = [{"serialNumber": vPC_add[f'B{i}'].value + "~" + vPC_add[f'C{i}'].value,
                       "ifName": "vPC" + vPC_add[f'D{i}'].value, "fabricName": "AZ-Phoenix"}]
    response = requests.post(url_vPC_add, headers=headers_token, verify=False, data=json.dumps(vPC_details))
    response = requests.post(url_vPC_deploy, headers=headers_token, verify=False, data=json.dumps(deploy_details))
    # print(json.dumps(network_details))
    print(response.text)

requests.post(url_logout, headers=headers_token, verify=False)
