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
dcnm_token = json.loads(response.text)['Dcnm-Token']   ##the token returned from dcnm is stored in this variable
headers_token = {'Content-Type': 'application/json', 'dcnm-token':str(dcnm_token)}

#### this section loads the Interface-Details.xlsx into the program

wb = load_workbook('/Users/krishna/python-projects/My_venvs/venv1-enterprise/DCNM-Python-Testing/Create Interfaces/Interface-Details.xlsx')
PC_add = wb["non-vPC-Port-Channel-Details"]
url_PC_add =  f"https://{dcnm_ip}/rest/interface"
url_PC_deploy = f"https://{dcnm_ip}/rest/interface/deploy"

###### Create PC-Interface #####
for i in range (2, len(PC_add['A'])+1):
    PC_details = {
                    "policy": PC_add[f'A{i}'].value,
                    "interfaceType": "INTERFACE_PORT_CHANNEL",
                    "interfaces": [
                        {
                        "serialNumber": PC_add[f'B{i}'].value,
                        "interfaceType": "INTERFACE_PORT_CHANNEL",
                        "ifName": "Port-channel"+PC_add[f'C{i}'].value,
                        "fabricName": "AZ-Phoenix",
                        "nvPairs": {
                            "MEMBER_INTERFACES": PC_add[f'D{i}'].value,
                            "PC_MODE": PC_add[f'I{i}'].value,
                            "BPDUGUARD_ENABLED": PC_add[f'G{i}'].value,
                            "PORTTYPE_FAST_ENABLED": PC_add[f'H{i}'].value,
                            "MTU": "jumbo",
                            "SPEED": "Auto",
                            "ALLOWED_VLANS": "none",
                            "DESC": PC_add[f'E{i}'].value,
                            "CONF": "",
                            "ADMIN_STATE": 'true',
                            "PO_ID": "Port-channel"+PC_add[f'C{i}'].value
                        }
                        }
                    ],
                    "skipResourceCheck": 'false'
                    }
    deploy_details = [{"serialNumber":PC_add[f'B{i}'].value,"ifName":"Port-channel"+PC_add[f'C{i}'].value,"fabricName":"AZ-Phoenix"}]
    response = requests.post(url_PC_add, headers=headers_token, verify=False, data=json.dumps(PC_details))
    response = requests.post(url_PC_deploy, headers=headers_token, verify=False, data=json.dumps(deploy_details)) 
    #print(json.dumps(network_details))
    print(response.text)

requests.post(url_logout, headers=headers_token, verify=False)