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

#### this section loads the VLAN-Details.xlsx into the program

wb = load_workbook('/Users/krishna/python-projects/My_venvs/venv1-enterprise/DCNM-Python-Testing/Create-VLANs/VLAN-Details.xlsx')
network_attach = wb["VLAN-Association"]
url_network_attach =  f"https://{dcnm_ip}/rest/top-down/fabrics/{fabric_name}/networks/attachments"
url_network_deploy =  f"https://{dcnm_ip}/rest/top-down/fabrics/{fabric_name}/networks/deployments"
#url_role = f"https://{dcnm_ip}/rest/control/switches/roles" ####default role is Leaf

###### Attach Networks/VLANs #####
for i in range (2, len(network_attach['A'])+1):
    network_attach_details = [{
                            "networkName": network_attach[f'A{i}'].value,
                            "lanAttachList": [
                            {
                                "fabric": fabric_name,
                                "networkName": network_attach[f'A{i}'].value,
                                "serialNumber": network_attach[f'C{i}'].value,
                                "switchPorts": network_attach[f'E{i}'].value,
                                "detachSwitchPorts": "",
                                "vlan": network_attach[f'F{i}'].value,
                                "untagged": 'false',
                                "freeformConfig": "",
                                "deployment": 'true',
                                "extensionValues": "",
                                "instanceValues": ""
                            },
                            {
                                "fabric": fabric_name,
                                "networkName": network_attach[f'A{i}'].value,
                                "serialNumber": network_attach[f'D{i}'].value,
                                "switchPorts": network_attach[f'E{i}'].value,
                                "detachSwitchPorts": "",
                                "vlan": network_attach[f'F{i}'].value,
                                "untagged": 'false',
                                "freeformConfig": "",
                                "deployment": 'true',
                                "extensionValues": "",
                                "instanceValues": ""
                            }
                            ]
                        }]
    network_deploy = {"networkNames":network_attach[f'A{i}'].value}
    response = requests.post(url_network_attach, headers=headers_token, verify=False, data=json.dumps(network_attach_details))
    response_deploy = requests.post(url_network_deploy, headers=headers_token, verify=False, data=json.dumps(network_deploy))
    print(response_deploy.text)
 
requests.post(url_logout, headers=headers_token, verify=False)