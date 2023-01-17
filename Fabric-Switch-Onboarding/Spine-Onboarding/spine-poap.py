import json
import requests
from requests.auth import HTTPBasicAuth
from openpyxl import Workbook, load_workbook

##### this section gathers the token from DCNM which will be used for the subsequent Get or Post requests####
dcnm_ip = "10.122.104.50"
fabric_name = "AZ-Phoenix"
url = f"https://{dcnm_ip}/rest/logon"
requests.packages.urllib3.disable_warnings()
dcnm_creds = HTTPBasicAuth('admin', 'C!sc0123')
headers = {'Content-Type': 'application/json'}
request_body = {"expirationTime": "999999"}

response = requests.post(url, headers=headers, auth=dcnm_creds, verify=False, data=json.dumps(request_body))
dcnm_token = json.loads(response.text)['Dcnm-Token']   ##the token returned from dcnm is stored in this variable
headers_token = {'Content-Type': 'application/json', 'dcnm-token':str(dcnm_token)}

#### this section loads the Spine-Details.xlsx into the program

wb = load_workbook('/Users/krishna/python-projects/My_venvs/venv1-enterprise/DCNM-Python-Testing/Fabric-Switch-Onboarding/Spine-Onboarding/Spine-Details.xlsx')
spine_onboarding = wb["Spine_Onboarding"]
url_poap = f"https://{dcnm_ip}/rest/control/fabrics/{fabric_name}/inventory/poap"
url_role = f"https://{dcnm_ip}/rest/control/switches/roles"

for i in range (2, len(spine_onboarding['A'])+1):
    spine_details = [{
                        "serialNumber": spine_onboarding[f'A{i}'].value,
                        "model": spine_onboarding[f'B{i}'].value,
                        "version": spine_onboarding[f'D{i}'].value,
                        "hostname": spine_onboarding[f'E{i}'].value,
                        "ipAddress": spine_onboarding[f'F{i}'].value,
                        "password": spine_onboarding[f'G{i}'].value,
                        "discoveryUsername": "admin",
                        "discoveryPassword": spine_onboarding[f'G{i}'].value,
                        "discoveryAuthProtocol": 0,
                        "data": "{\"gateway\": \"192.168.201.1/24\", \"modulesModel\": [\"N9K-X9364v\", \"N9K-vSUP\"]}"}]
    requests.post(url_poap, headers=headers_token, verify=False, data=json.dumps(spine_details))  #### Poaps the Spine Switch
    spine_role= [{"serialNumber": spine_onboarding[f'A{i}'].value,"role": "Spine"}]
    requests.post(url_role, headers=headers_token, verify=False, data=json.dumps(spine_role)) #### Assigns Spine Role to the switch