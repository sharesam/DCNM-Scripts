import json
import requests
from openpyxl import Workbook, load_workbook
from DCNM_Authentication import headers_token, dcnm_ip, url_logout, fabric_name

# this section loads the Interface-Details.xlsx into the program

wb = load_workbook('Interface-Details-Retirement.xlsx')
trunk_interface_add = wb["Trunk-Interface-Details"]
url_trunk_add = f"https://{dcnm_ip}/rest/interface"
url_trunk_deploy = f"https://{dcnm_ip}/rest/interface/deploy"

# Create trunk-Interface #####
for i in range(2, len(trunk_interface_add['A']) + 1):
    trunk_details = {
        "policy": "int_trunk_host_11_1",
        "interfaceType": "INTERFACE_ETHERNET",
        "interfaces": [
            {
                "serialNumber": trunk_interface_add[f'B{i}'].value,
                "ifName": "Ethernet" + trunk_interface_add[f'C{i}'].value,
                "fabricName": fabric_name,
                "nvPairs": {
                    "BPDUGUARD_ENABLED": trunk_interface_add[f'F{i}'].value,
                    "PORTTYPE_FAST_ENABLED": trunk_interface_add[f'G{i}'].value,
                    "MTU": trunk_interface_add[f'H{i}'].value,
                    "SPEED": trunk_interface_add[f'E{i}'].value,
                    "ALLOWED_VLANS": trunk_interface_add[f'I{i}'].value,
                    "DESC": trunk_interface_add[f'D{i}'].value,
                    "CONF": "",
                    "ADMIN_STATE": "true",
                    "INTF_NAME": "Ethernet" + trunk_interface_add[f'C{i}'].value
                }
            }
        ]
    }
    # deploy_details = [{"serialNumber": trunk_interface_add[f'B{i}'].value + "~" + trunk_interface_add[f'C{i}'].value,
    #                    "ifName": "vPC" + trunk_interface_add[f'D{i}'].value, "fabricName": "AZ-Phoenix"}]
    print(trunk_details)
    response = requests.post(url_trunk_add, headers=headers_token, verify=False, data=json.dumps(trunk_details))
    # response = requests.post(url_vPC_deploy, headers=headers_token, verify=False, data=json.dumps(deploy_details))
    # print(json.dumps(network_details))
    print(response.text)

requests.post(url_logout, headers=headers_token, verify=False)
