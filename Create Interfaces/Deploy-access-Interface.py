import json
import requests
from openpyxl import load_workbook
from DCNM_Authentication import headers_token, dcnm_ip, url_logout

# this section loads the Interface-Details.xlsx into the program

wb = load_workbook('Interface-Details.xlsx')
access_interface_add = wb["Access-Interface-Details"]
url_access_int_add = f"https://{dcnm_ip}/rest/interface"
url_access_int_deploy = f"https://{dcnm_ip}/rest/interface/deploy"

# Create Access-Interface #####
for i in range(2, len(access_interface_add['A']) + 1):
    int_details = {
        "policy": "int_access_host_11_1",
        "interfaces": [
            {
                "serialNumber": access_interface_add[f'A{i}'].value,
                "interfaceType": "INTERFACE_ETHERNET",
                "ifName": "Ethernet" + access_interface_add[f'B{i}'].value,
                "fabricName": "SoAZ-PHX",
                "nvPairs": {
                    "BPDUGUARD_ENABLED": str.lower(access_interface_add[f'E{i}'].value),
                    "PORTTYPE_FAST_ENABLED": str.lower(access_interface_add[f'F{i}'].value),
                    "MTU": str.lower(access_interface_add[f'G{i}'].value),
                    "SPEED": access_interface_add[f'D{i}'].value,
                    "ACCESS_VLAN": "",
                    "DESC": access_interface_add[f'C{i}'].value,
                    "CONF": "",
                    "ADMIN_STATE": str.lower(access_interface_add[f'H{i}'].value),
                    "INTF_NAME": "Ethernet" + access_interface_add[f'B{i}'].value
                }
            }
        ]
    }
    print(int_details)
    # deploy_details = [{"serialNumber": access_interface_add[f'B{i}'].value + "~" + access_interface_add[f'C{
    # i}'].value,"ifName": "vPC" + access_interface_add[f'D{i}'].value, "fabricName": "AZ-Phoenix"}]
    response = requests.put(url_access_int_add, headers=headers_token, verify=False, data=json.dumps(int_details))
    # requests.post(url_access_int_deploy, headers=headers_token, verify=False,data=json.dumps(deploy_details))
    # print(json.dumps(network_details))
    print(response.text)

requests.post(url_logout, headers=headers_token, verify=False)
