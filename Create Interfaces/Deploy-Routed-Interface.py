import json
import requests
from openpyxl import load_workbook
from DCNM_Authentication import headers_token, fabric_name, url_logout, dcnm_ip

wb = load_workbook('Interface-Details.xlsx')
routed_int_attach = wb["Routed-Interface-Details"]
url_routed_interface_attach = f"https://{dcnm_ip}/rest/interface"

# Create Routed-Interface #####

for i in range(2, len(routed_int_attach['A']) + 1):
    routed_interface_details = {
        "policy": "int_routed_host_11_1",
        "interfaceType": "INTERFACE_ETHERNET",
        "interfaces": [
            {
                "serialNumber": routed_int_attach[f'B{i}'].value,
                "interfaceType": "INTERFACE_ETHERNET",
                "ifName": "Ethernet" + routed_int_attach[f'C{i}'].value,
                "fabricName": fabric_name,
                "nvPairs": {
                    "INTF_VRF": routed_int_attach[f'D{i}'].value,
                    "IP": routed_int_attach[f'E{i}'].value,
                    "PREFIX": str(routed_int_attach[f'F{i}'].value),
                    "ROUTING_TAG": "12345",
                    "MTU": str(routed_int_attach[f'G{i}'].value),
                    "SPEED": str(routed_int_attach[f'H{i}'].value),
                    "DESC": routed_int_attach[f'I{i}'].value,
                    "CONF": "",
                    "ADMIN_STATE": "true",
                    "INTF_NAME": "Ethernet" + routed_int_attach[f'C{i}'].value
                }
            }
        ]
    }
    print(json.dumps(routed_interface_details))
    response = requests.post(url_routed_interface_attach, headers=headers_token, verify=False,
                             data=json.dumps(routed_interface_details))
    print(response.text)

requests.post(url_logout, headers=headers_token, verify=False)
