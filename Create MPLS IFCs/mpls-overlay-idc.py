import json
import requests
from openpyxl import load_workbook
from DCNM_Authentication import headers_token, fabric_name, url_logout, dcnm_ip

wb = load_workbook('ifc.xlsx')
ifc_create_overlay = wb["overlay"]
url_ifc_create = f"https://{dcnm_ip}/rest/control/links"

for i in range(2, len(ifc_create_overlay['A']) + 1):
    ifc_overlay_details = {
        "sourceFabric": fabric_name,
        "destinationFabric": "SoAZ-EXT",
        "sourceDevice": ifc_create_overlay[f'B{i}'].value,
        "destinationDevice": "",
        "sourceSwitchName": ifc_create_overlay[f'A{i}'].value,
        "destinationSwitchName": ifc_create_overlay[f'C{i}'].value,
        "sourceInterface": "Loopback101",
        "destinationInterface": "Loopback0",
        "templateName": "ext_vxlan_mpls_overlay_setup",
        "nvPairs": {
            "asn": ifc_create_overlay[f'D{i}'].value,
            "NEIGHBOR_IP": ifc_create_overlay[f'F{i}'].value,
            "NEIGHBOR_ASN": ifc_create_overlay[f'E{i}'].value
        }
    }
    print(ifc_overlay_details)
    response = requests.post(url_ifc_create, headers=headers_token, verify=False,
                             data=json.dumps(ifc_overlay_details))
    print(response.text)

requests.post(url_logout, headers=headers_token, verify=False)
