import json
import requests
from openpyxl import load_workbook
from DCNM_Authentication import headers_token, fabric_name, url_logout, dcnm_ip

wb = load_workbook('ifc.xlsx')
ifc_create_underlay = wb["underlay"]
url_ifc_create = f"https://{dcnm_ip}/rest/control/links"

for i in range(2, len(ifc_create_underlay['A']) + 1):
    ifc_underlay_details = {
        "sourceFabric": fabric_name,
        "destinationFabric": "SoAZ-EXT",
        "sourceDevice": ifc_create_underlay[f'B{i}'].value,
        "destinationDevice": "",
        "sourceSwitchName": ifc_create_underlay[f'A{i}'].value,
        "destinationSwitchName": ifc_create_underlay[f'C{i}'].value,
        "sourceInterface": ifc_create_underlay[f'D{i}'].value,
        "destinationInterface": ifc_create_underlay[f'E{i}'].value,
        "templateName": "ext_vxlan_mpls_underlay_setup",
        "nvPairs": {
            "IP_MASK": ifc_create_underlay[f'F{i}'].value,
            "NEIGHBOR_IP": ifc_create_underlay[f'G{i}'].value,
            "MPLS_FABRIC": ifc_create_underlay[f'H{i}'].value,
            "PEER1_SR_MPLS_INDEX": "",
            "PEER2_SR_MPLS_INDEX": "",
            "GB_BLOCK_RANGE": "",
            "DCI_ROUTING_PROTO": ifc_create_underlay[f'I{i}'].value,
            "OSPF_AREA_ID": "0.0.0.0",
            "DCI_ROUTING_TAG": "MPLS_UNDERLAY",
            "MTU": "1500",
            "PEER1_DESC": f"To-{ifc_create_underlay[f'C{i}'].value}-{ifc_create_underlay[f'E{i}'].value}",
            "PEER2_DESC": "",
            "PEER1_CONF": "",
            "PEER2_CONF": ""
        }
    }
    print(ifc_underlay_details)
    response = requests.post(url_ifc_create, headers=headers_token, verify=False,
                             data=json.dumps(ifc_underlay_details))
    print(response.text)

requests.post(url_logout, headers=headers_token, verify=False)
