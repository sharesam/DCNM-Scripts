import json
import requests
from openpyxl import load_workbook
from DCNM_Authentication import headers_token, fabric_name, url_logout, dcnm_ip

wb = load_workbook('/Users/krishna/python-projects/My_venvs/venv1-enterprise/DCNM-Python/VRFs/VRF-Details.xlsx')
vrf_attach = wb["VRF-Association"]
url_vrf_attach = f"https://{dcnm_ip}/rest/top-down/fabrics/{fabric_name}/vrfs/attachments"

# Attach VRFs #####
for i in range(2, len(vrf_attach['A']) + 1):
    if vrf_attach[f'B{i}'].value == 'Yes':
        vrf_attach_details = [{"vrfName": vrf_attach[f'A{i}'].value,
                               "lanAttachList": [
                                   {
                                       "fabric": fabric_name,
                                       "vrfName": vrf_attach[f'A{i}'].value,
                                       "serialNumber": vrf_attach[f'C{i}'].value,
                                       "vlan": vrf_attach[f'E{i}'].value,
                                       "deployment": 'true'
                                   },
                                   {
                                       "fabric": fabric_name,
                                       "vrfName": vrf_attach[f'A{i}'].value,
                                       "serialNumber": vrf_attach[f'D{i}'].value,
                                       "vlan": vrf_attach[f'E{i}'].value,
                                       "deployment": 'true'
                                   },
                               ]}]
        requests.post(url_vrf_attach, headers=headers_token, verify=False, data=json.dumps(vrf_attach_details))
    elif vrf_attach[f'B{i}'].value == 'No':
        vrf_attach_details = [{"vrfName": vrf_attach[f'A{i}'].value,
                               "lanAttachList": [
                                   {
                                       "fabric": fabric_name,
                                       "vrfName": vrf_attach[f'A{i}'].value,
                                       "serialNumber": vrf_attach[f'C{i}'].value,
                                       "vlan": vrf_attach[f'E{i}'].value,
                                       "freeformConfig": f"vrf context {vrf_attach[f'A{i}'].value}\n"
                                                         f"  address-family ipv4 unicast\n"
                                                         f"    route-target import 65200:{vrf_attach[f'F{i}'].value}\n"
                                                         f"    route-target export 65200:{vrf_attach[f'F{i}'].value}\n"
                                                         f"  address-family ipv6 unicast\n"
                                                         f"    route-target import 65200:{vrf_attach[f'F{i}'].value}\n"
                                                         f"    route-target export 65200:{vrf_attach[f'F{i}'].value}\n",
                                       "deployment": 'true'
                                   }
                               ]}]
        requests.post(url_vrf_attach, headers=headers_token, verify=False, data=json.dumps(vrf_attach_details))
        print(f"vrf {vrf_attach[f'A{i}'].value} attached")
        print(f"========================================")

requests.post(url_logout, headers=headers_token, verify=False)
