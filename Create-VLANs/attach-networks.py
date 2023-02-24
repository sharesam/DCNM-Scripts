import json
import requests
from requests.auth import HTTPBasicAuth
from openpyxl import Workbook, load_workbook
from DCNM_Authentication import headers_token, fabric_name, url_logout, dcnm_ip

# this section loads the VLAN-Details.xlsx into the program

wb = load_workbook("Interface-Details-Retirement.xlsx")
network_attach_trunk = wb["Trunk-Interface-Details"]
network_attach_access = wb["Access-Interface-Details"]
all_networks = wb["all-networks"]
final_vlan_list = []
vlan_network_name_mapper = {}
url_network_attach = f"https://{dcnm_ip}/rest/top-down/fabrics/{fabric_name}/networks/attachments"
url_network_deploy = f"https://{dcnm_ip}/rest/top-down/fabrics/{fabric_name}/networks/deployments"

for i in range(2, len(network_attach_trunk['A']) + 1):
    vlan_list = network_attach_trunk[f'J{i}'].value.split(',')
    final_vlan_list.extend(x for x in vlan_list if x not in final_vlan_list)
    for x in final_vlan_list:
        for y in range(2, len(all_networks['A']) + 1):
            if x == str(all_networks[f'AA{y}'].value):
                vlan_network_name_mapper.update({x: all_networks[f'G{y}'].value})
        network_attach_details = [{
            "networkName": vlan_network_name_mapper[x],
            "lanAttachList": [
                {
                    "fabric": fabric_name,
                    "networkName": vlan_network_name_mapper[x],
                    "serialNumber": network_attach_trunk[f'B{i}'].value,
                    "switchPorts": "",
                    "detachSwitchPorts": "",
                    "vlan": x,
                    "untagged": 'false',
                    "freeformConfig": "",
                    "deployment": 'true',
                    "extensionValues": "",
                    "instanceValues": ""
                },
                {
                    "fabric": fabric_name,
                    "networkName": vlan_network_name_mapper[x],
                    "serialNumber": network_attach_trunk[f'C{i}'].value,
                    "switchPorts": "",
                    "detachSwitchPorts": "",
                    "vlan": x,
                    "untagged": 'false',
                    "freeformConfig": "",
                    "deployment": 'true',
                    "extensionValues": "",
                    "instanceValues": ""
                }
            ]
        }]
        print(network_attach_details)
        response = requests.post(url_network_attach, headers=headers_token, verify=False,
                                 data=json.dumps(network_attach_details))
        # response_deploy = requests.post(url_network_deploy, headers=headers_token, verify=False,
        #                                  data=json.dumps(network_deploy))
        print(response.text)
    final_vlan_list = []
    vlan_network_name_mapper = {}

final_vlan_list = []
vlan_network_name_mapper = {}

for i in range(2, len(network_attach_access['A']) + 1):
    vlan_list = str(network_attach_access[f'K{i}'].value)
    final_vlan_list.append(vlan_list)
    # final_vlan_list.extend(x for x in vlan_list if x not in final_vlan_list)
    for x in final_vlan_list:
        for y in range(2, len(all_networks['A']) + 1):
            if x == str(all_networks[f'AA{y}'].value):
                vlan_network_name_mapper.update({x: all_networks[f'G{y}'].value})
        network_attach_details = [{
            "networkName": vlan_network_name_mapper[x],
            "lanAttachList": [
                {
                    "fabric": fabric_name,
                    "networkName": vlan_network_name_mapper[x],
                    "serialNumber": network_attach_access[f'B{i}'].value,
                    "switchPorts": "",
                    "detachSwitchPorts": "",
                    "vlan": x,
                    "untagged": 'false',
                    "freeformConfig": "",
                    "deployment": 'true',
                    "extensionValues": "",
                    "instanceValues": ""
                },
                {
                    "fabric": fabric_name,
                    "networkName": vlan_network_name_mapper[x],
                    "serialNumber": network_attach_access[f'C{i}'].value,
                    "switchPorts": "",
                    "detachSwitchPorts": "",
                    "vlan": x,
                    "untagged": 'false',
                    "freeformConfig": "",
                    "deployment": 'true',
                    "extensionValues": "",
                    "instanceValues": ""
                }
            ]
        }]
        print(network_attach_details)
        response = requests.post(url_network_attach, headers=headers_token, verify=False,
                                 data=json.dumps(network_attach_details))
        # response_deploy = requests.post(url_network_deploy, headers=headers_token, verify=False,
        #                                  data=json.dumps(network_deploy))
        print(response.text)
    final_vlan_list = []
    vlan_network_name_mapper = {}
requests.post(url_logout, headers=headers_token, verify=False)
