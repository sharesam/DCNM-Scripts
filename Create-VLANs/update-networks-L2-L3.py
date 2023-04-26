import json
import requests
from requests.auth import HTTPBasicAuth
from openpyxl import load_workbook

from DCNM_Authentication import headers_token, fabric_name, url_logout, dcnm_ip

wb = load_workbook("Interface-Details-az-data.xlsx")
update_networks = wb["all-networks"]

for i in range(2, len(update_networks['A']) + 1):
    url_update_networks = f"https://{dcnm_ip}/rest/top-down/fabrics/{fabric_name}" \
                          f"/networks/{update_networks[f'G{i}'].value}"
    network_details = {
        "fabric": fabric_name,
        "vrf": update_networks[f'B{i}'].value,
        "networkName": update_networks[f'G{i}'].value,
        "displayName": update_networks[f'G{i}'].value,
        "networkId": update_networks[f'D{i}'].value,
        "networkTemplateConfig": "{\"gatewayIpAddress\":\"" + str(update_networks[f'AB{i}'].value) + "\","
                                                                                                     "\"gatewayIpV6Address\":\"\","
                                                                                                     "\"vlanName\":\"" + str(
            update_networks[f'G{i}'].value) + "\","
                                              "\"intfDescription\":\"" + str(update_networks[f'AE{i}'].value) + "\","
                                                                                                                "\"mtu\":\"\","
                                                                                                                "\"secondaryGW1\":\"\","
                                                                                                                "\"secondaryGW2\":\"\","
                                                                                                                "\"secondaryGW3\":\"\","
                                                                                                                "\"secondaryGW4\":\"\","
                                                                                                                "\"suppressArp\":false,"
                                                                                                                "\"enableIR\":false,"
                                                                                                                "\"trmEnabled\":false,"
                                                                                                                "\"rtBothAuto\":true,"
                                                                                                                "\"enableL3OnBorder\":false,"
                                                                                                                "\"mcastGroup\":\"" + str(
            update_networks[f'AF{i}'].value) + "\","
                                               "\"dhcpServerAddr1\":\"\","
                                               "\"vrfDhcp\":\"\","
                                               "\"dhcpServerAddr2\":\"\","
                                               "\"vrfDhcp2\":\"\","
                                               "\"dhcpServerAddr3\":\"\","
                                               "\"vrfDhcp3\":\"\","
                                               "\"loopbackId\":\"\","
                                               "\"tag\":\"12345\","
                                               "\"vrfName\":\"" + str(update_networks[f'B{i}'].value) + "\","
                                                                                                        "\"isLayer2Only\":false,"
                                                                                                        "\"nveId\":1,"
                                                                                                        "\"vlanId\":\"" + str(
            update_networks[f'AA{i}'].value) + "\","
                                               "\"segmentId\":\"" + str(update_networks[f'D{i}'].value) + "\","
                                                                                                          "\"networkName\":\"" + str(
            update_networks[f'G{i}'].value) + "\"}",
        "networkTemplate": "Default_Network_Universal",
        "networkExtensionTemplate": "Default_Network_Extension_Universal"
    }
    print(network_details)
    response = requests.put(url_update_networks, headers=headers_token, verify=False, data=json.dumps(network_details))
    print(response.text)
