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
dcnm_token = json.loads(response.text)['Dcnm-Token']  ##the token returned from dcnm is stored in this variable
headers_token = {'Content-Type': 'application/json', 'dcnm-token': str(dcnm_token)}

#### this section loads the VLAN-Details.xlsx into the program

wb = load_workbook(
    '/Users/krishna/python-projects/My_venvs/venv1-enterprise/dcnm-python-Testing/Create-VLANs/VLAN-Details.xlsx')
network_add = wb["VLAN-Create-L2-Only"]
url_network_add = f"https://{dcnm_ip}/rest/top-down/fabrics/{fabric_name}/networks"
# url_role = f"https://{dcnm_ip}/rest/control/switches/roles" ####default role is Leaf

###### Create Networks/VLANs #####
for i in range(2, len(network_add['A']) + 1):
    network_details = {
        "fabric": fabric_name,
        "vrf": "NA",
        "networkName": network_add[f'A{i}'].value,
        "displayName": network_add[f'B{i}'].value,
        "networkId": network_add[f'C{i}'].value,
        "networkTemplateConfig": "{\"gatewayIpAddress\":\"\",\"gatewayIpV6Address\":\"\",\"vlanName\":\"\",\"intfDescription\":\"\",\"mtu\":\"9216\",\"secondaryGW1\":\"\",\"secondaryGW2\":\"\",\"secondaryGW3\":\"\",\"secondaryGW4\":\"\",\"suppressArp\":false,\"enableIR\":false,\"trmEnabled\":false,\"rtBothAuto\":true,\"enableL3OnBorder\":false,\"mcastGroup\":\"" + str(
            network_add[
                f'E{i}'].value) + "\",\"dhcpServerAddr1\":\"\",\"vrfDhcp\":\"\",\"dhcpServerAddr2\":\"\",\"vrfDhcp2\":\"\",\"dhcpServerAddr3\":\"\",\"vrfDhcp3\":\"\",\"loopbackId\":\"\",\"tag\":\"12345\",\"vrfName\":\"NA\",\"isLayer2Only\":true,\"nveId\":1,\"vlanId\":\"" + str(
            network_add[f'F{i}'].value) + "\",\"segmentId\":\"" + str(
            network_add[f'C{i}'].value) + "\",\"networkName\":\"" + str(network_add[f'A{i}'].value) + "\"}",
        "networkTemplate": "Default_Network_Universal",
        "networkExtensionTemplate": "Default_Network_Extension_Universal",
        # "source": "null",
        # "serviceNetworkTemplate": "null"
    }
    response = requests.post(url_network_add, headers=headers_token, verify=False, data=json.dumps(network_details))
    print(response.text)

requests.post(url_logout, headers=headers_token, verify=False)
