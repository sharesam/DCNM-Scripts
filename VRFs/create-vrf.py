import json
import requests
from requests.auth import HTTPBasicAuth
from openpyxl import Workbook, load_workbook

# this section gathers the token from DCNM which will be used for the subsequent Get or Post requests####
dcnm_ip = "10.122.104.55"
fabric_name = "SoAZ_Fabric"
url_login = f"https://{dcnm_ip}/rest/logon"
url_logout = f"https://{dcnm_ip}/rest/logout"
requests.packages.urllib3.disable_warnings()
dcnm_creds = HTTPBasicAuth('admin', 'C!sc0123')
headers = {'Content-Type': 'application/json'}
request_body = {"expirationTime": "999999"}

response = requests.post(url_login, headers=headers, auth=dcnm_creds, verify=False, data=json.dumps(request_body))
dcnm_token = json.loads(response.text)['Dcnm-Token']  ##the token returned from dcnm is stored in this variable
headers_token = {'Content-Type': 'application/json', 'dcnm-token': str(dcnm_token)}

# this section loads the VRF.xlsx into the program

wb = load_workbook('/Users/krishna/python-projects/My_venvs/venv1-enterprise/DCNM-Python-Testing/VRFs/VRF-Details.xlsx')
vrf_add = wb["VRF-Create"]
url_vrf_add = f"https://{dcnm_ip}/rest/top-down/fabrics/{fabric_name}/vrfs"
# url_role = f"https://{dcnm_ip}/rest/control/switches/roles" ####default role is Leaf

###### Create VRFs #####
for i in range(2, len(vrf_add['A']) + 1):
    vrf_details = {"fabric": fabric_name, "vrfName": vrf_add[f'A{i}'].value, "vrfId": str(vrf_add[f'B{i}'].value),
                   "vrfTemplate": "Default_VRF_Universal",
                   "vrfTemplateConfig": "{\"advertiseDefaultRouteFlag\":\"true\",\"vrfVlanId\":\"\",\"isRPExternal\":\"false\",\"vrfDescription\":\"\",\"L3VniMcastGroup\":\"\",\"maxBgpPaths\":\"1\",\"maxIbgpPaths\":\"2\",\"vrfSegmentId\":" + "\"" + str(
                       vrf_add[
                           f'B{i}'].value) + "\"," + "\"ipv6LinkLocalFlag\":\"true\",\"vrfRouteMap\":\"FABRIC-RMAP-REDIST-SUBNET\",\"configureStaticDefaultRouteFlag\":\"true\",\"trmBGWMSiteEnabled\":\"false\",\"tag\":\"12345\",\"rpAddress\":\"\",\"nveId\":\"1\",\"bgpPasswordKeyType\":\"3\",\"bgpPassword\":\"\",\"mtu\":\"9216\",\"multicastGroup\":\"\",\"advertiseHostRouteFlag\":\"false\",\"vrfVlanName\":\"\",\"trmEnabled\":\"false\",\"loopbackNumber\":\"\",\"asn\":\"65001\",\"vrfIntfDescription\":\"\",\"vrfName\":""\"" + str(
                       vrf_add[f'A{i}'].value) + "\"""}", "vrfExtensionTemplate": "Default_VRF_Extension_Universal"}
    # vrf_details = {"fabric":"AZ-Phoenix", "vrfName":"Agency-1-VRF-2", "vrfId":"50002", "vrfTemplate":"Default_VRF_Universal", "vrfTemplateConfig": "{\"advertiseDefaultRouteFlag\":\"true\",\"vrfVlanId\":\"\",\"isRPExternal\":\"false\",\"vrfDescription\":\"\",\"L3VniMcastGroup\":\"\",\"maxBgpPaths\":\"1\",\"maxIbgpPaths\":\"2\",\"vrfSegmentId\":\"50002\",\"ipv6LinkLocalFlag\":\"true\",\"vrfRouteMap\":\"FABRIC-RMAP-REDIST-SUBNET\",\"configureStaticDefaultRouteFlag\":\"true\",\"trmBGWMSiteEnabled\":\"false\",\"tag\":\"12345\",\"rpAddress\":\"\",\"nveId\":\"1\",\"bgpPasswordKeyType\":\"3\",\"bgpPassword\":\"\",\"mtu\":\"9216\",\"multicastGroup\":\"\",\"advertiseHostRouteFlag\":\"false\",\"vrfVlanName\":\"\",\"trmEnabled\":\"false\",\"loopbackNumber\":\"\",\"asn\":\"65001\",\"vrfIntfDescription\":\"\",\"vrfName\":\"Agency-1-VRF-2\"}","vrfExtensionTemplate":"Default_VRF_Extension_Universal"}
    response = requests.post(url_vrf_add, headers=headers_token, verify=False, data=json.dumps(vrf_details))
    # print(json.dumps(vrf_details))
    print(response.text)

requests.post(url_logout, headers=headers_token, verify=False)
