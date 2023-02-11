import json
import requests
from openpyxl import load_workbook
from DCNM_Authentication import headers_token, fabric_name, url_logout, dcnm_ip

wb = load_workbook('get-hostname-serial.xlsx')
sheet = wb['hostname-serial']

for i in range(2, len(sheet['A']) + 1):
    serial = sheet[f"B{i}"].value
    policy = sheet[f"C{i}"].value
    url = f"https://{dcnm_ip}/rest/control/policies/{policy}"
    hostname = sheet[f"E{i}"].value
    hostname_details = {
        "id": int(sheet[f"D{i}"].value),
        "source": "",
        "serialNumber": serial,
        "policyId": policy,
        "entityType": "SWITCH",
        "entityName": "SWITCH",
        "templateName": "host_11_1",
        "priority": "100",
        "nvPairs": {
            "SWITCH_NAME": hostname,
            "FABRIC_NAME": fabric_name
        }
    }
    print(hostname_details)
    response = requests.put(url=url, headers=headers_token, verify=False, data=json.dumps(hostname_details))
    print(response.text)

requests.post(url_logout, headers=headers_token, verify=False)
