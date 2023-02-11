import json
import requests
from openpyxl import load_workbook
from DCNM_Authentication import headers_token, fabric_name, url_logout, dcnm_ip

wb = load_workbook('get-hostname-serial.xlsx')
sheet = wb['hostname-serial']

for i in range(2, len(sheet['A']) + 1):
    serial = sheet[f"B{i}"].value
    url = f"https://{dcnm_ip}/rest/control/policies/switches?serialNumber={serial}"
    response = requests.get(url=url, headers=headers_token, verify=False)
    for x in json.loads(response.text):
        if x["templateName"] == "host_11_1":
            sheet[f"C{i}"].value = x["policyId"]
            sheet[f"D{i}"].value = x["id"]
            wb.save('get-hostname-serial.xlsx')

requests.post(url_logout, headers=headers_token, verify=False)
