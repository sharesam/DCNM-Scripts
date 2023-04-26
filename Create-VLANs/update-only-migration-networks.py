from openpyxl import load_workbook

wb = load_workbook("Interface-Details-az-data.xlsx")
all_networks = wb["all-networks"]
update_networks = wb["L3-SVI-Details"]
final_vlan_list = []
vlan_network_name_mapper = {}

# update the VLAN number with Network Name
for i in range(2, len(update_networks['A']) + 1):
    for j in range(2, len(all_networks['A']) + 1):
        if update_networks[f'A{i}'].value == all_networks[f'AA{j}'].value:
            print(update_networks[f'A{i}'].value, all_networks[f'AA{j}'].value)
            all_networks[f'AH{j}'].value = 'yes'
            all_networks[f'AB{j}'].value = update_networks[f'D{i}'].value + update_networks[f'E{i}'].value
            all_networks[f'AE{j}'].value = update_networks[f'F{i}'].value
            all_networks[f'B{j}'].value = update_networks[f'B{i}'].value
            all_networks[f'P{j}'].value = "FALSE"

# update all-networks tab with migration only vlans
for k in range(2, len(all_networks['A']) + 1):
    if all_networks[f'AH{k}'].value is None:
        all_networks.delete_rows(k)
wb.save("Interface-Details-az-data.xlsx")

if len(update_networks['A']) == len(all_networks['A']):
    print("seems ok")
else:
    print("check again. rows not matching")
# url_update_network = f"https://{dcnm_ip}/rest/top-down/fabrics/{fabric_name}/networks/attachments"
