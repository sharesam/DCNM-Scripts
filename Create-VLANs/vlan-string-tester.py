from openpyxl import load_workbook

wb = load_workbook("Interface-Details-Mainframe.xlsx")
network_attach_trunk = wb["Trunk-Interface-Details"]
network_attach_access = wb["Access-Interface-Details"]
all_networks = wb["all-networks"]
final_vlan_list = []
vlan_network_name_mapper = {}

for i in range(2, len(network_attach_trunk['A']) + 1):
    vlan_list = network_attach_trunk[f'I{i}'].value.split(',')
    final_vlan_list.extend(x for x in vlan_list if x not in final_vlan_list)

for x in final_vlan_list:
    for y in range(2, len(all_networks['A']) + 1):
        if x == str(all_networks[f'AA{y}'].value):
            vlan_network_name_mapper.update({x: all_networks[f'G{y}'].value})

for i in range(2, len(network_attach_access['A']) + 1):
    vlan_list = network_attach_access[f'J{i}'].value.split(',')
    final_vlan_list.extend(x for x in vlan_list if x not in final_vlan_list)

for x in final_vlan_list:
    for y in range(2, len(all_networks['A']) + 1):
        if x == str(all_networks[f'AA{y}'].value):
            vlan_network_name_mapper.update({x: all_networks[f'G{y}'].value})

print(final_vlan_list)
print(vlan_network_name_mapper.keys())
