import json
import os
import sys
import xlwt
from xlwt import Workbook


# Get Virtual servers that are enabled
def GetVirtualServer(data):
	virtual_servers_enabled = []
	for k,v in data.items():
		if k == "virtual_servers":
			for i in range(len(v)):
				status = v[i]["properties"]["basic"]["enabled"]
				if status == True:
					virtual_servers_enabled.append(v[i]["name"])
					# Print name of enabled virtual server
					# print(v[i]["name"]) 

					# Print name of ports
					# print(v[i]["properties"]["basic"]["port"])

					# Print enabled virtual servers
					# print(virtual_servers_enabled)
					# Print names of Traffic IP 
					# print(v[i]["properties"]["basic"]["listen_on_traffic_ips"][0])
	return virtual_servers_enabled

# Get all the rules
def GetRules(data):
	rules = []
	contents = []

	wb = Workbook()
	sheet1=wb.add_sheet('Sheet 1')
	sheet1.write(0,0,'Rules')
	sheet1.write(0,1,'Contents')

	for k,v in data.items():
		if k == "rules":
			for i in range(len(v)):
				rules.append(v[i]["name"])
				contents.append(v[i]["content"])
				# Print Rules
				# print(v[i]["name"])
				# Print Content
				# print(contents])

	for x in range(len(rules)):
		# exceptions=["Rule_HTTPS_preprod.aitsl","rule_http_aitsl.edu.au","rule_http_preprod.aitsl.edu.au","Rule_HTTPS_preprod.aitsl"]
		# for items in exceptions:
		# 	if items == rules[x]:
		# 		break
		# 	else:
		# 		sheet1.write(x+1,0,rules[x])
		# 		sheet1.write(x+1,1,contents[x])
		# 		break
		if len(contents[x]) < 32767:
			print(contents[x])

	wb.save('rules.xls')



def main():
	with open('H:\\config.json', encoding="utf-8") as config_json:
		data_decode = json.load(config_json)
		data_dict = dict(data_decode)
	VirtualServers = GetVirtualServer(data_dict)
	GetRules(data_dict)
	


main()