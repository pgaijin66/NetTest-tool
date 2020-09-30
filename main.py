import json
import os
import sys
import xlwt
from xlwt import Workbook
import subprocess
import re
import socket


def get_virtual_servers(data):
	"""
	Searches dictionary and parses virtual servers.

	Arguments:
	data : Dictionary created by parsing Config.json file.

	Returns:
	ruvirtual_servers_enabled : List of virtual servers which are enabled
	
	"""
	virtual_servers = []
	for k,v in data.items():
		if k == "virtual_servers":
			for i in range(len(v)):
    				virtual_servers.append(v[i]["name"])
	return virtual_servers




def get_rules(data):
	"""
	Searches dictionary and parses rules and contents.

	Arguments:
	data : Dictionary created by parsing Config.json file.

	Returns:
	rules : List of rules.
	Contents: All the configurations from the rule.

	"""
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
				# print(v[i]["name"]) # Print Rules
				# print(contents]) # Print Content
	for x in range(len(rules)):
		exceptions=["Rule_HTTPS_preprod.aitsl","rule_http_aitsl.edu.au","rule_http_preprod.aitsl.edu.au","Rule_HTTPS_preprod.aitsl","rule_asiaeducation_redirect"]
		for items in exceptions:
			if len(contents[x]) < 32767:
				if items == rules[x]:
					break
				else:
					sheet1.write(x+1,0,rules[x])
					sheet1.write(x+1,1,contents[x])
					break
	wb.save(r'C:\Users\pthapa\Workspace\ESA\rules.xls')
	return rules, contents




def get_pool_node(data):
	"""
	Searches dictionary and parses pools and exports into excel sheet.

	Arguments:
	data : Dictionary created by parsing Config.json file.

	Returns:
	pool : Returns list of pool.

	"""
	pool = []
	node = []
	# wb = Workbook()
	# sheet1=wb.add_sheet('Sheet 1', cell_overwrite_ok=False)
	# sheet1.write(0,0,'Pools')
	# sheet1.write(0,1,'Nodes')
	# sheet1.write(0,2,'State')
	# sheet1.write(0,3,'Status Check')
	for k,v in data.items():
		if k == "pools":
			for i in range(len(v)):
				# print(v[i]["name"])
				for j in range(len(v[i]["properties"]["basic"]["nodes_table"])):
					# print(v[i]["properties"]["basic"]["nodes_table"][j]["state"])
					# print(v[i]["properties"]["basic"]["nodes_table"][j]["node"])
					node.append(v[i]["properties"]["basic"]["nodes_table"][j]["node"])
					# print("[*] Checking Ping status")
					check_stat(v[i]["properties"]["basic"]["nodes_table"][j]["node"])
					# print(v[i]["properties"]["basic"]["nodes_table"][j]["state"])
					# print(v[i]["name"])	
					# sheet1.write(i+1,0,v[i]["name"])
					# sheet1.write(i+1,1,v[i]["properties"]["basic"]["nodes_table"][j]["node"])
					# sheet1.write(i+1,2,v[i]["properties"]["basic"]["nodes_table"][j]["state"])
					pool.append(v[i]["name"])
	return pool





def check_stat(node):
	"""
	Splits node into ip and port and check ping test and telnet test.

	Arguments:
	node : Socket containing IP and PORT.

	"""
	ip, port = node.split(":")
	check_ping(ip)
	check_telnet(ip,port)




def check_ping(ip):
	"""
	Checks ping.

	Arguments:
	IP : IP address.

	"""
	prog = subprocess.run(["ping", ip], stdout=subprocess.PIPE)
	out = str(prog.stdout)
	if "Destination host unreachable" in out:
		print("Destination host unreachable")
	elif "\\r\\nRequest timed out." in out:
		print("Request time out.")
	else:
		print("Pingable")




def check_telnet(ip, port):
	"""
	Checks telnet connection.

	Arguments:
	ip : IP address.
	port: Port number.

	"""
	s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
	try:
		s.connect((ip,int(port)))
		s.shutdown(2)
		return True
	except:
		return False



def main():
	"""
	Main Program flow.

	Calls function to return list of rules and contents

	"""
	with open(r'C:\Users\pthapa\Workspace\ESA\config.json', encoding="utf-8") as config_json:
		data_decode = json.load(config_json)
		data_dict = dict(data_decode)
	get_virtual_servers(data_dict)
	get_rules(data_dict)
	get_pool_node(data_dict)



if __name__=="__main__":
	main()
