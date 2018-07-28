from netmiko import ConnectHandler
import re

cisco_ios1 = {
    'device_type': 'cisco_ios',
    'ip':   '172.16.100.203',
    'username': 'cisco',
    'password': 'cisco123'
}

net_connect = ConnectHandler(**cisco_ios1)

output = net_connect.send_command('show running-config')
output = output + net_connect.send_command('show inventory')
output = output + net_connect.send_command('show version')
output = output + net_connect.send_command('show memory statistics')
output = output + net_connect.send_command('show processes cpu sorted')
output = output + net_connect.send_command('show buffers')

#Parse
hostname = re.findall(r'^hostname (.*)\s+', output, re.I | re.M)
PID = re.findall(r'\s*PID:(.*)\s*,\sVID', output, re.I | re.M)
Description = re.findall(r',\s+DESCR:(.*)\s+', output, re.I | re.M)
SN = re.findall(r',\s+SN:(.*)\s+', output, re.I | re.M)
IOS = re.findall(r'^Cisco IOS Software,.*\s+\((.*)\),\s+', output, re.I | re.M)
Version = re.findall(r'^Cisco IOS Software.*,\s+Version\s+(.*),\s+', output, re.I | re.M)
Uptime = re.findall(r'\s*uptime is\s+(.*)\s*', output, re.I | re.M)
DRAM = re.findall(r'^cisco.*processor.*\swith\s(.*)\/', output, re.I | re.M)
MemTot = re.findall(r'^Processor\s+\w+\s+(\d+)\s+', output, re.I|re.M)
MemUse = re.findall(r'^Processor\s+\w+\s+\d+\s+(\d+)\s+', output, re.I|re.M)
MemFree = re.findall(r'^Processor\s+\w+\s+\d+\s+\d+\s+(\d+)\s+', output, re.I|re.M)
Cpu5sec = re.findall(r'\sfive\sseconds:\s(.*);\sone', output, re.I|re.M)
Cpu1min = re.findall(r'\sone\sminute:\s(.*);\sfive', output, re.I|re.M)
Cpu5min = re.findall(r'\sfive\sminutes:\s(.*)\s*', output, re.I|re.M)
smallhits = re.findall(r'^Small.*\s+.*\s+(\d+)\s+hits,', output, re.I|re.M)
smallmiss = re.findall(r'^Small.*\s+.*\s+.*(\d+)\s+misses,', output, re.I|re.M)
middlehits = re.findall(r'^Middle.*\s+.*\s+(\d+)\s+hits,', output, re.I|re.M)
middlemiss = re.findall(r'^Middle.*\s+.*\s+.*(\d+)\s+misses,', output, re.I|re.M)
bighits = re.findall(r'^Big.*\s+.*\s+(\d+)\s+hits,', output, re.I|re.M)
bigmiss = re.findall(r'^Big.*\s+.*\s+.*(\d+)\s+misses,', output, re.I|re.M)
verybighits = re.findall(r'^VeryBig.*\s+.*\s+(\d+)\s+hits,', output, re.I|re.M)
verybigmiss = re.findall(r'^VeryBig.*\s+.*\s+.*(\d+)\s+misses,', output, re.I|re.M)
largehits = re.findall(r'^Large.*\s+.*\s+(\d+)\s+hits,', output, re.I|re.M)
largemiss = re.findall(r'^Large.*\s+.*\s+.*(\d+)\s+misses,', output, re.I|re.M)
hugehits = re.findall(r'^Huge.*\s+.*\s+(\d+)\s+hits,', output, re.I|re.M)
hugemiss = re.findall(r'^Huge.*\s+.*\s+.*(\d+)\s+misses,', output, re.I|re.M)

# print(hostname)
# print(PID)
# print(Description)
# print(SN)
# print(IOS)
# print(Version)
# print(Uptime)
# print(DRAM)
# print(MemTot)
# print(MemUse)
# print(MemFree)
# print(Cpu5sec)
# print(Cpu1min)
# print(Cpu5min)
# print(smallhits)