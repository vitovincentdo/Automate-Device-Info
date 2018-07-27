import openpyxl
from collections import OrderedDict

wb = openpyxl.load_workbook('DeviceList.xlsx', data_only=True)
rs = wb["DeviceList"]

mylist = []

for cols in rs.iter_cols(min_row=1, min_col=1):
    for cell in cols:
        mylist.append(cell.value)
        # print('cell %s' % (cell.value))
        # print('cell %s %s' % (cell.coordinate,cell.value))

ip_index = mylist.index("ip")
user_index = mylist.index("username")
pass_index = mylist.index("password")
mylistIP = mylist[ip_index+1:user_index]
mylistUser = mylist[user_index+1:pass_index]
mylistPass = mylist[pass_index+1:]

dictcomb = OrderedDict([('IP', mylistIP), ('username', mylistUser), ('passwordd', mylistPass)])

# mylistDevType = []
# for i, count in enumerate(mylistIP, 1):
#     mylistDevType.append("cisco_ios")

# combine = list(zip(mylistDevType,mylistIP,mylistUser,mylistPass))
# print(combine)

for x, y, z in zip(mylistIP, mylistUser, mylistPass):
    print(x)
    print(y)
    print(z)