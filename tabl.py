import requests
import openpyxl
from pathlib import Path
from openpyxl import Workbook
dls = "https://enrol.hse.ru/storage/public_report/moscow/Bachelors/Mark.xlsx"
resp = requests.get(dls)
with open('test.xlsx', 'wb') as output:
    output.write(resp.content)
xlsx_file = Path('test.xlsx')
wb = openpyxl.load_workbook(xlsx_file)
wb.active = 0
sheet = wb.active
print(sheet['D20'].value)
N = 80
M = 15
A = []
for i in range(N):
    A.append([0]*M)
for i in range(80):
    A[i][0] = i+1
for i in range(80):
    A[i][1] = sheet['B'+str(i+16)].value
for i in range(80):
    A[i][2] = sheet['F'+str(i+16)].value
for i in range(80):
    A[i][3] = sheet['H'+str(i+16)].value
for i in range(80):
    A[i][4] = sheet['K'+str(i+16)].value
for i in range(80):
    A[i][5] = sheet['L'+str(i+16)].value
for i in range(80):
    A[i][6] = sheet['N'+str(i+16)].value
for i in range(80):
    A[i][7] = sheet['P'+str(i+16)].value
for i in range(80):
    A[i][8] = sheet['R'+str(i+16)].value # 18
for i in range(80):
    A[i][9] = sheet['T'+str(i+16)].value
for i in range(80):
    A[i][10] = sheet['V'+str(i+16)].value
for i in range(80):
    A[i][11] = sheet['X'+str(i+16)].value
for i in range(80):
    A[i][12] = sheet['Z'+str(i+16)].value
for i in range(80):
    A[i][13] = sheet['AB'+str(i+16)].value
for i in range(80):
    A[i][14] = sheet['AD'+str(i+16)].value


#for i in range(len(A)):
 #   for j in range(len(A[i])):
  #    print(A[i][j], end = ' ')
   # print()
count = 0
for i in range(80):
    if A[i][1]=='15778490223':
        pos=i+1
        bal=A[i][10]
for i in range(80):
    if A[i][3]!=None:
        count += 1
count+=3
first_pos=0
for i in range(80):
    if A[i][10]==bal:
        if first_pos==0:
            first_pos=i+1
        else:
            last_pos=i+1
qop=0
for i in range(80):
    if (A[i][5]=='+'):
        qop=i+1


#print(last_pos-qop+count)
sona=count+pos-qop
posl=last_pos-qop+count
#print(son,  posl)
