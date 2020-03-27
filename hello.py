import requests, os
import json
import openpyxl
print('hello')
g = input('did it worky')
print(g)




import glob
files = glob.glob('./*.xlsx')

for file in files:
    print('a')
    xl = openpyxl.load_workbook(file)
    sheet_name = xl.sheetnames[0]

    sheet = xl[sheet_name]
    print(sheet['D2'].value)
    print('Make sure latitude is in column D and longitude in E'