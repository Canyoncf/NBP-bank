from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
import requests
import json
import datetime
import os

x = datetime.date.today()
y = x.month - 2
z = x.year - 1

if x.month == 1:
    y = 11

if x.month == 2:
    y = 12
q = datetime.date(z,y,x.day)
p = datetime.date(z,12,31)
print(q,p)

if y== 12 or y == 11:
    url = f"https://api.nbp.pl/api/exchangerates/tables/A/2023-01-01/{x}/?format=json"
    url2 = f"https://api.nbp.pl/api/exchangerates/tables/A/{q}/{z}-12-31/?format=json"
    response = requests.get(url)
    print(url)
    response2 = requests.get(url2)
    data = json.loads(response.text)
    data2 = json.loads(response2.text)

    print(len({x.day}))
    if len(str({x.day})) == 1:
        m = f"0{x.day}"
    else:
        m = f"{x.day}"

    if len(str({y})) == 1:
        y = f"0{y}"
    else:
        y = f"{y}"
    print(f"{x.year}-{y}-{m}")

    url3 = f"https://api.nbp.pl/api/exchangerates/tables/A/{x.year}-0{y}-{m}/{x}/?format=json"
    response3 = requests.get(url3)
    data3 = json.loads(response3.text)