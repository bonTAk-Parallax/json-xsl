import openpyxl
import json
from openpyxl import Workbook, load_workbook

json_data = {}

with open("books.json", encoding="utf8") as json_file:
    json_data = json.load(json_file)

wb = Workbook()
ws = wb.active
ws.title = "First Sheet"
lst = []
n = len(json_data["books"][0])-1
# print(json_data["books"][0])
lst = list(json_data["books"][0].keys())
print(lst)
for i in range(len(lst)):
    ws.cell(1, i+1, lst[i])
    for x in range(n):
        ws.cell(x+2, i+1, json_data["books"][x][lst[i]])

wb.save("test_json.xlsx")