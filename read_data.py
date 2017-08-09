import json
from pprint import pprint
from openpyxl import Workbook, load_workbook

with open('./datas/warriors.json') as data_file:
    datas = json.load(data_file)
    data_title = datas.get('name')
    data_value = datas.get('value')

    wb = load_workbook('NBA.xlsx')
    sheet = wb.create_sheet()
    sheet.title = data_title
    row_num = 1

    for data in data_value:
        row_num += 1
        sheet.cell(column=2, row=row_num, value=data['x'])
        sheet.cell(column=3, row=row_num, value=data['y'])
        if 'p' in data.keys():
            sheet.cell(column=5, row=row_num, value=data['p'])
        if 'p' in data.keys() and data['p'] != 'f':
            sheet.cell(column=4, row=row_num, value=data['t'])
            sheet.cell(column=6, row=row_num, value=data['o'])
            sheet.cell(column=7, row=row_num, value=data['d'])

    wb.save('NBA.xlsx')
    print(data_title)

