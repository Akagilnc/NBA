import json
from pprint import pprint
from openpyxl import Workbook, load_workbook
from os import listdir
from os.path import isfile, join


class ReadData:

    def read_file_name(self):
        onlyfiles = [f for f in listdir(".\datas") if isfile(join('.\datas', f))]
        row_num = 1
        for filename in onlyfiles:
            print('begin to make ' + str(filename))
            row_num = self.make_datas(filename, row_num)
            print('end of ' + str(filename))

        print('done')

    def make_datas(self, name, row_num):
        with open('./datas/' + name) as data_file:
            datas = json.load(data_file)
            data_title = datas.get('name')
            data_value = datas.get('value')

            wb = load_workbook('NBA.xlsx')
            sheet = wb.active
            # sheet.title = data_title
            row_num += 1

            for data in data_value:
                row_num += 1
                sheet.cell(column=1, row=row_num, value=data_title)
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
        return row_num

process = ReadData()
process.read_file_name()
