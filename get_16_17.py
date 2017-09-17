from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows


class Get1617:
    def read_file(self):
        wb = load_workbook('./docs/2015-2017NBA&BAA&ABA ELO.xlsx')
        sheet = wb.active
        rownum = 0
        data = list()
        target = list()
        for line in sheet.values:
            rownum += 1
            if rownum == 1:
                continue
            if line[9] != '？':
                data.append(line[8:10])
            else:
                target.append(line[8:10])

        data.sort()
        return data, target

    def get_elo(self):
        data, target = self.read_file()

        for index, line in enumerate(target):
            y = self.calculator(line[0])
            target[index] = (line[0], y)

        print(target)

    def calculator(self, x):
        y = 0
        return y





process = Get1617()
process.get_elo()