from openpyxl import load_workbook


class Get1617:
    def __init__(self):
        self.data = list()

    def read_file(self):
        wb = load_workbook('./docs/2015-2017NBA&BAA&ABA ELO.xlsx')
        sheet = wb.active
        target = list()
        for rownum, line in enumerate(sheet.values):
            if rownum == 0:
                continue
            if line[9] != '？':
                self.data.append(line[8:10])
            else:
                target.append((rownum, line[8:10]))

        self.data.sort()
        return target

    def get_elo(self):
        target = self.read_file()

        for rownum, line in enumerate(target):
            y = self.calculator(line[1][0])
            target[rownum] = (line[0], (line[1][0], y))

    def calculator(self, x):
        y = 0
        self.find_nearest(x)
        return y

    def find_nearest(self, x):
        hi = len(self.data)
        low = 0
        x1, x2 = -1, -1
        while True:
            if hi == low:
                x1 = x2 = hi
            if hi - 1 == low:
                x1, x2 = hi, low

            if x1 != -1 and x2 != -1:
                print (str(x1) + "=======" + str(x2))
                return x1, x2
            mid = low + (hi - low)//2
            #print(mid)
            if x < (self.data[mid])[0]:
                #print('high')
                hi = mid
            elif x > self.data[mid][0]:
                #print('low')
                low = mid
            else:
                hi = low = mid

            if mid == 0:
                x1, x2 = 0, 1




process = Get1617()
process.get_elo()
