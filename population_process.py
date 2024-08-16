import xlwings as xl
import numpy as np
import json


class PopAndIncomeProcess:
    def __init__(self):
        np.printoptions(suppress=True)  # 输出设置
        self.app = xl.App(visible=False)
        self.app.display_alerts = False
        self.path = './data/Social and economic characteristics/'
        self.year = 2010
        self.name1 = 'Web_ACS'
        self.name2 = '_Pop-Race.xlsx'
        self.name3 = '_Inc-Pov-Emp.xlsx'
        self.sheet_name_pop = 'Total Pop & Median Age'
        self.sheet_name_income = 'Income'
        self.place_list = [None, '00000']
        self.mark = ['Place', 'place']

        self.dic = {}

    def data_get(self, name1, name2, sheet_name, data_type, position):
        for y in range(1, 20):
            try:
                path = self.path + name1 + str(self.year + y - 1) + name2
                wb_read = self.app.books.open(path)
                print(path + '-' + data_type + '-COMPLETE!')
                wb_read_sht = wb_read.sheets[sheet_name]

                column_mark = 0
                for q in range(4, 8):
                    for k in range(1, 20):
                        p = wb_read_sht.range((q, k)).value
                        if p in self.mark:
                            # print(p)
                            column_mark = k
                            break
                # print(column_mark)
                for j in range(7, 200):
                    # print((j, column_mark))
                    # print(wb_read_sht.range((j, a)).value)
                    if wb_read_sht.range((j, column_mark)).value in self.place_list:
                        if wb_read_sht.range('B' + str(j)).value is not None:
                            # print(wb_read_sht.range('B' + str(j)).value)
                            # print(wb_read_sht.range('A' + str(j)).value)
                            # self.dic[y + self.year] = {wb_read_sht.range('A' + str(j)).value}
                            county = wb_read_sht.range('A' + str(j)).value
                            year = y + self.year
                            if year in self.dic.keys():
                                if county in self.dic[year]:
                                    self.dic[year][county][data_type] = wb_read_sht.range(position + str(j)).value
                                else:
                                    self.dic[year][county] = {data_type: wb_read_sht.range(position + str(j)).value}
                            else:
                                self.dic[year] = {county: {data_type: wb_read_sht.range(position + str(j)).value}}
                wb_read.close()
            except FileNotFoundError:
                print(data_type + str(self.year + y) + 'File Not Found')
                continue

    def main(self):
        self.data_get(self.name1, self.name2, self.sheet_name_pop, 'pop', 'B')
        self.data_get(self.name1, self.name3, self.sheet_name_income, 'median_income', 'F')
        self.data_get(self.name1, self.name3, self.sheet_name_income, 'mean_income', 'H')
        print(self.dic)
        with open('./results/pop_and_income.json', 'w') as f:
            json.dump(self.dic, f)


if __name__ == '__main__':
    p = PopAndIncomeProcess()
    p.main()

