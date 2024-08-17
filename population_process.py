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
                            year = y + self.year -1
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

    def data_process(self):
        # delete incomplete county
        l = []
        lost = []
        for year in self.dic.keys():
            l.append(list(self.dic[year].keys()))
        t = 0
        for i in range(1, len(l)):
            # print(set(l[t]).symmetric_difference(set(l[i])))
            for j in set(l[t]).symmetric_difference(set(l[i])):
                lost.append(j)
            if len(l[i]) < len(l[t]):
                t = i
        # print(lost)

        for year in self.dic.keys():
            for i in lost:
                k = list(self.dic[year].keys())
                if i in k:
                    self.dic[year].pop(i)
        # ll = []
        # for year in self.dic.keys():
        #     ll.append(list(self.dic[year].keys()))
        # for i in ll:
        #     print(i)

        # 2020 data complete
        k = list(self.dic[2011].keys())
        self.dic[2020] = {}
        for county in k:
            self.dic[2020][county] = {}
            self.dic[2020][county]['pop'] = (self.dic[2019][county]['pop'] + self.dic[2021][county]['pop']) / 2
            self.dic[2020][county]['median_income'] = (self.dic[2019][county]['median_income'] + self.dic[2021][county]['median_income']) / 2
            self.dic[2020][county]['mean_income'] = (self.dic[2019][county]['mean_income'] + self.dic[2021][county]['mean_income']) / 2

    def main(self):
        self.data_get(self.name1, self.name2, self.sheet_name_pop, 'pop', 'B')
        self.data_get(self.name1, self.name3, self.sheet_name_income, 'median_income', 'F')
        self.data_get(self.name1, self.name3, self.sheet_name_income, 'mean_income', 'H')

        self.data_process()
        print(self.dic)

        with open('./results/pop_and_income.json', 'w') as f:
            json.dump(self.dic, f)

        # with open('./results/pop_and_income.json') as f:
        #     self.dic = json.load(f)
        # self.data_process()


if __name__ == '__main__':
    p = PopAndIncomeProcess()
    p.main()
