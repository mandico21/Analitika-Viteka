import json

import openpyxl
from fuzzywuzzy import fuzz


class Analitica:

    def __init__(self, file: str, pf_sheet, pf_wb, tk: str):
        self.af_wb = openpyxl.load_workbook(file)
        self.af_sheet = self.af_wb.worksheets[0]
        self.tk = tk
        self.pf_sheet = pf_sheet
        self.pf_wb = pf_wb

    def sheet_paser(self, row: int, columns: int):
        with open('src/json/data.json', 'r', encoding='utf-8') as json_file:
            data = json.load(json_file)
        data_tk = data[self.tk]['tk']
        data_shb = data[self.tk]['shb']
        rows = row + int(data_tk["row_app"])

        self.pf_sheet[data_shb["convert_as"] + str(rows)] = self.af_sheet[
            str(data_tk["convert"]) + str(columns + int(data_tk["row_2"]))].value

        self.pf_sheet[str(data_shb["minimum_1_as"]) + str(rows)] = self.af_sheet[
            str(data_tk["minimum_1"]) + str(columns + int(data_tk["row_3"]))].value

        self.pf_sheet[str(data_shb["minimum_2_as"]) + str(rows)] = self.af_sheet[
            str(data_tk["minimum_2"]) + str(int(columns) + int(data_tk["row_4"]))].value

        self.pf_sheet[str(data_shb["objem_as"]) + str(rows)] = self.af_sheet[
            str(data_tk["objem"]) + str(int(columns) + int(data_tk["row_5"]))].value

        self.pf_sheet[str(data_shb["ves_100_as"]) + str(rows)] = self.af_sheet[
            str(data_tk["ves_100"]) + str(int(columns) + int(data_tk["row_6"]))].value

        self.pf_sheet[str(data_shb["ves_3000_as"]) + str(rows)] = self.af_sheet[
            str(data_tk["ves_3000"]) + str(int(columns) + int(data_tk["row_7"]))].value

    def run(self):
        with open('src/json/city.json', 'r', encoding='utf-8') as json_file:
            city = json.load(json_file)
        with open('src/json/data.json', 'r', encoding='utf-8') as json_file:
            data = json.load(json_file)

        for my_city in city.keys():
            for i in range(1, self.af_sheet.max_row + 1):
                city_tk = self.af_sheet[data[self.tk]['tk']['city'] + str(i)].value
                check_site = fuzz.WRatio(my_city.lower(), city_tk)
                if check_site >= 95:
                    if self.tk == 'Энергия':
                        if city_tk == 'Владивосток':
                            if self.af_sheet[f'B{i}'].value != 'Авто':
                                i += 1
                    self.sheet_paser(city[my_city], i)
