import json
import os

import openpyxl
from PyQt5 import QtCore
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog

from src.analitica import Analitica
from src.file import Ui_Analiz


class Excel_file(QtCore.QThread):
    sgn = QtCore.pyqtSignal(str)

    def __init__(self, mainwindow, parent=None):
        super().__init__()
        self.mainwindow = mainwindow

    def run(self):
        with open('src/json/data.json', 'r', encoding='utf-8') as json_file:
            data = json.load(json_file)
        pf_wb = openpyxl.load_workbook(data[' ']['path_pattern'])
        pf_sheet = pf_wb.active
        for tk in data:
            if data[tk]['tk']['check'] is True:
                res = Analitica(data[tk]['path'], pf_sheet, pf_wb, tk)
                res.run()
                self.sgn.emit(f"{tk} - готов!")
        pf_wb.save(data[' ']['path_pattern'])


class Analitics(QMainWindow):

    def __init__(self):
        super(Analitics, self).__init__()
        self.thread = None
        self.ui = Ui_Analiz()
        self.ui.setupUi(self)
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('Анализ городов')
        # app.setStyle("Fusion")

        self.ui.pushButton_2.clicked.connect(self.on_clicked_save)
        self.ui.pushButton_4.clicked.connect(self.on_clicked_city_list)
        self.ui.pushButton_3.clicked.connect(self.on_clicked_add_city)
        self.ui.pushButton_6.clicked.connect(self.on_clicked_path)
        self.ui.pushButton_5.clicked.connect(self.on_clicked_path_pattern)
        self.ui.pushButton.clicked.connect(self.on_run_script)

        with open('src/json/data.json', 'r', encoding='utf-8') as json_file:
            data = json.load(json_file)

        self.ui.comboBox.addItems(data.keys())

        self.ui.comboBox.currentTextChanged.connect(self.on_combox)

    def on_combox(self, text):
        with open('src/json/data.json', 'r', encoding='utf-8') as json_file:
            data = json.load(json_file)
        # TK
        self.ui.lineEdit.setText(str(data[text]['tk']['city']))
        self.ui.lineEdit_2.setText(str(data[text]['tk']['convert']))
        self.ui.lineEdit_3.setText(str(data[text]['tk']['minimum_1']))
        self.ui.lineEdit_4.setText(str(data[text]['tk']['minimum_2']))
        self.ui.lineEdit_5.setText(str(data[text]['tk']['objem']))
        self.ui.lineEdit_6.setText(str(data[text]['tk']['ves_100']))
        self.ui.lineEdit_7.setText(str(data[text]['tk']['ves_3000']))
        self.ui.spinBox_15.setValue(data[text]['tk']['row_app'])
        self.ui.spinBox.setValue(data[text]['tk']['row_1'])
        self.ui.spinBox_2.setValue(data[text]['tk']['row_2'])
        self.ui.spinBox_3.setValue(data[text]['tk']['row_3'])
        self.ui.spinBox_4.setValue(data[text]['tk']['row_4'])
        self.ui.spinBox_5.setValue(data[text]['tk']['row_5'])
        self.ui.spinBox_6.setValue(data[text]['tk']['row_6'])
        self.ui.spinBox_7.setValue(data[text]['tk']['row_7'])
        self.ui.checkBox.setChecked(data[text]['tk']["check"])
        # SHB
        self.ui.lineEdit_16.setText(str(data[text]['shb']['convert_as']))
        self.ui.lineEdit_17.setText(str(data[text]['shb']['minimum_1_as']))
        self.ui.lineEdit_18.setText(str(data[text]['shb']['minimum_2_as']))
        self.ui.lineEdit_19.setText(str(data[text]['shb']['objem_as']))
        self.ui.lineEdit_20.setText(str(data[text]['shb']['ves_100_as']))
        self.ui.lineEdit_21.setText(str(data[text]['shb']['ves_3000_as']))

    def on_clicked_save(self):  # sourcery no-metrics skip
        with open('src/json/data.json', 'r', encoding='utf-8') as json_file:
            data = json.load(json_file)

        tk = self.ui.comboBox.currentText()

        city = self.ui.lineEdit.text()  # Город
        convert = self.ui.lineEdit_2.text()  # Конверт
        minimum_1 = self.ui.lineEdit_3.text()  # Минималка_1
        minimum_2 = self.ui.lineEdit_4.text()  # Минималка_2
        objem = self.ui.lineEdit_5.text()  # Объем
        ves_100 = self.ui.lineEdit_6.text()  # Вес до 100 кг
        ves_3000 = self.ui.lineEdit_7.text()  # Вес от 3000 кг
        row_app = self.ui.spinBox_15.text()  # добавление к строке Общее
        row_1 = self.ui.spinBox.text()  # добавление к строке
        row_2 = self.ui.spinBox_2.text()  # добавление к строке
        row_3 = self.ui.spinBox_3.text()  # добавление к строке
        row_4 = self.ui.spinBox_4.text()  # добавление к строке
        row_5 = self.ui.spinBox_5.text()  # добавление к строке
        row_6 = self.ui.spinBox_6.text()  # добавление к строке
        row_7 = self.ui.spinBox_7.text()  # добавление к строке
        check = self.ui.checkBox.isChecked()  # Чекбокс
        # ** Шаблок
        convert_as = self.ui.lineEdit_16.text()  # Конверт
        minimum_1_as = self.ui.lineEdit_17.text()  # Минималка_1
        minimum_2_as = self.ui.lineEdit_18.text()  # Минималка_2
        objem_as = self.ui.lineEdit_19.text()  # Объем
        ves_100_as = self.ui.lineEdit_20.text()  # Вес 100 кг`
        ves_3000_as = self.ui.lineEdit_21.text()  # Вес 3000 кг
        if tk != ' ':
            data[tk]['tk'] = {
                'city': city,
                'convert': convert,
                'minimum_1': minimum_1,
                'minimum_2': minimum_2,
                'objem': objem,
                'ves_100': ves_100,
                'ves_3000': ves_3000,
                'row_app': int(row_app),
                'row_1': int(row_1),
                'row_2': int(row_2),
                'row_3': int(row_3),
                'row_4': int(row_4),
                'row_5': int(row_5),
                'row_6': int(row_6),
                'row_7': int(row_7),
                'check': check}
            data[tk]['shb'] = {
                'convert_as': convert_as,
                'minimum_1_as': minimum_1_as,
                'minimum_2_as': minimum_2_as,
                'objem_as': objem_as,
                'ves_100_as': ves_100_as,
                'ves_3000_as': ves_3000_as,
            }

            with open('src/json/data.json', 'w', encoding='utf-8') as outfile:
                json.dump(data, outfile, ensure_ascii=False, indent=2)

    def on_clicked_city_list(self):
        os.startfile(r'src\json\city.json')

    def on_clicked_add_city(self):
        os.startfile(r'src\json\data.json')

    def on_clicked_path(self):
        tk = self.ui.comboBox.currentText()
        if tk == ' ':
            return
        with open('src/json/data.json', 'r', encoding='utf-8') as json_file:
            data = json.load(json_file)

        wb_patch = QFileDialog.getOpenFileName(directory=data[tk]['path'], filter='Model file (*.xlsx)')[0]

        if wb_patch:
            data[tk]['path'] = wb_patch
            with open('src/json/data.json', 'w', encoding='utf-8') as outfile:
                json.dump(data, outfile, ensure_ascii=False, indent=2)

    def on_clicked_path_pattern(self):
        with open('src/json/data.json', 'r', encoding='utf-8') as json_file:
            data = json.load(json_file)

        wb_patch = QFileDialog.getOpenFileName(directory=data[' ']['path_pattern'], filter='Model file (*.xlsx)')[0]
        if wb_patch:
            data[' ']['path_pattern'] = wb_patch
            with open('src/json/data.json', 'w', encoding='utf-8') as outfile:
                json.dump(data, outfile, ensure_ascii=False, indent=2)

    def on_run_script(self):
        self.ui.statusbar.showMessage("Скрипт запущен", 3000)
        self.thread = Excel_file(self)
        self.thread.start()
        self.thread.sgn.connect(self.getMsg)

    def getMsg(self, msg):
        self.ui.statusbar.showMessage(msg, 3000)


if __name__ == '__main__':
    app = QApplication([])
    applications = Analitics()
    applications.show()
    app.exec_()
