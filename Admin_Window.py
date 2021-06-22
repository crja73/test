import sys
import re
from datetime import timedelta, datetime
from docxtpl import DocxTemplate # редактирование шаблона *.docx
import sqlite3
from PySide6 import QtCore, QtWidgets, QtGui
from config import *
import os
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog, QMainWindow
import subprocess


class MyWidget(QtWidgets.QWidget, QtGui.QMainWindow):
    def __init__(self):
        super().__init__()
        ####### Формирование основного окна:

        ####### название основного окна
        self.setWindowTitle("Satis Expert")

        ####### логотип основного окна
        self.setWindowIcon(QtGui.QIcon("logo_company.png"))

        ####### компоновка основного окна и размеров
        self.centralwidget = QtWidgets.QWidget()
        self.centralwidget.setObjectName("centralwidget")
        self.tableWidget = QtWidgets.QTableWidget(self.centralwidget)
        self.tableWidget.setGeometry(QtCore.QRect(0, 500, 1878, 478)) # 1-левый край горизонт, 2-вертикаль от верха, 3-ширина, 4-высота
        self.tableWidget.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(31)

        ####### дерево фильтра
        self.treeView = QtWidgets.QTreeView(self.centralwidget)
        self.treeView.setGeometry(QtCore.QRect(0, 0, 260, 500))
        self.treeView.setObjectName("treeView")

        ####### отображать скролл
        self.tableWidget.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOn)
        self.tableWidget.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOn)

        ####### воткнуть таблицу
        self.layout = QtWidgets.QVBoxLayout()
        self.layout.addWidget(self.centralwidget)
        self.setLayout(self.layout)

        ####### Выделять всю строку
        self.tableWidget.setSelectionBehavior(QtWidgets.QTableWidget.SelectRows)

        ####### Кнопки
        self.btn_protokol = QtWidgets.QPushButton("Подготовить протокол", self)
        self.btn_protokol.setFixedWidth(200)
        self.btn_protokol.move(280, 11)

        self.btn_company = QtWidgets.QPushButton("Добавить компанию", self)
        self.btn_company.setFixedWidth(200)
        self.btn_company.move(280, 50)

        self.btn_pdf = QtWidgets.QPushButton("Прикрепить pdf", self)
        self.btn_pdf.setFixedWidth(200)
        self.btn_pdf.move(280, 89)

        ####### действие при нажатии кнопок
        self.btn_protokol.clicked.connect(self.Dannie)
        self.btn_company.clicked.connect(self.addPlochadka)
        self.btn_pdf.clicked.connect(self.add_pdf)

        ####### переменная по названию компании
        short_name_company = "clients" # неправильная сортировка

        ####### подключиться к sqlite
        conn = sqlite3.connect(Adress_DB)
        cur = conn.cursor()
        rowcount = cur.execute('''SELECT COUNT(*) FROM clients''').fetchone()[0]
        self.tableWidget.setRowCount(rowcount)
        cur.execute(f"SELECT * FROM '{short_name_company}' WHERE ROWID IS NOT ''") # неправильная сортировка

        ####### перебрать и выровнять по центру ячейки, вставить их
        for row, form in enumerate(cur):
            for column, item in enumerate(form):
                item = QtWidgets.QTableWidgetItem(str(item))
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                self.tableWidget.setItem(row, column, item)

                ####### переименовка столбцов
                self.tableWidget.setHorizontalHeaderLabels(['Название компании:',
                                                            'Номер заявки:',
                                                            'Номер протокола:',
                                                            'Лаборант:',
                                                            'Площадка:',
                                                            'Ответственное лицо:',
                                                            'Вид работ:',
                                                            'Дата бет./лнк:',
                                                            'Дата испытания:',
                                                            'Класс бетона/Ку:',
                                                            'Наименование конструкции:',
                                                            'Кол-во, объём:',
                                                            'Сутки испытаний:',
                                                            '№ акта отбора проб:',
                                                            '№ проекта:',
                                                            'Завод БСГ:',
                                                            'Проектная прочность:',
                                                            'Примечания:',
                                                            'Протокол:',
                                                            'Фото',
                                                            'ПТО:',
                                                            'ЛАБ:',
                                                            'П_3',
                                                            'П_4',
                                                            'П_5',
                                                            'П_6',
                                                            'П_7',
                                                            'П_8',
                                                            'П_9',
                                                            'П_10',
                                                            'Протокол_2:'])
                self.tableWidget.verticalHeader().hide()
        ####### запрос списка для дерева фильтра


        ####### закрытие соединения с SQLite
        conn.commit()
        conn.close()

        ####### ширина столбца согласно содержимому
        self.tableWidget.resizeColumnsToContents()

        ####### фиксированная ширина столбца
        header = self.tableWidget.horizontalHeader()
        header.resizeSection(17, 250)
        header.resizeSection(10, 300)

        ####### перенос столбца "Ключ" в первый столбец
        #header.moveSection(16, 0)


    ####### получение данных выделенной строки
    def Dannie(self):

        ####### переменная по названию компании
        short_name_company = "CRCC_RUS"

        ####### получение ID выбранной строки
        rows = sorted(set(index.row() for index in
                          self.tableWidget.selectedIndexes()))
        for row in rows:
            print('Строка %d считана' % row)

        ####### получение данных строки
        a1 = self.tableWidget.item(row, 0).text()    # номер заявки
        a2 = self.tableWidget.item(row, 1).text()    # номер протокола
        a3 = self.tableWidget.item(row, 2).text()    # ФИО лаборанта
        a4 = self.tableWidget.item(row, 3).text()    # короткое название площадки
        a5 = self.tableWidget.item(row, 4).text()    # должность/ФИО ПТО'шника
        a6 = self.tableWidget.item(row, 5).text()    # короткое название работ
        a7 = self.tableWidget.item(row, 6).text()    #(не приемлема->a23) дата бетонирования
        a8 = self.tableWidget.item(row, 7).text()    # короткая характеристика материала
        a9 = self.tableWidget.item(row, 8).text()    # назавние, оси, отметки конструкции
        a10 = self.tableWidget.item(row, 9).text()   # лабораторный объём выполненных работ
        a11 = self.tableWidget.item(row, 10).text()  # сутки испытаний
        a12 = self.tableWidget.item(row, 11).text()  # акт отбора проб
        a13 = self.tableWidget.item(row, 12).text()  # номер проектной документации
        a14 = self.tableWidget.item(row, 13).text()  # название бетонного завода
        a15 = self.tableWidget.item(row, 14).text()  # проектная прочность по ДКБС
        a16 = self.tableWidget.item(row, 15).text()  # примечания
        a17 = self.tableWidget.item(row, 16).text()  # КЛЮЧ
        a18 = self.tableWidget.item(row, 17).text()  # отметка от ПТО'шника о готовности
        a19 = self.tableWidget.item(row, 18).text()  # отметка лаборанта о готовности протокола
        a20 = self.tableWidget.item(row, 19).text()  # прикреплённый PDF

        ####### данные из sql листов
        conn = sqlite3.connect(r'//192.168.1.12/Scan/SqLite/lab.db')
        cur = conn.cursor()

        ####### получение полного названия компании
        cur.execute(f"SELECT * FROM Company WHERE short_name='{short_name_company}'")
        full_name_company = cur.fetchone()[2]

        ####### получение полного названия площадки
        short_name_Plochadka = short_name_company + "_" + a4
        cur.execute(f"SELECT * FROM Plochadka WHERE short_name='{short_name_Plochadka}'")
        full_name_Plochadka = cur.fetchone()[2]

        a21 = full_name_company                       # полное название компании + договор
        a22 = full_name_Plochadka                     # полное название площадки
        conn.commit()
        conn.close()

        ####### отделение цифр от текста (объём работ)
        a10 = re.sub('\D', '', a10)                   # общий объём работ

        ####### определение суток испытания
        day_list = ['«01»', '«02»', '«03»', '«04»', '«05»', '«06»', '«07»', '«08»', '«09»', '«10»', '«11»', '«12»',
                    '«13»', '«14»', '«15»', '«16»', '«17»', '«18»', '«19»', '«20»', '«21»', '«22»', '«23»', '«24»',
                    '«25»', '«26»', '«27»', '«28»', '«29»', '«30»', '«31»']
        month_list = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября',
                      'ноября', 'декабря']

        date_list = a7.split('.')
        full_Day = datetime(int(date_list[2]), int(date_list[1]), int(date_list[0]))
        full_Day = full_Day.strftime("%d.%m.%Y")
        Day_7 = datetime(int(date_list[2]), int(date_list[1]), int(date_list[0])) + timedelta(7)
        Day_7 = Day_7.strftime("%d.%m.%Y")
        Day_28 = datetime(int(date_list[2]), int(date_list[1]), int(date_list[0])) + timedelta(28)
        Day_28 = Day_28.strftime("%d.%m.%Y")

        date_list = full_Day.split('.')
        a23 = (day_list[int(date_list[0]) - 1] + ' ' + month_list[int(date_list[1]) - 1] + ' ' + date_list[2])
        date_list_7 = Day_7.split('.')
        a24 = (day_list[int(date_list_7[0]) -1] + ' ' + month_list[int(date_list_7[1]) - 1] + ' ' + date_list_7[2])
        date_list_28 = Day_28.split('.')
        a25 = (day_list[int(date_list_28[0]) -1] + ' ' + month_list[int(date_list_28[1]) - 1] + ' ' + date_list_28[2])

        # a23 дата бетонирования/производства работ
        # a24 дата + 7 суток
        # a25 дата + 28 суток

        ####### подготовка данных для протокола
        doc = DocxTemplate(r'//192.168.1.12/Scan/SqLite/Шаблон_кубы_100.docx')
        context = {'a2': a2,
                   'a21': a21,
                   'a22': a22,
                   'a9': a9,
                   'a14': a14,
                   'a7': a7,
                   'a23': a23,
                   'a24': a24}
        doc.render(context)
        doc.save(r'//192.168.1.12/Scan/SqLite/' + a2 + 'К7.docx')

    def addCompany(self):
        conn = sqlite3.connect(r'//192.168.1.12/Scan/SqLite/lab.db')
        cur = conn.cursor()
        #cur.execute('''INSERT INTO Company
        #                (ID, short_name, full_name) VALUES
        #                (2, "CRCC_RUS_2", "ООО «СиАрСиСи Рус», договор № 11/19")''')
        conn.commit()
        conn.close()
        print("Компания добавлена")

    def addPlochadka(self):
        conn = sqlite3.connect(r'//192.168.1.12/Scan/SqLite/lab.db')
        cur = conn.cursor()
        cur.execute("""CREATE TABLE IF NOT EXISTS Plochadka (ID INT PRIMARY KEY, short_name TEXT, full_name TEXT)""")
        #cur.execute('''INSERT INTO Plochadka
        #                (ID, short_name, full_name) VALUES
        #                (1., "CRCC_RUS_...", "Юго-Западный ...")''')
        conn.commit()
        conn.close()
        print("Площадка добавлена")


    def add_pdf(self):
        path = "C:/Users"
        path = os.path.realpath(path)
        os.startfile(path)

    def showDialog(self):  
        subprocess.Popen('explorer ""') 

if __name__ == "__main__":
    app = QtWidgets.QApplication([])
    app.setStyle('Fusion')
    widget = MyWidget()
    widget.setFixedSize(1900, 1000)
    widget.show()
    sys.exit(app.exec_())