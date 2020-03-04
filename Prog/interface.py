import datetime, sys
from os import path, system
from tkinter import filedialog as fd
from PyQt5.QtCore import QCoreApplication
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton
from win32com.client import Dispatch
import choose_group, create_group_file, excel_group_parser, excel_rasp_parser


# Регистрация карт
def registration():
    reg_file = fd.askopenfilename()

    abs_path = path.abspath(reg_file)

    reg_excel = Dispatch("Excel.Application")
    reg_excel.Visible = True

    wb = reg_excel.Workbooks.Open(abs_path)
    appl = QApplication(sys.argv)
    sheet = wb.ActiveSheet

    system('PlaceCard.exe')

    def check_text():
        date = datetime.date.today().strftime('%d.%m')
        c = 5

        while sheet.Cells(c, 1).Value:
            c += 1

        for x in range(5, c):
            if int(appl.clipboard().text()[:3] + appl.clipboard().text()[4:]) == (sheet.Cells(x, 2).Value):
                sheet.Cells(x, date).Interior.ColorIndex = 6
                sheet.Cells(x, date).Value = 1
                wb.Save()

    def reg():
        print(appl.clipboard().text())

        reg_excel.ActiveCell.Value = (appl.clipboard().text()[:3] + appl.clipboard().text()[4:])
        wb.Save()
        check_text()

    appl.clipboard().changed.connect(reg)

    print(appl.clipboard().text())


app = QApplication(sys.argv)

# Основное окно
start_window = QWidget()
start_window.resize(300, 250)
start_window.move(500, 200)
start_window.setWindowTitle('Учет посещаемости')

# Функциональные кнопки
group_choose_btn = QPushButton('Выбрать группу', start_window)
group_choose_btn.resize(group_choose_btn.sizeHint())
group_choose_btn.move(30, 10)
group_choose_btn.clicked.connect(choose_group.choose_group_file)

group_reg_btn = QPushButton('Регистрация карт', start_window)
group_reg_btn.resize(group_reg_btn.sizeHint())
group_reg_btn.move(30, 50)
group_reg_btn.clicked.connect(registration)

rasp_add_btn = QPushButton('Добавить расписание', start_window)
rasp_add_btn.resize(rasp_add_btn.sizeHint())
rasp_add_btn.move(30, 90)
rasp_add_btn.clicked.connect(excel_rasp_parser.rasp_excel_reader)

group_add_btn = QPushButton('Добавить группу', start_window)
group_add_btn.resize(group_add_btn.sizeHint())
group_add_btn.move(30, 130)
group_add_btn.clicked.connect(excel_group_parser.group_excel_reader)

create_group_file_btn = QPushButton('Создать журнал посещений', start_window)
create_group_file_btn.resize(create_group_file_btn.sizeHint())
create_group_file_btn.move(30, 170)
create_group_file_btn.clicked.connect(create_group_file.creation_group_file)

exit_btn = QPushButton('Выйти', start_window)
exit_btn.resize(exit_btn.sizeHint())
exit_btn.move(200, 190)
exit_btn.clicked.connect(QCoreApplication.instance().quit)

start_window.show()
sys.exit(app.exec_())
system('taskkill /IM EXCEL.EXE /F')
