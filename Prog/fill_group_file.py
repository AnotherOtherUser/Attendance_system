from os import system, path
from win32com.client import Dispatch
import win32com
from pathlib import Path
from tkinter import filedialog as fd, messagebox as mb
import sqlite3


def fill_file_group():
    conn = sqlite3.connect('BD.db')
    cursor = conn.cursor()

    # Прерывание всех процессов Micrisoft Excel
    system('taskkill /IM EXCEL.EXE /F')

    mb.askokcancel('Журнал посещений', 'Выберите журнал посещений для загрузки расписания группы.')

    open_group_file = fd.askopenfilename()

    group_num = Path(open_group_file).stem

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = 0

    wb = excel.Workbooks.Open(open_group_file)
    sheet = wb.ActiveSheet

    rasp_bd_list = cursor.execute("""SELECT НЕДЕЛЯ,ДАТА,НАЗВАНИЕ,ТИП,ВРЕМЯ FROM ДИСЦИПЛИНА WHERE ГРУППА==?""",
                                  (group_num.replace('_', '-'),)).fetchall()
    i = 1

    name_discip = cursor.execute(
        """SELECT НАЗВАНИЕ FROM ДИСЦИПЛИНА WHERE ГРУППА=='ИСТ-731' GROUP BY ГРУППА""").fetchall()

    sheet.Cells(1, 12).Value = name_discip

    for dis in rasp_bd_list:
        # номер учебной недели
        sheet.Cells(3, i + 3).Value = dis[0]

        # Дата проведения занятия
        sheet.Cells(4, i + 3).Value = dis[1]

        # Тип занятия (практика, лабораторная работа)
        if dis[3] == 'Практические занятия':
            sheet.Cells(5, i + 3).Value = 'пр'
        elif dis[3] == 'Лабораторная работа':
            sheet.Cells(5, i + 3).Value = 'лр'

        # Время проведения занятия (пара)
        sheet.Cells(6, i + 3).Value = dis[4]

        i += 1
    wb.Save()

    mb.askokcancel('Журнал посещений', 'Создание журнала посещений завершено.')
    system('taskkill /IM EXCEL.EXE /F')
