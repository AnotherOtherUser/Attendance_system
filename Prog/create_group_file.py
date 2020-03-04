from os import system, path
from win32com.client import Dispatch
import win32com
from pathlib import Path
from tkinter import filedialog as fd, messagebox as mb
import sqlite3
import fill_group_file


def creation_group_file():
    conn = sqlite3.connect('BD.db')
    cursor = conn.cursor()

    # Прерывание всех процессов Micrisoft Excel
    system('taskkill /IM EXCEL.EXE /F')

    # Создание файла с посещаемостью группы
    file_path = fd.asksaveasfilename(defaultextension='.xls')

    system('copy Shape.xls {}'.format(path.abspath(file_path)))

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = 0

    wb = excel.Workbooks.Open(file_path)
    sheet = wb.ActiveSheet

    group_num = Path(file_path).stem

    sheet.Cells(1, 3).Value = group_num

    students_bd_list = cursor.execute("""SELECT * FROM СТУДЕНТ WHERE ГРУППА==?""", (group_num, )).fetchall()
    i = 1

    # Заполнения столбца ФИО студента в таблице с посещаемостью
    for student in students_bd_list:
        sheet.Cells(i + 6, 3).Value = student[0] + ' ' + student[1] + ' ' + student[2]
        i += 1
    wb.Save()

    mb.askokcancel('Создание файла посещений', 'Создание файла посещений группы {group_num} завершено.')

    fill_group_file.fill_file_group()

    system('taskkill /IM EXCEL.EXE /F')

creation_group_file()