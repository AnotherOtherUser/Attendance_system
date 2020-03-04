from os import path, system
from win32com.client import Dispatch
from tkinter import filedialog as fd


# выбор файла с посещаемостью группы
def choose_group_file():
    system('taskkill /IM EXCEL.EXE /F')
    file = fd.askopenfilename()

    global absPath
    absPath = path.abspath(file)

    global excel
    excel = Dispatch("Excel.Application")
    excel.Visible = True

    wb = excel.Workbooks.Open(absPath)
    sheet = excel.ActiveSheet

    system('PlaceCard.exe')
