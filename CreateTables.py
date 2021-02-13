import os
import xlrd
import openpyxl
from tkinter import *

if not os.path.exists('Exel_docs'):
    os.makedirs('Exel_docs')
if not os.path.exists('Result'):
    os.makedirs('Result')


def get_column(file_name):
    rb = xlrd.open_workbook('Exel_docs\\' + file_name, formatting_info=True)
    sheets = rb.sheet_by_index(0)
    nums = []

    for rownum in range(7, sheets.nrows):
        row = sheets.row_values(rownum)
        nums.append(row[-1])
    return nums


def get_name_column(all_cols, first_file):
    rb = xlrd.open_workbook('Exel_docs\\' + first_file, formatting_info=True)
    sheets = rb.sheet_by_index(0)
    nums = []

    for rownum in range(7, sheets.nrows):
        row = sheets.row_values(rownum)
        nums.append(row[2])
    all_cols.insert(0, nums)


def create_file():
    files = os.listdir('Exel_docs')
    if len(files) == 0:
        info['text'] = 'Папка пуста'
        return
    all_cols = [get_column(i) for i in files]

    get_name_column(all_cols, files[0])
    book = openpyxl.Workbook()
    sheet = book.active
    for i, col in enumerate(all_cols):
        for j, row in enumerate(col):
            sheet.cell(row=j + 1, column=i + 1).value = row
    sheet.column_dimensions['A'].width = 95
    book.save('Result\\результат.xlsx')
    book.close()
    info['text'] = 'Отчет создан'


def open_file():
    pat = os.getcwd()
    os.startfile(f'{pat}\\Result\\результат.xlsx')


def docs_path():
    pat = os.getcwd()
    os.startfile(f'{pat}\\Exel_docs')


def docs_remove():
    docs = os.listdir('Exel_docs')
    for i in docs:
        os.remove(f'Exel_docs\\{i}')
    info['text'] = 'Файлы удалены'


def close():
    return root.quit()


root = Tk()
root.title('Помощник Бухгалтера')
root.geometry('500x150')
root.resizable(width=False, height=False)

create_doc = Button(text='Сформировать документ', command=create_file)
create_doc.grid(row=1, column=1, sticky=W + E, pady=25, padx=5)
open_res = Button(text='Открыть документ', command=open_file)
open_res.grid(row=1, column=2, sticky=W, pady=25, padx=5)
path_docs = Button(text='Папка для добавления файлов', command=docs_path)
path_docs.grid(row=1, column=3, sticky=E, pady=25, padx=5)
remove_docs = Button(text='Удалить документы из папки', command=docs_remove)
remove_docs.grid(row=4, column=1, sticky=E, pady=15, padx=5)
info = Label(text='', font=20)
info.grid(row=2, column=2, sticky=E + W)
close_tkinter = Button(text='Закрыть', command=close)
close_tkinter.grid(row=4, column=3, sticky=E, pady=15, padx=5)

root.mainloop()
