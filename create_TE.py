from win32com import client
import ctypes
import sys
import tkinter
from tkinter import simpledialog
from word_printer import *


Word_File_Name = "D:\\im\\IMWork\\Шаблон ТЭ.docx"
Word_New_File = "D:\\im\\IMWork\\temp\\СОПС ТЭ0.docx"
PUS = {
    "item_id": "индекс",
    "item_type": "ПУС",
    "item_ref": "A?"
}

def get_item(excel, start_row):
    item = {}
    item["item_id"] = int(excel.Rows(start_row).Cells(2).Value) # item_id
    item["item_type"] = excel.Rows(start_row).Cells(3).Value # item_type
    item["PUS"] = excel.Rows(start_row).Cells(4).Value # PUS
    item["harness_num"] = int(excel.Rows(start_row).Cells(5).Value) # harness_num
    item["num_in_harness"] = int(excel.Rows(start_row).Cells(15).Value) # num_in_harness
    item["item_ref"] = excel.Rows(start_row).Cells(16).Value # item_ref
    item["cable_num"] = int(excel.Rows(start_row).Cells(17).Value) # cable_num
    item["cable_type"] = excel.Rows(start_row).Cells(33).Value # cable_type
    
    return item


def get_cables(excel_app):
    try:
        excel = excel_app.ActiveWorkBook.WorkSheets("Кабели")
    except AttributeError as e:
        MessageBox = ctypes.windll.user32.MessageBoxW
        MessageBox(None, 'В исходной таблице отсутствует лист "Кабели"', 'TE0 Creator', 0)
        sys.exit()

    cable = {}
    cables = {}
    start_row = 2
    current_row = start_row
    while excel.Rows(current_row).Cells(1).Value != None:
        cable = {
            'name': excel.Rows(current_row).Cells(2).Value,
            'cross_sec': excel.Rows(current_row).Cells(3).Value,
            'core1': excel.Rows(current_row).Cells(4).Value,
            'core2': excel.Rows(current_row).Cells(5).Value,
            'core3': excel.Rows(current_row).Cells(6).Value,
            'core4': excel.Rows(current_row).Cells(7).Value,
            'core5': excel.Rows(current_row).Cells(8).Value,
            'core6': excel.Rows(current_row).Cells(9).Value,
            'core7': excel.Rows(current_row).Cells(10).Value       
        }
        cables[excel.Rows(current_row).Cells(1).Value] = cable
        current_row += 1
    return cables


def create_word_doc(items, cables):
    # word_app = client.Dispatch("Word.Application")
    word_app = client.gencache.EnsureDispatch("Word.Application")
    try:
        word = word_app.Documents.Open(Word_File_Name)
    except Exception as e:
        MessageBox = ctypes.windll.user32.MessageBoxW
        MessageBox(None, "Не могу найти файл шаблон: 'Шаблон ТЭ.docx'!!!", 'TE0 Creator', 0)
        sys.exit()
    
    word_app.Visible = False
    word_app.ScreenUpdating = False

    current_row = 3

    for item in items:
        if item.get("num_in_harness") == 1 or item == items[0]:
            prev_item = PUS
            current_item = item
            current_row = print_end_item(word_app, current_row)
        else:
            prev_item = current_item
            current_item = item
            
        
        item_type = item.get("item_type")
        if item_type == 'ИПК' or item_type == 'ИПР' or item_type == 'ИПД' or item_type == 'ИПП' or item_type == 'ИПТ':
            current_row = print_IPK_IPR_IPD_IPP(word_app, current_row, current_item, prev_item, cables)
        elif item_type == 'ИП 216-001Ех':
            current_row = print_IPP_216(word_app, current_row, current_item, prev_item, cables)
        else:
            current_row = print_IPES(word_app, current_row, current_item, prev_item, cables)

    current_row = print_end_item(word_app, current_row)

    word.SaveAs(Word_New_File)
    word_app.Visible = True
    word_app.ScreenUpdating = True

    

def create_table_of_connections():
    
    excel_app = client.Dispatch("Excel.Application")
    # excel_app.Visible = True
    # excel.ScreenUpdating = False
    MessageBox = ctypes.windll.user32.MessageBoxW
    try:
        excel = excel_app.ActiveWorkBook.WorkSheets("КопияИсхСОПС")
    except AttributeError as e:
        MessageBox(None, 'Не открыта таблица Excel с исходными данными! \n(Отсутствует лист "КопияИсхСОПС")', 'TE0 Creator', 0)
        sys.exit()
    
    
    # MessageBox(None, 'Дождитесь всплывающего окна окончания работы скрипта! Вычисляю, вычисляю...', 'TE0 Creator', 0)
    items = []
    start_row = 2
    end_row = 100000

    # User Dialog
    ROOT = tkinter.Tk()
    ROOT.tk.eval(f'tk::PlaceWindow {ROOT._w} center')
    ROOT.withdraw()

    output_range = simpledialog.askstring(
        title='TE0 Creator', 
        prompt="""Введите номера первой и последней строк для вывода в формате: начало-конец.
        Если ничего не вводить, будет выведен весь документ.
        Дождитесь всплывающего окна окончания работы скрипта!""", 
        parent=ROOT) 

    if output_range:
        start_row, end_row = output_range.split('-')

    current_row = int(start_row)
    cable_number = excel.Rows(current_row).Cells(17).Value
    try:
        while cable_number != None and current_row <= int(end_row):
            item = get_item(excel, current_row)
            items.append(item)
            
            current_row += 1
            cable_number = excel.Rows(current_row).Cells(17).Value

    except TypeError as e:
        MessageBox(None, 'Achtung! Что-то не так с данными в таблице Excel!', 'TE0 Creator', 0)
        sys.exit() 
    
    cables = get_cables(excel_app)

    create_word_doc(items, cables)
    
    MessageBox(None, 'Готово! Документ создан: D:\im\IMWork\\temp\СОПС ТЭ0.docx', 'TE0 Creator', 0)

    sys.exit()


# create_table_of_connections()