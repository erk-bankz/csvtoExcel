__author__ = 'Llorenç Suau Cànaves'
__copyright__ = 'Copyright (c) 2019. All rights are reserved.'
__license__ = 'GPL 3'

import csv
from xlsxwriter import Workbook
from pathlib import Path
import pandas as pd
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QDialog, QApplication
from PyQt5 import QtWidgets

from ui_main_dialog import Ui_Dialog

import openpyxl
import csv


def read_csv(file: str, encoding: str='utf-8-sig') -> list:
    """
    """
    fpath = Path(file)
    if fpath.suffix.lower() == '.csv':
        with open(fpath, newline='', encoding=encoding) as csvfile:
            reader = csv.reader(csvfile, delimiter=";")
            print(reader)
            return [row for row in reader]


def write_excel(file_path: str, rows: list)-> None:
    """
    """
    wb = Workbook(file_path)
    ws = wb.add_worksheet()
    nrow = 1
    for row in rows:
        ws.write_row(f'A{nrow}',  row)
        nrow += 1
    wb.close()


def convert_to_csv(file: str, delimiter):
    fpath = Path(file)
    folder = fpath.parent
    new_csv = Path(folder,f'{fpath.name}.csv')
    xlsx = pd.read_excel(file)
    #xlsx.dropna(how='all', axis='columns')
    xlsx.to_csv(str(new_csv),encoding="utf-8",sep=delimiter, index= False)

# def set_delimiters(file:str):


def convert_to_csv_two(file,delimit):
    wb = openpyxl.load_workbook(file)
    sh = wb.active
    with open(file+".csv", 'w', newline="") as f:
        col = csv.writer(f, delimiter=delimit)
        for row in sh.rows:
            col.writerow([cell.value for cell in row])

# file = r"c:\Users\ehom\Documents\IdeaProjects\Python\Projects\csvToExcel\sample\Excel\ATT_COW_0003405530.CSV.xlsx"
# convert_to_csv_two(file)