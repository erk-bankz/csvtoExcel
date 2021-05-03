__author__ = 'Llorenç Suau Cànaves'
__copyright__ = 'Copyright (c) 2019. All rights are reserved.'
__license__ = 'GPL 3'

from pathlib import Path
import modules as csv2ex  # calls/imports the other py file called "csv_to_excel"
from argparse import ArgumentParser
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QDialog, QApplication
from PyQt5 import QtWidgets

from ui_main_dialog import Ui_Dialog
import openpyxl
import docx
import os
import re
import modules
import sys
import pandas


class Dialog(QDialog, Ui_Dialog):
    """
    Converts CSV to Excel and Excel to CSV with user choice of delimiter
    """
    def __init__(self, parent=None):
        super(Dialog, self).__init__(parent)
        self.setupUi(self)
        self.buttonBox.clicked.connect(self.buttonClicked)  # alternative way to call your method
    # @pyqtSlot()

    def buttonClicked(self):
        #print(self.lineEdit.text())
        if self.lineEdit.text() == "":
            QtWidgets.QMessageBox.critical(self, "Error", "Please enter correct path")
            sys.exit()
        pp= Path(self.lineEdit.text())
        if pp.is_dir():
            try:
                if self.radioButton.isChecked():
                    files = []
                    files = [file for file in pp.glob('**/*.*') if file.suffix.lower() == ".csv"]
                    for file in files:
                        folder = file.parent
                        excel_file = Path(folder,f'{file.name}.xlsx')
                        rows = modules.read_csv(str(file), encoding="utf-8")
                        for row in rows:
                            for column in row:
                                length = column.count('\n')
                                if length > 253:
                                    QtWidgets.QMessageBox.critical(self, "Error", "Cell contains over 253 linebreaks, manual prep is needed")
                                    sys.exit()
                        modules.write_excel((str(excel_file)),rows)
                if self.radioButton_2.isChecked():
                    files = []
                    files = [file for file in pp.glob('**/*.*') if file.suffix.lower() == ".xlsx"]
                    delimiter = ","
                    if self.radioButton_3.isChecked():
                        delimiter = "\t"
                    if self.radioButton_4.isChecked():
                        delimiter = "|"
                    if self.radioButton_5.isChecked():
                        delimiter = ";"
                    for file in files:
                        modules.convert_to_csv(str(file), delimiter)
                QtWidgets.QMessageBox.information(self, "Macro finished", "Files has finished processing")
            except:
                e= sys.exc_info()
                QtWidgets.QMessageBox.critical(Dialog(), "Error", str(e[1]))



#flag if cell exceeds character limit
if __name__ == "__main__":
    import sys
    application = QApplication(sys.argv)
    macro_dialog = Dialog() # create object of dialog, **use the name of your class (ie class Dialog)**
    macro_dialog.show()
    sys.exit(application.exec_())
