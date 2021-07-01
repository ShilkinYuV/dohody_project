import openpyxl
from PyQt5.QtCore import QThread, pyqtSignal
from openpyxl import Workbook
import re


class SaveExcel(QThread):

    def __init__(self):
        super(SaveExcel, self).__init__()
        self.result_one = {}
        self.result_two = {}

    def save_excel(self, result_one, result_two):
        self.result_one = result_one
        self.result_two = result_two
        wb = openpyxl.load_workbook('maket.xlsx')
        sheet = wb.get_sheet_names()[0]
        ws = wb[sheet]
        ws['B6'] = result_one['total_received_by_account_40101_03100']
        ws['C6'] = result_one['total_received_by_account_40101_03100']
        print(result_one['total_received_by_account_40101_03100'])
        print(result_two['total_received_by_account_40101_03100'])
