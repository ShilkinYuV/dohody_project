import openpyxl
from openpyxl import Workbook

class Read:

    def __init__(self):
        self.filename = ''
        self.total_received_by_account_40101_03100 = ''
        self.refund_of_overpaid_amounts = ''
        self.total_transferred_to_the_budget = ''
        self.consolidated_budget = ''

    def read_excel(self, arg):
        self.filename = arg
        wb = openpyxl.load_workbook(self.filename)
        sheet = wb.get_sheet_names()[2]
        this_sheet = wb[sheet]
        i = 0
        for cell in this_sheet['C']:
            i = i+1
            if cell.value == 'Всего по разделам I и II':
                self.total_received_by_account_40101_03100 = this_sheet['D' + str(i)].value
                print(self.total_received_by_account_40101_03100)

                self.refund_of_overpaid_amounts = this_sheet['F' + str(i)].value
                print(self.refund_of_overpaid_amounts)

                self.total_transferred_to_the_budget = this_sheet['H' + str(i)].value
                print(self.total_transferred_to_the_budget)

                self.consolidated_budget = float(this_sheet['J' + str(i)].value) + float(this_sheet['N' + str(i)].value) + float(this_sheet['L' + str(i)].value)
                print(self.consolidated_budget)

                print(this_sheet['J' + str(i)].value)
                break


