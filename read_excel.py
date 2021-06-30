import openpyxl
from openpyxl import Workbook
import re


class Read:

    def __init__(self):
        self.filename = ''
        self.total_received_by_account_40101_03100 = ''
        self.refund_of_overpaid_amounts = ''
        self.total_transferred_to_the_budget = ''
        self.consolidated_budget = ''
        self.article_i_federal_budget_including = ''
        self.vat_on_goods_sold_on_the_territory_of_the_Russian_Federation = 0

    def read_excel(self, arg):
        self.filename = arg
        wb = openpyxl.load_workbook(self.filename)
        sheet_two = wb.get_sheet_names()[2]
        sheet_one = wb.get_sheet_names()[1]
        sheet_three = wb.get_sheet_names()[3]
        this_sheet = wb[sheet_two]
        i = 0
        for cell in this_sheet['C']:
            i = i + 1
            if cell.value == 'Всего по разделам I и II':
                self.total_received_by_account_40101_03100 = this_sheet['D' + str(i)].value

                self.refund_of_overpaid_amounts = this_sheet['F' + str(i)].value

                self.total_transferred_to_the_budget = this_sheet['H' + str(i)].value

                self.consolidated_budget = float(this_sheet['J' + str(i)].value) + float(
                    this_sheet['N' + str(i)].value) + float(this_sheet['L' + str(i)].value)

                self.article_i_federal_budget_including = this_sheet['J' + str(i)].value
                break

        this_sheet = wb[sheet_one]
        i = 0
        for cell in this_sheet['B']:
            i = i + 1
            result = re.match(r'10301', str(cell.value))
            if result:
                self.vat_on_goods_sold_on_the_territory_of_the_Russian_Federation = self.vat_on_goods_sold_on_the_territory_of_the_Russian_Federation + float(
                    this_sheet['J' + str(i)].value)

        this_sheet = wb[sheet_three]
        i = 0
        for cell in this_sheet['B']:
            i = i + 1
            result = re.match(r'10301', str(cell.value))
            if result:
                self.vat_on_goods_sold_on_the_territory_of_the_Russian_Federation = self.vat_on_goods_sold_on_the_territory_of_the_Russian_Federation + float(
                    this_sheet['F' + str(i)].value)

        self.total_received_by_account_40101_03100 = round(self.total_received_by_account_40101_03100/1000000, 2)
        print(self.total_received_by_account_40101_03100)

        self.refund_of_overpaid_amounts = round(self.refund_of_overpaid_amounts/1000000, 2)
        print(self.refund_of_overpaid_amounts)

        self.total_transferred_to_the_budget = round(self.total_transferred_to_the_budget/1000000, 2)
        print(self.total_transferred_to_the_budget)

        self.consolidated_budget = round(self.consolidated_budget/1000000, 2)
        print(self.consolidated_budget)

        self.article_i_federal_budget_including = round(self.article_i_federal_budget_including/1000000, 2)
        print(self.article_i_federal_budget_including)

        self.vat_on_goods_sold_on_the_territory_of_the_Russian_Federation = round(self.vat_on_goods_sold_on_the_territory_of_the_Russian_Federation/1000000, 2)
        print(self.vat_on_goods_sold_on_the_territory_of_the_Russian_Federation)
