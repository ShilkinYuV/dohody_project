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
        self.vat_on_goods_imported_into_the_territory_of_the_Russian_Federation = 0
        self.income_tax = 0
        self.article_II_consolidated_regional_budget = 0
        self.regional_budgets = ''
        self.regional_budgets_NDFL = 0
        self.regional_budgets_land_tax_from_organizations = 0
        self.local_budgets = ''
        self.local_budgets_NDFL = 0
        self.local_budgets_land_tax_from_organizations = 0
        self.local_budgets_comprehensive_income_taxes = 0
        self.article_III_state_off_budget_funds = 0
        self.pension_fund = 0
        self.social_insurance_fund = 0
        self.federal_health_insurance_fund = 0
        self.territorial_health_insurance_fund = 0
        self.article_IY_other_recipients_MOU_FC = 0
        self.account_balance_40101 = 0
        self.NVS_chapter_100 = 0
        self.total_for_section_III = 0
        self.total_for_section_III_federal_budgets = 0
        self.total_for_section_III_regional_budgets = 0
        self.total_for_section_III_local_budgets = 0
        self.GVF = 0

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

                self.article_II_consolidated_regional_budget = this_sheet['L' + str(i)].value
                self.article_II_consolidated_regional_budget = float(self.article_II_consolidated_regional_budget) + this_sheet['N' + str(i)].value

                self.regional_budgets = this_sheet['L' + str(i)].value
                self.local_budgets = this_sheet['N' + str(i)].value
                break

        i = 0
        for cell in this_sheet['C']:
            i = i + 1
            if cell.value == 'Всего по разделам I и II':
                self.pension_fund = float(this_sheet['D' + str(i)].value)
                self.social_insurance_fund = float(this_sheet['F' + str(i)].value)
                self.federal_health_insurance_fund = float(this_sheet['H' + str(i)].value)
                self.territorial_health_insurance_fund = float(this_sheet['J' + str(i)].value)
                self.article_IY_other_recipients_MOU_FC = float(this_sheet['L' + str(i)].value)
                self.account_balance_40101 = float(this_sheet['O' + str(i)].value)

        this_sheet = wb[sheet_one]
        i = 0
        for cell in this_sheet['B']:
            i = i + 1
            result = re.match(r'10301', str(cell.value))
            if result:
                self.vat_on_goods_sold_on_the_territory_of_the_Russian_Federation = self.vat_on_goods_sold_on_the_territory_of_the_Russian_Federation + float(
                    this_sheet['J' + str(i)].value)

            result_2 = re.match(r'10401', str(cell.value))
            if result_2:
                self.vat_on_goods_imported_into_the_territory_of_the_Russian_Federation = self.vat_on_goods_imported_into_the_territory_of_the_Russian_Federation + float(
                    this_sheet['J' + str(i)].value)

            result_3 = re.match(r'10101', str(cell.value))
            if result_3:
                self.income_tax = self.income_tax + float(
                    this_sheet['J' + str(i)].value)
                self.regional_budgets_land_tax_from_organizations = self.regional_budgets_land_tax_from_organizations + float(
                    this_sheet['L' + str(i)].value)

            result_4 = re.match(r'10102', str(cell.value))
            if result_4:
                self.regional_budgets_NDFL = self.regional_budgets_NDFL + float(this_sheet['L' + str(i)].value)
                self.local_budgets_NDFL = self.local_budgets_NDFL + float(this_sheet['N' + str(i)].value)

            result_5 = re.match(r'1060603', str(cell.value))
            if result_5:
                self.local_budgets_land_tax_from_organizations = self.local_budgets_land_tax_from_organizations + float(this_sheet['N' + str(i)].value)

            result_6 = re.match(r'105', str(cell.value))
            if result_6:
                self.local_budgets_comprehensive_income_taxes = self.local_budgets_comprehensive_income_taxes + float(this_sheet['N' + str(i)].value)

            result_7 = re.match(r'11701010016000180', str(cell.value))
            if result_7:
                self.NVS_chapter_100 = self.NVS_chapter_100 + float(this_sheet['J' + str(i)].value)

            if this_sheet['D' + str(i)].value == 'В том числе по бюджетам:':
                break

        this_sheet = wb[sheet_three]
        i = 0
        for cell in this_sheet['B']:
            i = i + 1
            result = re.match(r'10301', str(cell.value))
            if result:
                self.vat_on_goods_sold_on_the_territory_of_the_Russian_Federation = self.vat_on_goods_sold_on_the_territory_of_the_Russian_Federation + float(
                    this_sheet['F' + str(i)].value)

            result_2 = re.match(r'11701010016000180', str(cell.value))
            if result_2:
                self.NVS_chapter_100 = self.NVS_chapter_100 + float(this_sheet['F' + str(i)].value)

            if this_sheet['C' + str(i)].value == 'Всего по разделу III':
                self.total_for_section_III = float(this_sheet['D' + str(i)].value)
                self.total_for_section_III_federal_budgets = float(this_sheet['F' + str(i)].value)
                self.total_for_section_III_regional_budgets = float(this_sheet['H' + str(i)].value)
                self.total_for_section_III_local_budgets = float(this_sheet['J' + str(i)].value)
                self.GVF = float(this_sheet['L' + str(i)].value) + float(this_sheet['N' + str(i)].value) + float(this_sheet['P' + str(i)].value) + float(this_sheet['R' + str(i)].value)

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

        self.vat_on_goods_imported_into_the_territory_of_the_Russian_Federation = round(self.vat_on_goods_imported_into_the_territory_of_the_Russian_Federation/1000000, 2)
        print(self.vat_on_goods_imported_into_the_territory_of_the_Russian_Federation)

        self.income_tax = round(self.income_tax/1000000, 2)
        print(self.income_tax)

        self.article_II_consolidated_regional_budget = round(self.article_II_consolidated_regional_budget/1000000, 2)
        print(self.article_II_consolidated_regional_budget)

        self.regional_budgets = round(self.regional_budgets/1000000, 2)
        print(self.regional_budgets)

        self.regional_budgets_NDFL = round(self.regional_budgets_NDFL/1000000, 2)
        print(self.regional_budgets_NDFL)

        self.regional_budgets_land_tax_from_organizations = round(self.regional_budgets_land_tax_from_organizations/1000000, 2)
        print(self.regional_budgets_land_tax_from_organizations)

        self.local_budgets = round(self.local_budgets/1000000, 2)
        print(self.local_budgets)

        self.local_budgets_NDFL = round(self.local_budgets_NDFL/1000000, 2)
        print(self.local_budgets_NDFL)

        self.local_budgets_land_tax_from_organizations = round(self.local_budgets_land_tax_from_organizations/1000000, 2)
        print(self.local_budgets_land_tax_from_organizations)

        self.local_budgets_comprehensive_income_taxes = round(self.local_budgets_comprehensive_income_taxes/1000000, 2)
        print(self.local_budgets_comprehensive_income_taxes)

        self.article_III_state_off_budget_funds = float(self.pension_fund) + float(self.social_insurance_fund) + float(self.federal_health_insurance_fund) + float(self.territorial_health_insurance_fund)
        self.article_III_state_off_budget_funds = round(self.article_III_state_off_budget_funds/1000000, 2)
        print(self.article_III_state_off_budget_funds)

        self.pension_fund = round(self.pension_fund/1000000, 2)
        print(self.pension_fund)

        self.social_insurance_fund = round(self.social_insurance_fund/1000000, 2)
        print(self.social_insurance_fund)

        self.federal_health_insurance_fund = round(self.federal_health_insurance_fund/1000000, 2)
        print(self.federal_health_insurance_fund)

        self.territorial_health_insurance_fund = round(self.territorial_health_insurance_fund/1000000, 2)
        print(self.territorial_health_insurance_fund)

        self.article_IY_other_recipients_MOU_FC = round(self.article_IY_other_recipients_MOU_FC/1000000, 2)
        print(self.article_IY_other_recipients_MOU_FC)

        self.account_balance_40101 = round(self.account_balance_40101/1000000, 2)
        print(self.account_balance_40101)

        self.NVS_chapter_100 = round(self.NVS_chapter_100/1000000, 2)
        print(self.NVS_chapter_100)

        self.total_for_section_III = round(self.total_for_section_III/1000000, 2)
        print(self.total_for_section_III)

        self.total_for_section_III_federal_budgets = round(self.total_for_section_III_federal_budgets/1000000, 2)
        print(self.total_for_section_III_federal_budgets)

        self.total_for_section_III_regional_budgets = round(self.total_for_section_III_regional_budgets/1000000, 2)
        print(self.total_for_section_III_regional_budgets)

        self.total_for_section_III_local_budgets = round(self.total_for_section_III_local_budgets/1000000, 2)
        print(self.total_for_section_III_local_budgets)

        self.GVF = round(self.GVF/1000000, 2)
        print(self.GVF)




