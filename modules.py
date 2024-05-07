import openpyxl as xl


def process_income_workbook(filename):
    wb = xl.load_workbook(filename)
    sheet = wb['Sheet1']

    for row in range(2, sheet.max_row + 1):
        gross_income_cell = sheet.cell(row, column=2)
        tax_owed_cell = sheet.cell(row, 3)
        net_income_cell = sheet.cell(row, 4)
        if gross_income_cell.value or net_income_cell.value or tax_owed_cell.value is not None:
            tax_owed = gross_income_cell.value - net_income_cell.value
            tax_owed_cell.value = tax_owed

    wb.save(filename)
    print(f"\nThe operation on {filename} has finished!")


def process_mtd_figures(income_wb_name,
                        expenses_wb_name,
                        mtd_wb_name):
    income_wb = xl.load_workbook(income_wb_name)
    expenses_wb = xl.load_workbook(expenses_wb_name)
    mtd_wb = xl.load_workbook(mtd_wb_name)
    income_sheet, expenses_sheet, mtd_sheet = (income_wb['Sheet1'],
                                               expenses_wb['Sheet1'],
                                               mtd_wb['Sheet1'])
    total_gross_pay = 0
    total_tax_paid = 0
    total_net_pay = 0
    total_expenses = 0

    for row in range(2, income_sheet.max_row + 1):
        gross_cell = income_sheet.cell(row, column=2)
        tax_cell = income_sheet.cell(row, 3)
        net_cell = income_sheet.cell(row, 4)
        if gross_cell.value is None:
            total_gross_pay += 0
        else:
            total_gross_pay += gross_cell.value
        if tax_cell.value is None:
            total_tax_paid += 0
        else:
            total_tax_paid += tax_cell.value
        if net_cell.value is None:
            total_net_pay += 0
        else:
            total_net_pay += net_cell.value

    for row in range(2, expenses_sheet.max_row + 1):
        cost_cell = expenses_sheet.cell(row, 3)
        total_expenses += cost_cell.value

    total_savings = total_net_pay - total_expenses

    mtd_gross_cell = mtd_sheet.cell(row=2, column=1)
    mtd_tax_cell = mtd_sheet.cell(2, 2)
    mtd_net_cell = mtd_sheet.cell(2, 3)
    mtd_expenses_cell = mtd_sheet.cell(2, 4)
    mtd_savings_cell = mtd_sheet.cell(2, 5)
    mtd_gross_cell.value = total_gross_pay
    mtd_tax_cell.value = total_tax_paid
    mtd_net_cell.value = total_net_pay
    mtd_expenses_cell.value = total_expenses
    mtd_savings_cell.value = total_savings

    mtd_wb.save(mtd_wb_name)
    print(f"\nThe operation on {mtd_wb_name} has finished!")
