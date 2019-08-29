"""
https://www.upwork.com/jobs/~01828e8462d1731552
"""
import datetime
from xlrd import open_workbook
import xlsxwriter


def set_format(worksheet, workbook):
    """
    Sets format to worksheet
    :param worksheet:
    :param workbook:
    :return:
    """
    worksheet.set_column('A:A', 10.38)
    worksheet.set_column('B:B', 10.38)
    worksheet.set_column('C:C', 24)
    worksheet.set_column('D:D', 31.75)
    worksheet.set_column('D:D', 34.5)


def rework_file():
    xls_read_filename = 'names.xlsx'
    xl_workbook = open_workbook(xls_read_filename)
    if not xl_workbook.nsheets:
        print("Empty book")
        return
    first_sheet = xl_workbook.sheet_by_index(0)
    data = []
    num_cols = first_sheet.ncols
    for row_idx in range(0, first_sheet.nrows):
        row_data = []
        for col_idx in range(0, num_cols):
            cell_obj = first_sheet.cell(row_idx, col_idx)
            row_data.append(cell_obj.value)
        data.append(row_data)

    excel_filename = 'AAA_Result_' + datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S') + '.xlsx'
    workbook = xlsxwriter.Workbook(excel_filename)
    worksheet_all = workbook.add_worksheet()
    cell_format = workbook.add_format()

    set_format(worksheet_all, workbook)

    row = 0
    for rec in data:
        print(rec)
        for i in range(0, num_cols):
            worksheet_all.write(row, i, rec[i], cell_format)
        row += 1
    workbook.close()


if __name__ == "__main__":
    rework_file()
