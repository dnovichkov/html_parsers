"""This file contains result for parsing names from
Project link: https://www.upwork.com/jobs/~0112736463cc4cc4a9"""

import os
import datetime

import xlsxwriter
import requests


def create_headers(worksheet, workbook):
    """
    Add headers to Excel-file
    :param worksheet:
    :param workbook:
    :return:
    """
    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': True})

    # Write some data headers.
    worksheet.set_column('A:A', 32)
    worksheet.set_column('B:B', 28)
    worksheet.set_column('C:C', 51)

    worksheet.write('A1', 'link', bold)
    worksheet.write('B1', 'ACCESSION NUMBER', bold)
    worksheet.write('C1', 'CONFORMED SUBMISSION TYPE', bold)
    worksheet.write('D1', 'PUBLIC DOCUMENT COUNT', bold)
    worksheet.write('E1', 'CONFORMED PERIOD OF REPORT', bold)
    worksheet.write('F1', 'FILED AS OF DATE', bold)
    worksheet.write('G1', 'DATE AS OF CHANGE', bold)
    worksheet.write('H1', 'BUSINESS ADDRESS - (STREET 1)', bold)
    worksheet.write('I1', 'BUSINESS ADDRESS - (CITY)', bold)
    worksheet.write('J1', 'BUSINESS ADDRESS - (STATE)', bold)
    worksheet.write('K1', 'BUSINESS ADDRESS - (ZIP)', bold)


def get_text_filenames():
    """
    Get html-file names in 'html'-directory
    :return:
    """
    filenames = []
    file_folder = os.getcwd() + "\\txts"
    for file in os.listdir(file_folder):
        if file.endswith(".txt"):
            filenames.append('txts\\' + file)
    return filenames


def find_value(strings, tag):
    for line in strings:
        if tag in line:
            res = line.split('\t')[-1]
            # print(res)
            return res
    return ''


def parse_pages():
    """
    Parse saved pages and save data to Excel-file.
    :return:
    """
    excel_filename = 'Result_' + datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S') + '.xlsx'
    workbook = xlsxwriter.Workbook(excel_filename)
    worksheet_all = workbook.add_worksheet()

    create_headers(worksheet_all, workbook)

    row = 1
    col = 0

    cell_format = workbook.add_format()
    cell__wrapped_format = workbook.add_format()
    cell__wrapped_format.set_text_wrap()
    # site_url = 'http://medsalltheworld.com/'
    index = 1
    link_file = 'Adresses.txt'
    with open(link_file, "r", encoding="utf-8") as text_file:
        addr = text_file.read().split('\n')
        for line in addr:
            # print(line)
            print(index, '  ', line)
            index += 1
            content = requests.get(line).text
            # print(content)
            # res = text_file.read()
            lines = content.split('\n')
            ac_number = find_value(lines, 'ACCESSION NUMBER')
            cs_type = find_value(lines, 'CONFORMED SUBMISSION TYPE')

            pd_count = find_value(lines, 'PUBLIC DOCUMENT COUNT')
            cp_report = find_value(lines, 'CONFORMED PERIOD OF REPORT')
            filled_date = find_value(lines, 'FILED AS OF DATE')

            dt_change = find_value(lines, 'DATE AS OF CHANGE')
            street1 = find_value(lines, 'STREET 1')
            city = find_value(lines, 'CITY')
            state = find_value(lines, 'STATE:')
            zip_ = find_value(lines, 'ZIP')

            worksheet_all.write(row, col, line, cell_format)
            worksheet_all.write(row, col + 1, ac_number, cell_format)
            worksheet_all.write(row, col + 2, cs_type, cell_format)
            worksheet_all.write(row, col + 3, pd_count, cell_format)
            worksheet_all.write(row, col + 4, cp_report, cell_format)
            worksheet_all.write(row, col + 5, filled_date, cell_format)
            worksheet_all.write(row, col + 6, dt_change, cell_format)
            worksheet_all.write(row, col + 7, street1, cell_format)
            worksheet_all.write(row, col + 8, city, cell_format)
            worksheet_all.write(row, col + 9, state, cell_format)
            worksheet_all.write(row, col + 10, zip_, cell_format)
            row += 1

    workbook.close()
    return


if __name__ == "__main__":
    parse_pages()
