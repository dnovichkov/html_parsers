"""This file contains result for parsing names from
Results of parsing you can download from https://cloud.mail.ru/public/NHUP/TkfDJY7Gz
Project link: https://kwork.ru/track?id=4063767"""

import os
import time
import datetime
import string

from bs4 import BeautifulSoup

from selenium.webdriver.common import desired_capabilities
from selenium.webdriver.opera import options
from selenium import webdriver

import xlsxwriter


def download_pages():
    """
    Download necessary pages from http://medsalltheworld.com
    :return:
    """
    # Replace this path with the actual path on your machine.
    current_dir = os.path.dirname(__file__)
    opera_driver_filename = os.path.join(current_dir, 'operadriver.exe')
    opera_driver_location = os.path.abspath(opera_driver_filename)

    # Replace this path with the actual path on your machine.
    oper_exe_file_location = os.path.abspath('D:\\Program Files\\Opera\\58.0.3135.127\\opera.exe')

    opera_capabilities = desired_capabilities.DesiredCapabilities.OPERA.copy()

    opera_options = options.ChromeOptions()
    opera_options.binary_location = oper_exe_file_location

    # Use the below argument if you want the Opera browser to be in the max. state when launching.
    opera_options.add_argument('--start-maximized')

    driver = webdriver.Chrome(executable_path=opera_driver_location, chrome_options=opera_options,
                              desired_capabilities=opera_capabilities)
    # Some sites don't open from Russia. So, you have 5 seconds to enable Opera VPN.
    vpn_settings_url = 'opera://settings/vpn'
    driver.get(vpn_settings_url)
    time.sleep(5)
    # You can use multithreading here but I don't like enable VPN some times:)
    for symbol in string.ascii_lowercase:
        letter_part = 'letter-' + symbol + '-eu.html'
        url = 'http://medsalltheworld.com/' + letter_part
        driver.get(url)

        time.sleep(2)
        page = driver.page_source
        current_dir = os.path.dirname(__file__)
        html_filename = 'htmls\\page_' + '_' + letter_part
        full_filename = os.path.join(current_dir, html_filename)
        with open(full_filename, "w", encoding="utf-8") as file:
            file.write(page)


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
    worksheet.set_column('A:A', 28)
    worksheet.set_column('B:B', 65)
    worksheet.set_column('C:C', 32)

    worksheet.write('A1', 'Name', bold)
    worksheet.write('B1', 'Url', bold)
    worksheet.write('C1', 'Saved page', bold)


def get_html_filenames():
    """
    Get html-file names in 'html'-directory
    :return:
    """
    filenames = []
    file_folder = os.getcwd() + "\\htmls"
    for file in os.listdir(file_folder):
        if file.endswith(".html"):
            filenames.append('htmls\\' + file)
    return filenames


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
    site_url = 'http://medsalltheworld.com/'
    for full_filename in get_html_filenames():
        with open(full_filename, "r", encoding="utf-8") as html_file:
            try:
                soup = BeautifulSoup(html_file.read(), "lxml")
                product_name_elements = soup.find_all("li", class_="col-xs-6 col-md-4")
                for elem in product_name_elements:
                    name = elem.select('h3')[0].text.replace('®', '')
                    elem_url = site_url + elem.select('h3')[0].find('a')['href']

                    worksheet_all.write(row, col, name, cell_format)
                    worksheet_all.write(row, col + 1, elem_url, cell_format)
                    worksheet_all.write(row, col + 2, full_filename, cell_format)
                    row += 1

            except AttributeError:
                print(full_filename)

    workbook.close()


if __name__ == "__main__":
    download_pages()
    parse_pages()
