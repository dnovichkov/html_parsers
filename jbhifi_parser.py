"""
Project link: https://www.upwork.com/jobs/~01dd8043d26914d97c
Results: https://cloud.mail.ru/public/MoXK/ve1GAGzx9
"""

import os
import time
import datetime

from bs4 import BeautifulSoup
from selenium.webdriver.common import desired_capabilities
from selenium.webdriver.opera import options
from selenium import webdriver


import xlsxwriter

ITERATION_COUNT = 1


def download_pages():
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
    driver.get("https://www.jbhifi.com.au/")
    while True:
        if 'JB Hi-Fi | JB Hi-Fi - Australia' in driver.page_source:
            break

    new_elems = driver.find_elements_by_class_name("info-wrapper")
    print(len(new_elems))
    for x in range(0, len(new_elems)):
        new_elems = driver.find_elements_by_class_name("info-wrapper")

        time.sleep(3)
        if new_elems[x].is_displayed():
            print('Parse elem ', x)
            print(new_elems[x].text)
            link_elem = new_elems[x].find_element_by_xpath('..')

            link_elem.click()
            time.sleep(3)
            page = driver.page_source
            current_dir = os.path.dirname(__file__)
            html_filename = 'htmls\\page_' + '_' + str(x) + '.html'
            filemame = os.path.join(current_dir, html_filename)
            with open(filemame, "w", encoding="utf-8") as f:
                f.write(page)
            driver.back()
            time.sleep(1)


def create_headers(worksheet, workbook):
    """
    Adding headers to worksheet
    :param worksheet:
    :param workbook:
    :return:
    """
    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': True})

    # Write some data headers.
    worksheet.set_column('A:A', 14)
    worksheet.set_column('B:B', 23)
    worksheet.set_column('C:C', 68)
    worksheet.set_column('D:D', 23)
    worksheet.set_column('D:D', 42)
    worksheet.set_column('E:E', 42)
    worksheet.set_column('F:F', 23)

    worksheet.write('A1', 'GTIN', bold)
    worksheet.write('B1', 'Manufacturer\'s warranty', bold)
    worksheet.write('C1', 'Model', bold)
    worksheet.write('D1', 'Brand', bold)
    worksheet.write('E1', 'key feature', bold)
    worksheet.write('F1', 'Photo 1', bold)
    worksheet.write('G1', 'Saved page', bold)


def parse_pages():
    filenames = []
    file_folder = os.getcwd() + "\\htmls"
    for file in os.listdir(file_folder):
        if file.endswith(".html"):
            filenames.append(file)

    excel_filename = 'Result_' + datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S') + '.xlsx'
    workbook = xlsxwriter.Workbook(excel_filename)
    worksheet_all = workbook.add_worksheet()

    create_headers(worksheet_all, workbook)

    row = 1
    col = 0

    cell_format = workbook.add_format()
    cell__wrapped_format = workbook.add_format()
    cell__wrapped_format.set_text_wrap()
    for filename in filenames:
        full_filename = 'htmls\\' + filename
        # print(full_filename)
        with open(full_filename, "r", encoding="utf-8") as html_file:
            try:
                soup = BeautifulSoup(html_file.read(), "lxml")
                productViewPLU = "None"
                try:
                    productViewPLU_elem = soup.find("input", id="productViewPLU")
                    if productViewPLU_elem:
                        productViewPLU = str(productViewPLU_elem)[48:-3]
                        print(productViewPLU)
                except AttributeError:
                    print("!!!")
                name = soup.find("meta", {'property': 'name'})
                # print(str(name)[15:-19])
                brand = soup.find("meta", {'property': 'brand'})
                # print(str(brand)[15:-20])
                logo = soup.find("meta", {'property': 'logo'})
                print(str(logo)[41:-19])
                description = soup.find("meta", {'property': 'description'})
                # # print(str(description))
                warranty_elem_res = "None"
                try:
                    warranty_elem = soup.find(text='Manufacturer\'s warranty').parent.parent
                    warranty_elem_res = str(warranty_elem)[44:-4]
                except AttributeError:
                    print("warranty_elem_res")

                # print(str(warranty_elem)[44:-4])
                #
                worksheet_all.write(row, col, productViewPLU, cell_format)
                worksheet_all.write(row, col + 1, str(warranty_elem_res), cell_format)
                worksheet_all.write(row, col + 2, str(name)[15:-19], cell_format)
                worksheet_all.write(row, col + 3, str(brand)[15:-20], cell_format)
                worksheet_all.write(row, col + 4, str(description), cell__wrapped_format)
                worksheet_all.write(row, col + 5, str(logo)[41:-19], cell_format)
                worksheet_all.write(row, col + 6, full_filename, cell_format)

            except AttributeError:
                print(full_filename)
            finally:
                row += 1

    workbook.close()


if __name__ == "__main__":

    download_pages()
    parse_pages()
