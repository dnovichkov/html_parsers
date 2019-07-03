"""This file contains result for parsing names from
Project link: https://www.upwork.com/jobs/Data-Scraping_~019c9ab656391fe44f"""

import os
import time
import datetime
import string

from bs4 import BeautifulSoup

from selenium.webdriver.common import desired_capabilities
from selenium.webdriver.opera import options
from selenium import webdriver

import xlsxwriter
from selenium.webdriver.support.ui import Select


def download_pages():
    """
    Download necessary pages from https://www.riai.ie/practice_directory/search-results/search&pCity=Dublin/
    :return:
    """
    # Replace this path with the actual path on your machine.
    current_dir = os.path.dirname(__file__)
    opera_driver_filename = os.path.join(current_dir, 'operadriver.exe')
    opera_driver_location = os.path.abspath(opera_driver_filename)

    # Replace this path with the actual path on your machine.
    oper_exe_file_location = os.path.abspath('D:\\Program Files\\Opera\\60.0.3255.170\\opera.exe')

    opera_capabilities = desired_capabilities.DesiredCapabilities.OPERA.copy()

    opera_options = options.ChromeOptions()
    opera_options.binary_location = oper_exe_file_location

    # Use the below argument if you want the Opera browser to be in the max. state when launching.
    opera_options.add_argument('--start-maximized')

    driver = webdriver.Chrome(executable_path=opera_driver_location, chrome_options=opera_options,
                              desired_capabilities=opera_capabilities)
    driver.get("https://www.riai.ie/practice_directory/search-results/search&pCity=Dublin/")
    select = Select(driver.find_element_by_name('sortable_length'))
    select.select_by_value('100')
    time.sleep(2)
    # return

    # next_elem = driver.find_element_by_class_name('paginate_enabled_next')
    counter = 0
    while True:
        # Parse
        new_elems = driver.find_elements_by_link_text('More Info')
        print(len(new_elems))
        for x in range(0, len(new_elems)):
            newest_elems = driver.find_elements_by_link_text('More Info')

            # time.sleep(3)
            if newest_elems[x].is_displayed():
                print('Parse elem ', x)

                newest_elems[x].click()
                time.sleep(1)
                page = driver.page_source
                current_dir = os.path.dirname(__file__)
                html_filename = 'htmls\\page_' + '_' + str(x) + '_' + str(counter) + '.html'
                filemame = os.path.join(current_dir, html_filename)
                with open(filemame, "w", encoding="utf-8") as f:
                    f.write(page)
                driver.back()

                select = Select(driver.find_element_by_name('sortable_length'))
                select.select_by_value('100')
                time.sleep(1)

                for ind in range(1, counter + 1):
                    next_elem = driver.find_element_by_class_name('paginate_enabled_next')
                    print('1 skip ', next_elem)
                    if not next_elem:
                        break
                    next_elem.click()
                    select = Select(driver.find_element_by_name('sortable_length'))
                    select.select_by_value('100')
                    time.sleep(2)

        counter += 1
        next_elem = driver.find_element_by_class_name('paginate_enabled_next')
        print('2 skip ', next_elem)
        if not next_elem:
            break
        next_elem.click()
        select = Select(driver.find_element_by_name('sortable_length'))
        select.select_by_value('100')
        time.sleep(2)
    return


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

    worksheet.write('A1', 'Practice Name', bold)
    worksheet.write('B1', 'Lead Partners', bold)
    worksheet.write('C1', 'Partners', bold)
    worksheet.write('D1', 'Address', bold)
    worksheet.write('E1', 'Telephone', bold)
    worksheet.write('F1', 'Mobile Telephone', bold)
    worksheet.write('G1', 'Fax', bold)
    worksheet.write('H1', 'Website', bold)
    worksheet.write('I1', 'Email', bold)
    worksheet.write('J1', 'Page', bold)


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
    # site_url = 'http://medsalltheworld.com/'
    html_filenames = get_html_filenames()
    print(len(html_filenames))

    for full_filename in html_filenames:
        with open(full_filename, "r", encoding="utf-8") as html_file:
            # print(row, ' ', full_filename)

            try:
                soup = BeautifulSoup(html_file.read(), "lxml")
                # body = soup.find('body')
                content = soup.find("div", id="content")
                # print(content)
                company = content.find('h3').text
                # print(company)
                #
                info = content.find_all('p')
                address = ' '.join(info[1].text.split('\n'))
                # print(address)
                phone_data = info[2].text.split('\n')
                phone = phone_data[0].strip('Tel: ')
                # print(phone)
                mob = phone_data[1].strip('Mob: ')
                # print(mob)
                fax = phone_data[2].strip('Fax: ')
                # print(fax)


                web_data = info[3].text.split('\n')
                # print(web_data)
                email = web_data[0].strip('E: ')
                # print(email)
                web_addr = web_data[1].strip('W: ')
                # print(web_addr)
                lead_partner = ''
                partners = ''
                if len(info) > 5 and 'Lead Partner:' in content.text:
                    lead_partner = info[5].text
                    # print(lead_partner)
                    if 'Lead Partner:' in lead_partner:
                        lead_partner = info[6].text
                    if 'Partners' in content:
                        partners = content.find('ul').text
                    # print(partners)

                worksheet_all.write(row, col, company, cell_format)
                worksheet_all.write(row, col + 1, lead_partner, cell_format)
                worksheet_all.write(row, col + 2, partners, cell_format)
                worksheet_all.write(row, col + 3, address, cell_format)
                worksheet_all.write(row, col + 4, phone, cell_format)
                worksheet_all.write(row, col + 5, mob, cell_format)
                worksheet_all.write(row, col + 6, fax, cell_format)
                worksheet_all.write(row, col + 7, web_addr, cell_format)
                worksheet_all.write(row, col + 8, email, cell_format)
                worksheet_all.write(row, col + 9, full_filename, cell_format)
                row += 1

            except AttributeError:
                print("!!!!!!!!", full_filename)

    workbook.close()


if __name__ == "__main__":
    # download_pages()
    parse_pages()
