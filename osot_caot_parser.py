"""
Parsers fow two sites: www.osot.on.ca and caot.ca
Project link: https://www.upwork.com/ab/proposals/1112592691613425665
Results: https://cloud.mail.ru/public/3qTk/WFqHdtZ49
"""

import os
import time
import datetime

from selenium import webdriver
from selenium.webdriver.remote.errorhandler import WebDriverException
from selenium.webdriver.support.ui import Select

from bs4 import BeautifulSoup
import xlsxwriter

ITERATION_COUNT = 1


def osot_download(start_pos=1, proc_count=ITERATION_COUNT):
    driver = webdriver.Chrome()

    driver.get("https://www.osot.on.ca/OSOT/Membership/Directory/PublicDirectory/OSOT/OT_Directory/"
               "PublicDirectory.aspx")
    select_kilometres = Select(driver.find_element_by_name(
        "ctl01$TemplateBody$WebPartManager1$gwpste_container_"
        "ctlPublicDirectory$cictlPublicDirectory$ddlPostalKilometers"))
    print(select_kilometres)

    select_kilometres.select_by_visible_text("All")

    time.sleep(2)
    search_elem = driver.find_element_by_name('ctl01$TemplateBody$WebPartManager1$gwpste_container_'
                                              'ctlPublicDirectory$cictlPublicDirectory$btnSearch2')
    search_elem.click()

    time.sleep(2)
    try:
        for page_ind in range(0, 26):
            print('Open page ', page_ind)
            if page_ind < 25:
                time.sleep(3)

                next_page_elem = driver.find_element_by_id('ctl01_TemplateBody_WebPartManager1_gwpste_'
                                                           'container_SearchResult_ciSearchResult_lbtnNext2')
                next_page_elem.click()
                continue

            for i in range(0, 25):
                if page_ind == 25 and i < 14:
                    continue
                time.sleep(3)
                link_name = 'ctl01_TemplateBody_WebPartManager1_gwpste_container_SearchResult_' \
                            'ciSearchResult_lvResults_ctrl' + str(i) + '_lbtnFullName'

                name_link = driver.find_element_by_id(link_name)
                name_link.click()
                time.sleep(4)
                page = driver.page_source
                current_dir = os.path.dirname(__file__)
                html_filename = 'htmls\\page_' + str(page_ind) + '_name_' + str(i) + '_' + '.html'
                filename = os.path.join(current_dir, html_filename)
                print(filename)
                with open(filename, "w", encoding="utf-8") as f:
                    f.write(page)
                return_elem = driver.find_element_by_id(
                    'ctl01_TemplateBody_WebPartManager1_gwpste_container_Details_ciDetails_'
                    'lbtnBackSearchResults')
                return_elem.click()
            time.sleep(3)
            next_page_elem = driver.find_element_by_id('ctl01_TemplateBody_WebPartManager1_gwpste_'
                                                       'container_SearchResult_ciSearchResult_lbtnNext2')
            next_page_elem.click()

    except WebDriverException as ex:
        print("Exception: ", ex)
        dirname = os.path.dirname(__file__)
        png_filename = 'bad_screens\\Screen_page_' + str(page_ind) + '.png'
        full_filename = os.path.join(dirname, png_filename)
        print("Saving bad screen to ", full_filename)
        driver.save_screenshot(full_filename)
    finally:
        driver.close()


def caot_download(start_pos=1, proc_count=ITERATION_COUNT):
    driver = webdriver.Chrome()

    driver.get("https://www.caot.ca/site/findot")
    select_province = Select(driver.find_element_by_id("clientForm.clientFilter.v586176"))
    print(select_province)

    select_province.select_by_visible_text("Ontario")

    time.sleep(2)
    search_elem = driver.find_element_by_name('update_results')
    search_elem.click()

    time.sleep(2)
    try:
        for page_ind in range(1, 27):
            print('Open page ', page_ind)

            time.sleep(3)
            page = driver.page_source
            current_dir = os.path.dirname(__file__)
            html_filename = 'htmls\\caot_page_' + str(page_ind) + '.html'
            filename = os.path.join(current_dir, html_filename)
            print(filename)
            with open(filename, "w", encoding="utf-8") as f:
                f.write(page)

            next_page_link = driver.find_element_by_link_text(str(page_ind + 1))
            next_page_link.click()

    except WebDriverException as ex:
        print("Exception: ", ex)
        dirname = os.path.dirname(__file__)
        png_filename = 'bad_screens\\Screen_page_' + str(page_ind) + '.png'
        full_filename = os.path.join(dirname, png_filename)
        print("Saving bad screen to ", full_filename)
        driver.save_screenshot(full_filename)
    finally:
        driver.close()


def parse_osot():
    filenames = []
    file_folder = os.getcwd() + "\\htmls"
    for file in os.listdir(file_folder):
        if file.endswith(".html"):
            filenames.append(file)

    excel_filename = 'Result_' + datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S') + '.xlsx'
    workbook = xlsxwriter.Workbook(excel_filename)
    worksheet_all = workbook.add_worksheet()

    row = 1
    col = 0

    cell_format = workbook.add_format()
    for filename in filenames:
        full_filename = 'htmls\\' + filename
        # print(full_filename)
        with open(full_filename, "r", encoding="utf-8") as html_file:
            try:

                soup = BeautifulSoup(html_file.read(), "lxml")
                span_name = soup.find("span", id="ctl01_TemplateBody_WebPartManager1_gwpste_container_"
                                                 "Details_ciDetails_lblName")
                # print(span_name.text)
                span_city = soup.find("span", id="ctl01_TemplateBody_WebPartManager1_gwpste_container_"
                                                 "Details_ciDetails_lblCity")
                # print(span_city.text)
                span_phone = soup.find("span", id="ctl01_TemplateBody_WebPartManager1_gwpste_container_"
                                                  "Details_ciDetails_lblWorkPhone")
                # print(span_phone.text)
                span_email = soup.find("span", id="ctl01_TemplateBody_WebPartManager1_gwpste_container_"
                                                  "Details_ciDetails_lblEmail")
                # print(span_email.text)

                span_company = soup.find("span", id="ctl01_TemplateBody_WebPartManager1_gwpste_container_"
                                                    "Details_ciDetails_lblCompanyName")
                # print(span_company.text)

                span_funded_by = soup.find("span", id="ctl01_TemplateBody_WebPartManager1_gwpste_container_"
                                                      "Details_ciDetails_lblFundedBy")
                # print(span_funded_by.text)

                worksheet_all.write(row, col, span_name.text, cell_format)
                worksheet_all.write(row, col + 1, span_city.text, cell_format)
                worksheet_all.write(row, col + 2, span_phone.text, cell_format)
                worksheet_all.write(row, col + 3, span_email.text, cell_format)
                worksheet_all.write(row, col + 4, span_company.text, cell_format)
                worksheet_all.write(row, col + 5, span_funded_by.text, cell_format)

            except AttributeError:
                print(full_filename)
            finally:
                row += 1

    workbook.close()


def parse_caot():

    filenames = []
    file_folder = os.getcwd() + "\\htmls"
    for file in os.listdir(file_folder):
        if file.endswith(".html"):
            filenames.append(file)

    excel_filename = 'Result_' + datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S') + '.xlsx'
    workbook = xlsxwriter.Workbook(excel_filename)
    worksheet_all = workbook.add_worksheet()

    row = 1
    col = 0

    cell_format = workbook.add_format()
    for filename in filenames:
        full_filename = 'htmls\\' + filename
        print(full_filename)
        with open(full_filename, "r", encoding="utf-8") as html_file:
            try:

                soup = BeautifulSoup(html_file.read(), "lxml")
                member_elements = soup.find_all("table", class_="otMember")
                # print(len(member_elements))

                for element in member_elements:
                    # print(element)
                    table_body = element.find("tbody")
                    rows = table_body.find_all("tr")
                    i  = 0
                    for tr in rows:
                        cols = tr.find_all("td")

                        worksheet_all.write(row, col, cols[0].find("h1").text.strip(), cell_format)
                        worksheet_all.write(row, col + 1, cols[0].find("p").text.strip(), cell_format)
                        worksheet_all.write(row, col + 2, cols[0].find("div").text.strip(), cell_format)
                        worksheet_all.write(row, col + 3, cols[0].find_all("p")[1].text.strip(), cell_format)
                        worksheet_all.write(row, col + 4, cols[0].find_all("p")[2].text.strip().replace('Areas of Practice:   ', ''), cell_format)

                        row += 1

            except AttributeError:
                print("!!! ", full_filename)
            # finally:
            #     row += 1

    workbook.close()


if __name__ == "__main__":
    # parse_osot()
    parse_caot()
    # pool = Pool()
    #
    # with Pool() as pool:
    #     pool.map(caot_download, range(1, ITERATION_COUNT + 1))
        # pool.map(osot_download, range(1, ITERATION_COUNT + 1))
