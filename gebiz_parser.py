"""
Contains multitherad code for parsing suppliers from www.gebiz.gov.sg
Project link: https://www.upwork.com/jobs/~01a76e809fbd3b3edf
"""
import os
import time
import datetime

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.errorhandler import WebDriverException, NoSuchElementException
import xlsxwriter
from multiprocessing import Pool


ITERATION_COUNT = 4


def main(start_pos=1, proc_count=ITERATION_COUNT):
    driver = webdriver.Chrome()

    driver.get("https://www.gebiz.gov.sg/ptn/supplier/directory/index.xhtml")
    elem = driver.find_element_by_name("contentForm:search")
    elem.send_keys(Keys.RETURN)

    page_count = 812
    filename = str(start_pos) + '_Result' + '_' + datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S') + '.xlsx'
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()

    row = 1
    col = 0

    cell_format = workbook.add_format()
    for i in range(start_pos, page_count + 1, proc_count):
        try:
            print('Parse page ', i)
            id_name = 'contentForm:j_idt184:j_idt196_' + str(i) + '_' + str(i)
            print('Search elem with id %s on page %s ', id_name, i)
            # print(id_name)
            try:
                page_elem = driver.find_element_by_id(id_name)
                page_elem.click()
                time.sleep(3)
                new_elems = driver.find_elements_by_class_name("commandLink_TITLE-BLUE")
                # print(new_elems)
                # print(len(new_elems))
                try:
                    for x in range(0, len(new_elems)):
                        new_elems = driver.find_elements_by_class_name("commandLink_TITLE-BLUE")
                        if new_elems[x].is_displayed():
                            print('Parse elem ', x)
                            new_elems[x].click()

                            ref_elem = driver.find_element_by_class_name("outputText_LABEL-GRAY")
                            ref_data = ref_elem.text
                            worksheet.write(row, col, ref_data, cell_format)

                            name_elem = driver.find_element_by_class_name("outputText_TITLE-BLACK")
                            name_text = name_elem.text
                            print('parsing data about ', name_text)
                            page = driver.page_source
                            dir = os.path.dirname(__file__)
                            html_filename = 'htmls\\page_' + str(i) + '_' + str(x) + '___' + name_text + '.html'
                            filemame = os.path.join(dir, html_filename)
                            with open(filemame, "w", encoding="utf-8") as f:
                                f.write(page)

                            col_ind = col + 1
                            worksheet.write(row, col_ind, name_text, cell_format)
                            print(ref_data)
                            for elem in driver.find_elements_by_xpath('.//div[@class = "form2_MAIN"]'):
                                col_ind = col_ind + 1
                                # print(col_ind)
                                worksheet.write(row, col_ind, elem.text, cell_format)

                            row += 1
                            return_elem = driver.find_element_by_name('contentForm:j_idt137')
                            return_elem.click()
                            time.sleep(1)

                except WebDriverException as ex:
                    print("Exception: ", ex)
                    dirname = os.path.dirname(__file__)
                    png_filename = 'bad_screens\\Screen_page_' + str(i) + '.png'
                    full_filename = os.path.join(dirname, png_filename)
                    print("Saving bad screen to ", full_filename)
                    driver.save_screenshot(full_filename)

            except NoSuchElementException as ex:
                print("Exception: %s for elem wih id ", ex, id_name)
                dirname = os.path.dirname(__file__)
                png_filename = 'bad_screens\\Screen_page_' + str(i) + '.png'
                full_filename = os.path.join(dirname, png_filename)
                print("Saving bad screen to ", full_filename)
                driver.save_screenshot(full_filename)
                continue

            except WebDriverException as ex:
                print("Exception: ", ex)
                dirname = os.path.dirname(__file__)
                png_filename = 'bad_screens\\Screen_page_' + str(i) + '_elem_' + str(x) + '.png'
                full_filename = os.path.join(dirname, png_filename)
                print("Saving bad screen to ", full_filename)
                driver.save_screenshot(full_filename)

        except WebDriverException as ex:
            print("Exception: ", ex)
            dirname = os.path.dirname(__file__)
            png_filename = 'bad_screens\\Screen_page_' + str(i) + '.png'
            full_filename = os.path.join(dirname, png_filename)
            print("Saving bad screen to ", full_filename)
            driver.save_screenshot(full_filename)

    driver.close()
    workbook.close()


if __name__ == "__main__":
    pool = Pool()

    with Pool() as pool:
        pool.map(main, range(1, ITERATION_COUNT + 1))
