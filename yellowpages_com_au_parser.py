"""
Parser for yellowpages.com.au search results
Project link: https://www.upwork.com/ab/proposals/1111540607826567168
Results: https://cloud.mail.ru/public/57if/ppXW8FRWh
"""
import os
import datetime
from multiprocessing import Pool
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.remote.errorhandler import WebDriverException
import xlsxwriter

ITERATION_COUNT = 1


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
    worksheet.set_column('A:A', 60)
    worksheet.set_column('B:B', 14)
    worksheet.set_column('C:C', 42)
    worksheet.set_column('D:D', 46)

    worksheet.write('A1', 'Company name', bold)
    worksheet.write('B1', 'Phone', bold)
    worksheet.write('C1', 'E-mail', bold)
    worksheet.write('D1', 'Website', bold)


def get_html_pages(start_pos=1, proc_count=ITERATION_COUNT):
    """
    Saving html pages locally
    :param start_pos:
    :param proc_count:
    :return:
    """
    driver = webdriver.Chrome()

    driver.get("https://www.yellowpages.com.au/")

    while True:
        if 'Servicing' in driver.page_source:
            break

    page_count = 29
    for i in range(start_pos, page_count + 1, proc_count):
        try:
            print('Parse page ', i)
            url = 'https://www.yellowpages.com.au/search/listings?clue=shopfitters&' \
                  'eventType=pagination&pageNumber=' + \
                  str(i) +'&referredBy=www.yellowpages.com.au'
            driver.get(url)
            page = driver.page_source
            dirname = os.path.dirname(__file__)
            html_filename = 'htmls\\page_' + str(i) + '___\'.html'
            filemame = os.path.join(dirname, html_filename)
            with open(filemame, "w", encoding="utf-8") as html_file:
                html_file.write(page)

        except WebDriverException as ex:
            print("Exception: ", ex)
            dirname = os.path.dirname(__file__)
            png_filename = 'bad_screens\\Screen_page_' + str(i) + '.png'
            full_filename = os.path.join(dirname, png_filename)
            print("Saving bad screen to ", full_filename)
            driver.save_screenshot(full_filename)

    driver.close()


def parse_page():
    """
    Extracting data from html_files
    :return:
    """
    excel_filename = 'Result_' + datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S') + '.xlsx'
    workbook = xlsxwriter.Workbook(excel_filename)
    worksheet_all = workbook.add_worksheet(name="All")
    worksheet_email = workbook.add_worksheet(name="Emails")
    create_headers(worksheet_all, workbook)
    create_headers(worksheet_email, workbook)

    row = 1
    row_email = 1
    col = 0

    cell_format = workbook.add_format()
    # page_number = 8
    for page_number in range(1, 30):
        filename = 'htmls\\page_' + str(page_number) + '___\'.html'
        print(filename)
        parsed_element_for_page = 0
        with open(filename, "r", encoding="utf-8") as html_file:

            soup = BeautifulSoup(html_file.read(), "lxml")

            elements = soup.find_all("div", class_="search-contact-card call-to-actions-5 feedback-feature-on")
            elements.extend(soup.find_all("div", class_="search-contact-card call-to-actions-4 feedback-feature-on"))
            elements.extend(soup.find_all("div", class_="search-contact-card call-to-actions-2 feedback-feature-on"))
            elements.extend(soup.find_all("div", class_="search-contact-card call-to-actions-3 feedback-feature-on"))
            elements.extend(soup.find_all("div", class_="flow-layout inside-gap inside-gap-large vertical"))

            print(len(elements))
            for elem in elements:
                print("________________________")
                clickable_listing_elem = elem.find("div", class_="listing-details clickable-listing")
                if not clickable_listing_elem:
                    clickable_listing_elem = elem.find("div", class_="srp-brand-bar-container")
                if not clickable_listing_elem:
                    clickable_listing_elem = elem.find("div", class_="cell first-cell")
                if clickable_listing_elem:
                    # print(clickable_listing_elem.text)
                    company_name_elem = clickable_listing_elem.find("a", class_="listing-name")
                    if not company_name_elem:
                        print(clickable_listing_elem)
                    company_name = company_name_elem.text
                    print(company_name)
                    parsed_element_for_page += 1
                    worksheet_all.write(row, col + 4, filename, cell_format)
                    worksheet_all.write(row, col, company_name, cell_format)

                phone_elem = elem.find("a", title="Phone")
                if phone_elem:
                    # print(phone_elem)
                    # print(clickable_listing_elem.text)
                    phone = phone_elem.text
                    print(phone)
                    worksheet_all.write(row, col + 1, phone, cell_format)

                malto_elem = elem.find("a", class_="contact contact-main contact-email ")
                if not malto_elem:
                    malto_elem = elem.find("a", class_="image middle contact-main contact-email target-media-contact ")

                url_elem = elem.find("a", class_="contact contact-main contact-url ")
                if not url_elem:
                    url_elem = elem.find("a", class_="image middle contact-main contact-url target-media-contact ")
                if url_elem:
                    website = url_elem.get('href')
                    print(website)
                    worksheet_all.write(row, col + 3, website, cell_format)

                if malto_elem:
                    mailto_raw = malto_elem.get('href').replace('mailto:', '')
                    subject_pos = mailto_raw.find('?subject=')
                    mailto = mailto_raw[0:subject_pos].replace('%40', '@')
                    print(mailto)
                    worksheet_email.write(row_email, col, company_name, cell_format)
                    worksheet_email.write(row_email, col + 1, phone, cell_format)
                    worksheet_email.write(row_email, col + 2, mailto, cell_format)
                    worksheet_all.write(row, col + 2, mailto, cell_format)
                    worksheet_email.write(row_email, col + 3, website, cell_format)
                    worksheet_email.write(row_email, col + 4, filename, cell_format)
                    row_email += 1
                row += 1
            print("Parsed companies", parsed_element_for_page)

    workbook.close()


if __name__ == "__main__":

    with Pool() as pool:
        pool.map(get_html_pages, range(1, ITERATION_COUNT + 1))

    parse_page()
