"""
https://www.upwork.com/jobs/~01e147996dbc5ce0fc
"""
import os
import requests
import datetime
from bs4 import BeautifulSoup
import xlsxwriter


ITERATION_COUNT = 8


def get_pages():
    base_url = 'https://api-g.weedmaps.com/discovery/v1/listings?filter%5Bany_retailer_services%5D%5B%5D=' \
               'storefront&filter%5Bbounding_box%5D=' \
               '24.647017162630366%2C-131.79199218750003%2C38.20365531807151%2C-109.29199218750001&page_size=100&' \
               'page='

    for i in range(1, 12):
        full_address = base_url + str(i)
        resp = requests.get(full_address)
        resp_json = resp.json()

        listings = resp_json.get('data').get('listings')
        print(len(listings))
        page_count = 0
        for data in listings:
            web_url = data.get('web_url')
            detail_url = web_url + '/about'
            print(detail_url)

            html_page_resp = requests.get(detail_url)

            name = data.get('name')

            html_filename = 'weedmaps_pages\\page-' + str(i) + '_' + str(page_count) + '_' + ''.join(e for e in name if e.isalnum()) + '.html'
            page_count += 1
            current_dir = os.path.dirname(__file__)
            filemame = os.path.join(current_dir, html_filename)
            with open(filemame, "w", encoding="utf-8") as f:
                f.write(html_page_resp.text)


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
    worksheet.set_column('C:C', 14)
    worksheet.set_column('D:D', 20)
    worksheet.set_column('E:E', 14)
    worksheet.set_column('F:F', 42)
    worksheet.set_column('G:G', 20)

    worksheet.write('A1', 'Name', bold)
    worksheet.write('B1', 'Location', bold)
    worksheet.write('C1', 'Email', bold)
    worksheet.write('D1', 'Website', bold)
    worksheet.write('E1', 'Phone', bold)
    worksheet.write('F1', 'Day and Hours', bold)
    worksheet.write('G1', 'Saved page', bold)


def parse_pages():
    filenames = []
    file_folder = os.getcwd() + "\\weedmaps_pages"
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
        full_filename = 'weedmaps_pages\\' + filename
        # print(full_filename)
        with open(full_filename, "r", encoding="utf-8") as html_file:
            try:
                soup = BeautifulSoup(html_file.read(), "lxml")
                name = "None"
                try:
                    name = soup.find("h1", class_="styled-components__Name-soafp9-0 cWmvtr").text
                    print(name)
                except AttributeError:
                    print("Can't find name")
                location = ''
                try:
                    location = soup.find("p", class_="styled-components__AddressRow-sc-1k0lbjf-2 dwPNra").text
                    print(location)
                except AttributeError:
                    print("Can't find location")
                email = ''
                try:
                    email = soup.find("div", class_="src__Box-sc-1sbtrzs-0 styled-components__DetailGridItem-d53rlt-0 styled-components__Email-d53rlt-3 icSxPE").text
                    print(email)
                except AttributeError:
                    print("Can't find email")
                website = ''
                try:
                    website = soup.find("div", class_="src__Box-sc-1sbtrzs-0 styled-components__DetailGridItem-d53rlt-0 styled-components__Website-d53rlt-4 uWbmk").text
                    print(website)
                except AttributeError:
                    print("Can't find website")
                phone = ''
                try:
                    phone = soup.find("div", class_="src__Box-sc-1sbtrzs-0 styled-components__DetailGridItem-d53rlt-0 styled-components__PhoneNumber-d53rlt-8 cMfVkr").text
                    print(phone)
                except AttributeError:
                    print("Can't find phone")
                day_hours = ''
                try:
                    day_hours = soup.find("div", class_="src__Box-sc-1sbtrzs-0 styled-components__DetailGridItem-d53rlt-0 styled-components__OpenHours-d53rlt-1 cBJxsr").text
                    print(day_hours)
                except AttributeError:
                    print("Can't find day_hours")

                worksheet_all.write(row, col, name, cell_format)
                worksheet_all.write(row, col + 1, str(location), cell_format)
                worksheet_all.write(row, col + 2, email, cell_format)
                worksheet_all.write(row, col + 3, str(website), cell_format)
                worksheet_all.write(row, col + 4, str(phone), cell__wrapped_format)
                worksheet_all.write(row, col + 5, str(day_hours), cell_format)
                worksheet_all.write(row, col + 6, full_filename, cell_format)

            except AttributeError:
                print(full_filename)
            finally:
                row += 1

    workbook.close()


if __name__ == "__main__":
    # get_pages()
    parse_pages()
