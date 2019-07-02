"""
https://www.upwork.com/jobs/scrape-site-map-1999-pages_~01395f6ed73adb0ace
"""
import requests
from bs4 import BeautifulSoup
import datetime
import xlsxwriter


def get_pages():

    excel_filename = 'Result_imdb_' + datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S') + '.xlsx'
    workbook = xlsxwriter.Workbook(excel_filename, {'strings_to_urls': False})
    worksheet_all = workbook.add_worksheet()

    row = 1
    col = 0

    cell_format = workbook.add_format()

    for i in range(1, 2000):
        print(i)
        full_address = 'https://www.imdb.com/sitemap/title-' + str(i) + '.xml.gz'
        resp = requests.get(full_address)
        # print(resp.text)

        soup = BeautifulSoup(resp.text, "lxml")

        product_name_elements = soup.find_all("loc")
        for elem in product_name_elements:
            # print(elem)
            # print(elem.text)
            worksheet_all.write(row, col, elem.text, cell_format)
            row += 1
            # print(elem.loc)
            # print('__________')
            # name = elem.select('h3')[0].text.replace('®', '')
        # print(resp.content)
    workbook.close()

if __name__ == "__main__":
    get_pages()
    # parse_pages()
