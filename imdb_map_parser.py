"""
https://www.upwork.com/jobs/scrape-site-map-1999-pages_~01395f6ed73adb0ace
"""
import os
import requests
from bs4 import BeautifulSoup
import datetime
from multiprocessing import Pool


ITERATION_COUNT = 8


def get_pages(start_pos=1, proc_count=ITERATION_COUNT):

    page_count = 1999
    for i in range(start_pos, page_count + 1, proc_count):
        print(i)
        full_address = 'https://www.imdb.com/sitemap/title-' + str(i) + '.xml.gz'
        resp = requests.get(full_address)
        xml_filename = 'xmls\\title-' + str(i) + '.xml'

        current_dir = os.path.dirname(__file__)
        filemame = os.path.join(current_dir, xml_filename)
        with open(filemame, "w", encoding="utf-8") as f:
            f.write(resp.text)


def parse_data():

    # excel_filename = 'Result_imdb_' + datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S') + '.xlsx'
    csv_filename = 'Result_imdb_' + datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S') + '.csv'
    with open(csv_filename, "w", encoding="utf-8") as f:

        for page_number in range(1, 2000):
            print(page_number)
            filename = 'xmls\\title-' + str(page_number) + '.xml'
            with open(filename, "r", encoding="utf-8") as xml_file:
                soup = BeautifulSoup(xml_file.read(), "lxml")

                product_name_elements = soup.find_all("loc")
                for elem in product_name_elements:
                    f.write(elem.text)
                    f.write('\r')

if __name__ == "__main__":

    # pool = Pool()
    #
    # with Pool() as pool:
    #     pool.map(get_pages, range(1, ITERATION_COUNT + 1))
    # get_pages()
    parse_data()
