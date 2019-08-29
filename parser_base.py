"""

"""
import os
import requests
import datetime
import logging
from multiprocessing import Pool
from bs4 import BeautifulSoup


PROC_COUNT = 8


class ParserBase:
    """

    """
    def __init__(self):
        self.process_count = 1
        self.result_filename = None

    def run_parsing(self):
        pass

    def parse_page(self):
        pass

    def save_page(self):
        pass

    def save_results(self):
        pass


def get_pages(start_pos=1, proc_count=PROC_COUNT):
    """
    Download all pages in N threads
    :param start_pos:
    :param proc_count:
    :return:
    """

    page_count = 1999
    for i in range(start_pos, page_count + 1, proc_count):
        print(i)
        full_address = 'https://www.imdb.com/sitemap/title-' + str(i) + '.xml.gz'
        resp = requests.get(full_address)
        xml_filename = 'xmls\\title-' + str(i) + '.xml'

        current_dir = os.path.dirname(__file__)
        filename = os.path.join(current_dir, xml_filename)
        with open(filename, "w", encoding="utf-8") as f:
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
    # log_file_name = (
    #     "log_" + datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S") + ".log"
    # )
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        # filename=log_file_name,
    )
    pool = Pool()

    with Pool() as pool:
        pool.map(get_pages, range(1, PROC_COUNT + 1))
    # get_pages()
    # parse_data()
