"""
Code for parsing items from https://customers.unfi.com/
See https://www.upwork.com/jobs/~01c01c44cfe174ac49 for details.
"""

import datetime
import json
import os
import random
import time

import xlsxwriter
from bs4 import BeautifulSoup
from loguru import logger
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.errorhandler import WebDriverException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


def sleep_timeout():
    """
    Make random timeout in range [2;4] secs.
    :return:
    """
    val = random.random() * 4
    if val > 2:
        logger.debug(f'We sleep for {val} secs')
        time.sleep(val)
        return
    return sleep_timeout()


def parse_page(html_filename):
    """
    Exctract necessary data from saved file
    :param html_filename:
    :return:
    """
    result = []
    with open(html_filename, "r", encoding="utf-8") as html_file:
        logger.debug(f'Open page {html_filename}')
        try:
            soup = BeautifulSoup(html_file.read(), "lxml")
            logger.debug(f'Page {html_filename} was loaded')
            table = soup.find('table', attrs={'role': 'grid'})
            table_body = table.find('tbody')
            rows = table_body.find_all('tr')
            logger.debug(f'Found {len(rows)} rows')
            for row in rows:
                cols = row.find_all('td')
                cols = [ele.text.strip() for ele in cols]

                record = \
                    {
                        'brand': cols[2],
                        'UPC': cols[3],
                        'product_code': cols[5],
                        'product_desc': cols[6],
                        'pack_size': cols[7],
                        'min_qty': cols[9],
                        'quantity': cols[10],
                        'availability': cols[14],
                        'disc': cols[15],
                        'price': cols[16],
                        'total_price': cols[17],
                    }
                result.append(record)
        except AttributeError:
            print(html_filename)
    return result


def parse_results(json_filename):
    """
    Parse results from selected data-file.
    :param json_filename:
    :return:
    """
    with open(json_filename, encoding='utf-8') as f:
        data = f.read()
        json_data = json.loads(data)

        results = []
        for rec in json_data:
            url = rec.get('url')
            brand_file = rec.get('brand_file')
            logger.debug(f'Parse data for f{url}')
            items = rec.get('items')
            for html_filename in items:

                items_data = parse_page(html_filename)
                for data in items_data:
                    result_rec = \
                        {
                            'brand_url': url,
                            'brand_file': brand_file,
                            'item_page': html_filename,
                        }
                    result_rec.update(data)
                    results.append(result_rec)
        return results


def save_results(filename, results):
    """
    Save results to Excel
    :param filename:
    :param results:
    :return:
    """
    try:
        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet()
        bold = workbook.add_format({'bold': True})

        worksheet.freeze_panes(1, 0, None, 14)

        headers = \
            [
                "Category url",
                "Category file",
                'Items file',
                'Brand',
                'UPC',
                'Product Code',
                'Product Description',
                'Pack Size',
                'Min Qty',
                'Quantity',
                'Availability',
                'Disc.',
                'Price',
                'Total Price',
            ]

        for index, header in enumerate(headers):
            worksheet.write(0, index, header, bold)

        cell_format = workbook.add_format()
        res_positions = \
            {
                "brand_url": 0,
                "brand_file": 1,
                "item_page": 2,
                "brand": 3,
                "UPC": 4,
                "product_code": 5,
                "product_desc": 6,
                "pack_size": 7,
                "min_qty": 8,
                "quantity": 9,
                "availability": 10,
                "disc": 11,
                "price": 12,
                "total_price": 13,
            }

        for i, res in enumerate(results):
            row_count = i + 1
            for name, value in res.items():
                if name in res_positions:
                    worksheet.write(row_count, res_positions[name], value, cell_format)

        worksheet.autofilter(0, 0, len(results), 25)
        workbook.close()
    except Exception as ex:
        logger.error(f'Cannot save file {filename}: {ex}')


def save_json(json_result, filename_prefix: str):
    json_file_name = filename_prefix + "_" + datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S") + "_.json"

    with open(json_file_name, "w") as f:
        for chunk in json.JSONEncoder(indent=4, ensure_ascii=False).iterencode(
                json_result
        ):
            f.write(chunk)
    return json_file_name


def main():
    """
    Run parsing process.
    :return:
    """
    driver = webdriver.Chrome(ChromeDriverManager().install())

    login_to_site(driver)

    el_links = get_categories_url(driver)
    json_res = []

    try:
        logger.debug(f'Found {len(el_links)} categories')
        for i, brand_link in enumerate(el_links):
            res = {'url': brand_link}
            logger.debug(f'Open {i} brand-link: {brand_link}')
            sleep_timeout()
            res.update(extract_items_pages(brand_link, driver, i))
            sleep_timeout()
            json_res.append(res)
    except Exception as ex:
        logger.error(ex)
    json_file_name = save_json(json_res, 'Results')
    parsing_results = parse_results(json_file_name)
    save_results('Results.xlsx', parsing_results)
    driver.close()
    return


def extract_items_pages(brand_link, driver, i):
    """
    Exctract and save pages
    :param brand_link:
    :param driver:
    :param i:
    :return:
    """
    driver.get(brand_link)
    logger.debug(f'{i} brand-link: {brand_link} opened')
    brand_html_filename = 'htmls\\brand_' + str(i) + '_' + '.html'
    res = {'brand_file': brand_html_filename, 'items': []}
    save_html(driver, brand_html_filename)

    wait = WebDriverWait(driver, 10)
    wait.until(EC.element_to_be_clickable((By.XPATH,
                                           '//*[@id="gridBrandProducts"]/div/span[1]/span/span/span[2]/span')))
    sleep_timeout()

    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    sleep_timeout()

    is_next_page = True
    j = 0
    prev_page = driver.page_source
    while is_next_page:
        try:
            logger.debug(f'Open {j}-items page')
            html_filename = 'htmls\\page_' + str(i) + '_' + str(j) + '_' + '.html'
            logger.debug(f'Save to {html_filename}')
            res['items'].append(html_filename)
            save_html(driver, html_filename)

            contact_link = driver.find_element_by_link_text('Contact us')
            actions = ActionChains(driver)
            actions.move_to_element(contact_link).perform()
            logger.debug('We scrolled to contact_link')

            next_page_elem_click = driver.find_element_by_xpath('//*[@id="gridBrandProducts"]/div/a[3]/span')
            logger.debug('We found next-page element')
            next_page_elem_click.click()
            logger.debug('We clicked next-page element')

            sleep_timeout()
            this_page = driver.page_source
            if this_page == prev_page:
                break
            j += 1
            prev_page = this_page

        except WebDriverException:
            logger.debug("Element is not clickable")
            break

    return res


def get_categories_url(driver):
    """
    Parse categories links.
    :param driver:
    :return:
    """
    driver.get("https://customers.unfi.com/pages/categories.aspx")
    sleep_timeout()
    elems = driver.find_elements_by_css_selector(".brand [href]")
    el_links = [elem.get_attribute('href') for elem in elems]
    return el_links


def login_to_site(driver):
    """
    Process for site login. Currently login and password are hardcoded.
    :param driver:
    :return:
    """
    driver.get("https://customers.unfi.com/")
    logger.debug(f'We open main page')
    sleep_timeout()
    login_el = driver.find_element_by_xpath('//*[@id="userName"]')
    # TODO: Move login and password to command-line args or env-vars.
    login = 'SOME_LOGIN_NAME'
    login_el.send_keys(login)
    sleep_timeout()
    password = 'SOME_PASSWORD!'
    passw_el = driver.find_element_by_xpath('//*[@id="Password"]')
    passw_el.send_keys(password)
    sleep_timeout()
    passw_el.send_keys(Keys.ENTER)
    sleep_timeout()
    logger.debug(f'We login successfully')


def save_html(driver, html_filename):
    """
    Save current page to file
    :param driver:
    :param html_filename:
    :return:
    """
    page = driver.page_source
    folder = os.path.dirname(__file__)
    full_name = os.path.join(folder, html_filename)
    with open(full_name, "w", encoding="utf-8") as f:
        f.write(page)


if __name__ == "__main__":
    main()
