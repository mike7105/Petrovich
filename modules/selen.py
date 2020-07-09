# -*- coding: utf-8 -*-
"""functions for scrab Petrovich"""

import json
import string

from selenium.webdriver import Firefox
from selenium.webdriver.remote.webelement import WebElement


def toint(s) -> int:
    """конвертирует в целое, если не получилось в 0

    :param str s: строка
    :return: результируещее целое или 0
    """
    s = s.strip(string.whitespace).replace(' ', '').replace(',', '.')
    if s:
        return int(s)
    else:
        return 0


def tofloat(s) -> float:
    """конвертирует в число, если не получилось 0.0

    :param str s: строка
    :return: результирующее число с плавающей точкой или 0.0
    """
    s = s.strip(string.whitespace).replace(' ', '').replace(',', '.')
    if s:
        return float(s)
    else:
        return 0.0


def getProduct(product, browser) -> dict:
    """Собирает нужные атрибуты товара в словарь

    :param WebElement product: html элемент, где всё искать
    :param Firefox browser: браузер веб драйвер
    :return: словарь с нужными атрибутами товара img name price quant link
    """
    resdict = {}
    browser.execute_script("arguments[0].scrollIntoView();", product)
    art = product.find_element_by_css_selector('span[data-test="product-code"]').text
    art = toint(art)
    resdict[art] = {}

    img = product.find_element_by_class_name('listed-product-img').get_attribute('src')
    resdict[art]['img'] = img

    prname = product.find_element_by_css_selector('span[data-test="product-title"]').text
    resdict[art]['name'] = prname.strip(string.whitespace)

    price = product.find_element_by_css_selector('span[data-test="product-retail-price"]').text
    resdict[art]['price'] = tofloat(price)

    quant = product.find_element_by_css_selector('input[data-test="product-counter"]')
    resdict[art]['quant'] = toint(quant.get_attribute('value'))

    link = product.find_element_by_class_name('listed-product-link').get_attribute('href')
    resdict[art]['link'] = link

    return resdict


def getName(elem) -> str:
    """Получает имя текущего узла

    :param WebElement elem: текущий узел
    :return:
    """
    names = elem.find_elements_by_tag_name('input')
    if names:
        name = names[0].get_attribute('name').strip()
    else:
        name = ""
    return name


def getData(url, jsName):
    """Основн функция, которая парсит страницу

    :param str url: адрес к странице со списком
    :param str jsName: json файл, куда сохранять итог
    """
    # jsName = r'C:\Users\25430\Notebook_test\python-scraping\ремонт 70м2.json'
    # url = "https://moscow.petrovich.ru/estimate/11673276/"

    browser = Firefox()
    browser.get(url)
    browser.set_window_size(1000, 1100)  # maximize_window()
    # browser.implicitly_wait(1)

    button = browser.find_element_by_css_selector('button[data-test="renovation-tab"]')
    button.click()

    home = {}
    # все комнаты
    rooms = browser.find_elements_by_css_selector('div[data-test="room-dropdown"]')
    for i, room in enumerate(rooms):
        room.click()
        name = f"{i}. {getName(room)}"
        print(name)
        home[name] = {}
        # черновые/отделочные материалы
        drafts = room.find_elements_by_css_selector('div[data-test="draft-dropdown"]')
        if drafts:
            for draft in drafts:
                draft.click()
                draftname = getName(draft)
                print("\t" + draftname)
                home[name][draftname] = {}
                # части комнаты
                rparts = draft.find_elements_by_css_selector('div[data-test="room-part-dropdown"]')
                for rpart in rparts:
                    rpart.click()
                    rpartname = getName(rpart)
                    print("\t\t" + rpartname)
                    home[name][draftname][rpartname] = {}
                    # продукты
                    products = rpart.find_elements_by_css_selector('div[data-test="product-block"]')
                    for product in products:
                        for k, v in getProduct(product, browser).items():
                            home[name][draftname][rpartname][k] = v
        else:
            home[name][name] = {}
            home[name][name][name] = {}
            products = room.find_elements_by_css_selector('div[data-test="product-block"]')
            for product in products:
                for k, v in getProduct(product, browser).items():
                    home[name][name][name][k] = v

    # сохраняю собранную информацию в json
    with open(jsName, "w", encoding="utf8") as write_file:
        json.dump(home, write_file, ensure_ascii=False)

    browser.close()


if __name__ == "__main__":
    pass
