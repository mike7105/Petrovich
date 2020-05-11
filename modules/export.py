# -*- coding: utf-8 -*-
"""functions for export"""

import xlsxwriter
from xlsxwriter.worksheet import Worksheet
import json
from io import BytesIO
from urllib.request import urlopen
import urllib.error as urEr


def writeDict(k: str, v: object, r: int, c: int, wsCur: Worksheet, cfnorm: object, prev=None) -> int:
    """записывает словарь в переданный эксель

    :param k: ключ словаря
    :param v: значение ключа
    :param r: строка для записи
    :param c: колонка для записи
    :param wsCur: лист, на которы йнадо записывать
    :param cfnorm: формат закраски ячейки
    :param prev: список предыдущих листов
    :return: строка, в которую записывали
    """

    if prev is None:
        prev = []

    isDict = True
    if isinstance(v, dict):
        wsCur.write(r, c, k, cfnorm)
        for k1, v1 in v.items():
            r = writeDict(k1, v1, r, c + 1, wsCur, cfnorm, prev + [k])

            if not isinstance(v1, dict):
                c += 1
                isDict = False

        if not isDict:
            for i in range(len(prev)):
                wsCur.write(r, i, prev[i], cfnorm)
            wsCur.write_formula(r, 9, "=G{0} * H{0}".format(r + 1), cfnorm)
            r += 1
    else:
        if c == 4:
            try:
                image_data = BytesIO(urlopen(str(v)).read())
                wsCur.insert_image(r, c, v, {'image_data': image_data,
                                             'object_position': 2,
                                             'align': 'center',
                                             'valign': 'vcenter'})
            except (urEr.HTTPError, urEr.URLError) as e:
                print(v)
                print(e)
                wsCur.write_url(r, c, v, cfnorm)
        elif c == 8:
            wsCur.write_url(r, c, v, cfnorm)
        else:
            wsCur.write(r, c, v, cfnorm)
    return r


def toExcel(jsName: str, wbName: str, shName: str = "Sheet1"):
    """Экспортирует json в эксель

    :param jsName: путь к файлу json
    :param wbName: путь к файлу для экспорта
    :param shName: название рабочего листа
    """
    fbold = {
        'bold': 1,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'text_wrap': True
    }
    fnorm = {
        'valign': 'vcenter',
        'border': 1,
        'text_wrap': True
    }
    # Light red fill with dark red text.
    fcondred = {
        'bg_color': '#FFC7CE',
        'font_color': '#9C0006'
    }

    workbook = xlsxwriter.Workbook(wbName)
    cfbold = workbook.add_format(fbold)
    cfnorm = workbook.add_format(fnorm)
    cfcondred = workbook.add_format(fcondred)
    wsCur = workbook.add_worksheet(shName)
    wsCur.set_default_row(108)
    wsCur.set_column('A:C', 22.7)
    wsCur.set_column('D:D', 7.5)
    wsCur.set_column('E:E', 19)
    wsCur.set_column('F:F', 60)
    wsCur.set_column('G:H', 7.5)
    wsCur.set_column('I:I', 39.3)
    wsCur.set_column('J:J', 8)

    wsCur.write(0, 0, "Помещение", cfbold)
    wsCur.write(0, 1, "Отделка", cfbold)
    wsCur.write(0, 2, "Элемент", cfbold)
    wsCur.write(0, 3, "Артикул", cfbold)
    wsCur.write(0, 4, "Картинка", cfbold)
    wsCur.write(0, 5, "Наименование", cfbold)
    wsCur.write(0, 6, "Цена", cfbold)
    wsCur.write(0, 7, "Кол-во", cfbold)
    wsCur.write(0, 8, "Ссылка", cfbold)
    wsCur.write(0, 9, "Сумма", cfbold)

    with open(jsName, "r", encoding="utf8") as f:
        home = json.load(f)

    r = 1
    for k, v in home.items():
        r = writeDict(k, v, r, 0, wsCur, cfnorm)
    wsCur.conditional_format('G1:G{}'.format(r), {'type': 'cell', 'criteria': '=', 'value': 0, 'format': cfcondred})
    wsCur.conditional_format('J1:J{}'.format(r), {'type': 'cell', 'criteria': '=', 'value': 0, 'format': cfcondred})

    workbook.close()


if __name__ == "__main__":
    pass
