# -*- coding: utf-8 -*-
"""Собрает в эксель и json собранный заказ в петровиче
Mihail Chesnokov 11.05.2020"""
__version__ = 'Version:1.0'

import sys
import modules.selen
import modules.export
import time


if __name__ == '__main__':
    try:
        jsName = r'O:\Progs\Python\Petrovich\result\ремонт 66м2_v2.json'
        wbName = jsName.replace('json', 'xlsx')
        url = "https://moscow.petrovich.ru/estimate/11701096/"
        # "https://moscow.petrovich.ru/estimate/11701036/"
        # "https://moscow.petrovich.ru/estimate/11673276/"

        startTime = time.time()
        print("Read URL")
        print(f"start time: {time.ctime(startTime)}")
        modules.selen.getData(url, jsName)
        finishTime = time.time()
        print(f"finish time: {time.ctime(finishTime)}\nduration: {(finishTime - startTime) / 60} min")

        startTime = time.time()
        print("Export Excel")
        print(f"start time: {time.ctime(startTime)}")
        modules.export.toExcel(jsName, wbName, "66m2")
        finishTime = time.time()
        print(f"finish time: {time.ctime(finishTime)}\nduration: {(finishTime - startTime)} sec")

    except Exception as e:
        print(e)
        print('Unexpected error:', sys.exc_info()[0])
