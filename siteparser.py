# -*- coding: utf-8 -*-

import os
import warnings
from datetime import datetime

from bs4 import BeautifulSoup  # Парсинг полученного контента
from requests import get, exceptions    # Получение HTTP-запросов.
                                        # Удобнее и стабильнее в работе,
                                        # чем встроенная библиотека urllib.

# Константы
FSTEK = "https://reestr.fstec.ru/reg3"
FSB = "http://clsz.fsb.ru/clsz/certification.htm"
  
# Заголовки нужны для того, чтобы сервер
# не посчитал нас за ботов и не заблокировал.
HEADERS_FSTEK = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.150 Safari/537.36 OPR/74.0.3911.160",
    "Referer": FSTEK
    }

HEADERS_FSB = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.150 Safari/537.36 OPR/74.0.3911.160",
    "Referer": FSB
    }

warnings.filterwarnings('ignore', 'Unverified HTTPS request.*')

# Скачиваем файл с таблицей сертификатов с сайта
def download_file(website):
    if website == "FSTEK":
        url = "https://reestr.fstec.ru/reg3?option=com_rajax&module=rfiles&method=download&format=file&mod=209&file=1"
        headers = HEADERS_FSTEK
        file_name = "fstek_reg.csv"
    elif website == "FSB":
        url = "http://clsz.fsb.ru/files/download/svedeniya_po_sertifikatam_15032023.doc"
        headers = HEADERS_FSB
        file_name = "fsb_reg.doc"
    file = os.getcwd() + "\\" + file_name

    r = get(url, headers=headers, verify=False)
    with open(file, 'wb') as f:
        f.write(r.content)


# Получаем дату обновления таблицы
def get_update_date(html):
    # Второй параметр это тип документа, с которым мы работаем
    # (опциональный, но использование желательно)
    soup = BeautifulSoup(html, "html.parser")   
    label = soup.find("dd", class_="modified").text.strip()
    if "Обновлено: " in label:
        text = label.replace("Обновлено: ", "")
    try:
        actual_date = datetime.strptime(text, "%d.%m.%Y %H:%M")
    except:
        actual_date = text
    return actual_date


# Основная функция
def parse(website, get_date=False):
    html = None
    if website == "FSTEK":
        url = FSTEK
        headers = HEADERS_FSTEK
    elif website == "FSB":
        url = FSB
        headers = HEADERS_FSB

    try:
        html = get(url, headers=headers, verify=False)
        if html.status_code == 200: # Код ответа об успешном статусе
                                    # "The HTTP 200 OK"
            if get_date and website == "FSTEK":
                return get_update_date(html.text)
            else:
                download_file(website)
        else:
            return False
    except exceptions.ConnectionError:
        return False
