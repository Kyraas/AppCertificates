# import win32api

# print(win32api.FormatMessage(-2147221014))

from datetime import datetime

from bs4 import BeautifulSoup  # Парсинг полученного контента
from requests import get, exceptions    # Получение HTTP-запросов.
                                        # Удобнее и стабильнее в работе,
                                        # чем встроенная библиотека urllib.


FSB = "http://clsz.fsb.ru/clsz/certification.htm"
HEADERS_FSB = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.150 Safari/537.36 OPR/74.0.3911.160",
    "Referer": FSB
    }


# Основная функция
def parse():
    html = None
    url = FSB
    headers = HEADERS_FSB

    try:
        html = get(url, headers=headers, verify=False)
        if html.status_code == 200: # Код ответа об успешном статусе
                                    # "The HTTP 200 OK"
            get_file_name(html.text)
        else:
            return False
    except exceptions.ConnectionError:
        return False


def get_file_name(html):
    soup = BeautifulSoup(html, "html.parser").find("section", class_="main clearfix")
    data = soup.find("a", href=True)
    print(data['href'])

parse()
