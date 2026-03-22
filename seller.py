import io
import logging.config
import os
import re
import zipfile
from environs import Env

import pandas as pd
import requests

logger = logging.getLogger(__file__)


def get_product_list(last_id, client_id, seller_token):
    """Получить список товаров магазина Озон.

    Данный список имеет ограничение в 1000 позиций.
    Args:
        last_id (str): 	Идентификатор последнего значения списка.
        client_id (str): Клиентский идентификатор из личного кабинета продавца
        в разделе "Seller API".
        seller_token (str): Уникальный ключ из личного кабинета продавца
        в разделе "API key".

    Returns:
        dict: Словарь с товарами продавца Озон.

    Examples:
        Корректное исполнение:
        >>> get_product_list("", "123456", "token")
        {'items': [{'offer_id': '136748', 'product_id': 223681945}], 'total': 1}

        Некорректное исполнение (передан неверный параметр):
        >>> get_product_list("", "abc", "token")
        requests.exceptions.HTTPError: 404 Client Error: Not Found
    """
    url = "https://api-seller.ozon.ru/v2/product/list"
    headers = {
        "Client-Id": client_id,
        "Api-Key": seller_token,
    }
    payload = {
        "filter": {
            "visibility": "ALL",
        },
        "last_id": last_id,
        "limit": 1000,
    }
    response = requests.post(url, json=payload, headers=headers)
    response.raise_for_status()
    response_object = response.json()
    return response_object.get("result")


def get_offer_ids(client_id, seller_token):
    """Получить список артикулов товаров в магазине Озон.

    Args:
        client_id (str): Клиентский идентификатор из личного кабинета продавца
        в разделе "Seller API".
        seller_token (str): Уникальный ключ из личного кабинета продавца
        в разделе "API key".

    Returns:
        list: Список с артикулами товаров продавца Озон.

    Examples:
        Корректное исполнение:
        >>> get_offer_ids("123456", "token")
        ['136748', '136749', '136750']

        Некорректное исполнение (передан неверный параметр):
        >>> get_offer_ids("abc", "token")
        requests.exceptions.HTTPError: 404 Client Error: Not Found
    """
    last_id = ""
    product_list = []
    while True:
        some_prod = get_product_list(last_id, client_id, seller_token)
        product_list.extend(some_prod.get("items"))
        total = some_prod.get("total")
        last_id = some_prod.get("last_id")
        if total == len(product_list):
            break
    offer_ids = []
    for product in product_list:
        offer_ids.append(product.get("offer_id"))
    return offer_ids


def update_price(prices: list, client_id, seller_token):
    """Обновить цены товаров в магазине Озон.

    Принимает за раз до 1000 позиций.

    Args:
        prices (list): Список артикулов товаров с актуальными ценами.
        client_id (str): Клиентский идентификатор из личного кабинета продавца
        в разделе "Seller API".
        seller_token (str): Уникальный ключ из личного кабинета продавца
        в разделе "API key".

    Returns:
        dict: Словарь с артикулами, ценами товаров и статусом обновления.

    Examples:
        Корректное исполнение:
        >>> update_price([{"offer_id": "136748", "price": "5990"}], "123456",
        "token")
        {'result': [{'offer_id': 'PH8865', 'updated': True, 'errors': []}]}

        Некорректное исполнение (передан неверный параметр):
        >>> update_price([{"offer_id": "136748", "price": "5990"}], "abc",
        "token")
        requests.exceptions.HTTPError: 404 Client Error: Not Found
    """
    url = "https://api-seller.ozon.ru/v1/product/import/prices"
    headers = {
        "Client-Id": client_id,
        "Api-Key": seller_token,
    }
    payload = {"prices": prices}
    response = requests.post(url, json=payload, headers=headers)
    response.raise_for_status()
    return response.json()


def update_stocks(stocks: list, client_id, seller_token):
    """Обновить остатки товаров в магазине Озон.

    Принимает за раз до 100 позиций.

    Args:
        stocks (list): Список артикулов товаров с актуальными остатками.
        client_id (str): Клиентский идентификатор из личного кабинета продавца
        в разделе "Seller API".
        seller_token (str): Уникальный ключ из личного кабинета продавца
        в разделе "API key".

    Returns:
        dict: Словарь с артикулами, количеством товаров и статусом обновления.

    Examples:
        Корректное исполнение:
        >>> update_stocks([{"offer_id": "136748", "stock": 10}], "123456",
        "token")
        {'result': [{'offer_id': 'PH11042', 'updated': True, 'errors': []}]}

        Некорректное исполнение (передан неверный параметр):
        >>> update_stocks([{"offer_id": "136748", "stock": 10}], "abc",
        "token")
        requests.exceptions.HTTPError: 404 Client Error: Not Found
    """
    url = "https://api-seller.ozon.ru/v1/product/import/stocks"
    headers = {
        "Client-Id": client_id,
        "Api-Key": seller_token,
    }
    payload = {"stocks": stocks}
    response = requests.post(url, json=payload, headers=headers)
    response.raise_for_status()
    return response.json()


def download_stock():
    """Получить список остатков с сайта Casio.

    Returns:
        list: Список словарей с параметрами товаров.

    Examples:
        Корректное исполнение:
        >>> download_stock()
        [{'Код': 69791, 'Количество': '>10', 'Цена': '550.00 руб.'}]

        Некорректное исполнение (сервер блокирует доступ):
        >>> download_stock()
        requests.exceptions.HTTPError: 403 Client Error: Forbidden
    """
    # Скачать остатки с сайта
    casio_url = "https://timeworld.ru/upload/files/ostatki.zip"
    session = requests.Session()
    response = session.get(casio_url)
    response.raise_for_status()
    with response, zipfile.ZipFile(io.BytesIO(response.content)) as archive:
        archive.extractall(".")
    # Создаем список остатков часов:
    excel_file = "ostatki.xls"
    watch_remnants = pd.read_excel(
        io=excel_file,
        na_values=None,
        keep_default_na=False,
        header=17,
    ).to_dict(orient="records")
    os.remove("./ostatki.xls")  # Удалить файл
    return watch_remnants


def create_stocks(watch_remnants, offer_ids):
    """Создать остатки для магазина в Озон.

    Соотносятся остатки с сайта Casio по коду с остатками с Озон по id.
    Если товар с Озона отсутсвует в остатках Casio, то он обнуляется.

    Args:
        watch_remnants (list): Список остатков в виде словарей с параметрами
        товаров с сайта Casio.
        offer_ids (list): Список с артикулами товаров продавца Озон.

    Returns:
        list: Список артикулов товаров с актуальными остатками.

    Examples:
        Корректное исполнение:
        >>> create_stocks(watch_remnants, offer_ids)
        [{'offer_id': '136748', 'stock': 100}]

        Некорректное исполнение (передан неверный тип данных):
        >>> create_stocks(200, offer_ids)
        TypeError: 'int' object is not iterable
    """
    # Уберем то, что не загружено в seller
    stocks = []
    for watch in watch_remnants:
        if str(watch.get("Код")) in offer_ids:
            count = str(watch.get("Количество"))
            if count == ">10":
                stock = 100
            elif count == "1":
                stock = 0
            else:
                stock = int(watch.get("Количество"))
            stocks.append({"offer_id": str(watch.get("Код")), "stock": stock})
            offer_ids.remove(str(watch.get("Код")))
    # Добавим недостающее из загруженного:
    for offer_id in offer_ids:
        stocks.append({"offer_id": offer_id, "stock": 0})
    return stocks


def create_prices(watch_remnants, offer_ids):
    """Создать цены для товаров магазина в Озон.

    Переносятся цены с сайта Casio по коду в Озон по id.

    Args:
        watch_remnants (list): Список остатков в виде словарей с параметрами
        товаров с сайта Casio.
        offer_ids (list): Список с артикулами товаров продавца Озон.

    Returns:
        list: Список артикулов товаров с актуальными ценами.

    Examples:
        Корректное исполнение:
        >>> create_prices(watch_remnants, offer_ids)
        [{'auto_action_enabled': 'UNKNOWN', 'currency_code': 'RUB',
        'offer_id': '136748', 'old_price': '0', 'price': '5990'}]

        Некорректное исполнение (передан неверный тип данных):
        >>> create_prices(200, offer_ids)
        TypeError: 'int' object is not iterable
    """
    prices = []
    for watch in watch_remnants:
        if str(watch.get("Код")) in offer_ids:
            price = {
                "auto_action_enabled": "UNKNOWN",
                "currency_code": "RUB",
                "offer_id": str(watch.get("Код")),
                "old_price": "0",
                "price": price_conversion(watch.get("Цена")),
            }
            prices.append(price)
    return prices


def price_conversion(price: str) -> str:
    """Удалить лишние символы и дробную часть в строке с ценой.

    Args:
        price (str): Исходная строка с ценой товара, в которой десятичным
        разделителем является точка.

    Returns:
        str: Строка с ценой только из целого числа.

    Examples:
        Корректное исполнение:
        >>> price_conversion("5'990.00 руб.")
        '5990'

        Некорректное исполнение (неверный разделитель):
        >>> price_conversion("5'990,00 руб.")
        '599000'

        Некорректное исполнение (передано число вместо строки):
        >>> price_conversion(5990)
        AttributeError: 'int' object has no attribute 'split'
    """
    return re.sub("[^0-9]", "", price.split(".")[0])


def divide(lst: list, n: int):
    """Разделить список lst на части по n элементов.

    Args:
        lst (list): Список для разделения.
        n (int): Количество элементов, которые опрелеляют размер каждой части
        списка.

    Returns:
        generator: Генератор списков.

    Examples:
        Корректное исполнение:
        >>> divide([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11], 2)
        <generator object divide at 0x000001EBCF94CAC0>

        Некорректное исполнение (передан неверный тип данных):
        >>> divide([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11], "2")
        TypeError: 'str' object cannot be interpreted as an integer
    """
    for i in range(0, len(lst), n):
        yield lst[i : i + n]


async def upload_prices(watch_remnants, client_id, seller_token):
    offer_ids = get_offer_ids(client_id, seller_token)
    prices = create_prices(watch_remnants, offer_ids)
    for some_price in list(divide(prices, 1000)):
        update_price(some_price, client_id, seller_token)
    return prices


async def upload_stocks(watch_remnants, client_id, seller_token):
    offer_ids = get_offer_ids(client_id, seller_token)
    stocks = create_stocks(watch_remnants, offer_ids)
    for some_stock in list(divide(stocks, 100)):
        update_stocks(some_stock, client_id, seller_token)
    not_empty = list(filter(lambda stock: (stock.get("stock") != 0), stocks))
    return not_empty, stocks


def main():
    env = Env()
    seller_token = env.str("SELLER_TOKEN")
    client_id = env.str("CLIENT_ID")
    try:
        offer_ids = get_offer_ids(client_id, seller_token)
        watch_remnants = download_stock()
        # Обновить остатки
        stocks = create_stocks(watch_remnants, offer_ids)
        for some_stock in list(divide(stocks, 100)):
            update_stocks(some_stock, client_id, seller_token)
        # Поменять цены
        prices = create_prices(watch_remnants, offer_ids)
        for some_price in list(divide(prices, 900)):
            update_price(some_price, client_id, seller_token)
    except requests.exceptions.ReadTimeout:
        print("Превышено время ожидания...")
    except requests.exceptions.ConnectionError as error:
        print(error, "Ошибка соединения")
    except Exception as error:
        print(error, "ERROR_2")


if __name__ == "__main__":
    main()
