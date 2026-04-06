"""
HTTP-клиент WB API: авторизация, retry, эндпоинты.
"""

import json
import logging
import time
import urllib.request
from datetime import datetime, timedelta
from http.client import IncompleteRead
from urllib.error import URLError

import pandas as pd
import os

logger = logging.getLogger(__name__)

# API endpoints
API_SELLER_INFO = 'https://common-api.wildberries.ru/api/v1/seller-info'
API_PRICES = 'https://discounts-prices-api.wildberries.ru/api/v2/list/goods/filter'
API_CONTENT_CARDS = 'https://content-api.wildberries.ru/content/v2/get/cards/list'
API_STOCKS_REPORT = 'https://seller-analytics-api.wildberries.ru/api/v2/stocks-report/products/products'
API_WB_WAREHOUSES = 'https://seller-analytics-api.wildberries.ru/api/analytics/v1/stocks-report/wb-warehouses'


class WBTokenError(Exception):
    """Ошибка авторизации WB API (невалидный или истёкший токен)."""
    pass


class WBApiError(Exception):
    """Общая ошибка WB API."""
    pass


def get_token() -> str:
    """Возвращает токен WB API из переменной окружения WB_TOKEN."""
    token = os.getenv('WB_TOKEN')
    if not token:
        raise ValueError("WB_TOKEN не задан в переменных окружения")
    return token


def fetch_with_retry(url: str, token: str, retries: int = 3, delay: int = 5) -> list:
    """
    Выполняет HTTP запрос с повторными попытками при ошибках сети.

    Args:
        url: URL для запроса
        token: токен авторизации
        retries: количество попыток (по умолчанию 3)
        delay: пауза между попытками в секундах (по умолчанию 5)

    Returns:
        Список данных из JSON ответа

    Raises:
        WBTokenError: при ошибке авторизации (401/403)
        WBApiError: при других HTTP ошибках
    """
    req = urllib.request.Request(url)
    req.add_header('Authorization', token)

    for attempt in range(1, retries + 1):
        try:
            logger.info(f"Попытка {attempt}/{retries}...")
            with urllib.request.urlopen(req, timeout=180) as resp:
                chunks = []
                while True:
                    chunk = resp.read(64 * 1024)
                    if not chunk:
                        break
                    chunks.append(chunk)
                data = json.loads(b''.join(chunks).decode('utf-8'))
            return data if data else []

        except urllib.error.HTTPError as e:
            if e.code == 401:
                raise WBTokenError(
                    f"Токен невалиден или истёк (HTTP {e.code}). "
                    "Обновите токен в настройках бота."
                )
            if e.code == 403:
                raise WBTokenError(
                    f"Нет доступа (HTTP 403). Проверьте, что токен имеет "
                    "нужную категорию доступа и не истёк."
                )
            if e.code == 429:
                logger.warning(f"Rate limit (429), ожидание {delay * 2} сек...")
                time.sleep(delay * 2)
                continue
            raise WBApiError(f"Ошибка WB API: HTTP {e.code}") from e

        except (IncompleteRead, URLError, TimeoutError) as e:
            logger.warning(f"Попытка {attempt}/{retries} не удалась: {e}")
            if attempt < retries:
                logger.info(f"Ожидание {delay} сек перед повтором...")
                time.sleep(delay)
            else:
                logger.error(f"Все {retries} попытки исчерпаны")
                raise


def post_with_retry(url: str, token: str, body: dict, retries: int = 3, delay: int = 5):
    """
    Выполняет POST запрос с JSON body и повторными попытками.

    Args:
        url: URL для запроса
        token: токен авторизации
        body: тело запроса (будет сериализовано в JSON)
        retries: количество попыток
        delay: пауза между попытками в секундах

    Returns:
        Данные из JSON ответа

    Raises:
        WBTokenError: при ошибке авторизации (401/403)
        WBApiError: при других HTTP ошибках
    """
    payload = json.dumps(body).encode('utf-8')
    req = urllib.request.Request(url, data=payload, method='POST')
    req.add_header('Authorization', token)
    req.add_header('Content-Type', 'application/json')

    for attempt in range(1, retries + 1):
        try:
            with urllib.request.urlopen(req, timeout=180) as resp:
                chunks = []
                while True:
                    chunk = resp.read(64 * 1024)
                    if not chunk:
                        break
                    chunks.append(chunk)
                data = json.loads(b''.join(chunks).decode('utf-8'))
            return data if data else {}

        except urllib.error.HTTPError as e:
            if e.code == 401:
                raise WBTokenError(
                    f"Токен невалиден или истёк (HTTP {e.code}). "
                    "Обновите токен в настройках бота."
                )
            if e.code == 403:
                raise WBTokenError(
                    f"Нет доступа (HTTP 403). Проверьте, что токен имеет "
                    "категорию «Аналитика» и не истёк."
                )
            if e.code == 429:
                logger.warning(f"Rate limit (429), ожидание {delay * 2} сек...")
                time.sleep(delay * 2)
                continue
            raise WBApiError(f"Ошибка WB API: HTTP {e.code}") from e

        except (IncompleteRead, URLError, TimeoutError) as e:
            logger.warning(f"POST попытка {attempt}/{retries} не удалась: {e}")
            if attempt < retries:
                time.sleep(delay)
            else:
                logger.error(f"Все {retries} попытки исчерпаны")
                raise


def get_catalog(token: str = None) -> dict:
    """
    Получает полный каталог товаров через Content API (курсорная пагинация).

    Возвращает ВСЕ карточки товаров продавца, включая товары без остатков и продаж.

    Args:
        token: токен WB API

    Returns:
        dict {nmId: {'supplierArticle': str, 'subject': str, 'category': str}}
    """
    if token is None:
        token = get_token()

    catalog = {}
    limit = 100
    cursor = {"limit": limit}

    while True:
        body = {
            "settings": {
                "cursor": cursor,
                "filter": {"withPhoto": -1},
                "sort": {"ascending": False},
            },
        }
        data = post_with_retry(API_CONTENT_CARDS, token, body, retries=2)

        cards = data.get('cards', [])
        if not cards:
            break

        for card in cards:
            nm_id = card.get('nmID')
            if nm_id is None:
                continue
            # Извлекаем баркоды из sizes[].skus[]
            barcodes = []
            for size in card.get('sizes', []):
                barcodes.extend(size.get('skus', []))
            catalog[nm_id] = {
                'supplierArticle': card.get('vendorCode', ''),
                'subject': card.get('subjectName', ''),
                'category': card.get('subjectName', ''),
                'barcode': ', '.join(barcodes) if barcodes else '',
            }

        # Курсорная пагинация: берём cursor из ответа
        resp_cursor = data.get('cursor', {})
        total = resp_cursor.get('total', 0)

        if total < limit:
            break

        # Следующая страница
        cursor = {
            "limit": limit,
            "updatedAt": resp_cursor.get('updatedAt', ''),
            "nmID": resp_cursor.get('nmID', 0),
        }

        # Пауза между страницами (лимит 100 req/min)
        time.sleep(0.7)

    logger.info(f"Каталог Content API: {len(catalog)} карточек")
    return catalog


def get_seller_info(token: str) -> dict:
    """
    Получает информацию о продавце по токену.

    Args:
        token: токен WB API

    Returns:
        dict с информацией о продавце (name, sid, etc.)

    Raises:
        WBTokenError: при невалидном токене
    """
    data = fetch_with_retry(API_SELLER_INFO, token, retries=1)
    return data if isinstance(data, dict) else {}


def get_prices(nm_ids: list = None, token: str = None) -> tuple[dict, dict, dict]:
    """
    Получает цены товаров (после скидки) через Prices API.

    Args:
        nm_ids: список nmId для фильтрации (если None — все)
        token: токен WB API

    Returns:
        tuple (prices, articles, barcodes):
            prices: dict {nmID: discountedPrice} — минимальная цена среди размеров
            articles: dict {nmID: vendorCode} — артикулы продавца
            barcodes: dict {nmID: str} — баркоды (все SKU через ", ")
    """
    if token is None:
        token = get_token()

    nm_set = set(nm_ids) if nm_ids else None
    prices = {}
    articles = {}
    barcodes = {}
    limit = 1000
    offset = 0

    while True:
        url = f"{API_PRICES}?limit={limit}&offset={offset}"
        data = fetch_with_retry(url, token, retries=2)

        if not data:
            break

        goods = data.get('data', {}).get('listGoods', [])
        if not goods:
            break

        for item in goods:
            nm_id = item.get('nmID')
            if nm_set is not None and nm_id not in nm_set:
                continue
            articles[nm_id] = item.get('vendorCode', '')
            sizes = item.get('sizes', [])
            if sizes:
                discounted = [s.get('discountedPrice') for s in sizes
                              if s.get('discountedPrice') is not None]
                prices[nm_id] = min(discounted) if discounted else 0
                # Извлечение баркодов из sizes[].skus[]
                skus = []
                for s in sizes:
                    skus.extend(s.get('skus', []))
                if skus:
                    barcodes[nm_id] = ', '.join(skus)
            else:
                prices[nm_id] = 0

        if len(goods) < limit:
            break
        offset += limit

    logger.info(f"Загружено цен: {len(prices)}, баркодов: {len(barcodes)}")
    return prices, articles, barcodes


def get_stocks_report(token: str = None, period_days: int = 14) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Загружает данные через Seller Analytics Stocks Report.

    Один POST-запрос возвращает остатки и заказы по всем товарам.

    Args:
        token: токен WB API
        period_days: период для данных по заказам (по умолчанию 14)

    Returns:
        (stocks_df, orders_df) — два DataFrame:
        stocks_df: nmId, stock_qty, in_way_from_client, stock_qty_clean, supplierArticle, subject, category
        orders_df: nmId, supplierArticle, subject, category, orders_count_14d, orders_count_7d, avg_per_day

    Raises:
        WBApiError: при ошибках API
        WBTokenError: при невалидном токене
    """
    if token is None:
        token = get_token()

    date_to = datetime.now()
    date_from = date_to - timedelta(days=period_days)

    all_items = []
    offset = 0
    page_limit = 1000

    while True:
        body = {
            "currentPeriod": {
                "start": date_from.strftime("%Y-%m-%d"),
                "end": date_to.strftime("%Y-%m-%d"),
            },
            "stockType": "wb",
            "skipDeletedNm": True,
            "orderBy": {"field": "ordersCount", "mode": "desc"},
            "availabilityFilters": [],
            "offset": offset,
            "limit": page_limit,
        }

        logger.info(f"Stocks Report: offset={offset}, limit={page_limit}")
        data = post_with_retry(API_STOCKS_REPORT, token, body)

        items = []
        if isinstance(data, dict) and isinstance(data.get("data"), dict):
            items = data["data"].get("items", [])
        elif isinstance(data, dict) and isinstance(data.get("data"), list):
            items = data["data"]

        all_items.extend(items)
        logger.info(f"Stocks Report: получено {len(items)} товаров (всего {len(all_items)})")

        if len(items) < page_limit:
            break

        offset += page_limit
        time.sleep(21)  # rate limit: 3 req/min

    if not all_items:
        logger.warning("Stocks Report вернул 0 товаров")
        return pd.DataFrame(), pd.DataFrame()

    # Маппинг в формат stocks_df + orders_df
    stocks_rows = []
    orders_rows = []

    for item in all_items:
        nm_id = item.get("nmID")
        if not nm_id:
            continue

        metrics = item.get("metrics", {})

        stocks_rows.append({
            "nmId": nm_id,
            "stock_qty": metrics.get("stockCount", 0),
            "in_way_from_client": metrics.get("fromClientCount", 0),
            "stock_qty_clean": max(
                metrics.get("stockCount", 0) - metrics.get("fromClientCount", 0), 0
            ),
            "supplierArticle": item.get("vendorCode", ""),
            "subject": item.get("subjectName", ""),
            "category": "",
        })

        orders_rows.append({
            "nmId": nm_id,
            "supplierArticle": item.get("vendorCode", ""),
            "subject": item.get("subjectName", ""),
            "category": "",
            "orders_count_14d": metrics.get("ordersCount", 0),
            "orders_count_7d": 0,  # Stocks Report не разделяет 7д/14д
            "avg_per_day": metrics.get("avgOrders", 0),  # готовый avg от WB
        })

    stocks_df = pd.DataFrame(stocks_rows)
    orders_df = pd.DataFrame(orders_rows)

    logger.info(f"Stocks Report: {len(stocks_df)} товаров загружено")
    return stocks_df, orders_df


def get_warehouse_stocks(token: str = None, nm_ids: list = None) -> pd.DataFrame:
    """
    Загружает остатки по складам через современный wb-warehouses API.

    Замена deprecated GET /api/v1/supplier/stocks.
    Возвращает построчную детализацию: по одной строке на каждый склад+товар.

    Формат ответа API: {"data": {"items": [{nmId, warehouseName, quantity, inWayFromClient, ...}]}}

    Args:
        token: токен WB API (категория Analytics)
        nm_ids: список nmId для фильтрации (если None — все товары)

    Returns:
        DataFrame с колонками: nmId, warehouseName, quantity, inWayToClient, inWayFromClient
    """
    if token is None:
        token = get_token()

    all_rows = []
    offset = 0
    page_limit = 1000

    while True:
        body = {
            "nmIds": nm_ids or [],
            "limit": page_limit,
            "offset": offset,
        }

        logger.info(f"WB Warehouses: offset={offset}, limit={page_limit}")
        data = post_with_retry(API_WB_WAREHOUSES, token, body)

        # Ответ: {"data": {"items": [...]}}
        items = []
        if isinstance(data, dict):
            inner = data.get("data", {})
            if isinstance(inner, dict):
                items = inner.get("items", [])
            elif isinstance(inner, list):
                items = inner
        elif isinstance(data, list):
            items = data

        all_rows.extend(items)
        logger.info(f"WB Warehouses: получено {len(items)} строк (всего {len(all_rows)})")

        if len(items) < page_limit:
            break

        offset += page_limit
        time.sleep(21)  # rate limit: 1 req / 20 sec

    if not all_rows:
        logger.warning("WB Warehouses вернул 0 строк")
        return pd.DataFrame(columns=['nmId', 'warehouseName', 'quantity', 'inWayToClient', 'inWayFromClient'])

    df = pd.DataFrame(all_rows)

    # Проверяем наличие колонок с безопасным fallback
    for col in ['nmId', 'warehouseName', 'quantity', 'inWayToClient', 'inWayFromClient']:
        if col not in df.columns:
            df[col] = 0 if col != 'warehouseName' else ''

    logger.info(f"WB Warehouses: {len(df)} строк загружено")
    return df[['nmId', 'warehouseName', 'quantity', 'inWayToClient', 'inWayFromClient']]
