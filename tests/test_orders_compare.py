"""
Тестовый скрипт: сравнение orders vs sales vs funnel для одного товара.
Цель: выяснить почему avg_per_day в боте выше, чем в кабинете WB.
"""

import json
import os
import sqlite3
import base64
import urllib.request
from datetime import datetime, timedelta

# Получаем токен из БД бота
_PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(_PROJECT_ROOT, 'data', 'bot.db')
_OBF_KEY = os.getenv('TOKEN_OBF_KEY', 'wb-analiz-default-key').encode()


def _deobfuscate(stored: str) -> str:
    if not stored.startswith('obf:'):
        return stored
    xored = base64.b64decode(stored[4:])
    return bytes(b ^ _OBF_KEY[i % len(_OBF_KEY)] for i, b in enumerate(xored)).decode()


def get_token_from_db() -> str:
    conn = sqlite3.connect(DB_PATH)
    row = conn.execute('SELECT token FROM stores WHERE is_active = 1 LIMIT 1').fetchone()
    conn.close()
    if not row:
        raise ValueError("Нет активных магазинов в БД")
    return _deobfuscate(row[0])


TOKEN = os.getenv('WB_TOKEN') or get_token_from_db()
print(f"Токен: {TOKEN[:6]}***")

NM_ID = 198215260
DAYS = 14
DATE_FROM = (datetime.now() - timedelta(days=DAYS)).strftime('%Y-%m-%d')
DATE_TO = datetime.now().strftime('%Y-%m-%d')


def api_get(url: str) -> list:
    """GET-запрос к WB API."""
    req = urllib.request.Request(url)
    req.add_header('Authorization', TOKEN)
    with urllib.request.urlopen(req, timeout=120) as resp:
        data = json.loads(resp.read().decode('utf-8'))
    return data or []


def api_post(url: str, body: dict) -> dict:
    """POST-запрос к WB API."""
    req = urllib.request.Request(url, method='POST')
    req.add_header('Authorization', TOKEN)
    req.add_header('Content-Type', 'application/json')
    req.data = json.dumps(body).encode('utf-8')
    with urllib.request.urlopen(req, timeout=120) as resp:
        return json.loads(resp.read().decode('utf-8'))


print(f"=== Сравнение данных для nmID={NM_ID} за {DAYS} дней (с {DATE_FROM}) ===\n")

# 1. Orders API
print("1. Orders API (/api/v1/supplier/orders)...")
orders_url = f"https://statistics-api.wildberries.ru/api/v1/supplier/orders?dateFrom={DATE_FROM}"
orders_data = api_get(orders_url)
orders_for_nm = [o for o in orders_data if o.get('nmId') == NM_ID]
orders_count = len(orders_for_nm)
print(f"   Всего записей в API: {len(orders_data)}")
print(f"   Записей для nmID={NM_ID}: {orders_count}")
print(f"   avg_per_day (orders): {orders_count / DAYS:.2f} шт/день")

# Проверяем, есть ли поле isCancel или подобное
if orders_for_nm:
    sample = orders_for_nm[0]
    print(f"   Поля записи: {list(sample.keys())}")
    # Проверяем статусы отмен
    cancel_count = sum(1 for o in orders_for_nm if o.get('isCancel'))
    print(f"   Из них отменённых (isCancel=True): {cancel_count}")
    non_cancel = orders_count - cancel_count
    print(f"   Без отмен: {non_cancel}, avg_per_day: {non_cancel / DAYS:.2f} шт/день")
print()

# 2. Sales API
print("2. Sales API (/api/v1/supplier/sales)...")
try:
    sales_url = f"https://statistics-api.wildberries.ru/api/v1/supplier/sales?dateFrom={DATE_FROM}"
    sales_data = api_get(sales_url)
    sales_for_nm = [s for s in sales_data if s.get('nmId') == NM_ID]
    # В sales могут быть возвраты (saleID начинается с 'R')
    real_sales = [s for s in sales_for_nm if not str(s.get('saleID', '')).startswith('R')]
    returns = [s for s in sales_for_nm if str(s.get('saleID', '')).startswith('R')]
    print(f"   Всего записей в API: {len(sales_data)}")
    print(f"   Записей для nmID={NM_ID}: {len(sales_for_nm)}")
    print(f"   Из них продажи: {len(real_sales)}, возвраты: {len(returns)}")
    print(f"   avg_per_day (sales, без возвратов): {len(real_sales) / DAYS:.2f} шт/день")
    print(f"   avg_per_day (sales, нетто): {(len(real_sales) - len(returns)) / DAYS:.2f} шт/день")
    if sales_for_nm:
        print(f"   Поля записи: {list(sales_for_nm[0].keys())}")
except Exception as e:
    print(f"   Ошибка: {e}")
print()

# 3. Sales Funnel API
print("3. Sales Funnel API (/api/analytics/v3/sales-funnel/products)...")
try:
    funnel_url = "https://seller-analytics-api.wildberries.ru/api/analytics/v3/sales-funnel/products"
    funnel_body = {
        "nmIDs": [NM_ID],
        "selectedPeriod": {
            "begin": f"{DATE_FROM} 00:00:00",
            "end": f"{DATE_TO} 23:59:59"
        },
        "timezone": "Europe/Moscow",
        "aggregationLevel": "day",
        "page": 1
    }
    funnel_resp = api_post(funnel_url, funnel_body)
    print(f"   Ответ: {json.dumps(funnel_resp, indent=2, ensure_ascii=False)[:1500]}")
except urllib.error.HTTPError as e:
    body = e.read().decode('utf-8') if e.fp else ''
    print(f"   HTTP {e.code}: {body[:300]}")
    if e.code == 401:
        print("   (Возможно нужен отдельный токен типа Analytics)")
except Exception as e:
    print(f"   Ошибка: {e}")

print("\n=== Готово ===")
