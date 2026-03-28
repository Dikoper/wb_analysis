"""
Генерация Excel-отчёта по одному магазину (2 листа: повышение цены + нет на складе).
"""

import os
import re
import logging
from datetime import datetime

import pytz
import pandas as pd

from bot.config import THRESHOLD_A, THRESHOLD_B, REPORTS_DIR
from bot.excel_styles import (
    apply_header_style,
    apply_data_style,
    apply_hyperlinks,
    apply_group_colors,
    apply_price_increase_colors,
    apply_legend_style,
)

logger = logging.getLogger(__name__)


def generate_report_from_data(
    product_rows: list[dict],
    store_name: str = None,
    days_threshold: int = 7,
    threshold_a: float = THRESHOLD_A,
    threshold_b: float = THRESHOLD_B,
) -> str:
    """
    Генерирует Excel-отчёт из готовых данных (без обращения к API).

    Args:
        product_rows: список словарей с данными товаров (из БД или fetch_store_data)
        store_name: имя магазина для имени файла
        days_threshold: порог дней остатка (для аннотаций)
        threshold_a: порог группы A (для аннотаций)
        threshold_b: порог группы B (для аннотаций)

    Returns:
        Путь к сгенерированному файлу
    """
    logger.info("=== Генерация Excel-отчёта ===")

    df = pd.DataFrame(product_rows)

    if df.empty:
        logger.warning("Нет данных для отчёта")
        df = pd.DataFrame(columns=[
            'nm_id', 'supplier_article', 'subject', 'category',
            'product_group', 'stock_qty', 'in_way_from_client', 'stock_qty_clean',
            'orders_7d', 'orders_14d', 'orders_30d',
            'avg_per_day', 'days_remaining', 'price_increase_pct',
        ])

    # Лист 1: Товары для повышения цены (есть остаток И нужно повышение)
    report = df[(df['stock_qty'] > 0) & (df['price_increase_pct'] > 0)].copy()
    report = report.sort_values('days_remaining')
    logger.info(f"Товаров для повышения цены: {len(report)}")

    # Лист 2: Товары с нулевым остатком
    out_of_stock = df[df['stock_qty'] == 0].copy()
    out_of_stock = out_of_stock.sort_values(['product_group', 'avg_per_day'], ascending=[True, False])
    logger.info(f"Товаров с нулевым остатком: {len(out_of_stock)}")

    # Сохранение Excel
    os.makedirs(REPORTS_DIR, exist_ok=True)

    timestamp = datetime.now(pytz.timezone('Europe/Moscow')).strftime('%Y%m%d_%H%M')
    safe_name = re.sub(r'[^\w\s-]', '', store_name).strip()[:50] if store_name else ""
    name_part = f"_{safe_name}" if safe_name else ""
    output_path = os.path.join(REPORTS_DIR, f'price_report{name_part}_{timestamp}.xlsx')

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # ── Лист 1: Повысить цену ──────────────────────────────────────────
        report_export = report[[
            'supplier_article', 'nm_id', 'product_group',
            'stock_qty', 'in_way_from_client',
            'avg_per_day', 'days_remaining', 'price_increase_pct', 'price'
        ]].copy()
        report_export.columns = [
            'Артикул', 'ID (WB)', 'Группа',
            'Остаток (чист.)', 'В возвратах',
            'Продаж/день', 'Дней осталось', 'Повышение %', 'Цена ₽'
        ]
        report_export.to_excel(writer, sheet_name='Повысить цену', index=False)

        # ── Лист 2: Нет на складе ─────────────────────────────────────────
        out_of_stock_export = out_of_stock[[
            'supplier_article', 'nm_id', 'product_group', 'avg_per_day', 'price'
        ]].copy()
        out_of_stock_export.columns = ['Артикул', 'ID (WB)', 'Группа', 'Продаж/день', 'Цена ₽']
        out_of_stock_export.to_excel(writer, sheet_name='Нет на складе', index=False)

        # Итоги внизу листа 2
        ws2 = writer.sheets['Нет на складе']
        last_row = len(out_of_stock_export) + 3
        ws2.cell(row=last_row, column=1, value=f'Товаров с нулевым остатком: {len(out_of_stock)}')
        ws2.cell(row=last_row + 1, column=1, value=f'Упущенные продажи в день: {out_of_stock["avg_per_day"].sum():.2f} шт')

        # Форматирование
        logger.info("Форматирование Excel...")

        ws1 = writer.sheets['Повысить цену']

        # Лист 1
        apply_header_style(ws1, {1: 24, 2: 15, 3: 11, 4: 17, 5: 15, 6: 15, 7: 17, 8: 15, 9: 13})
        apply_data_style(ws1, float_cols=[6, 7, 9], int_cols=[4, 5, 8])
        apply_hyperlinks(ws1, link_col=2)
        apply_group_colors(ws1, group_col=3)
        apply_price_increase_colors(ws1, pct_col=8)

        # Лист 2
        apply_header_style(ws2, {1: 24, 2: 15, 3: 11, 4: 15, 5: 13})
        apply_data_style(ws2, float_cols=[4, 5], int_cols=[])
        apply_hyperlinks(ws2, link_col=2)
        apply_group_colors(ws2, group_col=3)

        # Аннотации под таблицами
        ws1_last = len(report_export) + 4
        ws1.cell(row=ws1_last, column=1, value='— Группы товаров —')
        ws1.cell(row=ws1_last + 1, column=1, value=f'A: ходовые (≥{threshold_a} шт/день за 30д) → среднее по 7 дням')
        ws1.cell(row=ws1_last + 2, column=1, value=f'B: средние (≥{threshold_b} шт/день) → среднее по 14 дням')
        ws1.cell(row=ws1_last + 3, column=1, value=f'C: редкие (<{threshold_b} шт/день) → среднее по 30 дням')
        ws1.cell(row=ws1_last + 4, column=1, value=f'Порог повышения цены: ≤{days_threshold} дней остатка')
        apply_legend_style(ws1, ws1_last, has_threshold_row=True)

        ws2_last = last_row + 4
        ws2.cell(row=ws2_last, column=1, value='— Группы товаров —')
        ws2.cell(row=ws2_last + 1, column=1, value=f'A: ходовые (≥{threshold_a} шт/день за 30д)')
        ws2.cell(row=ws2_last + 2, column=1, value=f'B: средние (≥{threshold_b} шт/день)')
        ws2.cell(row=ws2_last + 3, column=1, value=f'C: редкие (<{threshold_b} шт/день)')
        apply_legend_style(ws2, ws2_last, has_threshold_row=False)

    logger.info(f"✓ Отчёт сохранён: {output_path}")
    logger.info("=== Генерация завершена ===")

    return output_path
