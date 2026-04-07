"""
Генерация Excel-отчёта по одному магазину (3 листа: на исходе + нет на складе + все товары).
"""

import re
import logging

import pandas as pd
from openpyxl.comments import Comment
from openpyxl.styles import Font

from bot.config import (
    THRESHOLD_A, THRESHOLD_B, THRESHOLD_C,
    DEFAULT_REFILL_DAYS_1, DEFAULT_REFILL_DAYS_2, DEFAULT_REFILL_DAYS_3,
    DEFAULT_REFILL_RESERVE_PCT,
)
from bot.services.calculations import merge_wh_by_name, wh_compact_str, calc_refill_qty
from bot.reports.excel_styles import (
    HYPERLINK_FONT, LEFT,
    WB_PRODUCT_URL,
    apply_header_style,
    apply_data_style,
    apply_hyperlinks,
    apply_group_colors,
    apply_legend_style,
    add_stock_comment,
    add_refill_dropdown,
    make_report_path,
)

logger = logging.getLogger(__name__)


def _build_wh_index(warehouse_rows: list[dict], product_rows: list[dict]) -> dict:
    """
    Строит индекс складов по nm_id.

    Returns:
        {nm_id: [{warehouse_name, quantity, in_way_from_client}]}
    """
    index = {}
    for wr in warehouse_rows:
        nm_id = wr['nm_id']
        index.setdefault(nm_id, []).append({
            'warehouse_name': wr['warehouse_name'],
            'quantity': wr['quantity'],
            'in_way_from_client': wr.get('in_way_from_client', 0),
        })
    return index


def _apply_stock_comments(ws, df: pd.DataFrame, stock_col: int, wh_index: dict):
    """
    Добавляет Comment с детализацией по складам к ячейкам остатка.
    Значение ячейки (stock_qty_clean) уже записано как число через DataFrame export.
    """
    for i, (_, row) in enumerate(df.iterrows()):
        row_num = i + 2  # строка 1 — шапка
        nm_id = int(row['nm_id'])

        # Комментарий с детализацией по складам
        cell = ws.cell(row=row_num, column=stock_col)
        wh_list = wh_index.get(nm_id, [])
        if wh_list:
            wh_list = merge_wh_by_name(wh_list)
            add_stock_comment(cell, wh_list)


def _apply_price_comments(ws, df: pd.DataFrame, price_col: int):
    """
    Добавляет комментарии с рекомендацией повышения цены к ячейкам цены.
    """
    for i, (_, row) in enumerate(df.iterrows()):
        row_num = i + 2
        pct = row.get('price_increase_pct', 0)
        if pct and pct > 0:
            cell = ws.cell(row=row_num, column=price_col)
            comment = Comment(f"Рекомендация: повысить на {pct}%", "WB Analiz")
            comment.width = 200
            comment.height = 50
            cell.comment = comment


def _format_days_remaining(ws, df: pd.DataFrame, days_col: int):
    """Заменяет пустые ячейки 'Дней осталось' на '—' для товаров без продаж."""
    for i, (_, row) in enumerate(df.iterrows()):
        if row.get('avg_per_day', 0) == 0:
            cell = ws.cell(row=i + 2, column=days_col)
            cell.value = "—"
            cell.number_format = "@"


def _apply_product_links(ws, article_col: int, nm_id_col: int):
    """
    Добавляет кликабельные ссылки на карточку товара в ячейки артикула.
    """
    for row in ws.iter_rows(min_row=2):
        article_cell = row[article_col - 1]
        nm_id_cell = row[nm_id_col - 1]
        nm_id = nm_id_cell.value
        if nm_id is not None and article_cell.value is not None:
            article_cell.hyperlink = WB_PRODUCT_URL.format(nm_id)
            article_cell.font = HYPERLINK_FONT
            article_cell.alignment = LEFT


def _style_wh_column(ws, wh_col: int, num_rows: int):
    """Стилизует колонку детализации складов (font 9/gray, без переноса текста)."""
    wh_font = Font(name="Arial", size=9, color="555555")
    for row_num in range(2, num_rows + 2):
        cell = ws.cell(row=row_num, column=wh_col)
        if cell.value is not None:
            cell.font = wh_font
            cell.alignment = LEFT


def _build_refill_sheet(
    writer,
    df_running_out: pd.DataFrame,
    refill_days: list[int],
    reserve_pct: float,
    wh_index: dict,
):
    """
    Создаёт лист «Пополнение» с dropdown-ячейками объёма.

    Колонки:
        1 Артикул | 2 (пусто) | 3 Баркод |
        4 Объём пополнения ▼ | 5 (пусто) | 6 Остатки по складам

    df_running_out — тот же отфильтрованный набор, что на лист «На исходе».
    refill_days — три порога (возрастающие).
    reserve_pct — процент запаса надбавки.
    """
    d1, d2, d3 = refill_days[0], refill_days[1], refill_days[2]

    # Готовим плоский DataFrame для to_excel (значение по умолчанию — средний порог d2)
    rows_out = []
    per_row_options: list[list[str]] = []  # для dropdown
    for _, r in df_running_out.iterrows():
        avg = r.get('avg_per_day', 0) or 0
        q1 = calc_refill_qty(avg, d1, reserve_pct)
        q2 = calc_refill_qty(avg, d2, reserve_pct)
        q3 = calc_refill_qty(avg, d3, reserve_pct)
        default_val = f"{d2}д: {q2}"
        options = [f"{d1}д: {q1}", f"{d2}д: {q2}", f"{d3}д: {q3}"]
        per_row_options.append(options)

        rows_out.append({
            'Артикул': r.get('supplier_article', ''),
            '': '',  # разделитель (колонка 2)
            'Баркод': r.get('barcode', '') or '',
            'Объём пополнения': default_val,
            ' ': '',  # разделитель (колонка 5)
            'Остатки по складам': wh_compact_str(wh_index.get(int(r['nm_id']), [])),
        })

    export_df = pd.DataFrame(
        rows_out,
        columns=['Артикул', '', 'Баркод', 'Объём пополнения', ' ', 'Остатки по складам'],
    )
    export_df.to_excel(writer, sheet_name='Пополнение', index=False)
    ws = writer.sheets['Пополнение']

    # Стили шапки/ширины
    apply_header_style(
        ws,
        {1: 26, 2: 2, 3: 20, 4: 22, 5: 2, 6: 36},
        row_height=34,
        auto_filter=False,
    )
    # Базовые стили данных (LEFT в 1-й колонке, центр в остальных)
    apply_data_style(ws, float_cols=[], int_cols=[], row_height=None)

    # Ссылка на карточку товара по артикулу (пройдёмся по всем строкам)
    nm_ids = df_running_out['nm_id'].tolist()
    for i, nm_id in enumerate(nm_ids):
        row_num = i + 2
        article_cell = ws.cell(row=row_num, column=1)
        if article_cell.value and nm_id is not None:
            article_cell.hyperlink = WB_PRODUCT_URL.format(int(nm_id))
            article_cell.font = HYPERLINK_FONT
            article_cell.alignment = LEFT

    # Баркод — текстовый формат, чтобы длинные числа не ломались
    for i in range(len(df_running_out)):
        row_num = i + 2
        ws.cell(row=row_num, column=3).number_format = "@"

    # Dropdown на «Объём пополнения» (колонка 4) — один DataValidation на строку
    for i, options in enumerate(per_row_options):
        row_num = i + 2
        cell = ws.cell(row=row_num, column=4)
        add_refill_dropdown(ws, cell, options)

    # Колонка 6 — стиль складов (серый, мелкий, выравнивание слева)
    wh_font = Font(name="Arial", size=9, color="555555")
    for i in range(len(df_running_out)):
        row_num = i + 2
        c = ws.cell(row=row_num, column=6)
        if c.value is not None:
            c.font = wh_font
            c.alignment = LEFT
        # Комментарий с детализацией по складам
        nm_id = int(df_running_out.iloc[i]['nm_id'])
        raw_wh = wh_index.get(nm_id, [])
        if raw_wh:
            merged = merge_wh_by_name(raw_wh)
            add_stock_comment(c, merged)

    # Легенда под таблицей
    last = len(df_running_out) + 3
    ws.cell(
        row=last, column=1,
        value='— Объём пополнения —',
    )
    ws.cell(
        row=last + 1, column=1,
        value=f'Формула: avg × дни × (1 + {int(reserve_pct)}%/100), округление вверх',
    )
    ws.cell(
        row=last + 2, column=1,
        value=f'Пороги: {d1} / {d2} / {d3} дней',
    )
    ws.cell(
        row=last + 3, column=1,
        value='Для копирования выбранного значения используйте Специальная вставка → только значения',
    )
    legend_font = Font(name="Arial", size=9, italic=True, color="888888")
    title_font = Font(name="Arial", size=9, bold=True, color="888888")
    ws.cell(row=last, column=1).font = title_font
    for offset in (1, 2, 3):
        ws.cell(row=last + offset, column=1).font = legend_font


def generate_report_from_data(
    product_rows: list[dict],
    store_name: str = None,
    days_threshold: int = 7,
    threshold_a: float = THRESHOLD_A,
    threshold_b: float = THRESHOLD_B,
    threshold_c: float = THRESHOLD_C,
    warehouse_rows: list[dict] = None,
    refill_days: list[int] = None,
    refill_reserve_pct: float = DEFAULT_REFILL_RESERVE_PCT,
) -> str:
    """
    Генерирует Excel-отчёт из готовых данных (без обращения к API).

    Args:
        product_rows: список словарей с данными товаров (из БД или fetch_store_data)
        store_name: имя магазина для имени файла
        days_threshold: порог дней остатка (для аннотаций)
        threshold_a: порог группы A (для аннотаций)
        threshold_b: порог группы B (для аннотаций)
        warehouse_rows: данные по складам (опционально, для детализации остатков)

    Returns:
        Путь к сгенерированному файлу
    """
    logger.info("=== Генерация Excel-отчёта ===")

    # Дефолты для новых параметров (если вызвано из старого кода)
    if refill_days is None or len(refill_days) != 3:
        refill_days = [DEFAULT_REFILL_DAYS_1, DEFAULT_REFILL_DAYS_2, DEFAULT_REFILL_DAYS_3]

    df = pd.DataFrame(product_rows)

    if df.empty:
        logger.warning("Нет данных для отчёта")
        df = pd.DataFrame(columns=[
            'nm_id', 'supplier_article', 'barcode', 'subject', 'category',
            'product_group', 'stock_qty', 'in_way_from_client', 'stock_qty_clean',
            'orders_7d', 'orders_14d', 'orders_30d',
            'avg_per_day', 'days_remaining', 'price_increase_pct', 'price',
        ])

    # Fallback для barcode (старые данные из БД могут не содержать колонку)
    if 'barcode' not in df.columns:
        df['barcode'] = ''
    df['barcode'] = df['barcode'].fillna('')

    # Индекс складов для комментариев и детализации
    wh_index = _build_wh_index(warehouse_rows, product_rows) if warehouse_rows else {}

    # Детализация остатков по складам (компактная строка)
    df['wh_detail'] = df.apply(
        lambda r: wh_compact_str(wh_index.get(int(r['nm_id']), []))
        if (r.get('stock_qty_clean') or 0) > 0 else "—",
        axis=1,
    )

    # Лист 1: Товары на исходе (есть остаток И нужно повышение цены)
    report = df[(df['stock_qty'] > 0) & (df['price_increase_pct'] > 0)].copy()
    report = report.sort_values('days_remaining')
    logger.info(f"Товаров на исходе: {len(report)}")

    # Лист 2: Товары с нулевым остатком
    out_of_stock = df[df['stock_qty'] == 0].copy()
    out_of_stock = out_of_stock.sort_values(['product_group', 'avg_per_day'], ascending=[True, False])
    logger.info(f"Товаров с нулевым остатком: {len(out_of_stock)}")

    # Лист 3: Все товары
    all_products = df.copy()
    all_products = all_products.sort_values('avg_per_day', ascending=False)
    logger.info(f"Всего товаров: {len(all_products)}")

    # Путь к файлу
    safe_name = re.sub(r'[^\w\s-]', '', store_name).strip()[:50] if store_name else ""
    name_part = f"_{safe_name}" if safe_name else ""
    output_path = make_report_path(f'price_report{name_part}')

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # ── Лист 1: На исходе ────────────────────────────────────────────
        report_export = report[[
            'supplier_article', 'barcode', 'nm_id', 'product_group',
            'stock_qty_clean', 'in_way_from_client',
            'avg_per_day', 'days_remaining', 'price', 'wh_detail'
        ]].copy()
        report_export.columns = [
            'Артикул', 'Баркод', 'ID (WB)', 'Группа',
            'Остаток\n(чистый)', '_iwfc',
            'Продаж/день', 'Дней осталось', 'Цена ₽', 'Остатки по складам'
        ]
        report_export.to_excel(writer, sheet_name='На исходе', index=False)

        # ── Лист 2: Нет на складе ───────────────────────────────────────
        out_of_stock_export = out_of_stock[[
            'supplier_article', 'barcode', 'nm_id', 'product_group', 'avg_per_day', 'price'
        ]].copy()
        out_of_stock_export.columns = ['Артикул', 'Баркод', 'ID (WB)', 'Группа', 'Продаж/день', 'Цена ₽']
        out_of_stock_export.to_excel(writer, sheet_name='Нет на складе', index=False)

        # ── Лист 3: Пополнение (объёмы поставок) ────────────────────────
        _build_refill_sheet(
            writer, report, refill_days, refill_reserve_pct, wh_index,
        )

        # ── Лист 4: Все товары ───────────────────────────────────────────
        all_export = all_products[[
            'supplier_article', 'barcode', 'nm_id', 'product_group',
            'stock_qty_clean', 'in_way_from_client',
            'avg_per_day', 'days_remaining', 'price', 'wh_detail'
        ]].copy()
        all_export.columns = [
            'Артикул', 'Баркод', 'ID (WB)', 'Группа',
            'Остаток\n(чистый)', '_iwfc',
            'Продаж/день', 'Дней осталось', 'Цена ₽', 'Остатки по складам'
        ]
        all_export.to_excel(writer, sheet_name='Все товары', index=False)

        # Итоги внизу листа 2
        ws2 = writer.sheets['Нет на складе']
        last_row = len(out_of_stock_export) + 3
        ws2.cell(row=last_row, column=1, value=f'Товаров с нулевым остатком: {len(out_of_stock)}')
        ws2.cell(row=last_row + 1, column=1, value=f'Упущенные продажи в день: {out_of_stock["avg_per_day"].sum():.2f} шт')

        # ── Форматирование ───────────────────────────────────────────────
        logger.info("Форматирование Excel...")

        ws1 = writer.sheets['На исходе']
        ws3 = writer.sheets['Все товары']

        # --- Лист 1: На исходе ---
        # Колонки: 1=Артикул, 2=Баркод, 3=ID(WB), 4=Группа, 5=Остаток, 6=_iwfc, 7=Продаж/день, 8=Дней осталось, 9=Цена, 10=Остатки по складам
        apply_header_style(ws1, {1: 24, 2: 18, 3: 15, 4: 11, 5: 15, 6: 0, 7: 15, 8: 17, 9: 13, 10: 32})
        apply_data_style(ws1, float_cols=[7, 8, 9], int_cols=[5], row_height=None)
        apply_hyperlinks(ws1, link_col=3)
        apply_group_colors(ws1, group_col=4)

        # Скрыть вспомогательную колонку _iwfc (col 6)
        ws1.column_dimensions['F'].hidden = True

        # Прочерк для товаров без продаж (col 8)
        _format_days_remaining(ws1, report, days_col=8)

        # Комментарии к остаткам (col 5) и ценам (col 9)
        _apply_stock_comments(ws1, report, stock_col=5, wh_index=wh_index)
        _apply_price_comments(ws1, report, price_col=9)

        # Стилизация колонки складов (col 10)
        _style_wh_column(ws1, wh_col=10, num_rows=len(report_export))

        # --- Лист 2: Нет на складе ---
        # Колонки: 1=Артикул, 2=Баркод, 3=ID(WB), 4=Группа, 5=Продаж/день, 6=Цена
        apply_header_style(ws2, {1: 24, 2: 18, 3: 15, 4: 11, 5: 15, 6: 13})
        apply_data_style(ws2, float_cols=[5, 6], int_cols=[])
        apply_hyperlinks(ws2, link_col=3)
        apply_group_colors(ws2, group_col=4)

        # --- Лист 3: Все товары ---
        # Колонки: 1=Артикул, 2=Баркод, 3=ID(WB), 4=Группа, 5=Остаток, 6=_iwfc, 7=Продаж/день, 8=Дней осталось, 9=Цена, 10=Остатки по складам
        apply_header_style(ws3, {1: 24, 2: 18, 3: 15, 4: 11, 5: 15, 6: 0, 7: 15, 8: 17, 9: 13, 10: 32})
        apply_data_style(ws3, float_cols=[7, 8, 9], int_cols=[5], row_height=None)
        apply_hyperlinks(ws3, link_col=3)
        apply_group_colors(ws3, group_col=4)

        # Скрыть вспомогательную колонку _iwfc (col 6)
        ws3.column_dimensions['F'].hidden = True

        # Прочерк для товаров без продаж (col 8)
        _format_days_remaining(ws3, all_products, days_col=8)

        # Комментарии к остаткам (col 5)
        _apply_stock_comments(ws3, all_products, stock_col=5, wh_index=wh_index)

        # Ссылки на карточку товара по артикулу (col 1 → nmId из col 3)
        _apply_product_links(ws3, article_col=1, nm_id_col=3)

        # Стилизация колонки складов (col 10)
        _style_wh_column(ws3, wh_col=10, num_rows=len(all_export))

        # ── Аннотации под таблицами ──────────────────────────────────────
        ws1_last = len(report_export) + 4
        ws1.cell(row=ws1_last, column=1, value='— Группы товаров (продаж/день) —')
        ws1.cell(row=ws1_last + 1, column=1, value=f'A: ходовые (≥{threshold_a})')
        ws1.cell(row=ws1_last + 2, column=1, value=f'B: средние (≥{threshold_b})')
        ws1.cell(row=ws1_last + 3, column=1, value=f'C: редкие (≥{threshold_c})')
        ws1.cell(row=ws1_last + 4, column=1, value=f'D: почти не продаются (<{threshold_c})')
        ws1.cell(row=ws1_last + 5, column=1, value=f'Порог повышения цены: ≤{days_threshold} дней остатка')
        apply_legend_style(ws1, ws1_last, has_threshold_row=True)

        ws2_last = last_row + 4
        ws2.cell(row=ws2_last, column=1, value='— Группы товаров (продаж/день) —')
        ws2.cell(row=ws2_last + 1, column=1, value=f'A: ходовые (≥{threshold_a})')
        ws2.cell(row=ws2_last + 2, column=1, value=f'B: средние (≥{threshold_b})')
        ws2.cell(row=ws2_last + 3, column=1, value=f'C: редкие (≥{threshold_c})')
        ws2.cell(row=ws2_last + 4, column=1, value=f'D: почти не продаются (<{threshold_c})')
        apply_legend_style(ws2, ws2_last, has_threshold_row=False)

    logger.info(f"✓ Отчёт сохранён: {output_path}")
    logger.info("=== Генерация завершена ===")

    return output_path
