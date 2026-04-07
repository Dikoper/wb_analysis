"""
Генерация Excel-отчёта по одному магазину (3 листа: на исходе + нет на складе + все товары).
"""

import re
import logging

import pandas as pd
from openpyxl.comments import Comment
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from bot.config import (
    THRESHOLD_A, THRESHOLD_B, THRESHOLD_C,
    DEFAULT_REFILL_DAYS_1, DEFAULT_REFILL_DAYS_2, DEFAULT_REFILL_DAYS_3,
    DEFAULT_REFILL_RESERVE_PCT,
)
from bot.services.calculations import (
    merge_wh_by_name, wh_compact_str, calc_refill_qty,
    calc_smart_refill_qty, availability_ru, trend_arrow,
)
from bot.reports.excel_styles import (
    HYPERLINK_FONT, LEFT, CENTER, WRAP_LEFT,
    HEADER_BG, HEADER_FG,
    WB_PRODUCT_URL,
    apply_header_style,
    apply_data_style,
    apply_hyperlinks,
    apply_group_colors,
    apply_legend_style,
    add_stock_comment,
    add_days_dropdown,
    make_report_path,
    fill,
    thin_border,
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
    df_refill: pd.DataFrame,
    refill_days: list[int],
    reserve_pct: float,
    wh_index: dict,
):
    """
    Лист «Поставки» — два блока:

    Основной (col 1..7):
        1 Артикул | 2 gap | 3 Баркод | 4 Объём | 5 Срок | 6 gap | 7 Остатки

    Разделитель: col 8 (gap-separator).

    Предложение WB (col 9..15) — зеркально основному, баркод+объём впереди:
        9 Баркод | 10 Объём WB | 11 Оборотность | 12 Простой поставки |
        13 Упущено | 14 Срок распродажи остатка (WB) | 15 Тренд

    Скрытые P/Q/R (16..18) — предрассчитанные значения IF-формулы col 4.

    Шапка двухстрочная: row 1 — merge-баннер блоков, row 2 — подзаголовки.
    Данные начинаются с row 3, freeze_panes='A3'.
    """
    d1, d2, d3 = refill_days[0], refill_days[1], refill_days[2]
    default_days = d2

    sheet_name = 'Поставки'
    # Создаём пустой лист — будем заполнять вручную (двухстрочная шапка не
    # сочетается с pandas.to_excel).
    wb = writer.book
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    writer.sheets[sheet_name] = ws

    # ── Конфигурация колонок ─────────────────────────────────────────────
    col_widths = {
        1: 26, 2: 2, 3: 20, 4: 12, 5: 8, 6: 2, 7: 36,
        8: 3,
        9: 20, 10: 12, 11: 16, 12: 14, 13: 11, 14: 14, 15: 12,
        16: 0, 17: 0, 18: 0,  # скрытые P/Q/R
    }
    for col_num, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_num)].width = width
    for col_letter in ('P', 'Q', 'R'):
        ws.column_dimensions[col_letter].hidden = True

    # ── Row 1: merge-баннеры ─────────────────────────────────────────────
    banner_font = Font(name='Arial', size=11, bold=True, color='1A1A2E')
    banner_main_fill = PatternFill('solid', fgColor='E7EEF7')
    banner_wb_fill = PatternFill('solid', fgColor='FFE699')
    banner_align = Alignment(horizontal='center', vertical='center')

    ws.merge_cells('A1:H1')
    c = ws.cell(row=1, column=1, value='— Основной расчёт —')
    c.font = banner_font
    c.fill = banner_main_fill
    c.alignment = banner_align
    # Применяем заливку ко всем ячейкам merged-диапазона (для рамок)
    for col in range(1, 9):
        ws.cell(row=1, column=col).fill = banner_main_fill

    ws.merge_cells('I1:O1')
    c = ws.cell(row=1, column=9, value='ПРЕДЛОЖЕНИЕ WB')
    c.font = banner_font
    c.fill = banner_wb_fill
    c.alignment = banner_align
    for col in range(9, 16):
        ws.cell(row=1, column=col).fill = banner_wb_fill

    ws.row_dimensions[1].height = 22

    # ── Row 2: подзаголовки ──────────────────────────────────────────────
    headers = [
        'Артикул', '', 'Баркод', 'Объём', 'Срок (дн)', '', 'Остатки по складам',
        '',
        'Баркод', 'Объём WB', 'Оборотность', 'Простой поставки',
        'Упущено заказов', 'Срок рапродажи остатка (WB)', 'Тренд',
    ]
    header_font = Font(name='Arial', size=11, bold=True, color=HEADER_FG)
    header_fill = PatternFill('solid', fgColor=HEADER_BG)
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for idx, title in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=idx, value=title)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border('444444')
    ws.row_dimensions[2].height = 34

    # Комментарий к шапке «Срок распродажи (WB)» — объясняет, что это прогноз WB
    sale_rate_header = ws.cell(row=2, column=14)
    sale_rate_header.comment = Comment(
        "Прогноз от WB: за сколько дней при текущей динамике спроса "
        "распродастся текущий остаток.\n\n"
        "Источник: поле saleRate в Stocks Report API. Это собственный "
        "алгоритм WB (учитывает сезонность, тренды, дефицит), может "
        "сильно отличаться от простого остаток/ср.продажи.\n\n"
        "Показано справочно — в расчёте объёма не используется.",
        "WB Analiz",
    )

    ws.freeze_panes = 'A3'

    # ── Стили рамок и заливок ───────────────────────────────────────────
    medium = Side(style='medium', color='000000')
    thick_border = Border(left=medium, right=medium, top=medium, bottom=medium)
    gap_fill = PatternFill('solid', fgColor='F2F2F2')
    sep_fill = PatternFill('solid', fgColor='D9D9D9')
    burning_fill = PatternFill('solid', fgColor='FFCCCC')
    wh_font = Font(name='Arial', size=9, color='555555')
    data_font = Font(name='Arial', size=10, color='1A1A2E')
    data_border = thin_border()

    # Толстая рамка для шапок «Баркод»/«Объём» (col 3, 4 и 9, 10)
    for col in (3, 4, 9, 10):
        ws.cell(row=2, column=col).border = thick_border

    # ── Заполнение строк данных ──────────────────────────────────────────
    for i, (_, r) in enumerate(df_refill.iterrows()):
        row_num = i + 3  # данные с 3-й строки (1=баннер, 2=подзаголовки)
        avg = float(r.get('avg_per_day') or 0)
        nm_id = int(r['nm_id']) if r.get('nm_id') is not None else None
        barcode = r.get('barcode', '') or ''
        article = r.get('supplier_article', '') or ''

        # ── Основной блок (1..7) ─────────────────────────────────────
        # 1: Артикул (гиперссылка)
        a_cell = ws.cell(row=row_num, column=1, value=article)
        a_cell.font = data_font
        a_cell.alignment = LEFT
        a_cell.border = data_border
        if article and nm_id is not None:
            a_cell.hyperlink = WB_PRODUCT_URL.format(nm_id)
            a_cell.font = HYPERLINK_FONT

        # 3: Баркод
        b_cell = ws.cell(row=row_num, column=3, value=barcode)
        b_cell.font = data_font
        b_cell.alignment = CENTER
        b_cell.number_format = '@'
        b_cell.border = thick_border

        # 4: IF-формула на скрытые P/Q/R
        formula = (
            f"=IF(E{row_num}={d1},P{row_num},"
            f"IF(E{row_num}={d2},Q{row_num},R{row_num}))"
        )
        d_cell = ws.cell(row=row_num, column=4, value=formula)
        d_cell.font = data_font
        d_cell.alignment = CENTER
        d_cell.number_format = '0'
        d_cell.border = thick_border

        # 5: Срок (dropdown)
        e_cell = ws.cell(row=row_num, column=5, value=default_days)
        e_cell.font = data_font
        e_cell.alignment = CENTER
        e_cell.border = data_border
        add_days_dropdown(ws, e_cell, refill_days)

        # 7: Остатки по складам (wrap)
        wh_str = wh_compact_str(wh_index.get(nm_id, [])) if nm_id is not None else '—'
        wh_cell = ws.cell(row=row_num, column=7, value=wh_str)
        wh_cell.font = wh_font
        wh_cell.alignment = WRAP_LEFT
        wh_cell.border = data_border
        raw_wh = wh_index.get(nm_id, []) if nm_id is not None else []
        if raw_wh:
            add_stock_comment(wh_cell, merge_wh_by_name(raw_wh))

        # ── Блок «Предложение WB» (9..16) ────────────────────────────
        avail = (r.get('availability') or '') if isinstance(r.get('availability'), str) else ''
        miss = float(r.get('office_missing_days') or 0)
        lost = float(r.get('lost_orders') or 0)
        trend = float(r.get('trend_pct') or 0)
        sale_rate = float(r.get('sale_rate_days') or 0)

        smart_qty = calc_smart_refill_qty(
            avg, default_days, reserve_pct, avail, miss, trend,
        )

        # 9: Баркод (зеркально col 3)
        c = ws.cell(row=row_num, column=9, value=barcode)
        c.font = data_font
        c.alignment = CENTER
        c.number_format = '@'
        c.border = thick_border

        # 10: Объём WB
        c = ws.cell(row=row_num, column=10, value=smart_qty)
        c.font = data_font
        c.alignment = CENTER
        c.number_format = '0'
        c.border = thick_border

        # 11: Оборотность
        c = ws.cell(row=row_num, column=11, value=availability_ru(avail))
        c.font = data_font
        c.alignment = CENTER
        c.border = data_border

        # 12: Простой поставки
        c = ws.cell(row=row_num, column=12, value=f'{miss:.0f} / 14 дн')
        c.font = data_font
        c.alignment = CENTER
        c.border = data_border

        # 13: Упущено заказов (+ красная подсветка).
        # Проверяем округлённое значение: 0.3 → round=0 → '—' без заливки.
        lost_rounded = round(lost)
        c = ws.cell(row=row_num, column=13,
                    value=lost_rounded if lost_rounded > 0 else '—')
        c.font = data_font
        c.alignment = CENTER
        c.border = data_border
        if lost_rounded > 0:
            c.fill = burning_fill

        # 14: Срок продаж (WB) — прогноз WB (saleRate)
        sr_val = f'{int(max(0, sale_rate))} дн' if sale_rate > 0 else '—'
        c = ws.cell(row=row_num, column=14, value=sr_val)
        c.font = data_font
        c.alignment = CENTER
        c.border = data_border

        # 15: Тренд
        c = ws.cell(row=row_num, column=15, value=trend_arrow(trend))
        c.font = data_font
        c.alignment = CENTER
        c.border = data_border

        # ── Скрытые P/Q/R — предрассчитанные значения для IF ─────────
        ws.cell(row=row_num, column=16, value=calc_refill_qty(avg, d1, reserve_pct))
        ws.cell(row=row_num, column=17, value=calc_refill_qty(avg, d2, reserve_pct))
        ws.cell(row=row_num, column=18, value=calc_refill_qty(avg, d3, reserve_pct))

    # ── Заливка gap-колонок (для всех строк, включая шапку) ─────────────
    last_data_row = len(df_refill) + 2
    for r in range(2, last_data_row + 1):
        for c in (2, 6):
            ws.cell(row=r, column=c).fill = gap_fill
        ws.cell(row=r, column=8).fill = sep_fill

    # ── Легенда под таблицей ────────────────────────────────────────────
    last = last_data_row + 2
    legend_lines = [
        ('— Основной расчёт —', True),
        (f'Расчёт avg × срок × (1 + {int(reserve_pct)}%/100); срок выбирается '
         f'в колонке E: {d1} / {d2} / {d3} дней', False),
        ('Объём в колонке D обновляется автоматически при смене срока', False),
        ('', False),
        ('— Предложение WB —', True),
        (f'Умный расчёт с учётом WB-метрик (срок {d2} дней фиксированный):', False),
        ('  · Оборотность: дефицитный → ×1.25, стабильный → ×1.05, '
         'слабый → ×0.7, неликвид → 0 (не поставлять)', False),
        ('  · Простой поставки: >10д +20%, >5д +10%, >2д +5%', False),
        ('  · Тренд продаж: линейная корректировка, clamp [−15%..+15%]', False),
        ('🔴 Красная подсветка «Упущено заказов» = товар терял продажи в дефицит', False),
        ('Срок «Срок продаж (WB)» — справочный прогноз от WB (наведите курсор на шапку)', False),
    ]
    title_font = Font(name='Arial', size=9, bold=True, color='888888')
    legend_font = Font(name='Arial', size=9, italic=True, color='888888')
    for offset, (text, is_title) in enumerate(legend_lines):
        cell = ws.cell(row=last + offset, column=1, value=text)
        cell.font = title_font if is_title else legend_font


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

        # ── Лист 2: Поставки (объёмы пополнения) ────────────────────────
        # Фильтр: только товары с продажами и заполненным баркодом
        refill_df = report[
            (report['avg_per_day'] > 0) &
            (report['barcode'].astype(str).str.strip() != '')
        ].copy()
        _build_refill_sheet(
            writer, refill_df, refill_days, refill_reserve_pct, wh_index,
        )

        # ── Лист 3: Нет на складе ───────────────────────────────────────
        out_of_stock_export = out_of_stock[[
            'supplier_article', 'barcode', 'nm_id', 'product_group', 'avg_per_day', 'price'
        ]].copy()
        out_of_stock_export.columns = ['Артикул', 'Баркод', 'ID (WB)', 'Группа', 'Продаж/день', 'Цена ₽']
        out_of_stock_export.to_excel(writer, sheet_name='Нет на складе', index=False)

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
