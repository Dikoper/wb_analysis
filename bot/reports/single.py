"""
Генерация Excel-отчёта по одному магазину (3 листа: на исходе + нет на складе + все товары).
"""

import re
import logging

import pandas as pd
from openpyxl.comments import Comment
from openpyxl.styles import Font, Border, Side

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
    add_days_dropdown,
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
    df_refill: pd.DataFrame,
    refill_days: list[int],
    reserve_pct: float,
    wh_index: dict,
):
    """
    Создаёт лист «Поставки» — раздельные ячейки «Объём» и «Срок» с dropdown.

    Колонки (1..7) + 3 скрытые (8..10):
        1 Артикул | 2 (gap) | 3 Баркод | 4 Объём пополнения | 5 Срок (дн) |
        6 (gap) | 7 Остатки по складам | 8 _q1 | 9 _q2 | 10 _q3

    Объём в колонке 4 — простая lookup-формула IF на скрытые H/I/J,
    куда Python кладёт три предрассчитанных значения. Бизнес-логика расчёта
    остаётся в Python (calc_refill_qty), формула в ячейке только выбирает
    нужное значение по сроку из колонки E.
    """
    d1, d2, d3 = refill_days[0], refill_days[1], refill_days[2]
    default_days = d2  # средний порог как значение по умолчанию

    # Готовим плоский DataFrame для to_excel
    rows_out = []
    for _, r in df_refill.iterrows():
        rows_out.append({
            'Артикул': r.get('supplier_article', ''),
            '': '',  # разделитель (колонка 2)
            'Баркод': r.get('barcode', '') or '',
            'Объём пополнения': '',   # будет перезаписана формулой
            'Срок (дн)': default_days,
            ' ': '',  # разделитель (колонка 6)
            'Остатки по складам': wh_compact_str(wh_index.get(int(r['nm_id']), [])),
            '_q1': 0,
            '_q2': 0,
            '_q3': 0,
        })

    export_df = pd.DataFrame(
        rows_out,
        columns=['Артикул', '', 'Баркод', 'Объём пополнения', 'Срок (дн)',
                 ' ', 'Остатки по складам', '_q1', '_q2', '_q3'],
    )
    export_df.to_excel(writer, sheet_name='Поставки', index=False)
    ws = writer.sheets['Поставки']

    # Стили шапки/ширины (видимые колонки 1..7)
    apply_header_style(
        ws,
        {1: 26, 2: 2, 3: 20, 4: 20, 5: 12, 6: 2, 7: 36},
        row_height=34,
        auto_filter=False,
    )
    # Скрыть служебные колонки H/I/J
    for col_letter in ('H', 'I', 'J'):
        ws.column_dimensions[col_letter].hidden = True

    # Базовые стили данных
    apply_data_style(ws, float_cols=[], int_cols=[], row_height=None)

    # Толстая рамка для колонок «Баркод» (3) и «Объём пополнения» (4)
    medium = Side(style='medium', color='000000')
    thick_border = Border(left=medium, right=medium, top=medium, bottom=medium)
    ws.cell(row=1, column=3).border = thick_border
    ws.cell(row=1, column=4).border = thick_border

    # Заполнение строк: формулы, скрытые значения, dropdown, ссылки, бордюры
    for i, (_, r) in enumerate(df_refill.iterrows()):
        row_num = i + 2
        avg = float(r.get('avg_per_day') or 0)
        nm_id = int(r['nm_id']) if r.get('nm_id') is not None else None

        # Скрытые H/I/J — предрассчитанные числа от Python
        ws.cell(row=row_num, column=8).value = calc_refill_qty(avg, d1, reserve_pct)
        ws.cell(row=row_num, column=9).value = calc_refill_qty(avg, d2, reserve_pct)
        ws.cell(row=row_num, column=10).value = calc_refill_qty(avg, d3, reserve_pct)

        # Lookup-формула в D — без бизнес-логики
        formula = (
            f"=IF(E{row_num}={d1},H{row_num},"
            f"IF(E{row_num}={d2},I{row_num},J{row_num}))"
        )
        d_cell = ws.cell(row=row_num, column=4)
        d_cell.value = formula
        d_cell.number_format = "0"
        d_cell.border = thick_border

        # Баркод — текст + толстая рамка
        b_cell = ws.cell(row=row_num, column=3)
        b_cell.number_format = "@"
        b_cell.border = thick_border

        # Срок — dropdown 10/30/60 (из настроек)
        e_cell = ws.cell(row=row_num, column=5)
        add_days_dropdown(ws, e_cell, refill_days)

        # Артикул — гиперссылка на карточку WB
        a_cell = ws.cell(row=row_num, column=1)
        if a_cell.value and nm_id is not None:
            a_cell.hyperlink = WB_PRODUCT_URL.format(nm_id)
            a_cell.font = HYPERLINK_FONT
            a_cell.alignment = LEFT

        # Колонка 7 — стиль складов + комментарий
        wh_cell = ws.cell(row=row_num, column=7)
        if wh_cell.value is not None:
            wh_cell.font = Font(name="Arial", size=9, color="555555")
            wh_cell.alignment = LEFT
        raw_wh = wh_index.get(nm_id, []) if nm_id is not None else []
        if raw_wh:
            add_stock_comment(wh_cell, merge_wh_by_name(raw_wh))

    # Легенда под таблицей
    last = len(df_refill) + 3
    ws.cell(row=last, column=1, value='— Объём поставки —')
    ws.cell(
        row=last + 1, column=1,
        value=f'Расчёт в боте: avg × срок × (1 + {int(reserve_pct)}%/100), округление вверх',
    )
    ws.cell(
        row=last + 2, column=1,
        value=f'Срок выбирается из выпадающего списка в колонке E: {d1} / {d2} / {d3} дней',
    )
    ws.cell(
        row=last + 3, column=1,
        value='Объём в колонке D обновляется автоматически при смене срока',
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
