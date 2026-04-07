"""
Тесты генерации отчётов: проверка что Excel-файлы создаются корректно.
"""

import os

import pytest
from openpyxl import load_workbook

from bot.reports.single import generate_report_from_data
from bot.reports.comparison import generate_comparison_report
from bot.reports.summary import generate_summary_report


class TestSingleReport:
    def test_generates_file(self, product_rows, warehouse_rows, tmp_path, monkeypatch):
        monkeypatch.setattr('bot.config.REPORTS_DIR', str(tmp_path))
        path = generate_report_from_data(
            product_rows, store_name='Test Store', warehouse_rows=warehouse_rows,
        )
        assert path is not None
        assert os.path.exists(path)

    def test_sheets(self, product_rows, warehouse_rows, tmp_path, monkeypatch):
        monkeypatch.setattr('bot.config.REPORTS_DIR', str(tmp_path))
        path = generate_report_from_data(
            product_rows, store_name='Test', warehouse_rows=warehouse_rows,
        )
        wb = load_workbook(path)
        assert set(wb.sheetnames) == {'На исходе', 'Нет на складе', 'Поставки', 'Все товары'}

    def test_all_products_sheet_has_all_rows(self, product_rows, warehouse_rows, tmp_path, monkeypatch):
        monkeypatch.setattr('bot.config.REPORTS_DIR', str(tmp_path))
        path = generate_report_from_data(
            product_rows, store_name='Test', warehouse_rows=warehouse_rows,
        )
        wb = load_workbook(path)
        ws = wb['Все товары']
        # Строка 1 — шапка, данные начинаются со 2-й
        data_rows = [r for r in ws.iter_rows(min_row=2) if r[0].value is not None]
        assert len(data_rows) == len(product_rows)

    def test_empty_data(self, tmp_path, monkeypatch):
        monkeypatch.setattr('bot.config.REPORTS_DIR', str(tmp_path))
        path = generate_report_from_data([], store_name='Empty')
        assert os.path.exists(path)

    def test_refill_sheet_uses_lookup_formula(self, product_rows, warehouse_rows, tmp_path, monkeypatch):
        monkeypatch.setattr('bot.config.REPORTS_DIR', str(tmp_path))
        path = generate_report_from_data(
            product_rows, store_name='Test', warehouse_rows=warehouse_rows,
        )
        wb = load_workbook(path)
        ws = wb['Поставки']
        # Двухстрочная шапка: row 1 — баннеры, row 2 — подзаголовки, данные с row 3.
        # ART-1 в фикстуре имеет price_increase_pct>0, avg>0, баркод → попадает в refill.
        d_cell = ws.cell(row=3, column=4)
        assert isinstance(d_cell.value, str) and d_cell.value.startswith('=IF(')
        # Формула должна ссылаться на скрытые P/Q/R
        assert 'P3' in d_cell.value
        assert 'Q3' in d_cell.value
        assert 'R3' in d_cell.value
        e_cell = ws.cell(row=3, column=5)
        assert e_cell.value in (10, 30, 60)
        # P/Q/R — предрассчитанные числа
        assert isinstance(ws.cell(row=3, column=16).value, (int, float))
        assert isinstance(ws.cell(row=3, column=17).value, (int, float))
        assert isinstance(ws.cell(row=3, column=18).value, (int, float))
        assert ws.column_dimensions['P'].hidden is True
        assert ws.column_dimensions['Q'].hidden is True
        assert ws.column_dimensions['R'].hidden is True

    def test_refill_suggestion_block_headers(self, product_rows, warehouse_rows, tmp_path, monkeypatch):
        """Шапка блока «Предложение WB» содержит ожидаемые подзаголовки."""
        monkeypatch.setattr('bot.config.REPORTS_DIR', str(tmp_path))
        path = generate_report_from_data(
            product_rows, store_name='Test', warehouse_rows=warehouse_rows,
        )
        wb = load_workbook(path)
        ws = wb['Поставки']
        # Row 2 — подзаголовки
        assert ws.cell(row=2, column=9).value == 'Баркод'
        assert ws.cell(row=2, column=10).value == 'Объём WB'
        assert ws.cell(row=2, column=11).value == 'Оборотность'
        assert ws.cell(row=2, column=13).value == 'Упущено заказов'
        assert ws.cell(row=2, column=14).value == 'Распродать (WB)'
        assert ws.cell(row=2, column=15).value == 'Тренд'

    def test_refill_sale_rate_header_has_comment(self, product_rows, warehouse_rows, tmp_path, monkeypatch):
        """Шапка «Распродать (WB)» содержит поясняющий комментарий."""
        monkeypatch.setattr('bot.config.REPORTS_DIR', str(tmp_path))
        path = generate_report_from_data(
            product_rows, store_name='Test', warehouse_rows=warehouse_rows,
        )
        wb = load_workbook(path)
        ws = wb['Поставки']
        header_cell = ws.cell(row=2, column=14)
        assert header_cell.comment is not None
        assert 'saleRate' in header_cell.comment.text

    def test_nonliquid_zero_volume(self, product_rows, warehouse_rows, tmp_path, monkeypatch):
        """Товар с availability=nonLiquid → объём WB = 0."""
        monkeypatch.setattr('bot.config.REPORTS_DIR', str(tmp_path))
        path = generate_report_from_data(
            product_rows, store_name='Test', warehouse_rows=warehouse_rows,
        )
        wb = load_workbook(path)
        ws = wb['Поставки']
        # Найти строку с ART-2 (nonLiquid в фикстуре)
        for row in range(3, ws.max_row + 1):
            if ws.cell(row=row, column=1).value == 'ART-2':
                assert ws.cell(row=row, column=10).value == 0
                assert ws.cell(row=row, column=11).value == 'неликвид'
                return
        assert False, 'ART-2 не найден на листе Поставки'

    def test_burning_fill_when_lost_positive(self, product_rows, warehouse_rows, tmp_path, monkeypatch):
        """ART-1 имеет lost_orders=3.4 → ячейка «Упущено заказов» залита красным."""
        monkeypatch.setattr('bot.config.REPORTS_DIR', str(tmp_path))
        path = generate_report_from_data(
            product_rows, store_name='Test', warehouse_rows=warehouse_rows,
        )
        wb = load_workbook(path)
        ws = wb['Поставки']
        for row in range(3, ws.max_row + 1):
            if ws.cell(row=row, column=1).value == 'ART-1':
                lost_cell = ws.cell(row=row, column=13)
                assert lost_cell.value == 3
                assert lost_cell.fill.fgColor.rgb is not None
                assert 'FFCC' in (lost_cell.fill.fgColor.rgb or '')
                return
        assert False, 'ART-1 не найден на листе Поставки'

    def test_lost_below_half_shows_dash_no_fill(self, warehouse_rows, tmp_path, monkeypatch):
        """lost_orders=0.3 → round=0 → должно быть «—» и без красной заливки."""
        monkeypatch.setattr('bot.config.REPORTS_DIR', str(tmp_path))
        rows = [{
            'nm_id': 500, 'supplier_article': 'ART-LOW', 'barcode': '2000000000500',
            'subject': 'X', 'category': 'Y', 'product_group': 'A',
            'stock_qty': 5, 'in_way_from_client': 0, 'stock_qty_clean': 5,
            'orders_7d': 2, 'orders_14d': 4, 'orders_30d': None,
            'avg_per_day': 0.3, 'days_remaining': 16.0,
            'price_increase_pct': 10, 'price': 100.0,
            'availability': 'balanced', 'sale_rate_days': 16.0,
            'office_missing_days': 0.0, 'lost_orders': 0.3, 'trend_pct': 0.0,
        }]
        path = generate_report_from_data(rows, store_name='T', warehouse_rows=[])
        wb = load_workbook(path)
        ws = wb['Поставки']
        # Первая (и единственная) строка данных
        lost_cell = ws.cell(row=3, column=13)
        assert lost_cell.value == '—'
        # Без красной заливки
        rgb = lost_cell.fill.fgColor.rgb or ''
        assert 'FFCC' not in rgb

    def test_wh_column_wraps_text(self, product_rows, warehouse_rows, tmp_path, monkeypatch):
        """Колонка «Остатки по складам» (col 7) на «Поставках» имеет wrap_text."""
        monkeypatch.setattr('bot.config.REPORTS_DIR', str(tmp_path))
        path = generate_report_from_data(
            product_rows, store_name='Test', warehouse_rows=warehouse_rows,
        )
        wb = load_workbook(path)
        ws = wb['Поставки']
        # Берём первую строку данных
        cell = ws.cell(row=3, column=7)
        assert cell.alignment.wrap_text is True

    def test_gap_columns_have_fill(self, product_rows, warehouse_rows, tmp_path, monkeypatch):
        """Gap-колонки (2, 6) и separator (8) — залиты."""
        monkeypatch.setattr('bot.config.REPORTS_DIR', str(tmp_path))
        path = generate_report_from_data(
            product_rows, store_name='Test', warehouse_rows=warehouse_rows,
        )
        wb = load_workbook(path)
        ws = wb['Поставки']
        # Row 3 — первая строка данных
        for col in (2, 6, 8):
            cell = ws.cell(row=3, column=col)
            rgb = cell.fill.fgColor.rgb
            assert rgb is not None and rgb != '00000000'


class TestComparisonReport:
    def test_generates_file(self, product_rows, tmp_path, monkeypatch):
        monkeypatch.setattr('bot.config.REPORTS_DIR', str(tmp_path))
        # Берём часть товаров для каждого магазина (ART-1, ART-2 общие)
        store1 = product_rows[:2]
        store2 = [product_rows[0], product_rows[2]]  # ART-1 и ART-3
        path = generate_comparison_report(store1, store2, 'Магазин 1', 'Магазин 2')
        assert path is not None
        assert os.path.exists(path)

    def test_returns_none_if_no_common(self, tmp_path, monkeypatch):
        monkeypatch.setattr('bot.config.REPORTS_DIR', str(tmp_path))
        store1 = [{'nm_id': 1, 'supplier_article': 'UNIQUE-1', 'stock_qty': 10,
                    'stock_qty_clean': 10, 'in_way_from_client': 0,
                    'avg_per_day': 1.0, 'price': 100}]
        store2 = [{'nm_id': 2, 'supplier_article': 'UNIQUE-2', 'stock_qty': 5,
                    'stock_qty_clean': 5, 'in_way_from_client': 0,
                    'avg_per_day': 2.0, 'price': 200}]
        path = generate_comparison_report(store1, store2, 'S1', 'S2')
        assert path is None

    def test_common_articles_counted(self, product_rows, tmp_path, monkeypatch):
        monkeypatch.setattr('bot.config.REPORTS_DIR', str(tmp_path))
        path = generate_comparison_report(
            product_rows[:2], product_rows[:2], 'S1', 'S2',
        )
        wb = load_workbook(path)
        ws = wb.active
        # 2 общих артикула × 2 строки = 4 строки данных
        data_rows = [r for r in ws.iter_rows(min_row=2) if r[1].value is not None]
        assert len(data_rows) == 4


class TestSummaryReport:
    def test_generates_file(self, product_rows, warehouse_rows, tmp_path, monkeypatch):
        monkeypatch.setattr('bot.config.REPORTS_DIR', str(tmp_path))
        all_stores = {
            'Магазин 1': product_rows[:2],
            'Магазин 2': [product_rows[0], product_rows[2]],
        }
        all_wh = {
            'Магазин 1': warehouse_rows[:3],
            'Магазин 2': [warehouse_rows[0], warehouse_rows[3]],
        }
        path = generate_summary_report(all_stores, all_wh)
        assert path is not None
        assert os.path.exists(path)

    def test_returns_none_if_no_common(self, tmp_path, monkeypatch):
        monkeypatch.setattr('bot.config.REPORTS_DIR', str(tmp_path))
        store1 = [{'nm_id': 1, 'supplier_article': 'UNIQUE-1', 'stock_qty': 10,
                    'stock_qty_clean': 10, 'in_way_from_client': 0,
                    'avg_per_day': 1.0, 'price': 100}]
        store2 = [{'nm_id': 2, 'supplier_article': 'UNIQUE-2', 'stock_qty': 5,
                    'stock_qty_clean': 5, 'in_way_from_client': 0,
                    'avg_per_day': 2.0, 'price': 200}]
        path = generate_summary_report(
            {'S1': store1, 'S2': store2}, {'S1': [], 'S2': []},
        )
        assert path is None
