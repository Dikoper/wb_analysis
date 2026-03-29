"""
Константы стилей и функции форматирования Excel для отчётов WB Analiz.
"""

from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ── Цветовые константы ────────────────────────────────────────────────────────

# Шапка таблицы
HEADER_BG = "2E4057"
HEADER_FG = "FFFFFF"

# Цвета групп (фон ячейки)
GROUP_COLORS = {
    "A": "B7E4C7",  # мятный зелёный
    "B": "FFD6A5",  # персиковый
    "C": "C8B6E2",  # лавандовый
}

# Градиент повышения цены: {%: (фон, цвет_текста)}
PRICE_INCREASE_COLORS = {
    5:  ("FFF9C4", "333333"),
    10: ("FFE082", "333333"),
    15: ("FFB300", "333333"),
    25: ("FF8F00", "FFFFFF"),
    30: ("E65100", "FFFFFF"),
    40: ("BF360C", "FFFFFF"),
    50: ("7B1818", "FFFFFF"),
}

# Ссылка на кабинет продавца WB
WB_CABINET_URL = "https://www.wildberries.ru/catalog/{}/detail.aspx"

# Цвета рекомендаций сравнения
CMP_RAISE_BG = "E8F5E9"   # светло-зелёный
CMP_LOWER_BG = "FFEBEE"   # светло-красный
CMP_PAIR_ALT_BG = "F5F5F5"  # чередование пар
SUMMARY_PAIR_ALT_BG = "F5F5F5"

# ── Вспомогательные стили ─────────────────────────────────────────────────────

def thin_border(color="CCCCCC"):
    side = Side(style="thin", color=color)
    return Border(left=side, right=side, top=side, bottom=side)


def fill(hex_color):
    return PatternFill(fill_type="solid", fgColor=hex_color)


# ── Функции форматирования ────────────────────────────────────────────────────

def apply_header_style(ws, column_widths: dict, headers: list = None, row_height: int = 34):
    """
    Тёмная шапка, фриз, авто-фильтр, высота строк, ширина столбцов.

    Args:
        ws: openpyxl worksheet
        column_widths: {col_num: width}
        headers: если передан — записывает заголовки в строку 1
        row_height: высота строки шапки
    """
    header_font = Font(name="Arial", size=11, bold=True, color=HEADER_FG)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    if headers:
        for col_idx, header in enumerate(headers, 1):
            ws.cell(row=1, column=col_idx, value=header)

    ws.row_dimensions[1].height = row_height
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = fill(HEADER_BG)
        cell.alignment = center
        cell.border = thin_border("444444")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_num, width in column_widths.items():
        ws.column_dimensions[get_column_letter(col_num)].width = width


def apply_data_style(ws, float_cols: list, int_cols: list):
    """Шрифт, выравнивание, высота строк, числовые форматы для строк с данными."""
    data_font = Font(name="Arial", size=10, color="1A1A2E")
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")

    border = thin_border()

    for row in ws.iter_rows(min_row=2):
        # пропускаем строки аннотаций (нет значения во 2-м столбце)
        if row[0].value is None:
            continue
        ws.row_dimensions[row[0].row].height = 16
        for cell in row:
            cell.font = data_font
            cell.border = border
            cell.alignment = left if cell.column == 1 else center
            if cell.column in float_cols:
                cell.number_format = "0.00"
            elif cell.column in int_cols:
                cell.number_format = "0"


def apply_group_colors(ws, group_col: int, start_row: int = 2):
    """Окрашивает только ячейку группы товара."""
    for row in ws.iter_rows(min_row=start_row):
        cell = row[group_col - 1]
        group = cell.value
        if group in GROUP_COLORS:
            cell.fill = fill(GROUP_COLORS[group])
            cell.font = Font(name="Arial", size=10, bold=True, color="1A1A2E")


def apply_price_increase_colors(ws, pct_col: int):
    """Градиент фона + контрастный текст для ячейки % повышения."""
    for row in ws.iter_rows(min_row=2):
        cell = row[pct_col - 1]
        pct = cell.value
        if pct in PRICE_INCREASE_COLORS:
            bg, fg = PRICE_INCREASE_COLORS[pct]
            cell.fill = fill(bg)
            cell.font = Font(name="Arial", size=10, bold=True, color=fg)


def apply_legend_style(ws, start_row: int, has_threshold_row: bool = False):
    """Ненавязчивый стиль для аннотаций под таблицей."""
    title_font = Font(name="Arial", size=9, bold=True, color="888888")
    row_font = Font(name="Arial", size=9, italic=True, color="999999")
    note_font = Font(name="Arial", size=9, color="999999")
    summary_font = Font(name="Arial", size=9, bold=True, color="555555")

    # start_row: строка «— Группы товаров —»
    ws.cell(row=start_row, column=1).font = title_font
    ws.cell(row=start_row + 1, column=1).font = row_font
    ws.cell(row=start_row + 2, column=1).font = row_font
    ws.cell(row=start_row + 3, column=1).font = row_font
    if has_threshold_row:
        ws.cell(row=start_row + 4, column=1).font = note_font

    # Итоговые строки листа 2 (на 4 строки выше start_row: пустая + 2 итога + пустая)
    summary_row = start_row - 4
    if summary_row >= 1:
        ws.cell(row=summary_row, column=1).font = summary_font
        ws.cell(row=summary_row + 1, column=1).font = summary_font


def apply_hyperlinks(ws, link_col: int, start_row: int = 2):
    """Превращает nmId в кликабельные ссылки на кабинет WB."""
    hyperlink_font = Font(name="Arial", size=10, color="1155CC", underline="single")
    center = Alignment(horizontal="center", vertical="center")
    for row in ws.iter_rows(min_row=start_row):
        cell = row[link_col - 1]
        nm_id = cell.value
        if nm_id is not None:
            url = WB_CABINET_URL.format(nm_id)
            cell.hyperlink = url
            cell.value = nm_id
            cell.font = hyperlink_font
            cell.alignment = center


def write_legend_block(ws, start_row: int, lines: list[tuple[str, str]]):
    """
    Записывает блок легенды в произвольном формате.

    Args:
        lines: [(текст, тип), ...] где тип = 'title' | 'item' | 'note'
    """
    fonts = {
        'title': Font(name="Arial", size=9, bold=True, color="888888"),
        'item': Font(name="Arial", size=9, italic=True, color="999999"),
        'note': Font(name="Arial", size=9, color="999999"),
        'summary': Font(name="Arial", size=9, bold=True, color="555555"),
    }
    for i, (text, style) in enumerate(lines):
        cell = ws.cell(row=start_row + i, column=1, value=text)
        cell.font = fonts.get(style, fonts['item'])


def make_group_border(inner_thin_side, thick_side, is_top, is_bottom, col_idx, num_cols):
    """
    Универсальная рамка ячейки внутри группы (пара или N строк).

    Толстая граница сверху/снизу группы и по бокам таблицы, тонкая внутри.
    """
    top = thick_side if is_top else inner_thin_side
    bottom = thick_side if is_bottom else inner_thin_side
    left = thick_side if col_idx == 1 else inner_thin_side
    right = thick_side if col_idx == num_cols else inner_thin_side
    return Border(left=left, right=right, top=top, bottom=bottom)
