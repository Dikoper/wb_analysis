"""
Фасад обратной совместимости для модуля отчётов.

Все функции перенесены в отдельные модули:
- bot.excel_styles — константы стилей и форматирование Excel
- bot.calculations — бизнес-логика (группы, остатки, цены)
- bot.data_service — загрузка данных из API и кэширование
- bot.report_single — отчёт по одному магазину
- bot.report_comparison — сравнительный отчёт (два магазина)
- bot.report_summary — сводный отчёт (все магазины)
"""

# Бизнес-логика
from bot.calculations import (  # noqa: F401
    assign_group,
    calc_avg_by_group,
    calc_days_remaining,
    get_price_increase,
    aggregate_by_article,
    merge_wh_by_name as _merge_wh_by_name,
)

# Загрузка данных
from bot.data_service import (  # noqa: F401
    fetch_store_data,
    fetch_warehouse_data,
)

# Генерация отчётов
from bot.report_single import generate_report_from_data  # noqa: F401
from bot.report_comparison import generate_comparison_report  # noqa: F401
from bot.report_summary import generate_summary_report  # noqa: F401

# Стили (для обратной совместимости)
from bot.excel_styles import (  # noqa: F401
    HEADER_BG, HEADER_FG, GROUP_COLORS, PRICE_INCREASE_COLORS,
    WB_CABINET_URL, CMP_RAISE_BG, CMP_LOWER_BG,
    CMP_PAIR_ALT_BG, SUMMARY_PAIR_ALT_BG,
    thin_border as _thin_border,
    fill as _fill,
    apply_header_style,
    apply_data_style,
    apply_group_colors,
    apply_price_increase_colors,
    apply_legend_style,
    apply_hyperlinks,
)
