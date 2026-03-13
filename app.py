
from __future__ import annotations

from pathlib import Path

import streamlit as st

from engine import (
    CalculationError,
    TAX_SYSTEM_OPTIONS,
    build_result_workbook,
    calculate_unit_economics,
    load_commissions,
    load_keyword_rules,
    load_settings,
    load_tariffs,
    read_products_excel,
)

ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
TEMPLATES_DIR = ROOT / "templates"

st.set_page_config(page_title="Яндекс Маркет — юнит-экономика", layout="wide")
st.title("Яндекс Маркет — калькулятор юнит-экономики (FBS / FBY)")
st.caption("Excel на входе и на выходе. После изменения параметров пересчет выполняется сразу.")

commissions = load_commissions(DATA_DIR / "commissions.xlsx")
keyword_rules = load_keyword_rules(DATA_DIR / "keyword_rules.xlsx")
tariffs = load_tariffs(DATA_DIR / "yandex_tariffs.xlsx")
settings = load_settings(None)

upload_col, template_col = st.columns([4, 1.7])
with upload_col:
    st.subheader("Загрузите Excel с товарами")
    products_file = st.file_uploader(
        "Файл товаров (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=False,
        key="products",
        label_visibility="visible",
    )
with template_col:
    st.subheader(" ")
    path = TEMPLATES_DIR / "products_input_template.xlsx"
    with open(path, "rb") as fh:
        st.download_button(
            label="Скачать шаблон товаров",
            data=fh.read(),
            file_name="template_yandex_market_products.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

st.subheader("Управление параметрами в системе")
col1, col2, col3 = st.columns(3)

with col1:
    default_model = st.selectbox("Модель", options=["FBS", "FBY"], index=0)
    tax_system = st.selectbox(
        "Налоговая система",
        options=list(TAX_SYSTEM_OPTIONS.keys()),
        index=list(TAX_SYSTEM_OPTIONS.keys()).index(settings["tax_system"]),
    )
    target_margin_pct = st.number_input(
        "Целевая маржинальность, %",
        min_value=0.0,
        max_value=95.0,
        value=float(settings["target_margin_pct"]),
        step=0.5,
    )
    buyout_rate_pct = st.number_input(
        "Выкупаемость, %",
        min_value=1.0,
        max_value=100.0,
        value=float(settings["buyout_rate_pct"]),
        step=0.5,
    )

with col2:
    ad_rate_pct = st.number_input(
        "Реклама, %",
        min_value=0.0,
        max_value=80.0,
        value=float(settings["ad_rate_pct"]),
        step=0.5,
    )
    storage_days = st.number_input(
        "Срок хранения, дней",
        min_value=0,
        max_value=3650,
        value=int(settings["fby_storage_days"]),
        step=1,
    )
    fby_storage_cost_rub_per_day = st.number_input(
        "FBY: хранение, ₽ в день за 1 шт.",
        min_value=0.0,
        max_value=10000.0,
        value=float(settings["fby_storage_cost_rub_per_day"]),
        step=0.1,
    )
    fby_inbound_self_delivery_rub_per_unit = st.number_input(
        "FBY: наша доставка до склада ЯМ, ₽/шт.",
        min_value=0.0,
        max_value=50000.0,
        value=float(settings["fby_inbound_self_delivery_rub_per_unit"]),
        step=1.0,
    )

with col3:
    avg_orders_per_day = st.number_input(
        "FBS: среднее число заказов в день",
        min_value=1,
        max_value=10000,
        value=int(settings["fbs_avg_orders_per_day"]),
        step=1,
    )
    avg_items_per_order = st.number_input(
        "Среднее число товаров в заказе",
        min_value=1.0,
        max_value=20.0,
        value=float(settings["avg_items_per_order"]),
        step=0.1,
    )
    reverse_same_locality_share_pct = st.number_input(
        "Возвраты без обратной средней мили, %",
        min_value=0.0,
        max_value=100.0,
        value=float(settings["reverse_same_locality_share_pct"]),
        step=1.0,
    )
    defect_rate_pct = st.number_input(
        "Брак / потери, %",
        min_value=0.0,
        max_value=30.0,
        value=float(settings["defect_rate_pct"]),
        step=0.1,
    )

settings.update(
    {
        "default_model": default_model,
        "tax_system": tax_system,
        "target_margin_pct": target_margin_pct,
        "buyout_rate_pct": buyout_rate_pct,
        "ad_rate_pct": ad_rate_pct,
        "reverse_same_locality_share_pct": reverse_same_locality_share_pct,
        "defect_rate_pct": defect_rate_pct,
        "fbs_avg_orders_per_day": avg_orders_per_day,
        "avg_items_per_order": avg_items_per_order,
        "fby_storage_days": storage_days,
        "fby_storage_cost_rub_per_day": fby_storage_cost_rub_per_day,
        "fby_inbound_self_delivery_rub_per_unit": fby_inbound_self_delivery_rub_per_unit,
    }
)

if products_file is None:
    st.info("Загрузите файл товаров. Категория определится по названию, комиссия и логистика посчитаются автоматически.")
else:
    try:
        products_df = read_products_excel(products_file)
        result_df, warnings_df = calculate_unit_economics(
            products_df=products_df,
            settings=settings,
            tariffs=tariffs,
            commissions_df=commissions,
            keyword_rules_df=keyword_rules,
        )

        st.success(f"Расчет выполнен. Позиций: {len(result_df)}")
        st.dataframe(result_df, use_container_width=True, height=560)

        if not warnings_df.empty:
            with st.expander("Что проверить"):
                st.dataframe(warnings_df, use_container_width=True, height=220)

        result_bytes = build_result_workbook(
            result_df=result_df,
            warnings_df=warnings_df,
            settings=settings,
            tariffs=tariffs,
        )
        st.download_button(
            "Скачать результат Excel",
            data=result_bytes,
            file_name="yandex_market_unit_economics_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except CalculationError as exc:
        st.error(str(exc))
    except Exception as exc:
        st.exception(exc)
