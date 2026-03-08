
from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

from engine import (
    CalculationError,
    build_result_workbook,
    calculate_unit_economics,
    dataframe_to_excel_bytes,
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
st.caption("Загрузка и выгрузка — только в Excel. Пересчет происходит сразу после изменения параметров.")

with st.sidebar:
    st.header("Файлы шаблонов")
    for file_name, label in [
        ("products_input_template.xlsx", "Скачать шаблон товаров"),
        ("settings_input_template.xlsx", "Скачать шаблон настроек"),
    ]:
        path = TEMPLATES_DIR / file_name
        with open(path, "rb") as fh:
            st.download_button(
                label=label,
                data=fh.read(),
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

st.subheader("1) Загрузите Excel с товарами")
products_file = st.file_uploader(
    "Файл товаров",
    type=["xlsx"],
    accept_multiple_files=False,
    key="products",
)

st.subheader("2) При необходимости загрузите Excel с настройками")
settings_file = st.file_uploader(
    "Файл настроек",
    type=["xlsx"],
    accept_multiple_files=False,
    key="settings",
)

st.subheader("3) Управление параметрами в системе")
settings = load_settings(settings_file if settings_file is not None else (TEMPLATES_DIR / "settings_input_template.xlsx"))
tariffs = load_tariffs(DATA_DIR / "yandex_tariffs.xlsx")

col1, col2, col3 = st.columns(3)
with col1:
    default_model = st.selectbox("Модель по умолчанию", options=["FBS", "FBY"], index=0 if settings["default_model"] == "FBS" else 1)
    target_margin_pct = st.number_input("Целевая маржинальность, %", min_value=0.0, max_value=95.0, value=float(settings["target_margin_pct"]), step=0.5)
    buyout_rate_pct = st.number_input("Выкупаемость, %", min_value=1.0, max_value=100.0, value=float(settings["buyout_rate_pct"]), step=0.5)
    tax_rate_pct = st.number_input("Налог, %", min_value=0.0, max_value=50.0, value=float(settings["tax_rate_pct"]), step=0.5)
with col2:
    ad_rate_pct = st.number_input("Реклама, %", min_value=0.0, max_value=80.0, value=float(settings["ad_rate_pct"]), step=0.5)
    defect_rate_pct = st.number_input("Брак / потери, %", min_value=0.0, max_value=30.0, value=float(settings["defect_rate_pct"]), step=0.1)
    transfer_schedule = st.selectbox("График выплат", options=list(tariffs["payment_transfer"].keys()), index=list(tariffs["payment_transfer"].keys()).index(settings["transfer_schedule"]))
    reverse_same_locality_share_pct = st.number_input("Доля возвратов без обратной средней мили, %", min_value=0.0, max_value=100.0, value=float(settings["reverse_same_locality_share_pct"]), step=1.0)
with col3:
    avg_orders_per_day = st.number_input("FBS: среднее число заказов в день", min_value=1, max_value=10000, value=int(settings["fbs_avg_orders_per_day"]), step=1)
    avg_items_per_order = st.number_input("Среднее число товаров в заказе", min_value=1.0, max_value=20.0, value=float(settings["avg_items_per_order"]), step=0.1)
    fby_storage_cost_rub_per_unit = st.number_input("FBY: хранение, ₽/ед.", min_value=0.0, max_value=50000.0, value=float(settings["fby_storage_cost_rub_per_unit"]), step=1.0)
    fby_inbound_self_delivery_rub_per_unit = st.number_input("FBY: наша доставка до склада ЯМ, ₽/ед.", min_value=0.0, max_value=50000.0, value=float(settings["fby_inbound_self_delivery_rub_per_unit"]), step=1.0)

settings.update(
    {
        "default_model": default_model,
        "target_margin_pct": target_margin_pct,
        "buyout_rate_pct": buyout_rate_pct,
        "tax_rate_pct": tax_rate_pct,
        "ad_rate_pct": ad_rate_pct,
        "defect_rate_pct": defect_rate_pct,
        "transfer_schedule": transfer_schedule,
        "reverse_same_locality_share_pct": reverse_same_locality_share_pct,
        "fbs_avg_orders_per_day": avg_orders_per_day,
        "avg_items_per_order": avg_items_per_order,
        "fby_storage_cost_rub_per_unit": fby_storage_cost_rub_per_unit,
        "fby_inbound_self_delivery_rub_per_unit": fby_inbound_self_delivery_rub_per_unit,
    }
)

commissions = load_commissions(DATA_DIR / "commissions.xlsx")
keyword_rules = load_keyword_rules(DATA_DIR / "keyword_rules.xlsx")

if products_file is None:
    st.info("Загрузите файл товаров, чтобы получить расчет и скачать Excel-результат.")
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
        st.dataframe(result_df, use_container_width=True, height=520)

        if not warnings_df.empty:
            with st.expander("Предупреждения / что проверить"):
                st.dataframe(warnings_df, use_container_width=True)

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
