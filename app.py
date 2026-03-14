
import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Юнит‑экономика Яндекс Маркета (FBS / FBY)")

st.subheader("Загрузите Excel с товарами")

with open("templates/products_input_template.xlsx", "rb") as f:
    st.download_button(
        "📥 Скачать шаблон товаров",
        f,
        file_name="template_yandex_market_products.xlsx"
    )

uploaded = st.file_uploader("Файл товаров (.xlsx)", type=["xlsx"])

st.subheader("Параметры расчета")

tax = st.selectbox("Система налогообложения", ["ОСНО", "УСН 6%", "УСН 15%"])
target_margin = st.number_input("Целевая маржинальность %", value=30)

if uploaded:
    df = pd.read_excel(uploaded)

    df["Комиссия"] = df["Себестоимость (руб)"] * 0.08
    df["Логистика"] = 150
    df["Эквайринг"] = df["Себестоимость (руб)"] * 0.015

    df["Полная себестоимость"] = (
        df["Себестоимость (руб)"] +
        df["Комиссия"] +
        df["Логистика"] +
        df["Эквайринг"]
    )

    df["Рекомендованная цена"] = df["Полная себестоимость"] * (1 + target_margin/100)
    df["Маржа %"] = target_margin

    st.subheader("Результат расчета")
    st.dataframe(df, use_container_width=True)

    buffer = BytesIO()
    df.to_excel(buffer, index=False)

    st.download_button(
        "Скачать результат Excel",
        buffer.getvalue(),
        file_name="yandex_market_unit_economics.xlsx"
    )
