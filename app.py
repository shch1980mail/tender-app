import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime
from pptx import Presentation
from pptx.util import Inches
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

# ---------- НАСТРОЙКИ ----------
DATA_FILE = "claims.json"  # Файл будет создан автоматически

st.set_page_config(page_title="Отчёты", layout="wide")
st.title("📊 Ежемесячные отчёты")

# ---------- ЗАГРУЗКА И СОХРАНЕНИЕ ДАННЫХ ----------
def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return pd.DataFrame(json.load(f))
    return pd.DataFrame(columns=[
        "Месяц", "Сумма_продаж", "Контейнеры", "Продажи_шт",
        "Зав_брак", "Транс_бой", "Всего_претензий", "Браков_деталей"
    ])

def save_data(df):
    df.to_json(DATA_FILE, orient="records", force_ascii=False, indent=2)

if "df" not in st.session_state:
    st.session_state.df = load_data()

# ---------- БОКОВАЯ ПАНЕЛЬ: ВВОД ----------
with st.sidebar:
    st.header("➕ Добавить месяц")
    month = st.text_input("Месяц (ГГГГ-ММ)", value=datetime.now().strftime("%Y-%m"))
    sales_rub = st.number_input("Сумма продаж, руб.", value=0.0, step=1000.0)
    containers = st.number_input("Контейнеров, шт.", value=0, step=1)
    sales_units = st.number_input("Продажи, шт.", value=0, step=1)
    factory_defect = st.number_input("Заводской брак, шт.", value=0, step=1)
    transport_damage = st.number_input("Транспортный бой, шт.", value=0, step=1)
    total_claims = st.number_input("Всего претензий, шт.", value=0, step=1)
    defective_parts = st.number_input("Бракованных деталей, шт.", value=0, step=1)

    if st.button("💾 Сохранить"):
        new_row = {
            "Месяц": month,
            "Сумма_продаж": sales_rub,
            "Контейнеры": containers,
            "Продажи_шт": sales_units,
            "Зав_брак": factory_defect,
            "Транс_бой": transport_damage,
            "Всего_претензий": total_claims,
            "Браков_деталей": defective_parts,
        }
        # Если такой месяц уже есть — заменяем
        st.session_state.df = st.session_state.df[st.session_state.df["Месяц"] != month]
        st.session_state.df = pd.concat(
            [st.session_state.df, pd.DataFrame([new_row])],
            ignore_index=True
        ).sort_values("Месяц")
        save_data(st.session_state.df)
        st.success(f"Сохранено за {month}")

# ---------- ОСНОВНАЯ ЧАСТЬ ----------
st.subheader("📋 Данные")
if not st.session_state.df.empty:
    st.dataframe(st.session_state.df, use_container_width=True)

    st.subheader("📈 Динамика претензий")
    chart_df = st.session_state.df.sort_values("Месяц").set_index("Месяц")
    st.bar_chart(chart_df["Всего_претензий"])
else:
    st.info("Данных пока нет. Добавьте первый месяц слева.")

# ---------- ЭКСПОРТ В PPTX ----------
st.divider()
if st.button("📥 Скачать PPTX"):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Отчёт по претензиям"

    if not st.session_state.df.empty:
        df_sorted = st.session_state.df.sort_values("Месяц")
        chart_data = CategoryChartData()
        chart_data.categories = df_sorted["Месяц"].tolist()
        chart_data.add_series("Всего претензий, шт.", df_sorted["Всего_претензий"].tolist())

        slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            Inches(0.5), Inches(1.5), Inches(9), Inches(4.5),
            chart_data
        )

    prs.save("report.pptx")
    with open("report.pptx", "rb") as f:
        st.download_button(
            "💾 Сохранить файл",
            f,
            file_name="report.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
