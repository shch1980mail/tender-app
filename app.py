import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from pptx import Presentation
from pptx.util import Inches
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

st.set_page_config(page_title="Отчёты STIMUL", layout="wide")
st.title("📊 Ежемесячные отчёты: претензии, брак, логистика")

# ---------- ЗАГРУЗКА ДАННЫХ ----------
def load_json(path, default_cols=None):
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return pd.DataFrame(json.load(f))
    return pd.DataFrame(columns=default_cols or [])

claims = load_json("claims.json")
defects = load_json("defects.json")
freight = load_json("freight.json")

# ---------- РАСЧЁТ КОЭФФИЦИЕНТОВ ----------
def enrich(df):
    if df.empty:
        return df
    df = df.copy()
    df["Коэф_прет_млн"] = df["Всего_претензий"] / df["Сумма_продаж"] * 1_000_000
    df["Коэф_завбр_млн"] = df["Зав_брак"] / df["Сумма_продаж"] * 1_000_000
    df["Коэф_трансбой_млн"] = df["Транс_бой"] / df["Сумма_продаж"] * 1_000_000
    df["Доля_завбрака_%"] = df["Зав_брак"] / df["Всего_претензий"] * 100
    df["Доля_трансбоя_%"] = df["Транс_бой"] / df["Всего_претензий"] * 100
    df["Прет_на_конт"] = df.apply(lambda r: r["Всего_претензий"]/r["Контейнеры"] if r["Контейнеры"] else None, axis=1)
    df["Прет_на_ед"] = df["Всего_претензий"] / df["Продажи_шт"]
    df["Доля_брака_%"] = df["Всего_претензий"] / df["Продажи_шт"] * 100
    df["Брак_дет_на_прод_%"] = df["Браков_деталей"] / df["Продажи_шт"] * 100
    df["Брак_дет_на_млн"] = df["Браков_деталей"] / df["Сумма_продаж"] * 1_000_000
    return df

claims = enrich(claims).sort_values("Месяц")
defects = defects.sort_values("Месяц")
freight = freight.sort_values("Месяц")

# ---------- БОКОВАЯ ПАНЕЛЬ ----------
with st.sidebar:
    st.header("➕ Добавить месяц (претензии)")
    month = st.text_input("Месяц (ГГГГ-ММ)", value=datetime.now().strftime("%Y-%m"))
    sales_rub = st.number_input("Сумма продаж, руб.", value=0.0, step=1000.0)
    containers = st.number_input("Контейнеров, шт.", value=0, step=1)
    sales_units = st.number_input("Продажи, шт.", value=0, step=1)
    factory_defect = st.number_input("Заводской брак, шт.", value=0, step=1)
    transport_damage = st.number_input("Транспортный бой, шт.", value=0, step=1)
    total_claims = st.number_input("Всего претензий, шт.", value=0, step=1)
    defective_parts = st.number_input("Бракованных деталей, шт.", value=0, step=1)

    if st.button("💾 Сохранить запись"):
        new_row = {"Месяц": month, "Сумма_продаж": sales_rub, "Контейнеры": containers,
                   "Продажи_шт": sales_units, "Зав_брак": factory_defect,
                   "Транс_бой": transport_damage, "Всего_претензий": total_claims,
                   "Браков_деталей": defective_parts}
        df = claims[claims["Месяц"] != month]
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        df.to_json("claims.json", orient="records", force_ascii=False, indent=2)
        st.success(f"Сохранено за {month}. Обновите страницу.")
        st.rerun()

# ---------- KPI ----------
if not claims.empty:
    last = claims.iloc[-1]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Всего претензий", int(last["Всего_претензий"]))
    c2.metric("Доля брака, %", f"{last['Доля_брака_%']:.2f}")
    c3.metric("Прет/млн руб.", f"{last['Коэф_прет_млн']:.2f}")
    c4.metric("Брак. деталей", int(last["Браков_деталей"]))

# ---------- ГРАФИК 1: КОМБИНИРОВАННЫЙ (претензии + доля брака) ----------
st.subheader("📈 Претензии и доля брака по месяцам")
if not claims.empty:
    fig1 = make_subplots(specs=[[{"secondary_y": True}]])
    fig1.add_trace(go.Bar(x=claims["Месяц"], y=claims["Всего_претензий"],
                          name="Всего претензий, шт.", marker_color="#4C78A8"), secondary_y=False)
    fig1.add_trace(go.Scatter(x=claims["Месяц"], y=claims["Доля_брака_%"],
                              name="Доля брака, %", mode="lines+markers",
                              line=dict(color="#E45756", width=3)), secondary_y=True)
    fig1.update_yaxes(title_text="Претензии, шт.", secondary_y=False)
    fig1.update_yaxes(title_text="Доля брака, %", secondary_y=True)
    fig1.update_layout(height=450, hovermode="x unified")
    st.plotly_chart(fig1, use_container_width=True)

# ---------- ГРАФИК 2: КОЭФФИЦИЕНТЫ ----------
st.subheader("📉 Динамика коэффициентов на млн руб.")
if not claims.empty:
    fig2 = go.Figure()
    fig2.add_trace(go.Scatter(x=claims["Месяц"], y=claims["Коэф_прет_млн"], name="Прет/млн", mode="lines+markers"))
    fig2.add_trace(go.Scatter(x=claims["Месяц"], y=claims["Коэф_завбр_млн"], name="Зав.брак/млн", mode="lines+markers"))
    fig2.add_trace(go.Scatter(x=claims["Месяц"], y=claims["Коэф_трансбой_млн"], name="Транс.бой/млн", mode="lines+markers"))
    fig2.update_layout(height=400, hovermode="x unified")
    st.plotly_chart(fig2, use_container_width=True)

# ---------- ГРАФИК 3: БРАК ПО ТИПАМ (последний месяц) ----------
st.subheader("🔧 Брак по типам деталей (последний месяц)")
DEFECT_TYPES = ["Поддоны","Крыши","Экраны","Стекла","Профили_дуги","Панели","Некомплект","Центральная_стойка","Прочее"]
if not defects.empty:
    latest = defects.iloc[-1]
    values = [int(latest.get(t, 0)) for t in DEFECT_TYPES]
    fig3 = go.Figure(go.Bar(x=DEFECT_TYPES, y=values,
                            marker_color=["#E45756","#72B7B2","#4C78A8","#54A24B","#F58518","#B279A2","#9D755D","#BAB0AC","#EECA3B"],
                            text=values, textposition="outside"))
    fig3.update_layout(height=400, title=f"Месяц: {latest['Месяц']} (всего {int(latest['Всего'])})")
    st.plotly_chart(fig3, use_container_width=True)

# ---------- ГРАФИК 4: ДИНАМИКА БРАКА ПО ТИПАМ (stacked) ----------
st.subheader("📊 Динамика брака по типам (все месяцы)")
if not defects.empty:
    fig4 = go.Figure()
    for t in DEFECT_TYPES:
        if t in defects.columns:
            fig4.add_trace(go.Bar(x=defects["Месяц"], y=defects[t], name=t))
    fig4.update_layout(barmode="stack", height=450, hovermode="x unified")
    st.plotly_chart(fig4, use_container_width=True)

# ---------- ГРАФИК 5: ФРАХТ ----------
st.subheader("🚢 Ставки фрахта (USD и RUB)")
if not freight.empty:
    fig5 = make_subplots(specs=[[{"secondary_y": True}]])
    fig5.add_trace(go.Bar(x=freight["Месяц"], y=freight["Ставка_USD_НСК"],
                          name="USD НСК", marker_color="#4C78A8"), secondary_y=False)
    fig5.add_trace(go.Bar(x=freight["Месяц"], y=freight["Ставка_USD_КРД"],
                          name="USD КРД", marker_color="#E45756"), secondary_y=False)
    fig5.add_trace(go.Scatter(x=freight["Месяц"], y=freight["Ставка_RUB_НСК"],
                              name="RUB НСК", mode="lines+markers",
                              line=dict(color="#54A24B", width=3)), secondary_y=True)
    fig5.add_trace(go.Scatter(x=freight["Месяц"], y=freight["Ставка_RUB_КРД"],
                              name="RUB КРД", mode="lines+markers",
                              line=dict(color="#F58518", width=3)), secondary_y=True)
    fig5.add_trace(go.Scatter(x=freight["Месяц"], y=freight["Ставка_RUB_ЖД"],
                              name="RUB ЖД", mode="lines+markers",
                              line=dict(color="#B279A2", width=3, dash="dot")), secondary_y=True)
    fig5.update_yaxes(title_text="USD", secondary_y=False)
    fig5.update_yaxes(title_text="RUB", secondary_y=True)
    fig5.update_layout(height=450, hovermode="x unified")
    st.plotly_chart(fig5, use_container_width=True)

# ---------- ТАБЛИЦЫ ----------
with st.expander("📋 Претензии (полная таблица)"):
    st.dataframe(claims, use_container_width=True)
with st.expander("📋 Брак по типам"):
    st.dataframe(defects, use_container_width=True)
with st.expander("📋 Ставки фрахта"):
    st.dataframe(freight, use_container_width=True)

# ---------- ЭКСПОРТ PPTX ----------
st.divider()
if st.button("📥 Скачать PPTX"):
    prs = Presentation()
    # Слайд 1: Фрахт
    if not freight.empty:
        s1 = prs.slides.add_slide(prs.slide_layouts[5])
        s1.shapes.title.text = "Ставки фрахта из Китая"
        cd = CategoryChartData()
        cd.categories = freight["Месяц"].tolist()
        cd.add_series("USD НСК", freight["Ставка_USD_НСК"].fillna(0).tolist())
        cd.add_series("USD КРД", freight["Ставка_USD_КРД"].fillna(0).tolist())
        s1.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,
                            Inches(0.5), Inches(1.5), Inches(9), Inches(4.5), cd)
    # Слайд 2: Претензии
    s2 = prs.slides.add_slide(prs.slide_layouts[5])
    s2.shapes.title.text = "Претензии и доля брака"
    cd2 = CategoryChartData()
    cd2.categories = claims["Месяц"].tolist()
    cd2.add_series("Всего претензий", claims["Всего_претензий"].tolist())
    s2.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,
                        Inches(0.5), Inches(1.5), Inches(9), Inches(4.5), cd2)
    # Слайд 3: Коэффициенты
    s3 = prs.slides.add_slide(prs.slide_layouts[5])
    s3.shapes.title.text = "Коэффициенты на млн руб."
    cd3 = CategoryChartData()
    cd3.categories = claims["Месяц"].tolist()
    cd3.add_series("Прет/млн", claims["Коэф_прет_млн"].tolist())
    cd3.add_series("Зав.брак/млн", claims["Коэф_завбр_млн"].tolist())
    s3.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS,
                        Inches(0.5), Inches(1.5), Inches(9), Inches(4.5), cd3)
    # Слайд 4: Брак по типам (последний месяц)
    if not defects.empty:
        s4 = prs.slides.add_slide(prs.slide_layouts[5])
        s4.shapes.title.text = f"Брак по типам — {latest['Месяц']}"
        cd4 = CategoryChartData()
        cd4.categories = DEFECT_TYPES
        cd4.add_series("Кол-во", values)
        s4.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,
                            Inches(0.5), Inches(1.5), Inches(9), Inches(4.5), cd4)

    prs.save("report.pptx")
    with open("report.pptx", "rb") as f:
        st.download_button("💾 Сохранить файл", f, file_name="report.pptx",
                           mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
