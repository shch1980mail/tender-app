import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime
import plotly.graph_objects as go
from pptx import Presentation
from pptx.util import Inches
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

st.set_page_config(page_title="Отчёты STIMUL", layout="wide")
st.title("📊 Ежемесячные отчёты STIMUL")

# ---------- ЗАГРУЗКА ----------
def load_json(path, default):
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return default

def save_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

sales = pd.DataFrame(load_json("sales.json", []))
claims = load_json("claims.json", [])
items = load_json("claim_items.json", [])
types = load_json("defect_types.json", {})
freight = pd.DataFrame(load_json("freight.json", []))

MONTHS = ["янв","фев","мар","апр","май","июн","июл","авг","сен","окт","ноя","дек"]

# ---------- ПРЕТЕНЗИИ: агрегация ----------
def build_claims_df():
    rows = []
    for c in claims:
        cid = c["id"]
        c_items = [i for i in items if i["claim_id"] == cid]
        total_parts = sum(i["Количество"] for i in c_items)
        rows.append({
            "id": cid, "Месяц": c["Месяц"], "Причина": c["Причина"],
            "Комментарий": c.get("Комментарий", ""),
            "Деталей": total_parts
        })
    return pd.DataFrame(rows)

claims_df = build_claims_df()

# ---------- БОКОВАЯ ПАНЕЛЬ ----------
tab = st.sidebar.radio("Раздел", [
    "📊 Дашборд",
    "➕ Ввести претензию",
    "⚙️ Справочник",
    "🚢 Фрахт",
    "💾 Продажи",
    "📥 Экспорт PPTX"
])

# =====================================================
# ДАШБОРД
# =====================================================
if tab == "📊 Дашборд":
    st.header("📊 Дашборд")

    if not claims_df.empty:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Всего претензий", len(claims_df))
        c2.metric("Всего деталей", int(claims_df["Деталей"].sum()))
        c3.metric("Транспортный бой",
                  len(claims_df[claims_df["Причина"] == "Транспортный бой"]))
        c4.metric("Вина поставщика",
                  len(claims_df[claims_df["Причина"] == "Вина поставщика"]))

    # Претензии по причинам (stacked)
    st.subheader("Претензии по причинам (по месяцам)")
    if not claims_df.empty:
        pivot = claims_df.pivot_table(index="Месяц", columns="Причина",
                                      values="id", aggfunc="count").fillna(0)
        fig = go.Figure()
        for col in pivot.columns:
            fig.add_trace(go.Bar(x=pivot.index, y=pivot[col], name=col))
        fig.update_layout(barmode="stack", height=400)
        st.plotly_chart(fig, use_container_width=True)

    # Брак по типам (stacked по месяцам)
    st.subheader("Брак по типам (по месяцам)")
    if items:
        df_items = pd.DataFrame(items)
        df_items["Месяц"] = df_items["claim_id"].map(
            {c["id"]: c["Месяц"] for c in claims})
        pivot2 = df_items.pivot_table(index="Месяц", columns="Тип",
                                      values="Количество", aggfunc="sum").fillna(0)
        fig2 = go.Figure()
        for col in pivot2.columns:
            fig2.add_trace(go.Bar(x=pivot2.index, y=pivot2[col], name=col))
        fig2.update_layout(barmode="stack", height=450, hovermode="x unified")
        st.plotly_chart(fig2, use_container_width=True)

    # Таблица: Тип | Кол-во | Виды
    st.subheader("Сводка по типам брака")
    if items:
        last_month = max(c["Месяц"] for c in claims)
        last_ids = [c["id"] for c in claims if c["Месяц"] == last_month]
        last_items = [i for i in items if i["claim_id"] in last_ids]
        if last_items:
            df_last = pd.DataFrame(last_items)
            summary = df_last.groupby("Тип").agg(
                Кол_во=("Количество", "sum"),
                Виды=("Вид", lambda x: ", ".join(sorted(set(x))))
            ).reset_index().sort_values("Кол_во", ascending=False)
            st.caption(f"Месяц: {last_month}")
            st.dataframe(summary, use_container_width=True)

    with st.expander("📋 Все претензии"):
        st.dataframe(claims_df, use_container_width=True)

# =====================================================
# ВВОД ПРЕТЕНЗИИ
# =====================================================
elif tab == "➕ Ввести претензию":
    st.header("➕ Новая претензия")

    month = st.text_input("Месяц (ГГГГ-ММ)", value=datetime.now().strftime("%Y-%m"))
    reason = st.radio("Причина", ["Транспортный бой", "Вина клиента", "Вина поставщика"])
    comment = st.text_area("Комментарий", "")

    st.subheader("Детали претензии")
    if "n_items" not in st.session_state:
        st.session_state.n_items = 1

    new_items = []
    for i in range(st.session_state.n_items):
        c1, c2, c3 = st.columns([3, 3, 1])
        t = c1.selectbox(f"Тип #{i+1}", list(types.keys()), key=f"t{i}")
        v = c2.selectbox(f"Вид #{i+1}", types[t], key=f"v{i}")
