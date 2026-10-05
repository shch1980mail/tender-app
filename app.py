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

# ---------- РАСЧЁТ МЕТРИК ----------
def build_claims_df():
    rows = []
    for c in claims:
        cid = c["id"]
        month = c["Месяц"]
        reason = c["Причина"]
        c_items = [i for i in items if i["claim_id"] == cid]
        total_parts = sum(i["Количество"] for i in c_items)
        rows.append({
            "id": cid, "Месяц": month, "Причина": reason,
            "Комментарий": c.get("Комментарий", ""),
            "Деталей": total_parts,
            "Детали": "; ".join(f"{i['Тип']}/{i['Вид']}×{i['Количество']}" for i in c_items)
        })
    return pd.DataFrame(rows)

claims_df = build_claims_df()

# ---------- БОКОВАЯ ПАНЕЛЬ: ВКЛАДКИ ----------
tab = st.sidebar.radio("Раздел", [
    "📊 Дашборд",
    "➕ Ввести претензию",
    "⚙️ Справочник",
    "🚢 Фрахт",
    "📥 Экспорт PPTX",
    "💾 Продажи"
])

# =====================================================
# ВКЛАДКА: ДАШБОРД
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

    # График 1: претензии по причинам (stacked bar)
    st.subheader("Претензии по причинам")
    if not claims_df.empty:
        pivot = claims_df.pivot_table(index="Месяц", columns="Причина",
                                      values="id", aggfunc="count").fillna(0)
        fig = go.Figure()
        for col in pivot.columns:
            fig.add_trace(go.Bar(x=pivot.index, y=pivot[col], name=col))
        fig.update_layout(barmode="stack", height=400)
        st.plotly_chart(fig, use_container_width=True)

    # График 2: топ-10 видов брака
    st.subheader("Топ видов брака (последний месяц)")
    if items:
        last_month = max(c["Месяц"] for c in claims)
        last_ids = [c["id"] for c in claims if c["Месяц"] == last_month]
        last_items = [i for i in items if i["claim_id"] in last_ids]
        if last_items:
            df_top = pd.DataFrame(last_items).groupby(["Тип", "Вид"])["Количество"].sum().reset_index()
            df_top = df_top.sort_values("Количество", ascending=False).head(10)
            df_top["Метка"] = df_top["Тип"] + " / " + df_top["Вид"]
            fig2 = go.Figure(go.Bar(x=df_top["Метка"], y=df_top["Количество"],
                                    text=df_top["Количество"], textposition="outside"))
            fig2.update_layout(height=400, xaxis_tickangle=-40)
            st.plotly_chart(fig2, use_container_width=True)

    # График 3: динамика по типам (stacked)
    st.subheader("Динамика по типам брака")
    if items:
        df_items = pd.DataFrame(items)
        df_items["Месяц"] = df_items["claim_id"].map(
            {c["id"]: c["Месяц"] for c in claims})
        pivot2 = df_items.pivot_table(index="Месяц", columns="Тип",
                                      values="Количество", aggfunc="sum").fillna(0)
        fig3 = go.Figure()
        for col in pivot2.columns:
            fig3.add_trace(go.Bar(x=pivot2.index, y=pivot2[col], name=col))
        fig3.update_layout(barmode="stack", height=450)
        st.plotly_chart(fig3, use_container_width=True)

    # Таблица
    with st.expander("📋 Все претензии"):
        st.dataframe(claims_df, use_container_width=True)

# =====================================================
# ВКЛАДКА: ВВОД ПРЕТЕНЗИИ
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
        q = c3.number_input(f"Кол-во", min_value=0, value=1, key=f"q{i}")
        if q > 0:
            new_items.append({"Тип": t, "Вид": v, "Количество": int(q)})

    c1, c2 = st.columns(2)
    if c1.button("➕ Добавить деталь"):
        st.session_state.n_items += 1
        st.rerun()
    if c2.button("➖ Удалить последнюю"):
        if st.session_state.n_items > 1:
            st.session_state.n_items -= 1
            st.rerun()

    if st.button("💾 Сохранить претензию"):
        new_id = max([c["id"] for c in claims], default=0) + 1
        claims.append({"id": new_id, "Месяц": month, "Причина": reason,
                       "Комментарий": comment})
        for it in new_items:
            items.append({"claim_id": new_id, "Тип": it["Тип"],
                          "Вид": it["Вид"], "Количество": it["Количество"]})
        save_json("claims.json", claims)
        save_json("claim_items.json", items)
        st.success(f"Претензия №{new_id} сохранена ({len(new_items)} деталей)")
        st.session_state.n_items = 1

# =====================================================
# ВКЛАДКА: СПРАВОЧНИК
# =====================================================
elif tab == "⚙️ Справочник":
    st.header("⚙️ Справочник типов и видов брака")

    st.subheader("Добавить новый тип")
    new_type = st.text_input("Название типа")
    if st.button("➕ Добавить тип") and new_type:
        if new_type not in types:
            types[new_type] = []
            save_json("defect_types.json", types)
            st.success(f"Тип «{new_type}» добавлен")
            st.rerun()

    st.subheader("Добавить вид в тип")
    sel_type = st.selectbox("Тип", list(types.keys()))
    new_view = st.text_input("Название вида")
    if st.button("➕ Добавить вид") and new_view:
        if new_view not in types[sel_type]:
            types[sel_type].append(new_view)
            save_json("defect_types.json", types)
            st.success(f"Вид «{new_view}» добавлен в «{sel_type}»")
            st.rerun()

    st.subheader("Текущий справочник")
    st.json(types)

# =====================================================
# ВКЛАДКА: ФРАХТ (2 отдельных графика)
# =====================================================
elif tab == "🚢 Фрахт":
    st.header("🚢 Ставки фрахта из Китая")

    st.subheader("Ставки в USD")
    fig_usd = go.Figure()
    fig_usd.add_trace(go.Bar(x=freight["Месяц"], y=freight["USD_НСК"],
                             name="НСК", marker_color="#4C78A8"))
    fig_usd.add_trace(go.Bar(x=freight["Месяц"], y=freight["USD_КРД"],
                             name="КРД", marker_color="#E45756"))
    fig_usd.update_layout(barmode="group", height=400, yaxis_title="USD")
    st.plotly_chart(fig_usd, use_container_width=True)

    st.subheader("Ставки в рублях")
    fig_rub = go.Figure()
    fig_rub.add_trace(go.Scatter(x=freight["Месяц"], y=freight["RUB_НСК"],
                                 name="НСК", mode="lines+markers"))
    fig_rub.add_trace(go.Scatter(x=freight["Месяц"], y=freight["RUB_КРД"],
                                 name="КРД", mode="lines+markers"))
    fig_rub.add_trace(go.Scatter(x=freight["Месяц"], y=freight["RUB_ЖД"],
                                 name="ЖД", mode="lines+markers",
                                 line=dict(dash="dot")))
    fig_rub.update_layout(height=400, yaxis_title="RUB")
    st.plotly_chart(fig_rub, use_container_width=True)

    with st.expander("📋 Таблица ставок"):
        st.dataframe(freight, use_container_width=True)

# =====================================================
# ВКЛАДКА: ЭКСПОРТ PPTX
# =====================================================
elif tab == "📥 Экспорт PPTX":
    st.header("📥 Экспорт в PPTX")
    if st.button("Собрать PPTX"):
        prs = Presentation()

        # Слайд 1: фрахт USD
        s1 = prs.slides.add_slide(prs.slide_layouts[5])
        s1.shapes.title.text = "Ставки фрахта (USD)"
        cd1 = CategoryChartData()
        cd1.categories = freight["Месяц"].tolist()
        cd1.add_series("НСК", freight["USD_НСК"].fillna(0).tolist())
        cd1.add_series("КРД", freight["USD_КРД"].fillna(0).tolist())
        s1.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,
                            Inches(0.5), Inches(1.5), Inches(9), Inches(4.5), cd1)

        # Слайд 2: фрахт RUB
        s2 = prs.slides.add_slide(prs.slide_layouts[5])
        s2.shapes.title.text = "Ставки фрахта (RUB)"
        cd2 = CategoryChartData()
        cd2.categories = freight["Месяц"].tolist()
        cd2.add_series("НСК", freight["RUB_НСК"].fillna(0).tolist())
        cd2.add_series("КРД", freight["RUB_КРД"].fillna(0).tolist())
        cd2.add_series("ЖД", freight["RUB_ЖД"].fillna(0).tolist())
        s2.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS,
                            Inches(0.5), Inches(1.5), Inches(9), Inches(4.5), cd2)

        # Слайд 3: претензии по причинам
        if not claims_df.empty:
            s3 = prs.slides.add_slide(prs.slide_layouts[5])
            s3.shapes.title.text = "Претензии по причинам"
            pivot = claims_df.pivot_table(index="Месяц", columns="Причина",
                                          values="id", aggfunc="count").fillna(0)
            cd3 = CategoryChartData()
            cd3.categories = pivot.index.tolist()
            for col in pivot.columns:
                cd3.add_series(col, pivot[col].tolist())
            s3.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED,
                                Inches(0.5), Inches(1.5), Inches(9), Inches(4.5), cd3)

        # Слайд 4: топ видов брака
        if items:
            last_month = max(c["Месяц"] for c in claims)
            last_ids = [c["id"] for c in claims if c["Месяц"] == last_month]
            last_items = [i for i in items if i["claim_id"] in last_ids]
            if last_items:
                df_top = pd.DataFrame(last_items).groupby("Тип")["Количество"].sum().reset_index()
                s4 = prs.slides.add_slide(prs.slide_layouts[5])
                s4.shapes.title.text = f"Брак по типам — {last_month}"
                cd4 = CategoryChartData()
                cd4.categories = df_top["Тип"].tolist()
                cd4.add_series("Кол-во", df_top["Количество"].tolist())
                s4.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,
                                    Inches(0.5), Inches(1.5), Inches(9), Inches(4.5), cd4)

        prs.save("report.pptx")
        with open("report.pptx", "rb") as f:
            st.download_button("💾 Скачать PPTX", f, file_name="report.pptx",
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

# =====================================================
# ВКЛАДКА: ПРОДАЖИ
# =====================================================
elif tab == "💾 Продажи":
    st.header("💾 Продажи по месяцам")

    with st.form("sales_form"):
        m = st.text_input("Месяц (ГГГГ-ММ)", value=datetime.now().strftime("%Y-%m"))
        s = st.number_input("Сумма продаж, руб.", value=0.0, step=1000.0)
        c = st.number_input("Контейнеров", value=0, step=1)
        u = st.number_input("Продажи, шт.", value=0, step=1)
        if st.form_submit_button("💾 Сохранить"):
            new_row = {"Месяц": m, "Сумма_продаж": s,
                       "Контейнеры": c, "Продажи_шт": u}
            data = load_json("sales.json", [])
            data = [r for r in data if r["Месяц"] != m]
            data.append(new_row)
            data.sort(key=lambda r: r["Месяц"])
            save_json("sales.json", data)
            st.success(f"Продажи за {m} сохранены")

    st.subheader("Таблица продаж")
    st.dataframe(pd.DataFrame(load_json("sales.json", [])), use_container_width=True)
