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

MONTHS = ["янв","фев","мар","апр","май","июн","июл","авг","сен","окт","ноя","дек"]

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

if "usd_rate" not in st.session_state:
    st.session_state.usd_rate = 90.0

tab = st.sidebar.radio("Раздел", [
    "📊 Дашборд",
    "➕ Ввести претензию",
    "⚙️ Справочник",
    "🚢 Фрахт",
    "💾 Продажи",
    "📥 Экспорт PPTX"
])

# =====================================================
# ФРАХТ
# =====================================================
if tab == "🚢 Фрахт":
    st.header("🚢 Стоимость доставки в Новосибирск")

    st.sidebar.subheader("Курс USD → RUB")
    st.session_state.usd_rate = st.sidebar.number_input(
        "Курс", value=st.session_state.usd_rate, step=1.0)

    rate = st.session_state.usd_rate

    def calc_full(row):
        if row["Тип"] == "МОРЕ":
            more = (row.get("Море_USD") or 0) * rate
            gd = row.get("ЖД_RUB") or 0
            avto = row.get("Авто_RUB") or 0
            return more + gd + avto
        elif row["Тип"] == "ЖД":
            gd = (row.get("ЖД_USD") or 0) * rate
            avto = row.get("Авто_RUB") or 0
            return gd + avto
        return 0

    if not freight.empty:
        freight["Полная_RUB"] = freight.apply(calc_full, axis=1)

    def plot_by_year(df, title):
        fig = go.Figure()
        for year in sorted(df["Год"].unique()):
            sub = df[df["Год"] == year].copy()
            sub = sub.set_index("Месяц").reindex(MONTHS)
            sub["Полная_RUB"] = sub["Полная_RUB"].ffill()
            fig.add_trace(go.Scatter(
                x=MONTHS, y=sub["Полная_RUB"],
                name=str(year), mode="lines+markers"
            ))
        fig.update_layout(
            title=title, height=450,
            xaxis_title="Месяц", yaxis_title="RUB",
            hovermode="x unified"
        )
        return fig

    st.subheader("📈 Полная стоимость доставки через море (МОРЕ)")
    sub_more = freight[freight["Тип"] == "МОРЕ"]
    if not sub_more.empty:
        st.plotly_chart(
            plot_by_year(sub_more, "Полная стоимость: Море + ЖД + Авто (RUB)"),
            use_container_width=True
        )

    st.subheader("📈 Полная стоимость доставки через ЖД (ЖД)")
    sub_zhd = freight[freight["Тип"] == "ЖД"]
    if not sub_zhd.empty:
        st.plotly_chart(
            plot_by_year(sub_zhd, "Полная стоимость: ЖД + Авто (RUB)"),
            use_container_width=True
        )

    with st.expander("📋 Таблица ставок"):
        st.dataframe(freight, use_container_width=True)

    st.subheader("➕ Добавить ставку")
    with st.form("freight_form"):
        c1, c2, c3 = st.columns(3)
        m = c1.selectbox("Месяц", MONTHS)
        g = c2.number_input("Год", min_value=2024, max_value=2030, value=2026)
        t = c3.selectbox("Тип", ["МОРЕ", "ЖД"])
        c4, c5, c6 = st.columns(3)
        more_usd = c4.number_input("Море, USD", value=0.0, step=10.0)
        gd_usd = c5.number_input("ЖД, USD", value=0.0, step=10.0)
        gd_rub = c5.number_input("ЖД, RUB", value=0.0, step=1000.0)
        avto_rub = c6.number_input("Авто, RUB", value=0.0, step=1000.0)
        if st.form_submit_button("💾 Сохранить"):
            new_row = {"Месяц": m, "Год": int(g), "Тип": t,
                       "Море_USD": more_usd or None,
                       "ЖД_USD": gd_usd or None,
                       "ЖД_RUB": gd_rub or None,
                       "Авто_RUB": avto_rub or None}
            data = load_json("freight.json", [])
            data = [r for r in data if not (r["Месяц"] == m and r["Год"] == int(g) and r["Тип"] == t)]
            data.append(new_row)
            save_json("freight.json", data)
            st.success(f"Ставка за {m} {g} ({t}) сохранена")
            st.rerun()

# =====================================================
# ДАШБОРД
# =====================================================
elif tab == "📊 Дашборд":
    st.header("📊 Дашборд")

    claims_df_rows = []
    for c in claims:
        c_items = [i for i in items if i["claim_id"] == c["id"]]
        total_parts = sum(i["Количество"] for i in c_items)
        claims_df_rows.append({
            "id": c["id"], "Месяц": c["Месяц"],
            "Причина": c["Причина"], "Деталей": total_parts
        })
    claims_df = pd.DataFrame(claims_df_rows)

    if not claims_df.empty:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Всего претензий", len(claims_df))
        c2.metric("Всего деталей", int(claims_df["Деталей"].sum()))
        c3.metric("Транспортный бой",
                  len(claims_df[claims_df["Причина"] == "Транспортный бой"]))
        c4.metric("Вина поставщика",
                  len(claims_df[claims_df["Причина"] == "Вина поставщика"]))

    st.subheader("🔧 Брак по типам — по месяцам")

    if items:
        df_items = pd.DataFrame(items)
        df_items["Месяц"] = df_items["claim_id"].map(
            {c["id"]: c["Месяц"] for c in claims})

        months_sorted = sorted(df_items["Месяц"].unique())
        cols_per_row = 3
        for i in range(0, len(months_sorted), cols_per_row):
            cols = st.columns(cols_per_row)
            for j, month in enumerate(months_sorted[i:i+cols_per_row]):
                with cols[j]:
                    sub = df_items[df_items["Месяц"] == month]
                    pivot = sub.groupby("Тип")["Количество"].sum().reset_index()
                    pivot = pivot.sort_values("Количество", ascending=False)

                    fig = go.Figure(go.Bar(
                        x=pivot["Тип"], y=pivot["Количество"],
                        text=pivot["Количество"], textposition="outside",
                        marker_color="#4C78A8"
                    ))
                    fig.update_layout(
                        title=f"{month}", height=300,
                        xaxis_tickangle=-40, showlegend=False,
                        margin=dict(l=20, r=20, t=40, b=80)
                    )
                    st.plotly_chart(fig, use_container_width=True)

                    desc = []
                    for _, r in pivot.iterrows():
                        views = sub[sub["Тип"] == r["Тип"]]["Вид"].unique()
                        desc.append(f"**{r['Тип']}**: {', '.join(views)}")
                    st.caption(" | ".join(desc))

        st.subheader("📋 Сводка за последний месяц")
        last_month = months_sorted[-1]
        sub_last = df_items[df_items["Месяц"] == last_month]
        summary = sub_last.groupby("Тип").agg(
            Кол_во=("Количество", "sum"),
            Виды=("Вид", lambda x: ", ".join(sorted(set(x))))
        ).reset_index().sort_values("Кол_во", ascending=False)
        st.dataframe(summary, use_container_width=True)

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
        q = c3.number_input("Кол-во", min_value=0, value=1, key=f"q{i}")
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
        st.success(f"Претензия №{new_id} сохранена")
        st.session_state.n_items = 1

# =====================================================
# СПРАВОЧНИК
# =====================================================
elif tab == "⚙️ Справочник":
    st.header("⚙️ Справочник")
    new_type = st.text_input("Новый тип")
    if st.button("➕ Добавить тип") and new_type:
        if new_type not in types:
            types[new_type] = []
            save_json("defect_types.json", types)
            st.rerun()

    sel_type = st.selectbox("Тип", list(types.keys()))
    new_view = st.text_input("Новый вид")
    if st.button("➕ Добавить вид") and new_view:
        if new_view not in types[sel_type]:
            types[sel_type].append(new_view)
            save_json("defect_types.json", types)
            st.rerun()

    st.json(types)

# =====================================================
# ПРОДАЖИ
# =====================================================
elif tab == "💾 Продажи":
    st.header("💾 Продажи")
    with st.form("sales_form"):
        m = st.text_input("Месяц (ГГГГ-ММ)", value=datetime.now().strftime("%Y-%m"))
        s = st.number_input("Сумма продаж, руб.", value=0.0, step=1000.0)
        c = st.number_input("Контейнеров", value=0, step=1)
        u = st.number_input("Продажи, шт.", value=0, step=1)
        if st.form_submit_button("💾 Сохранить"):
            data = load_json("sales.json", [])
            data = [r for r in data if r["Месяц"] != m]
            data.append({"Месяц": m, "Сумма_продаж": s,
                         "Контейнеры": c, "Продажи_шт": u})
            data.sort(key=lambda r: r["Месяц"])
            save_json("sales.json", data)
            st.success("Сохранено")
    st.dataframe(pd.DataFrame(load_json("sales.json", [])), use_container_width=True)

# =====================================================
# ЭКСПОРТ PPTX
# =====================================================
elif tab == "📥 Экспорт PPTX":
    st.header("📥 Экспорт")
    if st.button("Собрать PPTX"):
        prs = Presentation()

        sub_more = freight[freight["Тип"] == "МОРЕ"]
        if not sub_more.empty:
            s1 = prs.slides.add_slide(prs.slide_layouts[5])
            s1.shapes.title.text = "Полная стоимость: МОРЕ (RUB)"
            cd1 = CategoryChartData()
            cd1.categories = MONTHS
            for year in sorted(sub_more["Год"].unique()):
                y = sub_more[sub_more["Год"] == year].set_index("Месяц").reindex(MONTHS)
                y["Полная_RUB"] = y["Полная_RUB"].ffill()
                cd1.add_series(str(year), y["Полная_RUB"].fillna(0).tolist())
            s1.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS,
                                Inches(0.5), Inches(1.5), Inches(9), Inches(4.5), cd1)

        sub_zhd = freight[freight["Тип"] == "ЖД"]
        if not sub_zhd.empty:
            s2 = prs.slides.add_slide(prs.slide_layouts[5])
            s2.shapes.title.text = "Полная стоимость: ЖД (RUB)"
            cd2 = CategoryChartData()
            cd2.categories = MONTHS
            for year in sorted(sub_zhd["Год"].unique()):
                y = sub_zhd[sub_zhd["Год"] == year].set_index("Месяц").reindex(MONTHS)
                y["Полная_RUB"] = y["Полная_RUB"].ffill()
                cd2.add_series(str(year), y["Полная_RUB"].fillna(0).tolist())
            s2.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS,
                                Inches(0.5), Inches(1.5), Inches(9), Inches(4.5), cd2)

        prs.save("report.pptx")
        with open("report.pptx", "rb") as f:
            st.download_button("💾 Скачать", f, file_name="report.pptx",
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
