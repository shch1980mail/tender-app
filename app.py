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
        st.success(f"Претензия №{new_id} сохранена ({len(new_items)} деталей)")
        st.session_state.n_items = 1

# =====================================================
# СПРАВОЧНИК
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
# ФРАХТ
# =====================================================
elif tab == "🚢 Фрахт":
    st.header("🚢 Ставки фрахта из Китая")

    if not freight.empty:
        # 5 графиков: Море-НСК, Море-КРД, ЖД-НСК, ЖД-КРД, Общий

        def plot_by_year(df, value_col, title):
            fig = go.Figure()
            for year in sorted(df["Год"].unique()):
                sub = df[df["Год"] == year].set_index("Месяц").reindex(MONTHS)
                fig.add_trace(go.Bar(
                    x=MONTHS, y=sub[value_col], name=str(year)
                ))
            fig.update_layout(title=title, barmode="group", height=400,
                              xaxis_title="Месяц", yaxis_title=value_col)
            return fig

        st.subheader("Море → НСК")
        sub = freight[(freight["Направление"] == "НСК") & (freight["Тип"] == "Море")]
        st.plotly_chart(plot_by_year(sub, "USD", "Море → НСК, USD"), use_container_width=True)
        st.plotly_chart(plot_by_year(sub, "RUB", "Море → НСК, RUB"), use_container_width=True)

        st.subheader("Море → Краснодар")
        sub = freight[(freight["Направление"] == "КРД") & (freight["Тип"] == "Море")]
        st.plotly_chart(plot_by_year(sub, "USD", "Море → КРД, USD"), use_container_width=True)
        st.plotly_chart(plot_by_year(sub, "RUB", "Море → КРД, RUB"), use_container_width=True)

        st.subheader("ЖД → НСК")
        sub = freight[(freight["Направление"] == "НСК") & (freight["Тип"] == "ЖД")]
        st.plotly_chart(plot_by_year(sub, "RUB", "ЖД → НСК, RUB"), use_container_width=True)

        st.subheader("ЖД → Краснодар")
        sub = freight[(freight["Направление"] == "КРД") & (freight["Тип"] == "ЖД")]
        if not sub.empty:
            st.plotly_chart(plot_by_year(sub, "RUB", "ЖД → КРД, RUB"), use_container_width=True)
        else:
            st.info("Данных по ЖД → КРД пока нет")

        st.subheader("Общий график (RUB, все направления)")
        fig_all = go.Figure()
        for (dir_, tip_), grp in freight.groupby(["Направление", "Тип"]):
            grp = grp.sort_values(["Год"])
            fig_all.add_trace(go.Scatter(
                x=grp["Месяц"] + " " + grp["Год"].astype(str),
                y=grp["RUB"], name=f"{tip_} → {dir_}",
                mode="lines+markers"
            ))
        fig_all.update_layout(height=450, xaxis_title="Месяц", yaxis_title="RUB")
        st.plotly_chart(fig_all, use_container_width=True)

        with st.expander("📋 Таблица ставок"):
            st.dataframe(freight, use_container_width=True)

# =====================================================
# ПРОДАЖИ
# =====================================================
elif tab == "💾 Продажи":
    st.header("💾 Продажи по месяцам")

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
            st.success(f"Продажи за {m} сохранены")

    st.dataframe(pd.DataFrame(load_json("sales.json", [])), use_container_width=True)

# =====================================================
# ЭКСПОРТ PPTX
# =====================================================
elif tab == "📥 Экспорт PPTX":
    st.header("📥 Экспорт в PPTX")
    if st.button("Собрать PPTX"):
        prs = Presentation()

        # Слайд 1: фрахт USD НСК
        s1 = prs.slides.add_slide(prs.slide_layouts[5])
        s1.shapes.title.text = "Фрахт Море → НСК (USD)"
        sub = freight[(freight["Направление"] == "НСК") & (freight["Тип"] == "Море")]
        cd1 = CategoryChartData()
        cd1.categories = MONTHS
        for year in sorted(sub["Год"].unique()):
            y = sub[sub["Год"] == year].set_index("Месяц").reindex(MONTHS)["USD"].fillna(0).tolist()
            cd1.add_series(str(year), y)
        s1.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,
                            Inches(0.5), Inches(1.5), Inches(9), Inches(4.5), cd1)

        # Слайд 2: фрахт RUB НСК
        s2 = prs.slides.add_slide(prs.slide_layouts[5])
        s2.shapes.title.text = "Фрахт Море → НСК (RUB)"
        cd2 = CategoryChartData()
        cd2.categories = MONTHS
        for year in sorted(sub["Год"].unique()):
            y = sub[sub["Год"] == year].set_index("Месяц").reindex(MONTHS)["RUB"].fillna(0).tolist()
            cd2.add_series(str(year), y)
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

        # Слайд 4: брак по типам
        if items:
            df_items = pd.DataFrame(items)
            df_items["Месяц"] = df_items["claim_id"].map({c["id"]: c["Месяц"] for c in claims})
            pivot2 = df_items.pivot_table(index="Месяц", columns="Тип",
                                          values="Количество", aggfunc="sum").fillna(0)
            s4 = prs.slides.add_slide(prs.slide_layouts[5])
            s4.shapes.title.text = "Брак по типам"
            cd4 = CategoryChartData()
            cd4.categories = pivot2.index.tolist()
            for col in pivot2.columns:
                cd4.add_series(col, pivot2[col].tolist())
            s4.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED,
                                Inches(0.5), Inches(1.5), Inches(9), Inches(4.5), cd4)

        # Слайд 5: сводка по видам (последний месяц)
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
                s5 = prs.slides.add_slide(prs.slide_layouts[5])
                s5.shapes.title.text = f"Виды брака — {last_month}"
                rows_t = len(summary) + 1
                table = s5.shapes.add_table(rows_t, 3,
                    Inches(0.5), Inches(1.5), Inches(9), Inches(0.4 * rows_t)).table
                table.cell(0, 0).text = "Тип"
                table.cell(0, 1).text = "Кол-во"
                table.cell(0, 2).text = "Виды"
                for i, row in summary.iterrows():
                    table.cell(i + 1, 0).text = str(row["Тип"])
                    table.cell(i + 1, 1).text = str(int(row["Кол_во"]))
                    table.cell(i + 1, 2).text = str(row["Виды"])

        prs.save("report.pptx")
        with open("report.pptx", "rb") as f:
            st.download_button("💾 Скачать PPTX", f, file_name="report.pptx",
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
