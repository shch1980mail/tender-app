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

COLORS = {
    "Поддон": "#E45756",
    "Крыша": "#F58518",
    "Экран": "#4C78A8",
    "Стекло": "#54A24B",
    "Профиль": "#EECA3B",
    "Панель": "#72B7B2",
    "Некомплект": "#B279A2",
    "Центральная стойка": "#9D755D",
    "Прочее": "#BAB0AC",
}

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
rates = load_json("rates.json", [])

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

    def get_rate(month, year):
        for r in rates:
            if r["Месяц"] == month and r["Год"] == year:
                return r["Курс"]
        return 90.0

    def calc_components(row):
        rate = get_rate(row["Месяц"], row["Год"])
        more_usd = row.get("Море_USD") or 0
        gd_usd = row.get("ЖД_USD") or 0
        gd_rub = row.get("ЖД_RUB") or 0
        avto_rub = row.get("Авто_RUB") or 0

        more_rub = more_usd * rate
        if row["Тип"] == "МОРЕ":
            gd_rub_calc = gd_rub
        else:
            gd_rub_calc = gd_usd * rate

        return pd.Series({
            "Море_USD_знач": more_usd,
            "Море_RUB": more_rub,
            "ЖД_USD_знач": gd_usd,
            "ЖД_RUB_пересчёт": gd_rub_calc,
            "Авто_RUB_знач": avto_rub,
        })

    if not freight.empty:
        freight[["Море_USD_знач", "Море_RUB",
                 "ЖД_USD_знач", "ЖД_RUB_пересчёт",
                 "Авто_RUB_знач"]] = freight.apply(calc_components, axis=1)
        freight["Полный_МОРЕ_RUB"] = freight["Море_RUB"] + freight["ЖД_RUB_пересчёт"] + freight["Авто_RUB_знач"]
        freight["Полный_ЖД_RUB"] = freight["ЖД_RUB_пересчёт"] + freight["Авто_RUB_знач"]

    def plot_by_year(df, value_col, title, y_title="RUB"):
        fig = go.Figure()
        for year in sorted(df["Год"].unique()):
            sub = df[df["Год"] == year].copy()
            sub = sub.set_index("Месяц").reindex(MONTHS)
            sub[value_col] = sub[value_col].ffill()
            fig.add_trace(go.Scatter(
                x=MONTHS, y=sub[value_col],
                name=str(year), mode="lines+markers",
                line=dict(width=2),
                marker=dict(size=6)
            ))
        fig.update_layout(
            title=title,
            height=300,
            xaxis_title=None,
            yaxis_title=y_title,
            hovermode="x unified",
            margin=dict(l=30, r=10, t=40, b=70)
        )
        fig.update_xaxes(
            tickangle=-90,
            tickmode="array",
            tickvals=MONTHS,
            ticktext=MONTHS,
            range=[-0.5, 11.5]
        )
        return fig

    st.subheader("1. Полная стоимость доставки через море (RUB)")
    sub = freight[freight["Тип"] == "МОРЕ"]
    if not sub.empty:
        st.plotly_chart(plot_by_year(sub, "Полный_МОРЕ_RUB",
                                     "Море + ЖД + Авто (RUB)"),
                        use_container_width=True)

    st.subheader("2. Полная стоимость доставки через ЖД (RUB)")
    sub = freight[freight["Тип"] == "ЖД"]
    if not sub.empty:
        st.plotly_chart(plot_by_year(sub, "Полный_ЖД_RUB",
                                     "ЖД + Авто (RUB)"),
                        use_container_width=True)

    st.subheader("3. Только морской фрахт (RUB)")
    sub = freight[freight["Тип"] == "МОРЕ"]
    if not sub.empty:
        st.plotly_chart(plot_by_year(sub, "Море_RUB",
                                     "Море (USD × курс месяца → RUB)"),
                        use_container_width=True)

    st.subheader("4. Только ЖД фрахт (RUB)")
    sub = freight[freight["Тип"] == "ЖД"]
    if not sub.empty:
        st.plotly_chart(plot_by_year(sub, "ЖД_RUB_пересчёт",
                                     "ЖД (USD × курс месяца → RUB)"),
                        use_container_width=True)

    st.subheader("5. Только морской фрахт (USD)")
    sub = freight[freight["Тип"] == "МОРЕ"]
    if not sub.empty:
        st.plotly_chart(plot_by_year(sub, "Море_USD_знач",
                                     "Море (USD)", y_title="USD"),
                        use_container_width=True)

    st.subheader("6. Только ЖД фрахт (USD)")
    sub = freight[freight["Тип"] == "ЖД"]
    if not sub.empty:
        st.plotly_chart(plot_by_year(sub, "ЖД_USD_знач",
                                     "ЖД (USD)", y_title="USD"),
                        use_container_width=True)

    with st.expander("📋 Таблица ставок"):
        show_cols = ["Месяц","Год","Тип","Море_USD","ЖД_USD","ЖД_RUB","Авто_RUB"]
        st.dataframe(
            freight[[c for c in show_cols if c in freight.columns]],
            use_container_width=True, height=300
        )

    with st.expander("📋 Таблица курсов"):
        rates_df = pd.DataFrame(rates)
        if not rates_df.empty:
            rates_df = rates_df.sort_values(["Год","Месяц"])
            st.dataframe(rates_df, use_container_width=True, height=250)

    st.subheader("💱 Курсы USD/RUB по месяцам")
    with st.form("rates_form"):
        c1, c2, c3 = st.columns(3)
        m = c1.selectbox("Месяц", MONTHS)
        g = c2.number_input("Год", min_value=2024, max_value=2030, value=2026)
        k = c3.number_input("Курс", value=84.0, step=0.01)
        if st.form_submit_button("💾 Сохранить курс"):
            data = load_json("rates.json", [])
            data = [r for r in data if not (r["Месяц"] == m and r["Год"] == int(g))]
            data.append({"Месяц": m, "Год": int(g), "Курс": k})
            data.sort(key=lambda r: (r["Год"], MONTHS.index(r["Месяц"])))
            save_json("rates.json", data)
            st.success(f"Курс за {m} {g} сохранён")
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

                    colors = [COLORS.get(t, "#BAB0AC") for t in pivot["Тип"]]
                    fig = go.Figure(go.Bar(
                        x=pivot["Тип"], y=pivot["Количество"],
                        text=pivot["Количество"], textposition="outside",
                        marker_color=colors
                    ))
                    fig.update_layout(
                        title=f"{month}", height=260,
                        showlegend=False,
                        bargap=0.2,
                        margin=dict(l=10, r=10, t=35, b=80)
                    )
                    fig.update_xaxes(
                        tickangle=-90,
                        range=[-0.5, len(pivot) - 0.5]
                    )
                    st.plotly_chart(fig, use_container_width=True)

                    desc = []
                    for _, r in pivot.iterrows():
                        views = sub[sub["Тип"] == r["Тип"]]["Вид"].unique()
                        views_short = ", ".join(list(views)[:3])
                        if len(views) > 3:
                            views_short += "…"
                        desc.append(f"**{r['Тип']}**: {views_short}")
                    st.caption(" | ".join(desc))

        st.subheader("📋 Сводка за последний месяц")
        last_month = months_sorted[-1]
        sub_last = df_items[df_items["Месяц"] == last_month]
        summary = sub_last.groupby("Тип").agg(
            Кол_во=("Количество", "sum"),
            Виды=("Вид", lambda x: ", ".join(sorted(set(x))))
        ).reset_index().sort_values("Кол_во", ascending=False)
        st.dataframe(summary, use_container_width=True, height=250)

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
    st.dataframe(pd.DataFrame(load_json("sales.json", [])),
                 use_container_width=True, height=300)

# =====================================================
# ЭКСПОРТ PPTX
# =====================================================
elif tab == "📥 Экспорт PPTX":
    st.header("📥 Экспорт")
    if st.button("Собрать PPTX"):
        prs = Presentation()

        sub = freight[freight["Тип"] == "МОРЕ"]
        if not sub.empty:
            s1 = prs.slides.add_slide(prs.slide_layouts[5])
            s1.shapes.title.text = "Полная стоимость: Море (RUB)"
            cd1 = CategoryChartData()
            cd1.categories = MONTHS
            for year in sorted(sub["Год"].unique()):
                y = sub[sub["Год"] == year].set_index("Месяц").reindex(MONTHS)
                y["Полный_МОРЕ_RUB"] = y["Полный_МОРЕ_RUB"].ffill()
                cd1.add_series(str(year), y["Полный_МОРЕ_RUB"].fillna(0).tolist())
            s1.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS,
                                Inches(0.5), Inches(1.5), Inches(9), Inches(4.5), cd1)

        sub = freight[freight["Тип"] == "ЖД"]
        if not sub.empty:
            s2 = prs.slides.add_slide(prs.slide_layouts[5])
            s2.shapes.title.text = "Полная стоимость: ЖД (RUB)"
            cd2 = CategoryChartData()
            cd2.categories = MONTHS
            for year in sorted(sub["Год"].unique()):
                y = sub[sub["Год"] == year].set_index("Месяц").reindex(MONTHS)
                y["Полный_ЖД_RUB"] = y["Полный_ЖД_RUB"].ffill()
                cd2.add_series(str(year), y["Полный_ЖД_RUB"].fillna(0).tolist())
            s2.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS,
                                Inches(0.5), Inches(1.5), Inches(9), Inches(4.5), cd2)

        sub = freight[freight["Тип"] == "МОРЕ"]
        if not sub.empty:
            s3 = prs.slides.add_slide(prs.slide_layouts[5])
            s3.shapes.title.text = "Только море (USD)"
            cd3 = CategoryChartData()
            cd3.categories = MONTHS
            for year in sorted(sub["Год"].unique()):
                y = sub[sub["Год"] == year].set_index("Месяц").reindex(MONTHS)
                y["Море_USD_знач"] = y["Море_USD_знач"].ffill()
                cd3.add_series(str(year), y["Море_USD_знач"].fillna(0).tolist())
            s3.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS,
                                Inches(0.5), Inches(1.5), Inches(9), Inches(4.5), cd3)

        sub = freight[freight["Тип"] == "ЖД"]
        if not sub.empty:
            s4 = prs.slides.add_slide(prs.slide_layouts[5])
            s4.shapes.title.text = "Только ЖД (USD)"
            cd4 = CategoryChartData()
            cd4.categories = MONTHS
            for year in sorted(sub["Год"].unique()):
                y = sub[sub["Год"] == year].set_index("Месяц").reindex(MONTHS)
                y["ЖД_USD_знач"] = y["ЖД_USD_знач"].ffill()
                cd4.add_series(str(year), y["ЖД_USD_знач"].fillna(0).tolist())
            s4.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS,
                                Inches(0.5), Inches(1.5), Inches(9), Inches(4.5), cd4)

        prs.save("report.pptx")
        with open("report.pptx", "rb") as f:
            st.download_button("💾 Скачать", f, file_name="report.pptx",
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
