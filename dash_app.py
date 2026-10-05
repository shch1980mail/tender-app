import dash
from dash import dcc, html
import plotly.graph_objects as go
import pandas as pd
import json
import os

# --- Загрузка данных (как в вашем старом коде) ---
# Пока оставим заглушку, чтобы проверить работу
MONTHS = ["янв", "фев", "мар", "апр", "май", "июн"]
DATA_2025 = [100, 110, 120, 130, 125, 140]
DATA_2026 = [105, 115, 125, 135, 130, 145]

# --- Инициализация приложения ---
app = dash.Dash(__name__)
server = app.server # ЭТА СТРОКА ОЧЕНЬ ВАЖНА ДЛЯ ХОСТИНГА

# --- Верстка страницы ---
app.layout = html.Div([
    html.H1("📊 Мои отчёты на Dash", style={'textAlign': 'center', 'color': '#333'}),
    
    # Вкладки
    dcc.Tabs([
        # Вкладка 1: Фрахт
        dcc.Tab(label='🚢 Фрахт', children=[
            dcc.Graph(
                figure=go.Figure(
                    data=[
                        go.Scatter(x=MONTHS, y=DATA_2025, name="2025", mode="lines+markers", line=dict(width=3)),
                        go.Scatter(x=MONTHS, y=DATA_2026, name="2026", mode="lines+markers", line=dict(width=3)),
                    ],
                    layout=go.Layout(
                        title="Ставки фрахта по годам",
                        height=400,
                        margin=dict(l=40, r=20, t=50, b=40)
                    )
                )
            )
        ]),
        # Вкладка 2: Брак
        dcc.Tab(label='🔧 Брак', children=[
            dcc.Graph(
                figure=go.Figure(
                    data=[
                        go.Bar(
                            x=["Стекло", "Поддон", "Крыша"], 
                            y=[5, 3, 2], 
                            name="Количество",
                            marker_color=['#E45756', '#F58518', '#4C78A8']
                        )
                    ],
                    layout=go.Layout(
                        title="Брак по типам за месяц",
                        height=400
                    )
                )
            )
        ]),
    ])
])

# --- Запуск (для локальной отладки) ---
if __name__ == '__main__':
    # Мы укажем порт и хост, которые требует Render
    app.run(debug=True, host='0.0.0.0', port=8050)
