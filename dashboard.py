from dash_bootstrap_components._components.Col import Col
import dash
from dash_bootstrap_components._components.Card import Card
import dash_core_components as dcc
import dash_html_components as html
import dash_bootstrap_components as dbc
from dash.dependencies import Input, Output, State
import plotly.graph_objects as go
import plotly.express as px
import pandas as pd
import numpy as np
from sklearn import linear_model
import os
import webbrowser  # Importar webbrowser para abrir automáticamente el navegador

# AÑADIENDO RUTA Y ABRIENDO ARCHIVO CON PANDAS 
path = os.path.dirname(os.path.abspath(__file__))
path = os.path.join(path, "misc", "cars.csv")

df = pd.read_csv(path)

# CONSTRUYENDO TABLA DE DATOS
def draw_table():
    layout = go.Layout(
        showlegend=False,
        paper_bgcolor='rgba(0, 0, 0, 0)',
        margin=dict(r=10, l=10, b=10, t=10)
    )

    config = {'displayModeBar': False}

    # DIBUJANDO TABLA
    table = go.Figure(layout=layout)
    table.add_trace(
        go.Table(
            header=dict(values=list(df.columns),
                        line_color='black',
                        fill_color='paleturquoise',
                        align='center',
                        font_color='black'
                        ),
            cells=dict(
                values=[df.Marca, df.Modelo, df.Volumen, df.Peso, df.CO2],
                line_color='black',
                fill_color='lavender',
                align='left',
                font_color='black'
            )
        )
    )

    table_show = html.Div([
        dbc.Card(
            dbc.CardBody([
                html.H1("Tabla de datos"),
                dcc.Graph(
                    figure=table, config=config
                )
            ])
        )
    ])

    return table_show

# CONSTRUYENDO RADIO ITEMS DE GRÁFICO DE BARRAS
def radio_items_barras():
    cabeceras = list(df.columns)
    radio_items = html.Div([
        html.H1("Gráfico de barras"),
        dcc.RadioItems(
            id='radio_ejey_barras',
            options=[{'label': i, 'value': i} for i in cabeceras[2:]],
            value=cabeceras[2],
            labelStyle={
                'display': 'inline-block',
                'paddingRight': '20px'
            }
        )
    ])

    return radio_items

# CONTRUYENDO GRÁFICO DE BARRAS CON CALLBACKS
def grafico_barras():
    @app.callback(
        Output('grafico_barras', 'figure'),
        Input('radio_ejey_barras', 'value')
    )
    def act_grafico_barras(radio_ejey_barras):
        layout = go.Layout(
            showlegend=False,
            paper_bgcolor='rgba(0, 0, 0, 0)',
            yaxis_color='white',
            xaxis_color='white',
            xaxis={'title': 'Modelos de autos'},
            yaxis={'title': radio_ejey_barras},
            margin=dict(r=10, l=10, b=10, t=10)
        )

        fig_barras = go.Figure(layout=layout)
        fig_barras.add_trace(
            go.Bar(
                y=df[radio_ejey_barras],
                x=df['Modelo'],
                text=["Marca: {}".format(u) for u in df['Marca']]
            )
        )
        return fig_barras

    config = {'displayModeBar': False}
    barra = html.Div([
        dbc.Card(
            dbc.CardBody([
                radio_items_barras(),
                dcc.Graph(id='grafico_barras', config=config)
            ])
        )
    ])

    return barra

# GRÁFICO DE DISPERSIÓN CON REGRESIÓN MÚLTIPLE
def grafico_dispersion_regresion():
    X = df[['Peso', 'Volumen']]
    Y = df['CO2']

    regresion = linear_model.LinearRegression()
    regresion.fit(X, Y)

    r2 = regresion.score(X, Y)
    coef = regresion.coef_
    intercepto = regresion.intercept_

    y_pred = regresion.predict(X)

    fig_regresion = go.Figure()
    fig_regresion.add_trace(
        go.Scatter3d(
            x=df['Peso'],
            y=df['Volumen'],
            z=df['CO2'],
            mode='markers',
            marker=dict(size=5, color=df['CO2'], colorscale='Viridis'),
            name='Datos reales'
        )
    )

    fig_regresion.add_trace(
        go.Scatter3d(
            x=df['Peso'],
            y=df['Volumen'],
            z=y_pred,
            mode='lines',
            line=dict(color='red', width=2),
            name='Modelo de regresión'
        )
    )

    fig_regresion.update_layout(
        scene=dict(
            xaxis_title='Peso',
            yaxis_title='Volumen',
            zaxis_title='CO2'
        ),
        title=f"Regresión Múltiple (R² = {r2:.2f})"
    )

    interpretacion = html.Div([
        html.H4("Interpretación:"),
        html.P(f"La regresión múltiple indica que el modelo explica aproximadamente el {r2 * 100:.2f}% de la variabilidad de las emisiones de CO2."),
        html.P(f"Los coeficientes obtenidos sugieren que por cada kilogramo adicional de peso, el CO2 aumenta en aproximadamente {coef[0]:.2f} g, mientras que por cada cm³ adicional de volumen, el CO2 aumenta en {coef[1]:.2f} g.")
    ])

    grafico = html.Div([
        dbc.Card(
            dbc.CardBody([
                html.H1("Gráfico de Dispersión con Regresión Múltiple"),
                dcc.Graph(figure=fig_regresion),
                interpretacion
            ])
        )
    ])

    return grafico

# CONFIGURANDO BOOTSTRAP
external_stylesheets = [dbc.themes.SLATE]
app = dash.Dash(__name__, external_stylesheets=external_stylesheets)

# CONSTRUYENDO DASHBOARD
body = html.Div([
    dbc.Card(
        dbc.CardBody([
            dbc.Row([
                dbc.Col(
                    html.Div([
                        html.H1('Dashboard')
                    ]), width=2
                )
            ], align='center', justify='center'),
            html.Br(),
            dbc.Row([
                dbc.Col([
                    draw_table()
                ], width=6),
                dbc.Col([
                    grafico_barras()
                ], width=6)
            ]),
            html.Br(),
            dbc.Row([
                dbc.Col([
                    grafico_dispersion_regresion()
                ], width=12)
            ])
        ]), color='dark'
    )
], className="text-light")

# DIBUJAN EL BODY EN EL LAYOUT DE DASH
app.layout = html.Div([body])

# CORRER SERVIDOR
if __name__ == '__main__':
    port = 8050  # Puerto donde se ejecutará el servidor
    url = f"http://127.0.0.1:{port}/"
    webbrowser.open_new(url)  # Abre automáticamente el navegador
    app.run_server(debug=True, port=port)
