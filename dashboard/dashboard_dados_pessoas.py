import dash
from dash import dcc, html, Input, Output, dash_table
import pandas as pd
import plotly.express as px

# Carregar os dados
file_path = 'dados_pessoas.xlsx'
df = pd.read_excel(file_path)

# Limpar e formatar os dados
def clean_dinheiro(valor):
    if isinstance(valor, (int, float)):
        return valor
    # Remover o prefixo R$, substituir vírgulas por pontos e corrigir separadores de milhar
    valor = valor.replace("R$", "").replace(".", "").replace(",", ".")
    return float(valor)

df['Dinheiro_Num'] = df['Dinheiro'].apply(clean_dinheiro)  # Criar uma coluna numérica para gráficos

df['Dinheiro'] = df['Dinheiro_Num'].apply(lambda x: f"R${x:,.2f}".replace(",", "_").replace(".", ",").replace("_", "."))
df['Peso'] = df['Peso'].str.replace('kg', '').astype(float)

# Adicionando tendências de vendas fictícias
import numpy as np
np.random.seed(42)
df['Mês'] = pd.date_range(start="2023-01-01", periods=len(df), freq="M")
df['Vendas'] = np.random.randint(100, 1000, size=len(df))

# Adicionando coluna para calcular o valor restante até R$50.000
df['Valor_Restante'] = df['Dinheiro_Num'].apply(lambda x: max(0, 50000 - x))

# Inicializar o aplicativo Dash
app = dash.Dash(__name__, suppress_callback_exceptions=True)
app.title = "Dashboard - Dados Pessoas"

# Layout do Dashboard
app.layout = html.Div([
    html.Div([
        html.H1("Dashboard - Dados de Pessoas", style={"text-align": "center"}),
        html.P("Análise interativa dos dados da planilha.", style={"text-align": "center"}),
    ], style={"padding": "20px", "background-color": "#f4f4f4"}),

    html.Div([
        # Dropdown para seleção de gráfico
        html.Label("Selecione o gráfico:"),
        dcc.Dropdown(
            id="dropdown-grafico",
            options=[
                {"label": "Dinheiro por Pessoa", "value": "dinheiro_pessoa"},
                {"label": "Peso vs Idade", "value": "peso_idade"},
                {"label": "Distribuição do Dinheiro", "value": "distribuicao_dinheiro"},
                {"label": "Gráfico de Pizza - Valor Restante", "value": "pizza_restante"},
                {"label": "Tendência de Vendas", "value": "tendencia_vendas"}
            ],
            value="dinheiro_pessoa",
            clearable=False
        ),
    ], style={"width": "30%", "margin": "20px auto"}),

    html.Div([
        dcc.Graph(id="grafico-dinamico"),
    ], style={"padding": "20px"}),

    html.Div([
        html.Button("Tabela de Dados", id="botao-tabela", n_clicks=0, style={"padding": "10px 20px", "font-size": "16px", "cursor": "pointer"}),
        html.Div(
            id="container-tabela",
            style={"display": "none"},
            children=[
                html.H3("Tabela de Dados", style={"text-align": "center"}),
                dash_table.DataTable(
                    id="tabela-dados",
                    columns=[{"name": col, "id": col} for col in df.columns],
                    data=df.to_dict("records"),
                    style_table={"overflowX": "auto"},
                    style_cell={"textAlign": "center"},
                    style_header={"backgroundColor": "lightgrey", "fontWeight": "bold"},
                    style_as_list_view=True
                )
            ]
        )
    ], style={"padding": "20px", "background-color": "#f4f4f4"}),
])

# Callback para atualizar o gráfico dinamicamente
@app.callback(
    Output("grafico-dinamico", "figure"),
    [Input("dropdown-grafico", "value")]
)
def update_graph(tipo_grafico):
    if tipo_grafico == "dinheiro_pessoa":
        fig = px.bar(df, x="Pessoa", y="Dinheiro_Num", title="Dinheiro por Pessoa", labels={"Pessoa": "Pessoa", "Dinheiro_Num": "Dinheiro (R$)"})
    elif tipo_grafico == "peso_idade":
        fig = px.scatter(df, x="Idade", y="Peso", title="Peso vs Idade", labels={"Idade": "Idade", "Peso": "Peso (kg)"})
    elif tipo_grafico == "distribuicao_dinheiro":
        df["Categoria"] = df["Dinheiro_Num"].apply(lambda x: "Acima de R$50.000" if x >= 50000 else "Abaixo de R$50.000")
        fig = px.histogram(df, x="Categoria", y="Dinheiro_Num", color="Categoria", title="Distribuição do Dinheiro", labels={"Categoria": "Categoria", "Dinheiro_Num": "Total (R$)"})
    elif tipo_grafico == "pizza_restante":
        df_below_50k = df[df["Dinheiro_Num"] < 50000]
        fig = px.pie(df_below_50k, values="Valor_Restante", names="Pessoa")
    elif tipo_grafico == "tendencia_vendas":
        fig = px.line(df, x="Mês", y="Vendas", title="Tendência de Vendas", labels={"Mês": "Mês", "Vendas": "Vendas (R$)"})
    else:
        fig = px.bar(df, x="Pessoa", y="Dinheiro_Num", title="Dinheiro por Pessoa")

    fig.update_layout(title_x=0.5)
    return fig

# Callback para mostrar/esconder tabela
@app.callback(
    Output("container-tabela", "style"),
    [Input("botao-tabela", "n_clicks")]
)
def toggle_table(n_clicks):
    if n_clicks % 2 == 1:
        return {"display": "block"}
    return {"display": "none"}

# Executar o aplicativo
if __name__ == "__main__":
    app.run_server(debug=True)
