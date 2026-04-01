from datetime import datetime
from io import StringIO

import pandas as pd
from dash import dash_table, dcc, html

ANO_MIN, ANO_MAX = 2000, datetime.now().year

ACCENT = "#C9A84C"
BLUE = "#2563EB"
TEAL = "#0D9488"
ROSE = "#E11D48"


SELECT_FIELDS = [
    ("de_fase", "Fase"),
    ("de_tipo_operacao", "Tipo de Operação"),
    ("nm_proponente", "Proponente"),
    ("sg_fonte", "Fonte"),
    ("de_esfera", "Esfera"),
    ("nm_regiao", "Região"),
    ("nm_setor", "Setor"),
    ("nm_subsetor", "Subsetor"),
    ("sg_setor", "Coordenação"),
    ("sys", "SYS"),
    ("nm_limite", "Limite"),
]

DEFAULT_SELECTIONS = {
    "de_fase": [
        "Aprovado",
        "Em preparação",
        "Em execução",
        "Em negociação",
        "Negociação concluída",
        "Aguardando Assinatura",
    ],
    "de_tipo_operacao": [
        "Operação de Crédito Externo",
        "Operação de crédito externo",
    ],
}


def glass_card(*children, cls="") -> html.Div:
    return html.Div(className=f"glass-card {cls}".strip(), children=list(children))


def section_head(title: str, subtitle: str = "") -> html.Div:
    return html.Div(
        className="section-header-block",
        children=[
            html.Div(title, className="section-title"),
            html.Div(subtitle, className="section-subtitle") if subtitle else None,
        ],
    )


def _empty_state(title: str, text: str) -> html.Div:
    return glass_card(
        cls="empty-state",
        *[
            html.Div(title, className="empty-title"),
            html.P(text, className="empty-text"),
        ],
    )


def _clean_series_for_options(df: pd.DataFrame, col: str) -> pd.Series:
    if col not in df.columns:
        return pd.Series(dtype="string")

    s = df[col].astype("string")
    s = s.fillna("Não informado").replace(["<NA>", "nan", "None", ""], "Não informado")
    return s


def _get_options(df: pd.DataFrame, col: str) -> list[dict]:
    s = _clean_series_for_options(df, col)
    vals = sorted([str(v) for v in s.dropna().unique().tolist() if str(v).strip() != ""])
    return [{"label": v, "value": v} for v in vals]


def _selection_dropdown(df: pd.DataFrame, col: str, label: str) -> html.Div:
    options = _get_options(df, col)
    option_values = {opt["value"] for opt in options}
    default_values = [v for v in DEFAULT_SELECTIONS.get(col, []) if v in option_values]

    return html.Div(
        className="selection-item",
        children=[
            html.Label(label, className="filter-label"),
            dcc.Dropdown(
                id=f"carteira-select-{col}",
                options=options,
                value=default_values,
                multi=True,
                placeholder=f"Escolha {label.lower()}...",
                className="lovable-dropdown",
            ),
        ],
        style={"minWidth": "220px"},
    )


def _get_operacoes_default_columns(df: pd.DataFrame) -> list[str]:
    defaults = ["cd_pleito", "nu_processo_sei", "nm_pleito", "sg_pleito", "nm_proponente", "de_tipo_operacao", "de_fase", "de_fonte", "sg_fonte", "vl_financiamento_dolar",	"vl_contrapartida_dolar", "de_esfera", "nm_regiao", "nm_setor",	"nm_subsetor", "dt_primeiro_recebimento","dt_ultimo_recebimento","de_ult_versao_pleito","dt_primeira_cofiex","dt_validade_recomendacao","dt_reuniao_negociacao","dt_aprov_diretoria", "dt_aprovacao_senado","dt_publicacao_dou","dt_assinatura","dt_efetividade","dt_primeiro_desembolso","dt_ultimo_desembolso_original","dt_ultimo_desembolso_vigente","dt_primeira_amortizacao","dt_ultima_amortizacao", "dt_encerramento", "nu_resolucao_senado", "nu_operacao", "sg_setor","nm_tecnico", "de_email", "sys","nm_limite"]
    return [col for col in defaults if col in df.columns]


def painel_page_layout(df_json: str | None) -> html.Div:
    header = html.Div(
        className="page-header",
        children=[
            html.H1("Painel Analítico", className="page-title"),
            html.P(
                "Visão estruturada dos financiamentos autorizados pela Cofiex.",
                className="page-subtitle",
            ),
        ],
    )

    if not df_json:
        return html.Div(
            className="page-wrap fade-in",
            children=[
                header,
                _empty_state(
                    "Base não carregada",
                    "Acesse a página de Dados e carregue a base para visualizar o Painel.",
                ),
            ],
        )

    df = pd.read_json(StringIO(df_json), orient="split")

    selectors = [_selection_dropdown(df, col, label) for col, label in SELECT_FIELDS]
    operacoes_column_options = [{"label": col, "value": col} for col in df.columns]
    operacoes_default_columns = _get_operacoes_default_columns(df)

    return html.Div(
        className="page-wrap fade-in",
        children=[
            header,
            html.Div(
                style={
                    "display": "grid",
                    "gridTemplateColumns": "1fr",
                    "gap": "18px",
                    "marginBottom": "18px",
                },
                children=[
                    glass_card(
                        cls="section-card",
                        *[
                            html.Details(
                                open=False,
                                className="carteira-details",
                                children=[
                                    html.Summary(
                                        className="carteira-details-summary",
                                        children=[
                                            html.Div(
                                                [
                                                    html.Div("Seleção de dados", className="section-title"),
                                                    html.Div(
                                                        "Clique para expandir e escolher quais informações deseja exibir na Carteira Ativa.",
                                                        className="section-subtitle",
                                                    ),
                                                ],
                                                className="carteira-details-summary-text",
                                            ),
                                            html.Div(
                                                "Expandir seleção de dados",
                                                className="carteira-details-summary-action",
                                            ),
                                        ],
                                    ),
                                    html.Div(
                                        className="carteira-details-content",
                                        style={"paddingTop": "18px"},
                                        children=[
                                            html.Div(
                                                className="section-header",
                                                style={
                                                    "display": "flex",
                                                    "justifyContent": "space-between",
                                                    "alignItems": "flex-start",
                                                    "gap": "16px",
                                                    "flexWrap": "wrap",
                                                    "marginBottom": "16px",
                                                },
                                                children=[
                                                    html.Div(
                                                        [
                                                            html.Div("Filtros de exibição", className="section-title"),
                                                            html.Div(
                                                                "Escolha, nas listas abaixo, quais informações você deseja manter visíveis na análise.",
                                                                className="section-subtitle",
                                                            ),
                                                        ]
                                                    ),
                                                    html.Button(
                                                        "Limpar seleções",
                                                        id="carteira-btn-clear-selections",
                                                        n_clicks=0,
                                                        className="btn btn-outline",
                                                    ),
                                                ],
                                            ),
                                            html.Div(
                                                className="carteira-selection-grid",
                                                style={
                                                    "display": "grid",
                                                    "gridTemplateColumns": "repeat(auto-fit, minmax(220px, 1fr))",
                                                    "gap": "14px",
                                                },
                                                children=selectors,
                                            ),
                                        ],
                                    ),
                                ],
                            )
                        ],
                    ),
                    glass_card(
                        cls="section-card",
                        *[
                            html.Div(
                                className="section-header",
                                style={"marginBottom": "16px"},
                                children=[
                                    html.Div(
                                        [
                                            html.Div("Parâmetros gerais", className="section-title"),
                                            html.Div(
                                                "Defina o período de análise e a métrica de visualização dos gráficos.",
                                                className="section-subtitle",
                                            ),
                                        ]
                                    ),
                                ],
                            ),
                            html.Div(
                                style={
                                    "display": "grid",
                                    "gridTemplateColumns": "minmax(420px, 4fr) minmax(240px, 1fr)",
                                    "gap": "18px",
                                    "alignItems": "end",
                                },
                                children=[
                                    html.Div(
                                        className="filter-group",
                                        children=[
                                            html.Label("Período", className="filter-label"),
                                            dcc.RangeSlider(
                                                id="carteira-ano-range",
                                                min=ANO_MIN,
                                                max=ANO_MAX,
                                                step=1,
                                                value=[2016, ANO_MAX],
                                                marks={y: str(y) for y in range(ANO_MIN, ANO_MAX + 1, 5)},
                                                tooltip={"placement": "bottom", "always_visible": True},
                                                className="painel-slider",
                                            ),
                                        ],
                                    ),
                                    html.Div(
                                        className="filter-group",
                                        children=[
                                            html.Label("Métrica", className="filter-label"),
                                            dcc.RadioItems(
                                                id="carteira-metrica",
                                                options=[
                                                    {"label": "Valor USD", "value": "valor"},
                                                    {"label": "Operações", "value": "qtd"},
                                                ],
                                                value="valor",
                                                inline=True,
                                                className="painel-radio",
                                                inputClassName="painel-radio-input",
                                                labelClassName="painel-radio-label",
                                            ),
                                        ],
                                    ),
                                ],
                            ),
                        ],
                    ),
                ],
            ),
            html.Div(
                className="metrics-grid",
                children=[
                    html.Div(id="carteira-kpi-proj", className="metric-card", style={"borderTopColor": ACCENT}),
                    html.Div(id="carteira-kpi-ops", className="metric-card", style={"borderTopColor": BLUE}),
                    html.Div(id="carteira-kpi-fin", className="metric-card", style={"borderTopColor": TEAL}),
                    html.Div(id="carteira-kpi-contra", className="metric-card", style={"borderTopColor": "#7C3AED"}),
                    html.Div(id="carteira-kpi-ano", className="metric-card", style={"borderTopColor": ROSE}),
                ],
            ),
            dcc.Tabs(
                id="carteira-tabs",
                value="visao",
                className="painel-tabs",
                children=[
                    dcc.Tab(
                        label="Visão Geral",
                        value="visao",
                        className="painel-tab",
                        selected_className="painel-tab--active",
                        children=[
                            html.Div(
                                className="tab-content",
                                children=[
                                    glass_card(
                                        *[
                                            section_head(
                                                "Evolução Temporal das Autorizações",
                                                "Data da primeira recomendação na Cofiex",
                                            ),
                                            dcc.Loading(
                                                type="dot",
                                                color=ACCENT,
                                                children=[
                                                    dcc.Graph(
                                                        id="carteira-fig-temporal",
                                                        config={"displayModeBar": False},
                                                        style={"height": "280px"},
                                                    )
                                                ],
                                            ),
                                        ]
                                    ),
                                    html.Div(
                                        className="charts-row",
                                        children=[
                                            glass_card(
                                                cls="col-2",
                                                *[
                                                    section_head("Distribuição por Setor", "Top 10 por volume"),
                                                    dcc.Loading(
                                                        type="dot",
                                                        color=ACCENT,
                                                        children=[
                                                            dcc.Graph(
                                                                id="carteira-fig-setor",
                                                                config={"displayModeBar": False},
                                                                style={"height": "320px"},
                                                            )
                                                        ],
                                                    ),
                                                ],
                                            ),
                                            glass_card(
                                                cls="col-1",
                                                *[
                                                    section_head("Por Fonte Financiadora"),
                                                    dcc.Loading(
                                                        type="dot",
                                                        color=ACCENT,
                                                        children=[
                                                            dcc.Graph(
                                                                id="carteira-fig-fonte",
                                                                config={"displayModeBar": False},
                                                                style={"height": "320px"},
                                                            )
                                                        ],
                                                    ),
                                                ],
                                            ),
                                        ],
                                    ),
                                ],
                            )
                        ],
                    ),
                    dcc.Tab(
                        label="Fase de andamento",
                        value="ciclo",
                        className="painel-tab",
                        selected_className="painel-tab--active",
                        children=[
                            html.Div(
                                className="tab-content",
                                children=[
                                    html.Div(
                                        className="charts-row",
                                        children=[
                                            glass_card(
                                                cls="col-2",
                                                *[
                                                    section_head(
                                                        "Fase de Andamento",
                                                        "Distribuição das operações de acordo com a fase de andamento",
                                                    ),
                                                    dcc.Loading(
                                                        type="dot",
                                                        color=ACCENT,
                                                        children=[
                                                            dcc.Graph(
                                                                id="carteira-fig-fase",
                                                                config={"displayModeBar": False},
                                                                style={"height": "420px"},
                                                            )
                                                        ],
                                                    ),
                                                ],
                                            ),
                                            glass_card(
                                                cls="col-1",
                                                *[
                                                    section_head(
                                                        "Percentual por Fase",
                                                        "Participação relativa das fases selecionadas",
                                                    ),
                                                    dcc.Loading(
                                                        type="dot",
                                                        color=ACCENT,
                                                        children=[
                                                            dcc.Graph(
                                                                id="carteira-fig-fase-pct",
                                                                config={"displayModeBar": False},
                                                                style={"height": "420px"},
                                                            )
                                                        ],
                                                    ),
                                                ],
                                            ),
                                        ],
                                    ),
                                ],
                            )
                        ],
                    ),
                    dcc.Tab(
                        label="Geografia",
                        value="geografia",
                        className="painel-tab",
                        selected_className="painel-tab--active",
                        children=[
                            html.Div(
                                className="tab-content",
                                children=[
                                    html.Div(
                                        className="charts-row",
                                        children=[
                                            glass_card(
                                                cls="col-1",
                                                *[
                                                    section_head("Por Esfera de Governo"),
                                                    dcc.Loading(
                                                        type="dot",
                                                        color=ACCENT,
                                                        children=[
                                                            dcc.Graph(
                                                                id="carteira-fig-esfera",
                                                                config={"displayModeBar": False},
                                                                style={"height": "260px"},
                                                            )
                                                        ],
                                                    ),
                                                ],
                                            ),
                                            glass_card(
                                                cls="col-1",
                                                *[
                                                    section_head("Por Região"),
                                                    dcc.Loading(
                                                        type="dot",
                                                        color=ACCENT,
                                                        children=[
                                                            dcc.Graph(
                                                                id="carteira-fig-regiao",
                                                                config={"displayModeBar": False},
                                                                style={"height": "260px"},
                                                            )
                                                        ],
                                                    ),
                                                ],
                                            ),
                                        ],
                                    ),
                                    glass_card(
                                        *[
                                            section_head("Ranking de Unidades Federativas", "Top 15 por volume de financiamento"),
                                            dcc.Loading(
                                                type="dot",
                                                color=ACCENT,
                                                children=[
                                                    dcc.Graph(
                                                        id="carteira-fig-uf",
                                                        config={"displayModeBar": False},
                                                        style={"height": "360px"},
                                                    )
                                                ],
                                            ),
                                        ]
                                    ),
                                    glass_card(
                                        *[
                                            section_head("Mapa por Unidade Federativa", "Distribuição espacial dos financiamentos entre as UFs, excluindo-se os de abrangência federal"),
                                            dcc.Loading(
                                                type="dot",
                                                color=ACCENT,
                                                children=[
                                                    dcc.Graph(
                                                        id="carteira-fig-mapa-uf",
                                                        config={"displayModeBar": False},
                                                        style={"height": "520px"},
                                                    )
                                                ],
                                            ),
                                        ]
                                    ),
                                ],
                            )
                        ],
                    ),
                    dcc.Tab(
                        label="Operações",
                        value="operacoes",
                        className="painel-tab",
                        selected_className="painel-tab--active",
                        children=[
                            html.Div(
                                className="tab-content",
                                children=[
                                    glass_card(
                                        *[
                                            section_head(
                                                "Tabela de Operações",
                                                "Visualize os dados filtrados e escolha quais colunas deseja exibir",
                                            ),
                                            html.Div(
                                                className="section-header",
                                                style={
                                                    "display": "flex",
                                                    "justifyContent": "space-between",
                                                    "alignItems": "flex-start",
                                                    "gap": "16px",
                                                    "flexWrap": "wrap",
                                                    "marginBottom": "16px",
                                                },
                                                children=[
                                                    html.Div(
                                                        [
                                                            html.Div("Colunas da tabela", className="section-title"),
                                                            html.Div(
                                                                "Selecione as colunas que deseja exibir na tabela abaixo.",
                                                                className="section-subtitle",
                                                            ),
                                                        ]
                                                    ),
                                                    html.Div(
                                                        id="carteira-operacoes-selected-count",
                                                        className="section-counter",
                                                    ),
                                                ],
                                            ),
                                            dcc.Dropdown(
                                                id="carteira-operacoes-column-selector",
                                                options=operacoes_column_options,
                                                value=operacoes_default_columns,
                                                multi=True,
                                                placeholder="Escolha as colunas da tabela...",
                                                className="lovable-dropdown",
                                            ),
                                            html.Div(style={"height": "16px"}),
                                            html.Div(
                                                className="action-row",
                                                children=[
                                                    html.Button(
                                                        "Selecionar todas",
                                                        id="carteira-btn-select-all-operacoes",
                                                        n_clicks=0,
                                                        className="btn btn-outline",
                                                    ),
                                                    html.Button(
                                                        "Limpar seleção",
                                                        id="carteira-btn-clear-operacoes",
                                                        n_clicks=0,
                                                        className="btn btn-outline",
                                                    ),
                                                    html.Button(
                                                        "Exportar Excel",
                                                        id="carteira-btn-export-operacoes",
                                                        n_clicks=0,
                                                        className="btn btn-primary",
                                                    ),
                                                ],
                                            ),
                                            dcc.Download(id="carteira-download-excel-operacoes"),
                                            html.Div(
                                                className="table-shell",
                                                children=[
                                                    dash_table.DataTable(
                                                        id="carteira-operacoes-table",
                                                        page_size=15,
                                                        filter_action="native",
                                                        sort_action="native",
                                                        style_table={"overflowX": "auto"},
                                                        style_cell={
                                                            "textAlign": "left",
                                                            "padding": "10px 14px",
                                                            "minWidth": "120px",
                                                            "maxWidth": "320px",
                                                            "whiteSpace": "nowrap",
                                                            "overflow": "hidden",
                                                            "textOverflow": "ellipsis",
                                                            "fontSize": "12px",
                                                            "border": "none",
                                                        },
                                                        style_header={
                                                            "fontWeight": "700",
                                                            "fontSize": "11px",
                                                            "textTransform": "uppercase",
                                                            "letterSpacing": "0.04em",
                                                            "backgroundColor": "#F8FAFC",
                                                            "border": "none",
                                                            "color": "#64748B",
                                                        },
                                                        style_data={"backgroundColor": "#ffffff", "border": "none"},
                                                        style_data_conditional=[
                                                            {"if": {"row_index": "odd"}, "backgroundColor": "#FAFCFF"}
                                                        ],
                                                        css=[
                                                            {
                                                                "selector": ".dash-spreadsheet-container table",
                                                                "rule": "border-collapse: separate; border-spacing: 0; width: 100%;",
                                                            },
                                                            {
                                                                "selector": ".dash-spreadsheet-container tr:hover td",
                                                                "rule": "background-color: #F1F5F9; transition: background-color 0.15s;",
                                                            },
                                                        ],
                                                        data=[],
                                                        columns=[],
                                                    )
                                                ],
                                            ),
                                        ]
                                    ),
                                ],
                            )
                        ],
                    ),
                ],
            ),
        ],
    )