from io import StringIO

import pandas as pd
from dash import dash_table, dcc, html

ACCENT = "#C9A84C"
BLUE = "#2563EB"
TEAL = "#0D9488"
ROSE = "#E11D48"
VIOLET = "#7C3AED"

SELECT_FIELDS = [
    ("de_fase", "Fase"),
    ("de_tipo_operacao", "Tipo de Operação"),
    ("nm_proponente", "Proponente"),
    ("sg_fonte", "Fonte"),
    ("de_esfera", "Esfera"),
    ("nm_regiao", "Região"),
    ("nm_setor", "Setor"),
    ("nm_subsetor", "Subsetor"),
    ("sys", "Sistema"),
    ("nm_limite", "Limite"),
    ("nm_cg", "CG"),
    ("nm_tecnico", "Técnico"),
    ("cd_pleito", "Código"),
]

DEFAULT_SELECTIONS = {}


# ── HELPERS ───────────────────────────────────────────────────────────────────

def metric_card(label: str, value: str, sub: str, color: str = BLUE) -> html.Div:
    return html.Div(
        className="metric-card",
        style={"borderTopColor": color},
        children=[
            html.Div(label, className="metric-label"),
            html.Div(value, className="metric-value"),
            html.Div(sub, className="metric-subtitle"),
        ],
    )


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
                id=f"carteira-ca-select-{col}",
                options=options,
                value=default_values,
                multi=True,
                placeholder=f"Escolha {label.lower()}...",
                className="lovable-dropdown",
                disabled=df.empty,   # <- importante
            ),
        ],
        style={"minWidth": "220px"},
    )


def _get_tabela_default_columns(df: pd.DataFrame) -> list[str]:
    defaults = [
        "sg_pleito", "nm_cg", "nm_tecnico", "nm_proponente", "nm_pleito",
        "de_fase", "de_tipo_operacao", "sg_fonte", "vl_financiamento_dolar",
        "vl_contrapartida_dolar", "de_esfera", "nm_regiao", "nm_setor",
        "nm_subsetor", "sys", "nm_limite",
    ]
    return [col for col in defaults if col in df.columns]


def _get_resolucao_default_columns(df: pd.DataFrame) -> list[str]:
    defaults = [
        "nm_tecnico", "nm_cg", "cd_pleito", "nm_pleito", "nm_proponente",
        "dt_primeira_cofiex", "dt_validade_recomendacao",
        "sg_fonte", 
    ]
    return [col for col in defaults if col in df.columns]


# ── ABA ACOMPANHAMENTO ────────────────────────────────────────────────────────

def _acompanhamento_tab_content(df: pd.DataFrame) -> html.Div:
    """
    Aba de acompanhamento da carteira ativa por técnico e coordenação.
    Exibe KPIs dinâmicos e gráficos de distribuição, respondendo aos filtros globais.
    """
    return html.Div(
        className="tab-content",
        children=[
            # KPIs dinâmicos
            html.Div(
                className="metrics-grid",
                style={"marginBottom": "20px"},
                children=[
                    html.Div(id="ca-acomp-kpi-tecnicos", className="metric-card", style={"borderTopColor": ACCENT}),
                    html.Div(id="ca-acomp-kpi-ops", className="metric-card", style={"borderTopColor": BLUE}),
                    html.Div(id="ca-acomp-kpi-valor", className="metric-card", style={"borderTopColor": TEAL}),
                    html.Div(id="ca-acomp-kpi-cgs", className="metric-card", style={"borderTopColor": VIOLET}),
                ],
            ),

            # Linha 1 — distribuição por técnico + por CG
            html.Div(
                className="charts-row",
                children=[
                    glass_card(
                        cls="col-2",
                        *[
                            section_head(
                                "Operações por Analista",
                                "Top 15 analistas com maior volume na carteira filtrada",
                            ),
                            dcc.Loading(
                                type="dot",
                                color=ACCENT,
                                children=[
                                    dcc.Graph(
                                        id="ca-acomp-fig-tecnico",
                                        config={"displayModeBar": False},
                                        style={"height": "380px"},
                                    )
                                ],
                            ),
                        ],
                    ),
                    glass_card(
                        cls="col-1",
                        *[
                            section_head(
                                "Por Coordenação Geral",
                                "Distribuição das operações por CG",
                            ),
                            dcc.Loading(
                                type="dot",
                                color=ACCENT,
                                children=[
                                    dcc.Graph(
                                        id="ca-acomp-fig-cg",
                                        config={"displayModeBar": False},
                                        style={"height": "380px"},
                                    )
                                ],
                            ),
                        ],
                    ),
                ],
            ),

            # Linha 2 — fases por técnico em largura total
            html.Div(
                className="charts-row",
                children=[
                    glass_card(
                        cls="col-1",
                        *[
                            section_head(
                                "Distribuição de Fases por Analista",
                                "Analistas — composição da carteira por fase",
                            ),
                            dcc.Loading(
                                type="dot",
                                color=ACCENT,
                                children=[
                                    dcc.Graph(
                                        id="ca-acomp-fig-fase-tecnico",
                                        config={"displayModeBar": False},
                                        style={"height": "420px"},
                                    )
                                ],
                            ),
                        ],
                    ),
                ],
            ),

            # Linha 3 — setor por coordenação embaixo, largura total
            html.Div(
                className="charts-row",
                children=[
                    glass_card(
                        cls="col-1",
                        *[
                            section_head(
                                "Setor por Coordenação",
                                "Distribuição setorial das operações por CG",
                            ),
                            dcc.Loading(
                                type="dot",
                                color=ACCENT,
                                children=[
                                    dcc.Graph(
                                        id="ca-acomp-fig-setor-cg",
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



# ── ABA VALIDADE DA RESOLUÇÃO ─────────────────────────────────────────────────

def _resolucao_tab_content(df: pd.DataFrame) -> html.Div:
    """
    Aba de controle de validade das resoluções COFIEX.
    Alerta para operações com validade próxima ou vencida.
    """
    resolucao_col_options = [{"label": col, "value": col} for col in df.columns] if not df.empty else []
    resolucao_default_cols = _get_resolucao_default_columns(df)

    return html.Div(
        className="tab-content",
        children=[
            # KPIs de alerta
            html.Div(
                className="metrics-grid",
                style={"marginBottom": "20px"},
                children=[
                    html.Div(id="ca-res-kpi-vencidas",   className="metric-card", style={"borderTopColor": ROSE}),
                    html.Div(id="ca-res-kpi-30dias",      className="metric-card", style={"borderTopColor": "#F97316"}),
                    html.Div(id="ca-res-kpi-90dias",      className="metric-card", style={"borderTopColor": ACCENT}),
                    html.Div(id="ca-res-kpi-180dias",      className="metric-card", style={"borderTopColor": TEAL}),
                    html.Div(id="ca-res-kpi-ok",          className="metric-card", style={"borderTopColor": BLUE}),
                ],
            ),

            # Filtro de horizonte
            glass_card(
                cls="section-card",
                *[
                    html.Div(
                        style={"width": "100%"},
                        children=[
                            html.Label("Horizonte de alerta", className="filter-label"),
                            html.Div(
                                style={
                                    "display": "flex",
                                    "justifyContent": "center",
                                    "width": "100%",
                                    "marginTop": "12px",
                                },
                                children=[
                                    dcc.RadioItems(
                                        id="ca-res-horizonte",
                                        options=[
                                            {"label": "Vencidas", "value": "vencidas"},
                                            {"label": "Próximos 30 dias", "value": "30"},
                                            {"label": "Próximos 90 dias", "value": "90"},
                                            {"label": "Próximos 180 dias", "value": "180"},
                                            {"label": "Todas", "value": "todas"},
                                        ],
                                        value="180",
                                        inline=True,
                                        className="painel-radio painel-radio-responsivo",
                                        inputClassName="painel-radio-input",
                                        labelClassName="painel-radio-label",
                                    )
                                ],
                            ),
                        ],
                    ),
                ],
            ),

            # Gráfico de distribuição temporal
            glass_card(
                *[
                    section_head(
                        "Validade das Resoluções",
                        "Quantidade de resoluções em cada mês",
                    ),
                    dcc.Loading(
                        type="dot", color=ACCENT,
                        children=[
                            dcc.Graph(
                                id="ca-res-fig-timeline",
                                config={"displayModeBar": False},
                                style={"height": "240px"},
                            )
                        ],
                    ),
                ]
            ),

            # Tabela de alertas
            glass_card(
                *[
                    section_head(
                        "Operações em Alerta",
                        "Selecione as colunas e exporte o relatório de validade",
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
                            html.Div([
                                html.Div("Colunas da tabela", className="section-title"),
                                html.Div(
                                    "Selecione as colunas que deseja exibir.",
                                    className="section-subtitle",
                                ),
                            ]),
                            html.Div(id="ca-res-selected-count", className="section-counter"),
                        ],
                    ),
                    dcc.Dropdown(
                        id="ca-res-column-selector",
                        options=resolucao_col_options,
                        value=resolucao_default_cols,
                        multi=True,
                        placeholder="Escolha as colunas...",
                        className="lovable-dropdown",
                    ),
                    html.Div(style={"height": "16px"}),
                    html.Div(
                        className="action-row",
                        children=[
                            html.Button(
                                "Exportar Excel",
                                id="ca-res-btn-export",
                                n_clicks=0,
                                className="btn btn-primary",
                            ),
                        ],
                    ),
                    dcc.Download(id="ca-res-download-excel"),
                    html.Div(
                        className="table-shell",
                        children=[
                            dash_table.DataTable(
                                id="ca-res-table",
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
                                    {"if": {"row_index": "odd"}, "backgroundColor": "#FAFCFF"},
                                    # Destaque vermelho para vencidas
                                    {
                                        "if": {"filter_query": "{dias_para_vencer} < 0"},
                                        "backgroundColor": "#FFF1F2",
                                        "color": "#9F1239",
                                    },
                                    # Destaque laranja para < 30 dias
                                    {
                                        "if": {
                                            "filter_query": "{dias_para_vencer} >= 0 && {dias_para_vencer} < 30"
                                        },
                                        "backgroundColor": "#FFF7ED",
                                        "color": "#9A3412",
                                    },
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


# ── LAYOUT PRINCIPAL ──────────────────────────────────────────────────────────

def carteira_analistas_page_layout(
    df_json_ca: str | None = None,
    auth_component=None,
) -> html.Div:

    children = []

    if auth_component is not None:
        children.append(auth_component)

    if df_json_ca:
        try:
            df = pd.read_json(StringIO(df_json_ca), orient="split")
        except Exception:
            df = pd.DataFrame()
    else:
        df = pd.DataFrame()

    has_data = not df.empty

    # selectors = [_selection_dropdown(df, col, label) for col, label in SELECT_FIELDS] if has_data else []
    selectors = [_selection_dropdown(df, col, label) for col, label in SELECT_FIELDS]
    tabela_column_options = [{"label": col, "value": col} for col in df.columns] if has_data else []
    tabela_default_columns = _get_tabela_default_columns(df) if has_data else []

    children.extend([
        html.Div(
            className="page-header",
            style={"marginTop": "6px"},
            children=[
                html.H1("Carteira Ativa — Coordenações Gerais", className="page-title"),
                html.P(
                    "Carregue e visualize a planilha da Carteira Ativa das coordenações gerais.",
                    className="page-subtitle",
                ),
            ],
        ),
        html.Div(
            className="action-row",
            children=[
                html.Button(
                    "Carregar página",
                    id="btn-load-ca",
                    n_clicks=0,
                    className="btn btn-primary",
                ),
            ],
        ),

        # ── Seção principal (só aparece com dados) ──
        html.Div(
            id="preview-section-ca",
            style={"display": "block" if has_data else "none"},
            children=[
                # Seleção de dados / filtros
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
                                                    "Clique para expandir e filtrar os registros da Carteira Ativa.",
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
                                                html.Div([
                                                    html.Div("Filtros de exibição", className="section-title"),
                                                    html.Div(
                                                        "Escolha os filtros que deseja aplicar às abas e aos KPIs.",
                                                        className="section-subtitle",
                                                    ),
                                                ]),
                                                html.Button(
                                                    "Limpar seleções",
                                                    id="carteira-ca-btn-clear-selections",
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

                # Abas
                dcc.Tabs(
                    id="carteira-ca-tabs",
                    value="dados",
                    className="painel-tabs",
                    children=[

                        # ── Aba: Dados ──
                        dcc.Tab(
                            label="Dados",
                            value="dados",
                            className="painel-tab",
                            selected_className="painel-tab--active",
                            children=[
                                html.Div(
                                    className="tab-content",
                                    children=[
                                        html.Div(
                                            id="status-box-ca",
                                            className="status-box status-idle",
                                            children=[
                                                html.Div("Status", className="status-title"),
                                                html.Pre(
                                                    "Clique em 'Carregar página' para iniciar.",
                                                    className="status-message",
                                                ),
                                            ],
                                        ),
                                        html.Div(id="summary-cards-ca", className="metrics-grid"),
                                        glass_card(*[
                                            section_head(
                                                "Pré-visualização",
                                                "Primeiras 200 linhas da planilha carregada",
                                            ),
                                            html.Div(
                                                className="table-shell",
                                                children=[
                                                    dash_table.DataTable(
                                                        id="preview-table-ca-dados",
                                                        page_size=15,
                                                        style_table={"overflowX": "auto"},
                                                        style_cell={
                                                            "textAlign": "left", "padding": "10px 14px",
                                                            "minWidth": "120px", "maxWidth": "320px",
                                                            "whiteSpace": "nowrap", "overflow": "hidden",
                                                            "textOverflow": "ellipsis",
                                                            "fontSize": "12px", "border": "none",
                                                        },
                                                        style_header={
                                                            "fontWeight": "700", "fontSize": "11px",
                                                            "textTransform": "uppercase",
                                                            "letterSpacing": "0.04em",
                                                            "backgroundColor": "#F8FAFC",
                                                            "border": "none", "color": "#64748B",
                                                        },
                                                        style_data={"backgroundColor": "#ffffff", "border": "none"},
                                                        style_data_conditional=[
                                                            {"if": {"row_index": "odd"}, "backgroundColor": "#FAFCFF"}
                                                        ],
                                                        css=[
                                                            {"selector": ".dash-spreadsheet-container table",
                                                             "rule": "border-collapse: separate; border-spacing: 0; width: 100%;"},
                                                            {"selector": ".dash-spreadsheet-container tr:hover td",
                                                             "rule": "background-color: #F1F5F9; transition: background-color 0.15s;"},
                                                        ],
                                                        data=[], columns=[],
                                                    )
                                                ],
                                            ),
                                        ]),
                                    ],
                                )
                            ],
                        ),

                        # ── Aba: Acompanhamento ──
                        dcc.Tab(
                            label="Acompanhamento",
                            value="acompanhamento",
                            className="painel-tab",
                            selected_className="painel-tab--active",
                            children=[_acompanhamento_tab_content(df)],
                        ),

                        # ── Aba: Validade da Resolução ──
                        dcc.Tab(
                            label="Validade da Resolução",
                            value="resolucao",
                            className="painel-tab",
                            selected_className="painel-tab--active",
                            children=[_resolucao_tab_content(df)],
                        ),

                        # ── Aba: Tabela ──
                        dcc.Tab(
                            label="Tabela",
                            value="tabela",
                            className="painel-tab",
                            selected_className="painel-tab--active",
                            children=[
                                html.Div(
                                    className="tab-content",
                                    children=[
                                        glass_card(*[
                                            section_head(
                                                "Tabela da Carteira Ativa",
                                                "Visualize os dados filtrados e escolha quais colunas deseja exibir.",
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
                                                    html.Div([
                                                        html.Div("Colunas da tabela", className="section-title"),
                                                        html.Div(
                                                            "Selecione as colunas que deseja exibir na tabela abaixo.",
                                                            className="section-subtitle",
                                                        ),
                                                    ]),
                                                    html.Div(
                                                        id="carteira-ca-selected-count",
                                                        className="section-counter",
                                                    ),
                                                ],
                                            ),
                                            dcc.Dropdown(
                                                id="carteira-ca-column-selector",
                                                options=tabela_column_options,
                                                value=tabela_default_columns,
                                                multi=True,
                                                placeholder="Escolha as colunas da tabela...",
                                                className="lovable-dropdown",
                                            ),
                                            html.Div(style={"height": "16px"}),
                                            html.Div(
                                                className="action-row",
                                                children=[
                                                    html.Button("Selecionar todas",
                                                        id="carteira-ca-btn-select-all", n_clicks=0,
                                                        className="btn btn-outline"),
                                                    html.Button("Limpar seleção",
                                                        id="carteira-ca-btn-clear-columns", n_clicks=0,
                                                        className="btn btn-outline"),
                                                    html.Button("Exportar Excel",
                                                        id="carteira-ca-btn-export", n_clicks=0,
                                                        className="btn btn-primary"),
                                                ],
                                            ),
                                            dcc.Download(id="carteira-ca-download-excel"),
                                            html.Div(
                                                className="table-shell",
                                                children=[
                                                    dash_table.DataTable(
                                                        id="preview-table-ca",
                                                        page_size=15,
                                                        style_table={"overflowX": "auto"},
                                                        style_cell={
                                                            "textAlign": "left", "padding": "10px 14px",
                                                            "minWidth": "120px", "maxWidth": "320px",
                                                            "whiteSpace": "nowrap", "overflow": "hidden",
                                                            "textOverflow": "ellipsis",
                                                            "fontSize": "12px", "border": "none",
                                                        },
                                                        style_header={
                                                            "fontWeight": "700", "fontSize": "11px",
                                                            "textTransform": "uppercase",
                                                            "letterSpacing": "0.04em",
                                                            "backgroundColor": "#F8FAFC",
                                                            "border": "none", "color": "#64748B",
                                                        },
                                                        style_data={"backgroundColor": "#ffffff", "border": "none"},
                                                        style_data_conditional=[
                                                            {"if": {"row_index": "odd"}, "backgroundColor": "#FAFCFF"}
                                                        ],
                                                        css=[
                                                            {"selector": ".dash-spreadsheet-container table",
                                                             "rule": "border-collapse: separate; border-spacing: 0; width: 100%;"},
                                                            {"selector": ".dash-spreadsheet-container tr:hover td",
                                                             "rule": "background-color: #F1F5F9; transition: background-color 0.15s;"},
                                                        ],
                                                        data=[], columns=[],
                                                    )
                                                ],
                                            ),
                                        ]),
                                    ],
                                )
                            ],
                        ),
                    ],
                ),
            ],
        ),

        # ── Estado vazio ──
        html.Div(
            id="preview-section-ca-empty",
            style={"display": "block" if not has_data else "none"},
            children=[
                _empty_state(
                    "Base não carregada",
                    "Clique em 'Carregar página' para exibir a Carteira Ativa das coordenações gerais.",
                )
            ],
        ),
    ])

    return html.Div(className="page-wrap fade-in", children=children)
