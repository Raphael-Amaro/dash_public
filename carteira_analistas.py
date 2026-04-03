from dash import dash_table, dcc, html


def metric_card(label: str, value: str, sub: str, color: str = "#2563EB") -> html.Div:
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


def _selection_dropdown(label: str, component_id: str) -> html.Div:
    return html.Div(
        className="selection-item",
        children=[
            html.Label(label, className="filter-label"),
            dcc.Dropdown(
                id=component_id,
                options=[],
                value=[],
                multi=True,
                placeholder=f"Escolha {label.lower()}...",
                className="lovable-dropdown",
            ),
        ],
        style={"minWidth": "220px"},
    )


def carteira_analistas_page_layout(auth_component=None) -> html.Div:
    children = []

    if auth_component is not None:
        children.append(auth_component)

    children.extend(
        [
            html.Div(
                className="page-header",
                style={"marginTop": "6px"},
                children=[
                    html.H1("Carteira Ativa - Coordenações Gerais", className="page-title"),
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
            html.Div(
                id="preview-section-ca",
                style={"display": "none"},
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
                                                        "Clique para expandir e filtrar os registros da Carteira Ativa das coordenações gerais.",
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
                                                                "Escolha, nas listas abaixo, os filtros que deseja aplicar às abas e aos KPIs.",
                                                                className="section-subtitle",
                                                            ),
                                                        ]
                                                    ),
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
                                                children=[
                                                    _selection_dropdown("CG", "carteira-ca-select-nm_cg"),
                                                    _selection_dropdown("Técnico", "carteira-ca-select-nm_tecnico"),
                                                    _selection_dropdown("Sigla do Pleito", "carteira-ca-select-sg_pleito"),
                                                ],
                                            ),
                                        ],
                                    ),
                                ],
                            )
                        ],
                    ),
                    dcc.Tabs(
                        id="carteira-ca-tabs",
                        value="dados",
                        className="painel-tabs",
                        children=[
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
                                            glass_card(
                                                *[
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
                                                                style_data={
                                                                    "backgroundColor": "#ffffff",
                                                                    "border": "none",
                                                                },
                                                                style_data_conditional=[
                                                                    {
                                                                        "if": {"row_index": "odd"},
                                                                        "backgroundColor": "#FAFCFF",
                                                                    }
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
                            dcc.Tab(
                                label="Acompanhamento",
                                value="acompanhamento",
                                className="painel-tab",
                                selected_className="painel-tab--active",
                                children=[
                                    html.Div(
                                        className="tab-content",
                                        children=[
                                            glass_card(
                                                *[
                                                    section_head(
                                                        "Acompanhamento",
                                                        "Área reservada para evoluções futuras da Carteira Ativa das coordenações gerais.",
                                                    ),
                                                    html.Div(
                                                        "Conteúdo de acompanhamento será exibido aqui.",
                                                        className="section-subtitle",
                                                        style={"padding": "6px 0 2px 0"},
                                                    ),
                                                ]
                                            ),
                                        ],
                                    )
                                ],
                            ),
                            dcc.Tab(
                                label="Tabela",
                                value="tabela",
                                className="painel-tab",
                                selected_className="painel-tab--active",
                                children=[
                                    html.Div(
                                        className="tab-content",
                                        children=[
                                            glass_card(
                                                *[
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
                                                                id="carteira-ca-selected-count",
                                                                className="section-counter",
                                                            ),
                                                        ],
                                                    ),
                                                    dcc.Dropdown(
                                                        id="carteira-ca-column-selector",
                                                        options=[],
                                                        value=[],
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
                                                                id="carteira-ca-btn-select-all",
                                                                n_clicks=0,
                                                                className="btn btn-outline",
                                                            ),
                                                            html.Button(
                                                                "Limpar seleção",
                                                                id="carteira-ca-btn-clear-columns",
                                                                n_clicks=0,
                                                                className="btn btn-outline",
                                                            ),
                                                            html.Button(
                                                                "Exportar Excel",
                                                                id="carteira-ca-btn-export",
                                                                n_clicks=0,
                                                                className="btn btn-primary",
                                                            ),
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
                                                                style_data={
                                                                    "backgroundColor": "#ffffff",
                                                                    "border": "none",
                                                                },
                                                                style_data_conditional=[
                                                                    {
                                                                        "if": {"row_index": "odd"},
                                                                        "backgroundColor": "#FAFCFF",
                                                                    }
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
            ),
        ]
    )

    return html.Div(className="page-wrap fade-in", children=children)