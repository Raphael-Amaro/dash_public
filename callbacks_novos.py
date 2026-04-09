"""
Callbacks para as abas Acompanhamento e Validade da Resolução
do módulo carteira_analistas.

Cole esses callbacks no arquivo principal (run_msft_public.py ou equivalente),
junto com os outros callbacks existentes.

Dependências além das já importadas:
    from datetime import date
    import plotly.graph_objects as go
    from io import BytesIO
"""

# ─────────────────────────────────────────────────────────────────────────────
# IMPORTS NECESSÁRIOS (adicionar ao topo do arquivo principal se não existirem)
# ─────────────────────────────────────────────────────────────────────────────
# from datetime import date
# import plotly.graph_objects as go
# from io import BytesIO

# ─────────────────────────────────────────────────────────────────────────────
# PALETA (já deve existir no arquivo principal)
# ─────────────────────────────────────────────────────────────────────────────
# ACCENT = "#C9A84C"
# BLUE   = "#2563EB"
# TEAL   = "#0D9488"
# ROSE   = "#E11D48"
# VIOLET = "#7C3AED"

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS (copiar para o arquivo principal se não existirem)
# ─────────────────────────────────────────────────────────────────────────────

"""
def fmt_bi(v):
    if v >= 1e9:  return f"US$ {v/1e9:.1f}bi"
    if v >= 1e6:  return f"US$ {v/1e6:.0f}M"
    return f"US$ {v:,.0f}"

def kpi_block(label, value, sub):
    return [
        html.Div(label, className="metric-label"),
        html.Div(value, className="metric-value"),
        html.Div(sub,   className="metric-subtitle"),
    ]

XAXIS_DEF = dict(showgrid=False, zeroline=False, showline=False,
                 tickfont=dict(size=11, color="#64748B"))
YAXIS_DEF = dict(showgrid=True, gridcolor="#F1F5F9", zeroline=False,
                 showline=False, tickfont=dict(size=11, color="#64748B"))

PLOTLY_BASE = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Inter, system-ui, sans-serif", size=12, color="#334155"),
    margin=dict(t=30, r=20, b=50, l=20),
    hoverlabel=dict(bgcolor="#0F172A", bordercolor="#1E293B",
                    font=dict(color="#F8FAFC", size=12)),
)

FASE_COLORS = {
    "Em execução": "#0D9488", "Repagamento": "#2563EB",
    "Aprovado": "#C9A84C", "Aprovada COFIEX": "#C9A84C",
    "Aguardando Assinatura": "#F59E0B", "Em preparação": "#7C3AED",
    "Finalizada": "#94A3B8", "Arquivado": "#CBD5E1",
    "Reprovada": "#E11D48", "Devolvido": "#F97316",
    "Pautado na COFIEX": "#06B6D4", "Em negociação": "#6366F1",
}
"""

# ═════════════════════════════════════════════════════════════════════════════
# CALLBACKS — ABA ACOMPANHAMENTO
# ═════════════════════════════════════════════════════════════════════════════

ACOMPANHAMENTO_CALLBACKS = '''
# ── Callback: KPIs do Acompanhamento ─────────────────────────────────────────
@callback(
    Output("ca-acomp-kpi-tecnicos", "children"),
    Output("ca-acomp-kpi-ops",      "children"),
    Output("ca-acomp-kpi-valor",    "children"),
    Output("ca-acomp-kpi-cgs",      "children"),
    Input("carteira-ca-tabs", "value"),
    State("global-df-json-ca",      "data"),   # ajustar nome do Store conforme o projeto
    *(State(f"carteira-ca-select-{col}", "value") for col, _ in SELECT_FIELDS_CA),
)
def update_acomp_kpis(tab, df_json_ca, *filter_values):
    if tab != "acompanhamento" or not df_json_ca:
        raise PreventUpdate

    df = _apply_ca_filters(df_json_ca, filter_values)
    if df.empty:
        empty = kpi_block("—", "—", "sem dados")
        return empty, empty, empty, empty

    n_tecnicos = df["nm_tecnico"].nunique() if "nm_tecnico" in df.columns else 0
    n_ops      = len(df)
    val_fin    = df["vl_financiamento_dolar"].sum() if "vl_financiamento_dolar" in df.columns else 0
    n_cgs      = df["nm_cg"].nunique() if "nm_cg" in df.columns else 0

    return (
        kpi_block("Técnicos",            str(n_tecnicos),               "com operações na carteira"),
        kpi_block("Total de Operações",  f"{n_ops:,}".replace(",", "."), "no filtro atual"),
        kpi_block("Financiamento Total", fmt_bi(val_fin),               "soma do portfólio"),
        kpi_block("Coordenações Gerais", str(n_cgs),                    "CGs representadas"),
    )


# ── Callback: Gráfico Operações por Técnico ───────────────────────────────────
@callback(
    Output("ca-acomp-fig-tecnico", "figure"),
    Input("carteira-ca-tabs", "value"),
    State("global-df-json-ca",     "data"),
    *(State(f"carteira-ca-select-{col}", "value") for col, _ in SELECT_FIELDS_CA),
)
def update_acomp_fig_tecnico(tab, df_json_ca, *filter_values):
    if tab != "acompanhamento" or not df_json_ca:
        raise PreventUpdate

    df = _apply_ca_filters(df_json_ca, filter_values)
    if df.empty or "nm_tecnico" not in df.columns:
        return EMPTY_FIG

    col_val = "vl_financiamento_dolar"
    grp = (df.groupby("nm_tecnico")[col_val].sum()
             .reset_index().nlargest(15, col_val).sort_values(col_val))
    grp.columns = ["tecnico", "val"]

    n = len(grp)
    clrs = [f"rgba(37,99,235,{0.35 + 0.6*(i/max(n-1,1))})" for i in range(n)]

    fig = go.Figure(go.Bar(
        x=grp["val"] / 1e6, y=grp["tecnico"], orientation="h",
        marker_color=clrs, marker_line_width=0,
        hovertemplate="<b>%{y}</b><br>US$ %{x:,.0f}M<extra></extra>",
    ))
    layout = {**PLOTLY_BASE, "margin": dict(t=20, r=70, b=40, l=180)}
    layout["xaxis"] = dict(**XAXIS_DEF, showgrid=True, gridcolor="#F1F5F9",
                           tickformat=",.0f")
    layout["yaxis"] = dict(**YAXIS_DEF, showgrid=False, automargin=True)
    fig.update_layout(**layout)
    return fig


# ── Callback: Gráfico por CG ──────────────────────────────────────────────────
@callback(
    Output("ca-acomp-fig-cg", "figure"),
    Input("carteira-ca-tabs", "value"),
    State("global-df-json-ca",    "data"),
    *(State(f"carteira-ca-select-{col}", "value") for col, _ in SELECT_FIELDS_CA),
)
def update_acomp_fig_cg(tab, df_json_ca, *filter_values):
    if tab != "acompanhamento" or not df_json_ca:
        raise PreventUpdate

    df = _apply_ca_filters(df_json_ca, filter_values)
    if df.empty or "nm_cg" not in df.columns:
        return EMPTY_FIG

    grp = df.groupby("nm_cg")["vl_financiamento_dolar"].sum().reset_index()
    grp.columns = ["cg", "val"]
    grp = grp.sort_values("val", ascending=False)

    fig = go.Figure(go.Pie(
        labels=grp["cg"], values=grp["val"] / 1e6, hole=0.55,
        textinfo="label+percent", textfont=dict(size=11),
        marker=dict(line=dict(color="white", width=2)),
        hovertemplate="<b>%{label}</b><br>US$ %{value:,.0f}M<br>%{percent}<extra></extra>",
    ))
    fig.update_layout(
        **PLOTLY_BASE,
        xaxis=XAXIS_DEF, yaxis=YAXIS_DEF,
        legend=dict(orientation="v", x=1.02, y=0.5, font=dict(size=11)),
        margin=dict(t=20, r=130, b=20, l=20),
    )
    return fig


# ── Callback: Fase por Técnico (stacked) ─────────────────────────────────────
@callback(
    Output("ca-acomp-fig-fase-tecnico", "figure"),
    Input("carteira-ca-tabs", "value"),
    State("global-df-json-ca",          "data"),
    *(State(f"carteira-ca-select-{col}", "value") for col, _ in SELECT_FIELDS_CA),
)
def update_acomp_fig_fase_tecnico(tab, df_json_ca, *filter_values):
    if tab != "acompanhamento" or not df_json_ca:
        raise PreventUpdate

    df = _apply_ca_filters(df_json_ca, filter_values)
    if df.empty or "nm_tecnico" not in df.columns or "de_fase" not in df.columns:
        return EMPTY_FIG

    top10 = (df.groupby("nm_tecnico").size()
               .nlargest(10).index.tolist())
    df_top = df[df["nm_tecnico"].isin(top10)]

    pivot = (df_top.groupby(["nm_tecnico", "de_fase"])
                   .size().reset_index(name="qtd"))

    fases = pivot["de_fase"].unique().tolist()
    fig = go.Figure()
    for fase in fases:
        sub = pivot[pivot["de_fase"] == fase]
        fig.add_trace(go.Bar(
            name=fase,
            x=sub["nm_tecnico"],
            y=sub["qtd"],
            marker_color=FASE_COLORS.get(fase, "#94A3B8"),
            marker_line_width=0,
        ))

    fig.update_layout(
        **PLOTLY_BASE,
        barmode="stack",
        xaxis=dict(**XAXIS_DEF, tickangle=-30),
        yaxis=dict(**YAXIS_DEF),
        legend=dict(orientation="h", yanchor="bottom", y=1.02,
                    xanchor="right", x=1, font=dict(size=10)),
        margin=dict(t=30, r=20, b=90, l=50),
    )
    return fig


# ── Callback: Setor por CG ────────────────────────────────────────────────────
@callback(
    Output("ca-acomp-fig-setor-cg", "figure"),
    Input("carteira-ca-tabs", "value"),
    State("global-df-json-ca",        "data"),
    *(State(f"carteira-ca-select-{col}", "value") for col, _ in SELECT_FIELDS_CA),
)
def update_acomp_fig_setor_cg(tab, df_json_ca, *filter_values):
    if tab != "acompanhamento" or not df_json_ca:
        raise PreventUpdate

    df = _apply_ca_filters(df_json_ca, filter_values)
    if df.empty or "nm_cg" not in df.columns or "nm_setor" not in df.columns:
        return EMPTY_FIG

    grp = (df.groupby(["nm_cg", "nm_setor"])
             .size().reset_index(name="qtd"))
    setores = grp["nm_setor"].unique().tolist()

    fig = go.Figure()
    for setor in setores:
        sub = grp[grp["nm_setor"] == setor]
        fig.add_trace(go.Bar(
            name=setor, x=sub["nm_cg"], y=sub["qtd"],
            marker_line_width=0,
        ))
    fig.update_layout(
        **PLOTLY_BASE,
        barmode="stack",
        xaxis=dict(**XAXIS_DEF, tickangle=-20),
        yaxis=dict(**YAXIS_DEF),
        legend=dict(orientation="h", yanchor="bottom", y=1.02,
                    xanchor="right", x=1, font=dict(size=10)),
        margin=dict(t=30, r=20, b=70, l=50),
    )
    return fig
'''

# ═════════════════════════════════════════════════════════════════════════════
# CALLBACKS — ABA VALIDADE DA RESOLUÇÃO
# ═════════════════════════════════════════════════════════════════════════════

RESOLUCAO_CALLBACKS = '''
# ── Helper: calcular dias para vencer ────────────────────────────────────────
def _prep_resolucao_df(df_json_ca, filter_values, horizonte):
    """
    Retorna df filtrado com coluna adicional dias_para_vencer.
    horizonte: 'vencidas' | '30' | '90' | '180' | 'todas'
    """
    df = _apply_ca_filters(df_json_ca, filter_values)

    if df.empty or "dt_validade_recomendacao" not in df.columns:
        return pd.DataFrame()

    df = df.copy()
    df["dt_validade_recomendacao"] = pd.to_datetime(
        df["dt_validade_recomendacao"], errors="coerce"
    )
    df = df[df["dt_validade_recomendacao"].notna()]

    hoje = pd.Timestamp(date.today())
    df["dias_para_vencer"] = (df["dt_validade_recomendacao"] - hoje).dt.days

    if horizonte == "vencidas":
        df = df[df["dias_para_vencer"] < 0]
    elif horizonte != "todas":
        dias = int(horizonte)
        df = df[df["dias_para_vencer"] <= dias]

    return df.sort_values("dias_para_vencer")


# ── Callback: KPIs da Resolução ───────────────────────────────────────────────
@callback(
    Output("ca-res-kpi-vencidas", "children"),
    Output("ca-res-kpi-30dias",   "children"),
    Output("ca-res-kpi-90dias",   "children"),
    Output("ca-res-kpi-ok",       "children"),
    Input("carteira-ca-tabs",   "value"),
    Input("ca-res-horizonte",   "value"),
    State("global-df-json-ca",  "data"),
    *(State(f"carteira-ca-select-{col}", "value") for col, _ in SELECT_FIELDS_CA),
)
def update_res_kpis(tab, horizonte, df_json_ca, *filter_values):
    if tab != "resolucao" or not df_json_ca:
        raise PreventUpdate

    df = _apply_ca_filters(df_json_ca, filter_values)
    if df.empty or "dt_validade_recomendacao" not in df.columns:
        empty = kpi_block("—", "—", "sem dados")
        return empty, empty, empty, empty

    df = df.copy()
    df["dt_validade_recomendacao"] = pd.to_datetime(
        df["dt_validade_recomendacao"], errors="coerce"
    )
    hoje = pd.Timestamp(date.today())
    df["dias"] = (df["dt_validade_recomendacao"] - hoje).dt.days

    df_valid = df[df["dt_validade_recomendacao"].notna()]

    n_vencidas = (df_valid["dias"] < 0).sum()
    n_30       = ((df_valid["dias"] >= 0) & (df_valid["dias"] < 30)).sum()
    n_90       = ((df_valid["dias"] >= 0) & (df_valid["dias"] < 90)).sum()
    n_ok       = (df_valid["dias"] >= 90).sum()

    return (
        kpi_block("Vencidas",         str(int(n_vencidas)), "resolução expirada"),
        kpi_block("Vencem em 30 dias", str(int(n_30)),      "atenção imediata"),
        kpi_block("Vencem em 90 dias", str(int(n_90)),      "monitorar"),
        kpi_block("Dentro do prazo",   str(int(n_ok)),      "validade > 90 dias"),
    )


# ── Callback: Timeline de validades ──────────────────────────────────────────
@callback(
    Output("ca-res-fig-timeline", "figure"),
    Input("carteira-ca-tabs",    "value"),
    Input("ca-res-horizonte",    "value"),
    State("global-df-json-ca",   "data"),
    *(State(f"carteira-ca-select-{col}", "value") for col, _ in SELECT_FIELDS_CA),
)
def update_res_timeline(tab, horizonte, df_json_ca, *filter_values):
    if tab != "resolucao" or not df_json_ca:
        raise PreventUpdate

    df = _prep_resolucao_df(df_json_ca, filter_values, horizonte)
    if df.empty:
        return EMPTY_FIG

    df["mes_ano"] = df["dt_validade_recomendacao"].dt.to_period("M").astype(str)
    grp = df.groupby("mes_ano").size().reset_index(name="qtd")
    grp = grp.sort_values("mes_ano")

    hoje_str = pd.Timestamp(date.today()).to_period("M").strftime("%Y-%m")
    cores = ["#E11D48" if m <= hoje_str else "#C9A84C" for m in grp["mes_ano"]]

    fig = go.Figure(go.Bar(
        x=grp["mes_ano"], y=grp["qtd"],
        marker_color=cores, marker_line_width=0,
        hovertemplate="<b>%{x}</b><br>%{y} operações<extra></extra>",
    ))
    fig.update_layout(
        **PLOTLY_BASE,
        xaxis=dict(**XAXIS_DEF, tickangle=-30),
        yaxis=dict(**YAXIS_DEF, dtick=1),
        margin=dict(t=20, r=20, b=70, l=50),
    )
    return fig


# ── Callback: Tabela de alertas ───────────────────────────────────────────────
@callback(
    Output("ca-res-table",          "data"),
    Output("ca-res-table",          "columns"),
    Output("ca-res-selected-count", "children"),
    Input("carteira-ca-tabs",    "value"),
    Input("ca-res-horizonte",    "value"),
    Input("ca-res-column-selector", "value"),
    State("global-df-json-ca",   "data"),
    *(State(f"carteira-ca-select-{col}", "value") for col, _ in SELECT_FIELDS_CA),
)
def update_res_table(tab, horizonte, selected_cols, df_json_ca, *filter_values):
    if tab != "resolucao" or not df_json_ca:
        raise PreventUpdate

    df = _prep_resolucao_df(df_json_ca, filter_values, horizonte)
    if df.empty:
        return [], [], "0 colunas"

    # Garante que dias_para_vencer aparece
    if "dias_para_vencer" not in (selected_cols or []):
        selected_cols = (selected_cols or []) + ["dias_para_vencer"]

    valid_cols = [c for c in (selected_cols or []) if c in df.columns]
    df_out = df[valid_cols].copy()

    # Formata datas para exibição
    for col in df_out.select_dtypes(include=["datetime64[ns]", "datetimetz"]).columns:
        df_out[col] = df_out[col].dt.strftime("%d/%m/%Y")

    df_out = df_out.where(pd.notnull(df_out), None)

    columns = [{"name": c, "id": c} for c in df_out.columns]
    n = len(valid_cols)
    total = len(df.columns)

    return df_out.to_dict("records"), columns, f"{n} de {total} colunas"


# ── Callback: Exportar alerta de resolução ────────────────────────────────────
@callback(
    Output("ca-res-download-excel", "data"),
    Input("ca-res-btn-export",   "n_clicks"),
    State("ca-res-horizonte",    "value"),
    State("ca-res-column-selector", "value"),
    State("global-df-json-ca",   "data"),
    *(State(f"carteira-ca-select-{col}", "value") for col, _ in SELECT_FIELDS_CA),
    prevent_initial_call=True,
)
def export_res_excel(n_clicks, horizonte, selected_cols, df_json_ca, *filter_values):
    if not n_clicks or not df_json_ca:
        raise PreventUpdate

    df = _prep_resolucao_df(df_json_ca, filter_values, horizonte)
    if df.empty:
        raise PreventUpdate

    if "dias_para_vencer" not in (selected_cols or []):
        selected_cols = (selected_cols or []) + ["dias_para_vencer"]

    valid_cols = [c for c in (selected_cols or []) if c in df.columns]
    df_out = df[valid_cols].copy()

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df_out.to_excel(w, index=False, sheet_name="validade_resolucao")
    buf.seek(0)

    return dcc.send_bytes(buf.getvalue(), "validade_resolucao.xlsx")
'''

if __name__ == "__main__":
    print("=== Callbacks — Acompanhamento ===")
    print(ACOMPANHAMENTO_CALLBACKS)
    print("\n=== Callbacks — Validade da Resolução ===")
    print(RESOLUCAO_CALLBACKS)
