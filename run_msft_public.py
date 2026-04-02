from __future__ import annotations

import json
import locale
import os
import time
from datetime import datetime
from io import BytesIO, StringIO
from pathlib import Path
from typing import Any

import msal
import pandas as pd
import plotly.graph_objects as go
import pygwalker as pyg
import requests
from dash import Dash, Input, Output, State, callback, dash_table, dcc, html
from dash.exceptions import PreventUpdate
from dotenv import load_dotenv
from flask import redirect, request, session
from werkzeug.middleware.proxy_fix import ProxyFix

from painel import painel_page_layout

load_dotenv()

"""
SEAID / COFIEX — Painel de Financiamentos Externos
Versão 2.0 — Painel Analítico + PyGWalker + Gestão de Dados
"""

# ── CONFIGURAÇÃO ──────────────────────────────────────────────────────────────

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

TENANT_ID = os.environ["TENANT_ID"]
CLIENT_ID = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]

SHAREPOINT_HOSTNAME = os.environ["SHAREPOINT_HOSTNAME"]
SHAREPOINT_SITE_PATH = os.environ["SHAREPOINT_SITE_PATH"]
SHAREPOINT_FILE_PATH = os.environ["SHAREPOINT_FILE_PATH"]

FLASK_SECRET_KEY = os.environ["FLASK_SECRET_KEY"]

raw_app_base_url = os.getenv("APP_BASE_URL", "http://localhost:8050").strip()
if raw_app_base_url.startswith("http://") or raw_app_base_url.startswith("https://"):
    APP_BASE_URL = raw_app_base_url.rstrip("/")
else:
    APP_BASE_URL = f"https://{raw_app_base_url}".rstrip("/")

SCOPES = [
    "User.Read",
    "Files.Read",
    "Files.Read.All",
    "Sites.Read.All",
]

ANO_MIN, ANO_MAX = 2000, datetime.now().year

BASE_DIR = Path(__file__).resolve().parent
CACHE_DIR = BASE_DIR / "cache"
CACHE_DIR.mkdir(exist_ok=True)

BR_UF_GEOJSON_PATH = CACHE_DIR / "br_states_geojson.json"

# ── FORMATAÇÃO BRASILEIRA ─────────────────────────────────────────────────────

try:
    locale.setlocale(locale.LC_ALL, "pt_BR.UTF-8")
except locale.Error:
    try:
        locale.setlocale(locale.LC_ALL, "pt_BR.utf8")
    except locale.Error:
        locale.setlocale(locale.LC_ALL, "")


def _br_fallback_number(x, decimais=0):
    try:
        x = float(x)
    except (TypeError, ValueError):
        x = 0.0
    s = f"{x:,.{decimais}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def brazil_vlr(x, y=0):
    try:
        return locale.format_string(f"%.{y}f", float(x), grouping=True)
    except Exception:
        return _br_fallback_number(x, y)


def brazil_per(x, y=0):
    return f"{brazil_vlr(x, y)}%"


def fmt_int_br(x):
    try:
        return brazil_vlr(int(round(float(x))), 0)
    except Exception:
        return "0"


def fmt_mi_br(x, decimais=1):
    return f"{brazil_vlr(x, decimais)}M"


def fmt_usd_br(x, decimais=0):
    return f"US$ {brazil_vlr(x, decimais)}"


# ── PALETA ────────────────────────────────────────────────────────────────────

COLOR_SEQUENCE = [
    "#1D3C69",
    "#FFA300",
    "#FFD000",
    "#183EFF",
    "#66E266",
    "#00D000",
    "#FF0000",
    "#000000",
    "#4E79A7",
    "#A0CBE8",
    "#F28E2B",
    "#FFBE7D",
    "#EDC948",
    "#F6E58D",
    "#59A14F",
    "#8CD17D",
    "#B6992D",
    "#86BCB6",
    "#76B7B2",
    "#AFD5D2",
    "#E15759",
    "#FF9D9A",
    "#B07AA1",
    "#D4A6C8",
    "#9C755F",
    "#D7B5A6",
    "#BAB0AC",
    "#D5D0CC",
]

ACCENT = COLOR_SEQUENCE[0]
BLUE = COLOR_SEQUENCE[3]
TEAL = COLOR_SEQUENCE[18]
ROSE = COLOR_SEQUENCE[20]
VIOLET = COLOR_SEQUENCE[22]
MUTED = COLOR_SEQUENCE[26]

FONTE_COLORS = {
    "BIRD": COLOR_SEQUENCE[0],
    "BID": COLOR_SEQUENCE[1],
    "CAF": COLOR_SEQUENCE[2],
    "NDB": COLOR_SEQUENCE[3],
    "AFD": COLOR_SEQUENCE[4],
    "KfW": COLOR_SEQUENCE[5],
    "FONPLATA": COLOR_SEQUENCE[6],
    "Outras": COLOR_SEQUENCE[8],
}

FASE_COLORS = {
    "Em execução": COLOR_SEQUENCE[0],
    "Repagamento": COLOR_SEQUENCE[1],
    "Aprovado": COLOR_SEQUENCE[2],
    "Aprovada COFIEX": COLOR_SEQUENCE[2],
    "Aguardando Assinatura": COLOR_SEQUENCE[3],
    "Em preparação": COLOR_SEQUENCE[4],
    "Finalizada": COLOR_SEQUENCE[5],
    "Arquivado": COLOR_SEQUENCE[8],
    "Reprovada": COLOR_SEQUENCE[6],
    "Devolvido": COLOR_SEQUENCE[10],
    "Pautado na COFIEX": COLOR_SEQUENCE[9],
    "Em negociação": COLOR_SEQUENCE[18],
    "Em preenchimento": COLOR_SEQUENCE[26],
    "Em análise": COLOR_SEQUENCE[14],
    "Retornado": COLOR_SEQUENCE[20],
}

ESFERA_COLORS = {
    "Federal": COLOR_SEQUENCE[0],
    "Estadual": COLOR_SEQUENCE[1],
    "Municipal": COLOR_SEQUENCE[4],
}

# ── PLOTLY ────────────────────────────────────────────────────────────────────

PLOTLY_BASE = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Inter, system-ui, sans-serif", size=12, color="#334155"),
    hoverlabel=dict(
        bgcolor="#0F172A",
        bordercolor="#1E293B",
        font=dict(color="#F8FAFC", size=12),
    ),
)

DEFAULT_LEGEND = dict(
    orientation="h",
    yanchor="bottom",
    y=1.02,
    xanchor="right",
    x=1,
    font=dict(size=11),
)

XAXIS_DEF = dict(
    showgrid=False,
    zeroline=False,
    showline=False,
    tickfont=dict(size=11, color="#64748B"),
)

YAXIS_DEF = dict(
    showgrid=True,
    gridcolor="#F1F5F9",
    zeroline=False,
    showline=False,
    tickfont=dict(size=11, color="#64748B"),
)


def merge_dict(base: dict, **updates) -> dict:
    out = dict(base)
    out.update(updates)
    return out


def apply_layout(fig, *, xaxis=None, yaxis=None, margin=None, legend=None, **kwargs):
    layout = dict(PLOTLY_BASE)
    layout["separators"] = ",."

    if xaxis is not None:
        layout["xaxis"] = xaxis
    if yaxis is not None:
        layout["yaxis"] = yaxis
    if margin is not None:
        layout["margin"] = margin
    if legend is not None:
        layout["legend"] = legend

    layout.update(kwargs)
    fig.update_layout(**layout)
    return fig


EMPTY_FIG = go.Figure()
apply_layout(
    EMPTY_FIG,
    xaxis=XAXIS_DEF,
    yaxis=YAXIS_DEF,
    margin=dict(t=30, r=20, b=50, l=20),
)
EMPTY_FIG.update_layout(
    annotations=[
        dict(
            text="Carregue a base de dados para visualizar",
            showarrow=False,
            x=0.5,
            y=0.5,
            xref="paper",
            yref="paper",
            font=dict(size=13, color=MUTED),
        )
    ]
)

# ── AUTH / SHAREPOINT / GRAPH HELPERS ────────────────────────────────────────


def build_redirect_uri() -> str:
    return f"{APP_BASE_URL}/auth/callback"


def build_logout_redirect_uri() -> str:
    return APP_BASE_URL


def build_authority() -> str:
    return f"https://login.microsoftonline.com/{TENANT_ID}"


def build_msal_app() -> msal.ConfidentialClientApplication:
    return msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        authority=build_authority(),
        client_credential=CLIENT_SECRET,
    )


def is_authenticated() -> bool:
    access_token = session.get("access_token")
    expires_at = session.get("access_token_expires_at", 0)
    return bool(access_token and time.time() < expires_at)


def get_access_token_from_session() -> str:
    if not is_authenticated():
        raise RuntimeError("Usuário não autenticado ou sessão expirada.")
    return str(session["access_token"])


def clear_auth_session() -> None:
    for key in [
        "access_token",
        "access_token_expires_at",
        "user_name",
        "user_email",
        "id_token_claims",
        "auth_flow",
    ]:
        session.pop(key, None)


class GraphSharePointClient:
    def __init__(self, access_token: str, timeout: int = 60) -> None:
        self.access_token = access_token
        self.timeout = timeout

    def _headers(self) -> dict[str, str]:
        return {
            "Authorization": f"Bearer {self.access_token}",
            "Accept": "application/json",
        }

    def _get_json(self, url: str) -> dict[str, Any]:
        response = requests.get(url, headers=self._headers(), timeout=self.timeout)
        if not response.ok:
            raise RuntimeError(
                f"GET {url} falhou com status {response.status_code}: {response.text}"
            )
        return response.json()

    def _get_bytes(self, url: str) -> bytes:
        response = requests.get(url, headers=self._headers(), timeout=self.timeout)
        if not response.ok:
            raise RuntimeError(
                f"GET binário {url} falhou com status {response.status_code}: {response.text}"
            )
        return response.content

    def resolve_site(self, hostname: str, site_path: str) -> dict[str, Any]:
        url = f"{GRAPH_BASE_URL}/sites/{hostname}:{site_path}"
        return self._get_json(url)

    def list_drives(self, site_id: str) -> dict[str, Any]:
        url = f"{GRAPH_BASE_URL}/sites/{site_id}/drives"
        return self._get_json(url)

    def get_default_drive_id(self, site_id: str) -> str:
        drives = self.list_drives(site_id)
        values = drives.get("value", [])
        if not values:
            raise RuntimeError("Nenhum drive encontrado no site informado.")

        for drive in values:
            if str(drive.get("name", "")).lower() == "documents":
                return str(drive["id"])

        return str(values[0]["id"])

    def download_file_from_drive_path(self, drive_id: str, file_path: str) -> bytes:
        normalized_path = file_path if file_path.startswith("/") else f"/{file_path}"
        url = f"{GRAPH_BASE_URL}/drives/{drive_id}/root:{normalized_path}:/content"
        return self._get_bytes(url)

    def download_excel_from_site(
        self,
        hostname: str,
        site_path: str,
        file_path: str,
    ) -> bytes:
        site = self.resolve_site(hostname=hostname, site_path=site_path)
        site_id = str(site["id"])
        drive_id = self.get_default_drive_id(site_id=site_id)
        return self.download_file_from_drive_path(drive_id=drive_id, file_path=file_path)


# ── TRATAMENTO DE DATAS ───────────────────────────────────────────────────────


def date_date(df: pd.DataFrame) -> pd.DataFrame:
    colunas_com_dt = [col for col in df.columns if "dt" in str(col).lower()]
    for col in colunas_com_dt:
        df[col] = pd.to_datetime(df[col], errors="coerce")
    return df


def year_str(df: pd.DataFrame) -> pd.DataFrame:
    colunas_com_dt = [col for col in df.columns if "dt" in str(col).lower()]
    for col in colunas_com_dt:
        anos = pd.to_datetime(df[col], errors="coerce").dt.year
        df["Ano_" + str(col)] = anos.fillna(0).astype("int64")
    return df


def load_dataframe_from_sharepoint() -> pd.DataFrame:
    access_token = get_access_token_from_session()
    client = GraphSharePointClient(access_token=access_token)

    raw = client.download_excel_from_site(
        hostname=SHAREPOINT_HOSTNAME,
        site_path=SHAREPOINT_SITE_PATH,
        file_path=SHAREPOINT_FILE_PATH,
    )

    df = pd.read_excel(BytesIO(raw))
    df.columns = [str(c).strip() for c in df.columns]
    df = date_date(df)
    df = year_str(df)
    return df


def filter_df_by_columns(df_json: str | None, selected: list | None) -> pd.DataFrame:
    if not df_json:
        return pd.DataFrame()

    df = pd.read_json(StringIO(df_json), orient="split")
    if selected:
        valid = [c for c in selected if c in df.columns]
        df = df[valid] if valid else df.iloc[:, 0:0]
    return df


def get_full_df(df_json: str | None) -> pd.DataFrame:
    if not df_json:
        return pd.DataFrame()
    return pd.read_json(StringIO(df_json), orient="split")


# ── PROCESSAMENTO BASE ────────────────────────────────────────────────────────


def escolher_coluna_ano(df: pd.DataFrame) -> str | None:
    candidatas = [
        "Ano_dt_primeira_cofiex",
        "Ano_dt_recomendacao_cofiex",
        "Ano_dt_cofiex",
        "ano_cofiex",
    ]

    for col in candidatas:
        if col in df.columns:
            vals = pd.to_numeric(df[col], errors="coerce")
            if vals.notna().any() and (vals > 0).any():
                return col

    colunas_ano = [c for c in df.columns if str(c).startswith("Ano_dt")]
    for col in colunas_ano:
        vals = pd.to_numeric(df[col], errors="coerce")
        if vals.notna().any() and (vals > 0).any():
            return col

    return None


def prep_painel_df(df_json: str | None, ano_range: list) -> pd.DataFrame | None:
    if not df_json:
        return None

    df = get_full_df(df_json).copy()

    if df.empty:
        return df

    df = date_date(df)
    df = year_str(df)

    for col in ["vl_financiamento_dolar", "vl_contrapartida_dolar"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(float)

    text_defaults = {
        "nm_setor": "Não informado",
        "sg_fonte_resumo": "Não informado",
        "de_fase": "Não informado",
        "de_esfera": "Não informado",
        "nm_regiao": "Não informado",
        "sg_uf": "Não informado",
    }
    for col, default in text_defaults.items():
        if col in df.columns:
            df[col] = df[col].astype("string")
            df[col] = df[col].fillna(default).replace(["<NA>", "nan", "None", ""], default)

    col_ano = escolher_coluna_ano(df)

    if col_ano:
        df["ano_cofiex"] = pd.to_numeric(df[col_ano], errors="coerce").fillna(0).astype(int)
        df = df[df["ano_cofiex"] > 0]

        if ano_range:
            lo, hi = ano_range[0], ano_range[1]
            df = df[df["ano_cofiex"].between(lo, hi)]

    return df


def _normalize_text_col(df: pd.DataFrame, col: str, default: str = "Não informado") -> pd.DataFrame:
    if col in df.columns:
        df[col] = df[col].astype("string")
        df[col] = df[col].fillna(default).replace(["<NA>", "nan", "None", ""], default)
    return df


def prep_carteira_df(df_json: str | None, ano_range: list | None) -> pd.DataFrame | None:
    df = prep_painel_df(df_json, ano_range)
    if df is None or df.empty:
        return df

    cols_texto = [
        "de_fase",
        "de_tipo_operacao",
        "nm_proponente",
        "sg_fonte",
        "de_esfera",
        "nm_regiao",
        "nm_setor",
        "nm_subsetor",
        "sg_setor",
        "sys",
        "nm_limite",
        "sg_uf",
        "sg_fonte_resumo",
    ]

    for col in cols_texto:
        df = _normalize_text_col(df, col)

    return df


def apply_dropdown_filters(
    df: pd.DataFrame,
    de_fase=None,
    de_tipo_operacao=None,
    nm_proponente=None,
    sg_fonte=None,
    de_esfera=None,
    nm_regiao=None,
    nm_setor=None,
    nm_subsetor=None,
    sg_setor=None,
    sys=None,
    nm_limite=None,
) -> pd.DataFrame:
    if df is None or df.empty:
        return df

    filtros = {
        "de_fase": de_fase,
        "de_tipo_operacao": de_tipo_operacao,
        "nm_proponente": nm_proponente,
        "sg_fonte": sg_fonte,
        "de_esfera": de_esfera,
        "nm_regiao": nm_regiao,
        "nm_setor": nm_setor,
        "nm_subsetor": nm_subsetor,
        "sg_setor": sg_setor,
        "sys": sys,
        "nm_limite": nm_limite,
    }

    out = df.copy()

    for col, valores in filtros.items():
        if col in out.columns and valores:
            out = out[out[col].isin(valores)]

    return out


def contar_projetos_distintos(df: pd.DataFrame) -> int:
    for col in ["cd_pleito"]:
        if col in df.columns:
            return df[col].dropna().nunique()
    return len(df)


def fmt_bi(v: float) -> str:
    if v >= 1e9:
        return f"US$ {brazil_vlr(v / 1e9, 1)}bi"
    if v >= 1e6:
        return f"US$ {brazil_vlr(v / 1e6, 0)}M"
    return f"US$ {brazil_vlr(v, 0)}"


# ── CONSTRUTORES DE GRÁFICOS ──────────────────────────────────────────────────


def _gradient(base_rgb: tuple, n: int, lo=0.35, hi=0.95) -> list[str]:
    return [COLOR_SEQUENCE[i % len(COLOR_SEQUENCE)] for i in range(n)]


def chart_temporal(df: pd.DataFrame, metrica: str) -> go.Figure:
    if "ano_cofiex" not in df.columns or df.empty:
        return EMPTY_FIG

    grp = (
        df.groupby("ano_cofiex")
        .agg(
            valor=("vl_financiamento_dolar", "sum"),
            qtd=("ano_cofiex", "size"),
            proj=("cd_pleito", "nunique"),
        )
        .reset_index()
        .sort_values("ano_cofiex")
    )

    if grp.empty:
        return EMPTY_FIG

    grp["valor_mi"] = grp["valor"] / 1e6
    grp["valor_fmt"] = grp["valor_mi"].apply(lambda x: f"{brazil_vlr(x, 0)}M")
    grp["qtd_fmt"] = grp["qtd"].apply(fmt_int_br)
    grp["proj_fmt"] = grp["proj"].apply(fmt_int_br)

    y_vals = grp["valor_mi"] if metrica == "valor" else grp["qtd"]
    y_label = "Financiamento (US$ milhões)" if metrica == "valor" else "Operações"

    customdata = list(zip(grp["valor_fmt"], grp["qtd_fmt"], grp["proj_fmt"]))

    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            x=grp["ano_cofiex"],
            y=y_vals,
            name=y_label,
            marker_color=COLOR_SEQUENCE[0],
            marker_line_width=0,
            opacity=0.88,
            customdata=customdata,
            hovertemplate=(
                "<b>Ano: %{x}</b><br>"
                "Valor de financiamento: US$ %{customdata[0]}<br>"
                "Quantidade de operações: %{customdata[1]}<br>"
                "Quantidade de projetos: %{customdata[2]}"
                "<extra></extra>"
            ),
        )
    )

    if metrica == "valor":
        fig.add_trace(
            go.Scatter(
                x=grp["ano_cofiex"],
                y=grp["qtd"],
                name="Nº Operações",
                yaxis="y2",
                mode="lines+markers",
                line=dict(color=COLOR_SEQUENCE[1], width=2),
                marker=dict(size=4, color=COLOR_SEQUENCE[1]),
                customdata=customdata,
                hovertemplate=(
                    "<b>Ano: %{x}</b><br>"
                    "Valor de financiamento: US$ %{customdata[0]}<br>"
                    "Quantidade de operações: %{customdata[1]}<br>"
                    "Quantidade de projetos: %{customdata[2]}"
                    "<extra></extra>"
                ),
            )
        )

        apply_layout(
            fig,
            xaxis=merge_dict(XAXIS_DEF, dtick=1, tickangle=0),
            yaxis=merge_dict(YAXIS_DEF, tickformat=",.0f", ticksuffix="M"),
            margin=dict(t=30, r=20, b=50, l=20),
            legend=DEFAULT_LEGEND,
            barmode="group",
            yaxis2=dict(
                overlaying="y",
                side="right",
                showgrid=False,
                tickfont=dict(size=11, color=COLOR_SEQUENCE[1]),
                showline=False,
            ),
        )
    else:
        apply_layout(
            fig,
            xaxis=merge_dict(XAXIS_DEF, dtick=1, tickangle=0),
            yaxis=merge_dict(YAXIS_DEF, tickformat=",.0f"),
            margin=dict(t=30, r=20, b=50, l=20),
            legend=DEFAULT_LEGEND,
            barmode="group",
        )

    return fig


def chart_setor(df: pd.DataFrame, metrica: str) -> go.Figure:
    if "nm_setor" not in df.columns or df.empty:
        return EMPTY_FIG

    agg_dict = {
        "val": ("vl_financiamento_dolar", "sum") if "vl_financiamento_dolar" in df.columns else ("nm_setor", "size"),
        "qtd": ("nm_setor", "size"),
    }
    if "cd_pleito" in df.columns:
        agg_dict["proj"] = ("cd_pleito", pd.Series.nunique)
    else:
        agg_dict["proj"] = ("nm_setor", "size")

    grp = (
        df.groupby("nm_setor")
        .agg(**agg_dict)
        .reset_index()
        .rename(columns={"nm_setor": "setor"})
    )

    if grp.empty:
        return EMPTY_FIG

    grp = grp.nlargest(10, "val" if metrica == "valor" else "qtd").copy()
    grp = grp.sort_values("val" if metrica == "valor" else "qtd", ascending=False)

    grp["val_mi"] = grp["val"] / 1e6
    grp["val_fmt"] = grp["val_mi"].apply(lambda x: f"{brazil_vlr(x, 0)}M")
    grp["qtd_fmt"] = grp["qtd"].apply(fmt_int_br)
    grp["proj_fmt"] = grp["proj"].apply(fmt_int_br)

    xval = grp["val_mi"] if metrica == "valor" else grp["qtd"]
    clrs = [COLOR_SEQUENCE[i % len(COLOR_SEQUENCE)] for i in range(len(grp))]

    fig = go.Figure(
        go.Bar(
            x=xval,
            y=grp["setor"],
            orientation="h",
            marker_color=clrs,
            marker_line_width=0,
            customdata=list(zip(grp["val_fmt"], grp["qtd_fmt"], grp["proj_fmt"])),
            hovertemplate=(
                "<b>%{y}</b><br>"
                "Valor de financiamento: US$ %{customdata[0]}<br>"
                "Quantidade de operações: %{customdata[1]}<br>"
                "Quantidade de projetos: %{customdata[2]}"
                "<extra></extra>"
            ),
        )
    )

    apply_layout(
        fig,
        xaxis=merge_dict(
            XAXIS_DEF,
            showgrid=True,
            gridcolor="#F1F5F9",
            tickformat=",.0f",
            ticksuffix="M" if metrica == "valor" else "",
        ),
        yaxis=merge_dict(YAXIS_DEF, showgrid=False, automargin=True, autorange="reversed"),
        margin=dict(t=20, r=80, b=40, l=190),
        legend=DEFAULT_LEGEND,
    )
    return fig


def chart_fonte(df: pd.DataFrame, metrica: str) -> go.Figure:
    if "sg_fonte_resumo" not in df.columns or df.empty:
        return EMPTY_FIG

    agg_dict = {
        "val": ("vl_financiamento_dolar", "sum") if metrica == "valor" else ("sg_fonte_resumo", "size"),
        "valor_fin": ("vl_financiamento_dolar", "sum"),
        "qtd": ("sg_fonte_resumo", "size"),
    }
    if "cd_pleito" in df.columns:
        agg_dict["proj"] = ("cd_pleito", pd.Series.nunique)
    else:
        agg_dict["proj"] = ("sg_fonte_resumo", "size")

    grp = (
        df.groupby("sg_fonte_resumo")
        .agg(**agg_dict)
        .reset_index()
        .rename(columns={"sg_fonte_resumo": "fonte"})
    )

    if grp.empty:
        return EMPTY_FIG

    grp = grp.sort_values("val", ascending=False)
    grp["valor_fmt"] = (grp["valor_fin"] / 1e6).apply(lambda x: f"{brazil_vlr(x, 0)}M")
    grp["qtd_fmt"] = grp["qtd"].apply(fmt_int_br)
    grp["proj_fmt"] = grp["proj"].apply(fmt_int_br)

    clrs = [COLOR_SEQUENCE[i % len(COLOR_SEQUENCE)] for i in range(len(grp))]

    fig = go.Figure(
        go.Pie(
            labels=grp["fonte"],
            values=grp["val"],
            hole=0.62,
            marker=dict(colors=clrs, line=dict(color="white", width=2)),
            textinfo="label+percent",
            textfont=dict(size=11),
            customdata=list(zip(grp["valor_fmt"], grp["qtd_fmt"], grp["proj_fmt"])),
            hovertemplate=(
                "<b>%{label}</b><br>"
                "Valor de financiamento: US$ %{customdata[0]}<br>"
                "Quantidade de operações: %{customdata[1]}<br>"
                "Quantidade de projetos: %{customdata[2]}<br>"
                "%{percent}<extra></extra>"
            ),
        )
    )

    apply_layout(
        fig,
        xaxis=XAXIS_DEF,
        yaxis=YAXIS_DEF,
        margin=dict(t=20, r=130, b=20, l=20),
        legend=dict(orientation="v", x=1.02, y=0.5, font=dict(size=11)),
    )
    return fig


def chart_fase_percentual(df: pd.DataFrame, metrica: str) -> go.Figure:
    if "de_fase" not in df.columns or df.empty:
        return EMPTY_FIG

    agg_dict = {
        "val": ("vl_financiamento_dolar", "sum") if metrica == "valor" else ("de_fase", "size"),
        "valor_fin": ("vl_financiamento_dolar", "sum"),
        "qtd": ("de_fase", "size"),
    }
    if "cd_pleito" in df.columns:
        agg_dict["proj"] = ("cd_pleito", pd.Series.nunique)
    else:
        agg_dict["proj"] = ("de_fase", "size")

    grp = (
        df.groupby("de_fase")
        .agg(**agg_dict)
        .reset_index()
        .rename(columns={"de_fase": "fase"})
    )

    if grp.empty:
        return EMPTY_FIG

    grp = grp.sort_values("val", ascending=False)
    grp["valor_fmt"] = (grp["valor_fin"] / 1e6).apply(lambda x: f"{brazil_vlr(x, 0)}M")
    grp["qtd_fmt"] = grp["qtd"].apply(fmt_int_br)
    grp["proj_fmt"] = grp["proj"].apply(fmt_int_br)

    clrs = [COLOR_SEQUENCE[i % len(COLOR_SEQUENCE)] for i in range(len(grp))]

    fig = go.Figure(
        go.Pie(
            labels=grp["fase"],
            values=grp["val"],
            hole=0.62,
            marker=dict(colors=clrs, line=dict(color="white", width=2)),
            textinfo="label+percent",
            textfont=dict(size=11),
            customdata=list(zip(grp["valor_fmt"], grp["qtd_fmt"], grp["proj_fmt"])),
            hovertemplate=(
                "<b>%{label}</b><br>"
                "Valor de financiamento: US$ %{customdata[0]}<br>"
                "Quantidade de operações: %{customdata[1]}<br>"
                "Quantidade de projetos: %{customdata[2]}<br>"
                "%{percent}<extra></extra>"
            ),
        )
    )

    apply_layout(
        fig,
        xaxis=XAXIS_DEF,
        yaxis=YAXIS_DEF,
        margin=dict(t=20, r=20, b=90, l=20),
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.12,
            xanchor="center",
            x=0.5,
            font=dict(size=11),
        ),
    )
    return fig


def chart_fase(df: pd.DataFrame, metrica: str) -> go.Figure:
    if "de_fase" not in df.columns or df.empty:
        return EMPTY_FIG

    agg_dict = {
        "valor": ("vl_financiamento_dolar", "sum"),
        "qtd": ("de_fase", "size"),
    }
    if "cd_pleito" in df.columns:
        agg_dict["proj"] = ("cd_pleito", pd.Series.nunique)
    else:
        agg_dict["proj"] = ("de_fase", "size")

    grp = (
        df.groupby("de_fase")
        .agg(**agg_dict)
        .reset_index()
        .rename(columns={"de_fase": "fase"})
    )

    if grp.empty:
        return EMPTY_FIG

    grp = grp.sort_values("valor" if metrica == "valor" else "qtd", ascending=False)

    def quebra_linha(txt: str, limite: int = 16) -> str:
        txt = str(txt)
        if len(txt) <= limite:
            return txt

        partes = txt.split()
        linhas = []
        linha_atual = ""

        for parte in partes:
            teste = f"{linha_atual} {parte}".strip()
            if len(teste) <= limite:
                linha_atual = teste
            else:
                if linha_atual:
                    linhas.append(linha_atual)
                linha_atual = parte

        if linha_atual:
            linhas.append(linha_atual)

        return "<br>".join(linhas)

    grp["fase_label"] = grp["fase"].apply(quebra_linha)
    grp["valor_mi_fmt"] = grp["valor"].apply(lambda x: fmt_mi_br(x / 1e6, 1))
    grp["qtd_fmt"] = grp["qtd"].apply(fmt_int_br)
    grp["proj_fmt"] = grp["proj"].apply(fmt_int_br)

    y_bar = grp["valor"] / 1e6 if metrica == "valor" else grp["qtd"]
    clrs = [COLOR_SEQUENCE[i % len(COLOR_SEQUENCE)] for i in range(len(grp))]

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=grp["fase_label"],
            y=y_bar,
            marker_color=clrs,
            marker_line_width=0,
            name="Financiamento (US$ milhões)" if metrica == "valor" else "Operações",
            customdata=list(zip(grp["fase"], grp["qtd_fmt"], grp["valor_mi_fmt"], grp["proj_fmt"])),
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Valor de financiamento: US$ %{customdata[2]}<br>"
                "Quantidade de operações: %{customdata[1]}<br>"
                "Quantidade de projetos: %{customdata[3]}"
                "<extra></extra>"
            ),
        )
    )

    apply_layout(
        fig,
        xaxis=merge_dict(XAXIS_DEF, tickangle=0, automargin=True, ticklabelstandoff=18),
        yaxis=merge_dict(
            YAXIS_DEF,
            tickformat=",.0f",
            ticksuffix="M" if metrica == "valor" else "",
        ),
        margin=dict(t=20, r=20, b=110, l=60),
        legend=DEFAULT_LEGEND,
    )
    return fig


def chart_esfera(df: pd.DataFrame, metrica: str) -> go.Figure:
    if "de_esfera" not in df.columns or df.empty:
        return EMPTY_FIG

    agg_dict = {
        "valor": ("vl_financiamento_dolar", "sum"),
        "qtd": ("de_esfera", "size"),
    }
    if "cd_pleito" in df.columns:
        agg_dict["qtd_projetos"] = ("cd_pleito", pd.Series.nunique)
    else:
        agg_dict["qtd_projetos"] = ("de_esfera", "size")

    grp = (
        df.groupby("de_esfera")
        .agg(**agg_dict)
        .reset_index()
        .rename(columns={"de_esfera": "esfera"})
    )

    if grp.empty:
        return EMPTY_FIG

    grp = grp.sort_values("valor" if metrica == "valor" else "qtd", ascending=False)

    yval = grp["valor"] / 1e6 if metrica == "valor" else grp["qtd"]
    clrs = [COLOR_SEQUENCE[i % len(COLOR_SEQUENCE)] for i in range(len(grp))]

    grp["valor_fmt"] = (grp["valor"] / 1e6).apply(lambda x: f"{brazil_vlr(x, 0)}M")
    grp["qtd_fmt"] = grp["qtd"].apply(fmt_int_br)
    grp["qtd_projetos_fmt"] = grp["qtd_projetos"].apply(fmt_int_br)

    fig = go.Figure(
        go.Bar(
            x=grp["esfera"],
            y=yval,
            marker_color=clrs,
            marker_line_width=0,
            width=0.45,
            customdata=list(zip(grp["esfera"], grp["valor_fmt"], grp["qtd_fmt"], grp["qtd_projetos_fmt"])),
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Valor de financiamento: US$ %{customdata[1]}<br>"
                "Quantidade de operações: %{customdata[2]}<br>"
                "Quantidade de projetos: %{customdata[3]}<br>"
                "<extra></extra>"
            ),
        )
    )

    apply_layout(
        fig,
        xaxis=XAXIS_DEF,
        yaxis=merge_dict(
            YAXIS_DEF,
            tickformat=",.0f",
            ticksuffix="M" if metrica == "valor" else "",
        ),
        margin=dict(t=20, r=20, b=50, l=60),
        legend=DEFAULT_LEGEND,
    )
    return fig


def chart_regiao(df: pd.DataFrame, metrica: str) -> go.Figure:
    if "nm_regiao" not in df.columns or df.empty:
        return EMPTY_FIG

    agg_dict = {
        "valor": ("vl_financiamento_dolar", "sum"),
        "qtd": ("nm_regiao", "size"),
    }
    if "cd_pleito" in df.columns:
        agg_dict["qtd_projetos"] = ("cd_pleito", pd.Series.nunique)
    else:
        agg_dict["qtd_projetos"] = ("nm_regiao", "size")

    grp = (
        df.groupby("nm_regiao")
        .agg(**agg_dict)
        .reset_index()
        .rename(columns={"nm_regiao": "regiao"})
    )

    if grp.empty:
        return EMPTY_FIG

    grp = grp.sort_values("valor" if metrica == "valor" else "qtd", ascending=False)

    yval = grp["valor"] / 1e6 if metrica == "valor" else grp["qtd"]
    clrs = [COLOR_SEQUENCE[i % len(COLOR_SEQUENCE)] for i in range(len(grp))]

    grp["valor_fmt"] = (grp["valor"] / 1e6).apply(lambda x: f"{brazil_vlr(x, 0)}M")
    grp["qtd_fmt"] = grp["qtd"].apply(fmt_int_br)
    grp["qtd_projetos_fmt"] = grp["qtd_projetos"].apply(fmt_int_br)

    fig = go.Figure(
        go.Bar(
            x=grp["regiao"],
            y=yval,
            marker_color=clrs,
            marker_line_width=0,
            customdata=list(zip(grp["regiao"], grp["valor_fmt"], grp["qtd_fmt"], grp["qtd_projetos_fmt"])),
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Valor de financiamento: US$ %{customdata[1]}<br>"
                "Quantidade de operações: %{customdata[2]}<br>"
                "Quantidade de projetos: %{customdata[3]}<br>"
                "<extra></extra>"
            ),
        )
    )

    apply_layout(
        fig,
        xaxis=XAXIS_DEF,
        yaxis=merge_dict(
            YAXIS_DEF,
            tickformat=",.0f",
            ticksuffix="M" if metrica == "valor" else "",
        ),
        margin=dict(t=20, r=20, b=50, l=60),
        legend=DEFAULT_LEGEND,
    )
    return fig


def chart_uf(df: pd.DataFrame, metrica: str) -> go.Figure:
    if "sg_uf" not in df.columns or df.empty:
        return EMPTY_FIG

    df_plot = df.copy()

    if "de_esfera" in df_plot.columns:
        df_plot = df_plot[df_plot["de_esfera"].astype(str).str.strip().str.lower() != "federal"]

    if df_plot.empty:
        return EMPTY_FIG

    agg_dict = {
        "valor": ("vl_financiamento_dolar", "sum"),
        "qtd": ("sg_uf", "size"),
    }
    if "cd_pleito" in df_plot.columns:
        agg_dict["qtd_projetos"] = ("cd_pleito", pd.Series.nunique)
    else:
        agg_dict["qtd_projetos"] = ("sg_uf", "size")

    grp = (
        df_plot.groupby("sg_uf")
        .agg(**agg_dict)
        .reset_index()
        .rename(columns={"sg_uf": "uf"})
    )

    if grp.empty:
        return EMPTY_FIG

    grp = grp.nlargest(15, "valor" if metrica == "valor" else "qtd").sort_values(
        "valor" if metrica == "valor" else "qtd",
        ascending=False,
    )

    xval = grp["valor"] / 1e6 if metrica == "valor" else grp["qtd"]
    clrs = [COLOR_SEQUENCE[i % len(COLOR_SEQUENCE)] for i in range(len(grp))]

    grp["valor_fmt"] = (grp["valor"] / 1e6).apply(lambda x: f"{brazil_vlr(x, 0)}M")
    grp["qtd_fmt"] = grp["qtd"].apply(fmt_int_br)
    grp["qtd_projetos_fmt"] = grp["qtd_projetos"].apply(fmt_int_br)

    fig = go.Figure(
        go.Bar(
            x=xval,
            y=grp["uf"],
            orientation="h",
            marker_color=clrs,
            marker_line_width=0,
            customdata=list(zip(grp["uf"], grp["valor_fmt"], grp["qtd_fmt"], grp["qtd_projetos_fmt"])),
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Valor de financiamento: US$ %{customdata[1]}<br>"
                "Quantidade de operações: %{customdata[2]}<br>"
                "Quantidade de projetos: %{customdata[3]}<br>"
                "<extra></extra>"
            ),
        )
    )

    apply_layout(
        fig,
        xaxis=merge_dict(
            XAXIS_DEF,
            showgrid=True,
            gridcolor="#F1F5F9",
            tickformat=",.0f",
            ticksuffix="M" if metrica == "valor" else "",
        ),
        yaxis=merge_dict(YAXIS_DEF, showgrid=False, autorange="reversed"),
        margin=dict(t=20, r=70, b=40, l=55),
        legend=DEFAULT_LEGEND,
    )
    return fig


def load_br_states_geojson() -> dict:
    if BR_UF_GEOJSON_PATH.exists():
        with open(BR_UF_GEOJSON_PATH, "r", encoding="utf-8") as f:
            return json.load(f)

    import geobr

    br_uf = geobr.read_state(year=2020)
    geojson = json.loads(br_uf.to_json())

    with open(BR_UF_GEOJSON_PATH, "w", encoding="utf-8") as f:
        json.dump(geojson, f, ensure_ascii=False)

    return geojson


def chart_mapa_uf(df: pd.DataFrame, metrica: str) -> go.Figure:
    if df is None or df.empty or "cd_uf" not in df.columns:
        return EMPTY_FIG

    geojson = load_br_states_geojson()
    df_map = df.copy()

    if "de_esfera" in df_map.columns:
        df_map = df_map[df_map["de_esfera"].astype(str).str.strip().str.lower() != "federal"]

    if "de_fase" in df_map.columns:
        df_map = df_map[df_map["de_fase"].astype(str).str.strip().str.lower() != "federal"]

    df_map["cd_uf"] = pd.to_numeric(df_map["cd_uf"], errors="coerce")
    df_map = df_map[df_map["cd_uf"].notna()]

    if df_map.empty:
        return EMPTY_FIG

    agg_dict = {
        "valor_financiamento": ("vl_financiamento_dolar", "sum"),
        "qtd_operacoes": ("cd_uf", "size"),
    }
    if "cd_pleito" in df_map.columns:
        agg_dict["qtd_projetos"] = ("cd_pleito", pd.Series.nunique)
    else:
        agg_dict["qtd_projetos"] = ("cd_uf", "size")

    grp = df_map.groupby("cd_uf", as_index=False).agg(**agg_dict)
    grp["val"] = grp["valor_financiamento"] if metrica == "valor" else grp["qtd_operacoes"]

    zmin = float(grp["val"].min()) if not grp.empty else 0
    zmax = float(grp["val"].max()) if not grp.empty else 0

    if metrica == "valor":
        colorbar = dict(
            title=dict(
                text="US$ milhões",
                side="right",
                font=dict(size=12),
            ),
            thickness=14,
            len=0.75,
            y=0.5,
            tickmode="array",
            tickvals=[zmin, zmax],
            ticktext=[f"{brazil_vlr(zmin / 1e6, 0)}M", f"{brazil_vlr(zmax / 1e6, 0)}M"],
        )
    else:
        colorbar = dict(
            title=dict(
                text="Operações",
                side="right",
                font=dict(size=12),
            ),
            thickness=14,
            len=0.75,
            y=0.5,
            tickmode="array",
            tickvals=[zmin, zmax],
            ticktext=[fmt_int_br(zmin), fmt_int_br(zmax)],
        )

    grp["cd_uf"] = grp["cd_uf"].astype(int)
    grp["cd_uf_str"] = grp["cd_uf"].astype(str)

    nomes_estados = {
        int(feat["properties"]["code_state"]): feat["properties"]["abbrev_state"]
        for feat in geojson["features"]
    }
    grp["uf_sigla"] = grp["cd_uf"].map(nomes_estados).fillna("UF")

    grp["valor_financiamento_fmt"] = grp["valor_financiamento"].apply(
        lambda x: f"US$ {brazil_vlr(x / 1e6, 0)}M"
    )
    grp["qtd_operacoes_fmt"] = grp["qtd_operacoes"].apply(fmt_int_br)
    grp["qtd_projetos_fmt"] = grp["qtd_projetos"].apply(fmt_int_br)

    fig = go.Figure(
        go.Choropleth(
            geojson=geojson,
            locations=grp["cd_uf_str"],
            z=grp["val"],
            zmin=zmin,
            zmax=zmax,
            featureidkey="properties.code_state",
            colorscale=[[0.00, "#FFFFFF"], [1.00, "#1D3C69"]],
            marker_line_color="#94A3B8",
            marker_line_width=0.8,
            colorbar=colorbar,
            customdata=list(
                zip(
                    grp["uf_sigla"],
                    grp["valor_financiamento_fmt"],
                    grp["qtd_operacoes_fmt"],
                    grp["qtd_projetos_fmt"],
                )
            ),
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Valor do financiamento: %{customdata[1]}<br>"
                "Quantidade de operações: %{customdata[2]}<br>"
                "Quantidade de projetos: %{customdata[3]}"
                "<extra></extra>"
            ),
        )
    )

    fig.update_geos(
        fitbounds="locations",
        visible=False,
        bgcolor="rgba(0,0,0,0)",
        projection_type="mercator",
    )

    apply_layout(fig, margin=dict(t=10, r=20, b=10, l=20))
    return fig


# ── COMPONENTES REUTILIZÁVEIS ─────────────────────────────────────────────────


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


def kpi_block(label: str, value: str, sub: str) -> list:
    return [
        html.Div(label, className="metric-label"),
        html.Div(value, className="metric-value"),
        html.Div(sub, className="metric-subtitle"),
    ]


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


def auth_status_card() -> html.Div:
    if is_authenticated():
        user_name = session.get("user_name") or session.get("user_email") or "Usuário autenticado"
        subtitle = session.get("user_email") or "Sessão ativa no Microsoft Entra ID"

        return glass_card(
            cls="section-card",
            *[
                html.Div(
                    className="section-header",
                    style={"marginBottom": "12px"},
                    children=[
                        html.Div(
                            [
                                html.Div("Conexão Microsoft", className="section-title"),
                                html.Div(
                                    "Autenticação concluída com sucesso. Agora você pode carregar a base protegida do SharePoint.",
                                    className="section-subtitle",
                                ),
                            ]
                        )
                    ],
                ),
                html.Div(
                    style={
                        "padding": "18px 20px",
                        "borderRadius": "16px",
                        "background": "linear-gradient(135deg, rgba(201,168,76,0.18), rgba(29,60,105,0.10))",
                        "border": "1px solid rgba(201,168,76,0.45)",
                        "boxShadow": "0 14px 34px rgba(29,60,105,0.10)",
                        "display": "flex",
                        "justifyContent": "space-between",
                        "alignItems": "center",
                        "gap": "16px",
                        "flexWrap": "wrap",
                    },
                    children=[
                        html.Div(
                            [
                                html.Div(f"Conectado como: {user_name}", className="section-title"),
                                html.Div(subtitle, className="section-subtitle"),
                            ]
                        ),
                        html.A("Sair", href="/logout", className="btn btn-outline"),
                    ],
                ),
            ],
        )

    return glass_card(
        cls="section-card",
        *[
            html.Div(
                className="section-header",
                style={"marginBottom": "12px"},
                children=[
                    html.Div(
                        [
                            html.Div("Entrar com Microsoft", className="section-title"),
                            html.Div(
                                "Faça autenticação para acessar a planilha protegida no SharePoint e habilitar o painel.",
                                className="section-subtitle",
                            ),
                        ]
                    )
                ],
            ),
            html.Div(
                style={
                    "padding": "24px",
                    "borderRadius": "18px",
                    "background": "linear-gradient(135deg, rgba(201,168,76,0.20), rgba(29,60,105,0.10))",
                    "border": "1px solid rgba(201,168,76,0.55)",
                    "display": "flex",
                    "justifyContent": "space-between",
                    "alignItems": "center",
                    "gap": "20px",
                    "flexWrap": "wrap",
                    "boxShadow": "0 16px 40px rgba(29,60,105,0.12)",
                },
                children=[
                    html.Div(
                        style={"maxWidth": "650px"},
                        children=[
                            html.Div("Acesso protegido", className="section-title"),
                            html.Div(
                                "Esta aplicação utiliza login Microsoft para liberar o acesso ao arquivo no SharePoint. "
                                "Enquanto a autenticação não for concluída, nenhum conteúdo da aplicação será exibido.",
                                className="section-subtitle",
                            ),
                        ],
                    ),
                    html.A("Entrar com Microsoft", href="/login", className="btn btn-primary"),
                ],
            ),
        ],
    )


def unauthenticated_page_layout() -> html.Div:
    return html.Div(
        className="page-wrap fade-in",
        style={
            "minHeight": "100vh",
            "display": "flex",
            "alignItems": "flex-start",
            "justifyContent": "center",
            "paddingTop": "40px",
        },
        children=[
            html.Div(
                style={"width": "100%", "maxWidth": "980px"},
                children=[auth_status_card()],
            )
        ],
    )


# ── LAYOUTS DE PÁGINA ─────────────────────────────────────────────────────────


def home_page_layout(df_json: str | None = None) -> html.Div:
    has_data = bool(df_json)

    return html.Div(
        className="page-wrap fade-in",
        children=[
            auth_status_card(),
            html.Div(
                className="page-header",
                style={"marginTop": "6px"},
                children=[
                    html.H1("Dados", className="page-title"),
                    html.P(
                        "Carregue, filtre e exporte os dados da planilha compartilhada.",
                        className="page-subtitle",
                    ),
                ],
            ),
            html.Div(
                className="action-row",
                children=[
                    html.Button("Carregar base", id="btn-load", n_clicks=0, className="btn btn-primary"),
                    html.Button("Selecionar todas", id="btn-select-all", n_clicks=0, className="btn btn-outline"),
                    html.Button("Limpar seleção", id="btn-clear", n_clicks=0, className="btn btn-outline"),
                    html.Button("Exportar Excel", id="btn-export", n_clicks=0, className="btn btn-outline"),
                ],
            ),
            html.Div(
                id="status-box",
                className="status-box status-idle",
                children=[
                    html.Div("Status", className="status-title"),
                    html.Pre("Clique em 'Carregar base' para iniciar.", className="status-message"),
                ],
            ),
            html.Div(id="summary-cards", className="metrics-grid"),
            dcc.Download(id="download-excel"),
            html.Div(
                id="columns-section",
                style={"display": "block" if has_data else "none"},
                children=[
                    glass_card(
                        cls="section-card",
                        *[
                            html.Div(
                                className="section-header",
                                children=[
                                    html.Div(
                                        [
                                            html.Div("Colunas", className="section-title"),
                                            html.Div("Selecione as colunas a manter na base.", className="section-subtitle"),
                                        ]
                                    ),
                                    html.Div(id="selected-count", className="section-counter"),
                                ],
                            ),
                            dcc.Dropdown(
                                id="column-selector",
                                options=[],
                                value=[],
                                multi=True,
                                placeholder="Escolha as colunas...",
                                className="lovable-dropdown",
                            ),
                        ],
                    ),
                ],
            ),
            html.Div(
                id="preview-section",
                style={"display": "block" if has_data else "none"},
                children=[
                    html.Div(
                        className="table-section",
                        children=[
                            html.Div(
                                className="table-title-row",
                                children=[
                                    html.Div("Pré-visualização", className="section-title"),
                                    html.Div("Primeiras 50 linhas da base filtrada", className="section-subtitle"),
                                ],
                            ),
                            html.Div(
                                className="table-shell",
                                children=[
                                    dash_table.DataTable(
                                        id="preview-table",
                                        page_size=10,
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
                        ],
                    ),
                ],
            ),
        ],
    )


def bi_page_layout(df_json: str | None, selected: list | None, filename: str | None) -> html.Div:
    header = html.Div(
        className="page-header",
        children=[
            html.H1("Exploração Livre", className="page-title"),
            html.P(
                "Interface drag-and-drop para análises personalizadas sobre a base carregada.",
                className="page-subtitle",
            ),
        ],
    )

    if not df_json:
        return html.Div(
            className="page-wrap fade-in",
            children=[
                auth_status_card(),
                header,
                _empty_state(
                    "Base não carregada",
                    "Acesse a página de Dados e carregue a base para habilitar o PyGWalker.",
                ),
            ],
        )

    df = filter_df_by_columns(df_json, selected)

    if df.empty:
        return html.Div(
            className="page-wrap fade-in",
            children=[
                auth_status_card(),
                header,
                _empty_state(
                    "Nenhuma coluna selecionada",
                    "Volte para Dados e selecione pelo menos uma coluna.",
                ),
            ],
        )

    orig_df = pd.read_json(StringIO(df_json), orient="split")
    rows, cols = len(df), len(df.columns)
    orig_cols = len(orig_df.columns)
    completeness = (df.notna().sum().sum() / (rows * cols) * 100) if rows * cols > 0 else 100.0
    pyg_html = pyg.to_html(df, appearance="light")

    return html.Div(
        className="page-wrap fade-in",
        children=[
            auth_status_card(),
            header,
            html.Div(
                className="metrics-grid",
                children=[
                    metric_card("Registros", fmt_int_br(rows), "linhas", ACCENT),
                    metric_card("Colunas ativas", fmt_int_br(cols), f"de {fmt_int_br(orig_cols)}", BLUE),
                    metric_card("Preenchimento", brazil_per(completeness, 1), "células não nulas", TEAL),
                    metric_card("Arquivo ativo", filename or "base carregada", "", MUTED),
                ],
            ),
            glass_card(
                *[
                    section_head("PyGWalker", "Arraste campos para criar suas próprias visualizações"),
                    dcc.Loading(
                        type="circle",
                        color=ACCENT,
                        children=[html.Iframe(srcDoc=pyg_html, className="bi-iframe")],
                    ),
                ]
            ),
        ],
    )


# ── APP ───────────────────────────────────────────────────────────────────────

app = Dash(__name__, suppress_callback_exceptions=True, title="SEAID · COSID")
server = app.server
server.secret_key = FLASK_SECRET_KEY
server.config["SESSION_COOKIE_HTTPONLY"] = True
server.config["SESSION_COOKIE_SAMESITE"] = "Lax"
server.config["PREFERRED_URL_SCHEME"] = "https"
server.config["SESSION_COOKIE_SECURE"] = APP_BASE_URL.startswith("https://")
server.wsgi_app = ProxyFix(server.wsgi_app, x_proto=1, x_host=1, x_port=1)

app.layout = html.Div(
    className="app-shell",
    children=[
        dcc.Location(id="url"),
        dcc.Store(id="global-df-json"),
        dcc.Store(id="global-filename"),
        dcc.Store(id="global-selected-columns"),
        dcc.Store(id="global-load-status", data="idle"),
        dcc.Store(id="carteira-operacoes-selected-columns"),
        html.Aside(
            id="sidebar",
            className="sidebar sidebar-hidden",
            children=[
                html.Div(
                    className="sidebar-brand",
                    children=[
                        html.Div("FIN", className="sidebar-logo"),
                        html.Div(
                            [
                                html.Div("SEAID", className="sidebar-brand-title"),
                                html.Div("COSID", className="sidebar-brand-subtitle"),
                            ]
                        ),
                    ],
                ),
                html.Nav(
                    className="sidebar-nav",
                    children=[
                        dcc.Link(
                            [html.Span(className="nav-dot nav-dot-amber"), "Dados"],
                            href="/",
                            className="sidebar-link",
                            id="nav-dados",
                        ),
                        dcc.Link(
                            [html.Span(className="nav-dot nav-dot-blue"), "Painel Analítico"],
                            href="/painel",
                            className="sidebar-link",
                            id="nav-carteira-ativa",
                        ),
                        dcc.Link(
                            [html.Span(className="nav-dot nav-dot-teal"), "Exploração Livre"],
                            href="/bi",
                            className="sidebar-link",
                            id="nav-bi",
                        ),
                    ],
                ),
                html.Div(
                    className="sidebar-footer",
                    children=[
                        html.Div("Ministério do Planejamento", className="sidebar-footer-line"),
                        html.Div("e Orçamento · MPO", className="sidebar-footer-line"),
                    ],
                ),
            ],
        ),
        html.Main(id="page-content", className="main-content"),
    ],
)

# ── ROTAS DE AUTENTICAÇÃO ─────────────────────────────────────────────────────


@server.route("/login")
def login():
    msal_app = build_msal_app()
    flow = msal_app.initiate_auth_code_flow(
        scopes=SCOPES,
        redirect_uri=build_redirect_uri(),
    )
    session["auth_flow"] = flow
    return redirect(flow["auth_uri"])


@server.route("/auth/callback")
def auth_callback():
    flow = session.get("auth_flow")
    if not flow:
        return redirect("/")

    msal_app = build_msal_app()
    result = msal_app.acquire_token_by_auth_code_flow(
        flow,
        dict(request.args),
    )

    session.pop("auth_flow", None)

    if "access_token" not in result:
        error = result.get("error", "unknown_error")
        description = result.get("error_description", "Sem detalhes.")
        return f"Falha na autenticação: {error} - {description}", 400

    claims = result.get("id_token_claims", {}) or {}

    session["access_token"] = result["access_token"]
    session["access_token_expires_at"] = time.time() + int(result.get("expires_in", 3600)) - 60
    session["id_token_claims"] = claims
    session["user_name"] = claims.get("name", "")
    session["user_email"] = claims.get("preferred_username", "")

    return redirect("/")


@server.route("/logout")
def logout():
    clear_auth_session()
    logout_url = (
        f"{build_authority()}/oauth2/v2.0/logout"
        f"?post_logout_redirect_uri={build_logout_redirect_uri()}"
    )
    return redirect(logout_url)


@server.route("/")
def home_redirect():
    return app.index()


# ── CALLBACKS ─────────────────────────────────────────────────────────────────


@callback(
    Output("sidebar", "className"),
    Input("url", "pathname"),
    Input("global-df-json", "data"),
)
def toggle_sidebar(pathname, df_json):
    return "sidebar sidebar-visible" if is_authenticated() else "sidebar sidebar-hidden"


@callback(
    Output("nav-dados", "className"),
    Output("nav-carteira-ativa", "className"),
    Output("nav-bi", "className"),
    Input("url", "pathname"),
)
def update_nav_classes(pathname):
    base = "sidebar-link"
    return (
        base + (" active" if pathname == "/" else ""),
        base + (" active" if pathname == "/painel" else ""),
        base + (" active" if pathname == "/bi" else ""),
    )


@callback(
    Output("page-content", "children"),
    Input("url", "pathname"),
    State("global-df-json", "data"),
    State("global-selected-columns", "data"),
    State("global-filename", "data"),
)
def render_page(pathname, df_json, selected_columns, filename):
    if not is_authenticated():
        return unauthenticated_page_layout()

    if pathname == "/painel":
        return painel_page_layout(df_json)
    if pathname == "/bi":
        return bi_page_layout(df_json, selected_columns, filename)
    return home_page_layout(df_json)


@callback(
    Output("columns-section", "style"),
    Output("preview-section", "style"),
    Input("global-df-json", "data"),
)
def toggle_data_sections(df_json):
    if df_json:
        return {"display": "block"}, {"display": "block"}
    return {"display": "none"}, {"display": "none"}


@callback(
    Output("global-df-json", "data"),
    Output("global-filename", "data"),
    Output("status-box", "children"),
    Output("status-box", "className"),
    Output("column-selector", "options"),
    Output("column-selector", "value"),
    Output("global-load-status", "data"),
    Output("summary-cards", "children"),
    Input("btn-load", "n_clicks"),
    prevent_initial_call=True,
)
def load_shared_file(n_clicks):
    if not n_clicks:
        raise PreventUpdate

    if not is_authenticated():
        return (
            None,
            None,
            [
                html.Div("Status", className="status-title"),
                html.Pre(
                    "É necessário entrar com a Microsoft antes de carregar a base.",
                    className="status-message",
                ),
            ],
            "status-box status-error",
            [],
            [],
            "error",
            [],
        )

    try:
        df = load_dataframe_from_sharepoint()
        options = [{"label": c, "value": c} for c in df.columns]
        selected = list(df.columns)
        df_json = df.to_json(date_format="iso", orient="split")
        filename = Path(SHAREPOINT_FILE_PATH).name or "arquivo_compartilhado.xlsx"
        msg = f"Arquivo carregado com sucesso.\nLinhas: {fmt_int_br(len(df))}  ·  Colunas: {fmt_int_br(len(df.columns))}"

        val_total = df["vl_financiamento_dolar"].sum() if "vl_financiamento_dolar" in df.columns else 0
        n_fontes = df["sg_fonte"].nunique() if "sg_fonte" in df.columns else 0
        n_fases = df["de_fase"].nunique() if "de_fase" in df.columns else 0

        summary = [
            metric_card("Total de Registros", fmt_int_br(len(df)), "operações na base", ACCENT),
            metric_card("Colunas", fmt_int_br(len(df.columns)), "variáveis disponíveis", BLUE),
            metric_card("Financiamento Total", fmt_bi(val_total), "soma do portfólio", TEAL),
            metric_card("Diversidade", f"{fmt_int_br(n_fontes)} fontes", f"{fmt_int_br(n_fases)} fases distintas", ROSE),
        ]

        return (
            df_json,
            filename,
            [
                html.Div("Status", className="status-title"),
                html.Pre(msg, className="status-message"),
            ],
            "status-box status-success",
            options,
            selected,
            "success",
            summary,
        )

    except Exception as exc:
        return (
            None,
            None,
            [
                html.Div("Status", className="status-title"),
                html.Pre(f"Erro ao carregar arquivo:\n\n{exc}", className="status-message"),
            ],
            "status-box status-error",
            [],
            [],
            "error",
            [],
        )


@callback(
    Output("global-selected-columns", "data"),
    Input("column-selector", "value"),
    prevent_initial_call=True,
)
def sync_selected_columns(selected):
    return selected or []


@callback(
    Output("column-selector", "value", allow_duplicate=True),
    Input("btn-select-all", "n_clicks"),
    State("column-selector", "options"),
    prevent_initial_call=True,
)
def select_all_columns(n_clicks, options):
    if not n_clicks or not options:
        raise PreventUpdate
    return [o["value"] for o in options]


@callback(
    Output("column-selector", "value", allow_duplicate=True),
    Input("btn-clear", "n_clicks"),
    prevent_initial_call=True,
)
def clear_columns(n_clicks):
    if not n_clicks:
        raise PreventUpdate
    return []


@callback(
    Output("selected-count", "children"),
    Input("column-selector", "options"),
    Input("global-selected-columns", "data"),
)
def update_selected_count(options, selected):
    total = len(options or [])
    sel = len(selected or [])
    return f"{fmt_int_br(sel)} de {fmt_int_br(total)} selecionadas" if total else ""


@callback(
    Output("preview-table", "data"),
    Output("preview-table", "columns"),
    Input("global-df-json", "data"),
    Input("global-selected-columns", "data"),
)
def update_preview(df_json, selected):
    if not df_json:
        return [], []

    df = filter_df_by_columns(df_json, selected)
    preview = df.head(50).where(pd.notnull(df.head(50)), None)
    return preview.to_dict("records"), [{"name": str(c), "id": str(c)} for c in preview.columns]


@callback(
    Output("download-excel", "data"),
    Input("btn-export", "n_clicks"),
    State("global-df-json", "data"),
    State("global-selected-columns", "data"),
    State("global-filename", "data"),
    prevent_initial_call=True,
)
def export_selected_columns(n_clicks, df_json, selected, filename):
    if not n_clicks or not df_json:
        raise PreventUpdate

    df = filter_df_by_columns(df_json, selected)
    if df.empty:
        raise PreventUpdate

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="dados")
    buf.seek(0)
    name = (filename or "base").replace(".xlsx", "") + "_filtrada.xlsx"
    return dcc.send_bytes(buf.getvalue(), name)


# ── CALLBACKS DO DROPDOWN DE COLUNAS DA ABA OPERAÇÕES ────────────────────────


@callback(
    Output("carteira-operacoes-selected-columns", "data"),
    Input("carteira-operacoes-column-selector", "value"),
    prevent_initial_call=True,
)
def sync_carteira_operacoes_columns(selected):
    return selected or []


@callback(
    Output("carteira-operacoes-column-selector", "value", allow_duplicate=True),
    Input("carteira-btn-select-all-operacoes", "n_clicks"),
    State("carteira-operacoes-column-selector", "options"),
    prevent_initial_call=True,
)
def select_all_carteira_operacoes_columns(n_clicks, options):
    if not n_clicks or not options:
        raise PreventUpdate
    return [o["value"] for o in options]


@callback(
    Output("carteira-operacoes-column-selector", "value", allow_duplicate=True),
    Input("carteira-btn-clear-operacoes", "n_clicks"),
    prevent_initial_call=True,
)
def clear_carteira_operacoes_columns(n_clicks):
    if not n_clicks:
        raise PreventUpdate
    return []


@callback(
    Output("carteira-operacoes-selected-count", "children"),
    Input("carteira-operacoes-column-selector", "options"),
    Input("carteira-operacoes-column-selector", "value"),
)
def update_carteira_operacoes_selected_count(options, selected):
    total = len(options or [])
    sel = len(selected or [])
    return f"{fmt_int_br(sel)} de {fmt_int_br(total)} selecionadas" if total else ""


# ── CALLBACKS DA PÁGINA PAINEL ────────────────────────────────────────────────


@callback(
    Output("carteira-select-de_fase", "value"),
    Output("carteira-select-de_tipo_operacao", "value"),
    Output("carteira-select-nm_proponente", "value"),
    Output("carteira-select-sg_fonte", "value"),
    Output("carteira-select-de_esfera", "value"),
    Output("carteira-select-nm_regiao", "value"),
    Output("carteira-select-nm_setor", "value"),
    Output("carteira-select-nm_subsetor", "value"),
    Output("carteira-select-sg_setor", "value"),
    Output("carteira-select-sys", "value"),
    Output("carteira-select-nm_limite", "value"),
    Input("carteira-btn-clear-selections", "n_clicks"),
    prevent_initial_call=True,
)
def clear_carteira_selections(n_clicks):
    if not n_clicks:
        raise PreventUpdate
    return [], [], [], [], [], [], [], [], [], [], []


@callback(
    Output("carteira-fig-temporal", "figure"),
    Output("carteira-fig-setor", "figure"),
    Output("carteira-fig-fonte", "figure"),
    Output("carteira-fig-fase", "figure"),
    Output("carteira-fig-fase-pct", "figure"),
    Output("carteira-fig-esfera", "figure"),
    Output("carteira-fig-regiao", "figure"),
    Output("carteira-fig-uf", "figure"),
    Output("carteira-fig-mapa-uf", "figure"),
    Output("carteira-kpi-proj", "children"),
    Output("carteira-kpi-ops", "children"),
    Output("carteira-kpi-fin", "children"),
    Output("carteira-kpi-contra", "children"),
    Output("carteira-kpi-ano", "children"),
    Input("carteira-ano-range", "value"),
    Input("carteira-metrica", "value"),
    Input("carteira-select-de_fase", "value"),
    Input("carteira-select-de_tipo_operacao", "value"),
    Input("carteira-select-nm_proponente", "value"),
    Input("carteira-select-sg_fonte", "value"),
    Input("carteira-select-de_esfera", "value"),
    Input("carteira-select-nm_regiao", "value"),
    Input("carteira-select-nm_setor", "value"),
    Input("carteira-select-nm_subsetor", "value"),
    Input("carteira-select-sg_setor", "value"),
    Input("carteira-select-sys", "value"),
    Input("carteira-select-nm_limite", "value"),
    Input("url", "pathname"),
    State("global-df-json", "data"),
)
def update_carteira_charts(
    ano_range,
    metrica,
    de_fase,
    de_tipo_operacao,
    nm_proponente,
    sg_fonte,
    de_esfera,
    nm_regiao,
    nm_setor,
    nm_subsetor,
    sg_setor,
    sys,
    nm_limite,
    pathname,
    df_json,
):
    if pathname != "/painel":
        raise PreventUpdate

    df = prep_carteira_df(df_json, ano_range or [ANO_MIN, ANO_MAX])

    if df is None or df.empty:
        empty_kpi = kpi_block("—", "—", "sem dados")
        return (EMPTY_FIG,) * 9 + (empty_kpi,) * 5

    df = apply_dropdown_filters(
        df,
        de_fase=de_fase,
        de_tipo_operacao=de_tipo_operacao,
        nm_proponente=nm_proponente,
        sg_fonte=sg_fonte,
        de_esfera=de_esfera,
        nm_regiao=nm_regiao,
        nm_setor=nm_setor,
        nm_subsetor=nm_subsetor,
        sg_setor=sg_setor,
        sys=sys,
        nm_limite=nm_limite,
    )

    if df is None or df.empty:
        empty_kpi = kpi_block("—", "—", "sem dados após seleção")
        return (EMPTY_FIG,) * 9 + (empty_kpi,) * 5

    n_proj = contar_projetos_distintos(df)
    n_ops = len(df)
    val_fin = df["vl_financiamento_dolar"].sum() if "vl_financiamento_dolar" in df.columns else 0
    val_contra = df["vl_contrapartida_dolar"].sum() if "vl_contrapartida_dolar" in df.columns else 0

    ano_rec, ano_rec_n = "—", 0
    if "ano_cofiex" in df.columns and df["ano_cofiex"].notna().any():
        por_ano = df.groupby("ano_cofiex").size()
        if not por_ano.empty:
            ano_rec = str(int(por_ano.idxmax()))
            ano_rec_n = int(por_ano.max())

    return (
        chart_temporal(df, metrica),
        chart_setor(df, metrica),
        chart_fonte(df, metrica),
        chart_fase(df, metrica),
        chart_fase_percentual(df, metrica),
        chart_esfera(df, metrica),
        chart_regiao(df, metrica),
        chart_uf(df, metrica),
        chart_mapa_uf(df, metrica),
        kpi_block("Total de Projetos", fmt_int_br(n_proj), "projetos distintos"),
        kpi_block("Total de Operações", fmt_int_br(n_ops), "operações exibidas"),
        kpi_block("Financiamento Total", fmt_bi(val_fin), "soma em dólares"),
        kpi_block("Contrapartida Total", fmt_bi(val_contra), "soma em dólares"),
        kpi_block("Ano com Mais Aprovações", ano_rec, f"{fmt_int_br(ano_rec_n)} operações"),
    )


@callback(
    Output("carteira-operacoes-table", "data"),
    Output("carteira-operacoes-table", "columns"),
    Input("carteira-ano-range", "value"),
    Input("carteira-select-de_fase", "value"),
    Input("carteira-select-de_tipo_operacao", "value"),
    Input("carteira-select-nm_proponente", "value"),
    Input("carteira-select-sg_fonte", "value"),
    Input("carteira-select-de_esfera", "value"),
    Input("carteira-select-nm_regiao", "value"),
    Input("carteira-select-nm_setor", "value"),
    Input("carteira-select-nm_subsetor", "value"),
    Input("carteira-select-sg_setor", "value"),
    Input("carteira-select-sys", "value"),
    Input("carteira-select-nm_limite", "value"),
    Input("carteira-operacoes-selected-columns", "data"),
    Input("url", "pathname"),
    State("global-df-json", "data"),
)
def update_carteira_operacoes_table(
    ano_range,
    de_fase,
    de_tipo_operacao,
    nm_proponente,
    sg_fonte,
    de_esfera,
    nm_regiao,
    nm_setor,
    nm_subsetor,
    sg_setor,
    sys,
    nm_limite,
    selected_columns,
    pathname,
    df_json,
):
    if pathname != "/painel":
        raise PreventUpdate

    df = prep_carteira_df(df_json, ano_range or [ANO_MIN, ANO_MAX])

    if df is None or df.empty:
        return [], []

    df = apply_dropdown_filters(
        df,
        de_fase=de_fase,
        de_tipo_operacao=de_tipo_operacao,
        nm_proponente=nm_proponente,
        sg_fonte=sg_fonte,
        de_esfera=de_esfera,
        nm_regiao=nm_regiao,
        nm_setor=nm_setor,
        nm_subsetor=nm_subsetor,
        sg_setor=sg_setor,
        sys=sys,
        nm_limite=nm_limite,
    )

    if df is None or df.empty:
        return [], []

    if selected_columns:
        valid_cols = [c for c in selected_columns if c in df.columns]
        df = df[valid_cols] if valid_cols else df.iloc[:, 0:0]

    df_out = df.copy().where(pd.notnull(df), None)

    return (
        df_out.to_dict("records"),
        [{"name": str(c), "id": str(c)} for c in df_out.columns],
    )


@callback(
    Output("carteira-download-excel-operacoes", "data"),
    Input("carteira-btn-export-operacoes", "n_clicks"),
    State("carteira-ano-range", "value"),
    State("carteira-select-de_fase", "value"),
    State("carteira-select-de_tipo_operacao", "value"),
    State("carteira-select-nm_proponente", "value"),
    State("carteira-select-sg_fonte", "value"),
    State("carteira-select-de_esfera", "value"),
    State("carteira-select-nm_regiao", "value"),
    State("carteira-select-nm_setor", "value"),
    State("carteira-select-nm_subsetor", "value"),
    State("carteira-select-sg_setor", "value"),
    State("carteira-select-sys", "value"),
    State("carteira-select-nm_limite", "value"),
    State("carteira-operacoes-selected-columns", "data"),
    State("url", "pathname"),
    State("global-df-json", "data"),
    prevent_initial_call=True,
)
def export_carteira_operacoes_excel(
    n_clicks,
    ano_range,
    de_fase,
    de_tipo_operacao,
    nm_proponente,
    sg_fonte,
    de_esfera,
    nm_regiao,
    nm_setor,
    nm_subsetor,
    sg_setor,
    sys,
    nm_limite,
    selected_columns,
    pathname,
    df_json,
):
    if not n_clicks or pathname != "/painel" or not df_json:
        raise PreventUpdate

    df = prep_carteira_df(df_json, ano_range or [ANO_MIN, ANO_MAX])

    if df is None or df.empty:
        raise PreventUpdate

    df = apply_dropdown_filters(
        df,
        de_fase=de_fase,
        de_tipo_operacao=de_tipo_operacao,
        nm_proponente=nm_proponente,
        sg_fonte=sg_fonte,
        de_esfera=de_esfera,
        nm_regiao=nm_regiao,
        nm_setor=nm_setor,
        nm_subsetor=nm_subsetor,
        sg_setor=sg_setor,
        sys=sys,
        nm_limite=nm_limite,
    )

    if df is None or df.empty:
        raise PreventUpdate

    if selected_columns:
        valid_cols = [c for c in selected_columns if c in df.columns]
        df = df[valid_cols] if valid_cols else df.iloc[:, 0:0]

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="operacoes")
    buf.seek(0)

    return dcc.send_bytes(buf.getvalue(), "operacoes_filtradas.xlsx")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8050, debug=True, use_reloader=False)