"""
SEAID / COFIEX — Painel de Financiamentos Externos
=======================================================
Versão : 2.1
Projeto: PEI 19 / BI Gerencial SUFIN — Ministério do Planejamento e Orçamento

Histórico de versões
--------------------
v2.1 (abr/2026)
  - Aba Acompanhamento da Carteira Ativa: KPIs dinâmicos, gráficos por técnico,
    CG, fase e setor integrados ao arquivo principal
  - Aba Validade da Resolução: alertas por horizonte temporal (vencidas / 30 / 90
    / 180 dias / todas), timeline mensal e tabela exportável com highlight por urgência
  - Ajustes defensivos em callbacks com dados nulos
  - Docstring e versionamento adicionados

v2.0 (abr/2026)
  - Autenticação Microsoft via MSAL (App Registration)
  - Cache com Flask-Caching + Parquet (TTL 10 min)
  - Mapa de calor por UF (GeoJSON)
  - Carteira Ativa das CGs como módulo separado (carteira_analistas.py)
  - Carregamento lazy da base principal

v1.0 (mar/2026)
  - Versão inicial: Painel Analítico + PyGWalker + Gestão de Dados
"""

from __future__ import annotations

import json
import locale
import os
import time
from datetime import date, datetime
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
from flask_caching import Cache
from werkzeug.middleware.proxy_fix import ProxyFix

from painel import painel_page_layout
from carteira_analistas import carteira_analistas_page_layout

load_dotenv()

# ── CONFIGURAÇÃO ──────────────────────────────────────────────────────────────

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

TENANT_ID = os.environ["TENANT_ID"]
CLIENT_ID = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]

SHAREPOINT_HOSTNAME = os.environ["SHAREPOINT_HOSTNAME"]
SHAREPOINT_SITE_PATH = os.environ["SHAREPOINT_SITE_PATH"]
SHAREPOINT_FILE_PATH = os.environ["SHAREPOINT_FILE_PATH"]

SHAREPOINT_SITE_PATH_CA = os.environ["SHAREPOINT_SITE_PATH_CA"]
SHAREPOINT_FILE_PATH_CA = os.environ["SHAREPOINT_FILE_PATH_CA"]

FLASK_SECRET_KEY = os.environ["FLASK_SECRET_KEY"]
APP_BASE_URL = os.getenv("APP_BASE_URL", "http://localhost:8050").rstrip("/")

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
PARQUET_CACHE_PATH = CACHE_DIR / "base_preparada.parquet"
PARQUET_META_PATH = CACHE_DIR / "base_preparada_meta.json"

PARQUET_CACHE_PATH_CA = CACHE_DIR / "base_preparada_ca.parquet"
PARQUET_META_PATH_CA = CACHE_DIR / "base_preparada_ca_meta.json"

BASE_MEMORY_CACHE_KEY = "base_preparada_json_v3"
BASE_MEMORY_CACHE_KEY_CA = "base_preparada_ca_json_v1"
BASE_MEMORY_CACHE_TTL_SECONDS = 600

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

# ── HELPERS DA PÁGINA [CGs] CARTEIRA ATIVA ───────────────────────────────────

# Campos de filtro usados na página carteira_analistas.
# Mantidos separados aqui para facilitar o reaproveitamento nos callbacks
# das abas Acompanhamento e Validade da Resolução.
SELECT_FIELDS_CA = [
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

# Cores mais estáveis para a visualização das fases na aba Acompanhamento.
FASE_COLORS = {
    "Em execução": TEAL,
    "Repagamento": BLUE,
    "Aprovado": ACCENT,
    "Aprovada COFIEX": ACCENT,
    "Aguardando Assinatura": "#F59E0B",
    "Em preparação": VIOLET,
    "Finalizada": "#94A3B8",
    "Arquivado": "#CBD5E1",
    "Reprovada": ROSE,
    "Devolvido": "#F97316",
    "Pautado na COFIEX": "#06B6D4",
    "Em negociação": "#6366F1",
}

def _apply_ca_filters(df_json_ca: str | None, filter_values: tuple) -> pd.DataFrame:
    """
    Aplica, na mesma ordem de SELECT_FIELDS_CA, os filtros da página
    [CGs] Carteira Ativa sobre a base carregada em global-df-json-ca.
    """
    if not df_json_ca:
        return pd.DataFrame()

    kwargs = {}
    for (col, _), value in zip(SELECT_FIELDS_CA, filter_values):
        kwargs[col] = value

    return build_filtered_carteira_ca_df(df_json_ca, **kwargs)


def _prep_resolucao_df(
    df_json_ca: str | None,
    filter_values: tuple,
    horizonte: str,
) -> pd.DataFrame:
    """
    Prepara a base da aba 'Validade da Resolução':
    - aplica filtros globais da página;
    - mantém somente operações na fase Aprovado;
    - converte dt_validade_recomendacao;
    - calcula dias_para_vencer;
    - aplica o horizonte selecionado.
    """
    df = _apply_ca_filters(df_json_ca, filter_values)

    if df.empty or "dt_validade_recomendacao" not in df.columns:
        return pd.DataFrame()

    df = df.copy()

    # mantém somente fase Aprovado nesta aba
    if "de_fase" in df.columns:
        df = df[df["de_fase"].fillna("").astype(str).str.strip().str.lower().eq("aprovado")]

    if df.empty:
        return pd.DataFrame()

    df["dt_validade_recomendacao"] = pd.to_datetime(
        df["dt_validade_recomendacao"], errors="coerce"
    )
    df = df[df["dt_validade_recomendacao"].notna()]

    if df.empty:
        return pd.DataFrame()

    hoje = pd.Timestamp(date.today())
    df["dias_para_vencer"] = (df["dt_validade_recomendacao"] - hoje).dt.days

    if horizonte == "vencidas":
        df = df[df["dias_para_vencer"] < 0]
    elif horizonte == "30":
        df = df[(df["dias_para_vencer"] >= 0) & (df["dias_para_vencer"] < 30)]
    elif horizonte == "90":
        df = df[(df["dias_para_vencer"] >= 0) & (df["dias_para_vencer"] < 90)]
    elif horizonte == "180":
        df = df[(df["dias_para_vencer"] >= 0) & (df["dias_para_vencer"] < 180)]
    elif horizonte == "todas":
        pass

    return df.sort_values("dias_para_vencer")

# ── APP / SERVER / CACHE ─────────────────────────────────────────────────────

app = Dash(__name__, suppress_callback_exceptions=True, title="SEAID · COSID")
server = app.server
server.secret_key = FLASK_SECRET_KEY
server.config["SESSION_COOKIE_HTTPONLY"] = True
server.config["SESSION_COOKIE_SAMESITE"] = "Lax"
server.wsgi_app = ProxyFix(server.wsgi_app, x_proto=1, x_host=1)

cache = Cache(
    server,
    config={
        "CACHE_TYPE": "SimpleCache",
        "CACHE_DEFAULT_TIMEOUT": BASE_MEMORY_CACHE_TTL_SECONDS,
    },
)

# ── AUTH / SHAREPOINT / GRAPH HELPERS ────────────────────────────────────────


def build_redirect_uri() -> str:
    return f"{APP_BASE_URL}/auth/callback"


def build_logout_redirect_uri() -> str:
    return f"{APP_BASE_URL}/"


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


# ── CARGA OTIMIZADA DA BASE ───────────────────────────────────────────────────


def preprocess_base_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    df.columns = [str(c).strip() for c in df.columns]
    df = date_date(df)

    if "dt_primeira_cofiex" in df.columns:
        df["ano_cofiex"] = pd.to_datetime(df["dt_primeira_cofiex"], errors="coerce").dt.year
        df["ano_cofiex"] = df["ano_cofiex"].fillna(0).astype(int)
    else:
        df["ano_cofiex"] = 0

    for col in ["vl_financiamento_dolar", "vl_contrapartida_dolar"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(float)

    text_cols = [
        "nm_setor",
        "sg_fonte_resumo",
        "de_fase",
        "de_esfera",
        "nm_regiao",
        "sg_uf",
        "de_tipo_operacao",
        "nm_proponente",
        "sg_fonte",
        "nm_subsetor",
        "sg_setor",
        "sys",
        "nm_limite",
    ]

    for col in text_cols:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype("string")
                .fillna("Não informado")
                .replace(["<NA>", "nan", "None", ""], "Não informado")
            )

    return df


def _save_parquet_metadata() -> None:
    meta = {
        "updated_at_epoch": time.time(),
        "updated_at_iso": datetime.now().isoformat(),
    }
    PARQUET_META_PATH.write_text(
        json.dumps(meta, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _read_parquet_metadata() -> dict[str, Any] | None:
    if not PARQUET_META_PATH.exists():
        return None
    try:
        return json.loads(PARQUET_META_PATH.read_text(encoding="utf-8"))
    except Exception:
        return None


def _save_parquet_metadata_ca() -> None:
    meta = {
        "updated_at_epoch": time.time(),
        "updated_at_iso": datetime.now().isoformat(),
    }
    PARQUET_META_PATH_CA.write_text(
        json.dumps(meta, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _read_parquet_metadata_ca() -> dict[str, Any] | None:
    if not PARQUET_META_PATH_CA.exists():
        return None
    try:
        return json.loads(PARQUET_META_PATH_CA.read_text(encoding="utf-8"))
    except Exception:
        return None


def _is_parquet_fresh(ttl_seconds: int = BASE_MEMORY_CACHE_TTL_SECONDS) -> bool:
    if not PARQUET_CACHE_PATH.exists():
        return False

    meta = _read_parquet_metadata()
    if not meta:
        return False

    updated_at = float(meta.get("updated_at_epoch", 0))
    if updated_at <= 0:
        return False

    return (time.time() - updated_at) < ttl_seconds


def _is_parquet_fresh_ca(ttl_seconds: int = BASE_MEMORY_CACHE_TTL_SECONDS) -> bool:
    if not PARQUET_CACHE_PATH_CA.exists():
        return False

    meta = _read_parquet_metadata_ca()
    if not meta:
        return False

    updated_at = float(meta.get("updated_at_epoch", 0))
    if updated_at <= 0:
        return False

    return (time.time() - updated_at) < ttl_seconds


def download_and_prepare_base_df() -> pd.DataFrame:
    access_token = get_access_token_from_session()
    client = GraphSharePointClient(access_token=access_token)

    raw = client.download_excel_from_site(
        hostname=SHAREPOINT_HOSTNAME,
        site_path=SHAREPOINT_SITE_PATH,
        file_path=SHAREPOINT_FILE_PATH,
    )

    df = pd.read_excel(BytesIO(raw))
    df = preprocess_base_df(df)

    df.to_parquet(PARQUET_CACHE_PATH, index=False)
    _save_parquet_metadata()

    return df


# def download_and_prepare_base_df_ca() -> pd.DataFrame:
#     access_token = get_access_token_from_session()
#     client = GraphSharePointClient(access_token=access_token)

#     raw = client.download_excel_from_site(
#         hostname=SHAREPOINT_HOSTNAME,
#         site_path=SHAREPOINT_SITE_PATH_CA,
#         file_path=SHAREPOINT_FILE_PATH_CA,
#     )

#     df = pd.read_excel(BytesIO(raw), sheet_name='base')
#     df = df.copy()
#     df.columns = [str(c).strip() for c in df.columns]
#     df = date_date(df)

#     df.to_parquet(PARQUET_CACHE_PATH_CA, index=False)
#     _save_parquet_metadata_ca()

#     return df

def download_and_prepare_base_df_ca() -> pd.DataFrame:
    access_token = get_access_token_from_session()
    client = GraphSharePointClient(access_token=access_token)

    raw = client.download_excel_from_site(
        hostname=SHAREPOINT_HOSTNAME,
        site_path=SHAREPOINT_SITE_PATH_CA,
        file_path=SHAREPOINT_FILE_PATH_CA,
    )

    # Lê tudo como objeto para evitar conversões ruins automáticas
    df = pd.read_excel(BytesIO(raw), sheet_name="base", dtype=object)
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    date_cols_ca = [
        "dt_tx_cambio",
        "dt_primeiro_recebimento",
        "dt_ultimo_recebimento",
        "dt_primeira_cofiex",
        "dt_validade_recomendacao",
        "dt_reuniao_negociacao",
        "dt_aprov_diretoria",
        "dt_aprovacao_senado",
        "dt_publicacao_dou",
        "dt_assinatura",
        "dt_efetividade",
        "dt_primeiro_desembolso",
        "dt_ultimo_desembolso_original",
        "dt_ultimo_desembolso_vigente",
        "dt_primeira_amortizacao",
        "dt_ultima_amortizacao",
        "dt_encerramento",
    ]

    def parse_ca_date(value):
        if pd.isna(value):
            return pd.NaT

        # datetime/date já válidos
        if isinstance(value, (pd.Timestamp, datetime)):
            return pd.to_datetime(value, errors="coerce")

        # strings
        if isinstance(value, str):
            v = value.strip()
            if v == "":
                return pd.NaT

            # evita converter "0" em 1970
            if v in {"0", "0.0", "00/00/0000", "0000-00-00"}:
                return pd.NaT

            # tenta padrão brasileiro primeiro
            dt = pd.to_datetime(v, errors="coerce", dayfirst=True)
            if pd.notna(dt):
                return dt

            # fallback
            return pd.to_datetime(v, errors="coerce")

        # números
        if isinstance(value, (int, float)):
            # evita 0 -> 1970-01-01
            if value in (0, 0.0):
                return pd.NaT

            # trata serial do Excel
            if 1 <= float(value) <= 60000:
                return pd.to_datetime(value, unit="D", origin="1899-12-30", errors="coerce")

            # números fora da faixa não devem virar data
            return pd.NaT

        return pd.to_datetime(value, errors="coerce")

    for col in date_cols_ca:
        if col in df.columns:
            df[col] = df[col].apply(parse_ca_date)

    text_cols = [
        "cd_pleito",
        "nu_processo_sei",
        "nm_cg",
        "nm_tecnico",
        "nm_pleito",
        "sg_pleito",
        "nm_proponente",
        "de_tipo_operacao",
        "de_fase",
        "de_justificativa_arq",
        "de_fonte",
        "sg_fonte",
        "sg_fonte_resumo",
        "nm_moeda",
        "sg_moeda",
        "symbol_moeda",
        "de_esfera",
        "nm_regiao",
        "nm_uf",
        "sg_uf",
        "nm_municipio",
        "nm_setor",
        "nm_subsetor",
        "nm_classificacao",
        "nu_resolucao_senado",
        "nu_operacao",
        "sys",
        "tp_esfera",
        "nm_solicitante",
        "tp_solicitante",
        "nm_limite",
    ]

    for col in text_cols:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype("string")
                .fillna("Não informado")
                .replace(["<NA>", "nan", "None"], "Não informado")
            )

    numeric_cols = [
        "vl_financiamento_ref",
        "vl_contrapartida_ref",
        "tx_cambio",
        "vl_financiamento_dolar",
        "vl_contrapartida_dolar",
        "cd_uf",
        "cd_municipio",
        "id_pleito",
        "qtde_cofiex_resultado",
        "vl_financiamento_pleito",
        "vl_contrapartida_pleito",
        "taxa_cambio_pleito",
    ]

    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    df.to_parquet(PARQUET_CACHE_PATH_CA, index=False)
    _save_parquet_metadata_ca()

    return df


def load_prepared_base_df(force_refresh: bool = False) -> pd.DataFrame:
    if not force_refresh and _is_parquet_fresh():
        return pd.read_parquet(PARQUET_CACHE_PATH)

    return download_and_prepare_base_df()


def load_prepared_base_df_ca(force_refresh: bool = False) -> pd.DataFrame:
    if not force_refresh and _is_parquet_fresh_ca():
        return pd.read_parquet(PARQUET_CACHE_PATH_CA)

    return download_and_prepare_base_df_ca()


def get_prepared_base_json(force_refresh: bool = False) -> str:
    if not force_refresh:
        cached_json = cache.get(BASE_MEMORY_CACHE_KEY)
        if cached_json:
            return cached_json

    df = load_prepared_base_df(force_refresh=force_refresh)
    df_json = df.to_json(date_format="iso", orient="split")
    cache.set(BASE_MEMORY_CACHE_KEY, df_json, timeout=BASE_MEMORY_CACHE_TTL_SECONDS)
    return df_json


def get_prepared_base_json_ca(force_refresh: bool = False) -> str:
    if not force_refresh:
        cached_json = cache.get(BASE_MEMORY_CACHE_KEY_CA)
        if cached_json:
            return cached_json

    df = load_prepared_base_df_ca(force_refresh=force_refresh)
    df_json = df.to_json(date_format="iso", orient="split")
    cache.set(BASE_MEMORY_CACHE_KEY_CA, df_json, timeout=BASE_MEMORY_CACHE_TTL_SECONDS)
    return df_json


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


def prep_painel_df(df_json: str | None, ano_range: list) -> pd.DataFrame | None:
    if not df_json:
        return None

    df = get_full_df(df_json).copy()
    if df.empty:
        return df

    df = date_date(df)

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
            df[col] = (
                df[col]
                .astype("string")
                .fillna(default)
                .replace(["<NA>", "nan", "None", ""], default)
            )

    if "dt_primeira_cofiex" in df.columns:
        df["ano_cofiex"] = pd.to_datetime(df["dt_primeira_cofiex"], errors="coerce").dt.year
        df["ano_cofiex"] = df["ano_cofiex"].fillna(0).astype(int)
    elif "ano_cofiex" in df.columns:
        df["ano_cofiex"] = pd.to_numeric(df["ano_cofiex"], errors="coerce").fillna(0).astype(int)
    else:
        df["ano_cofiex"] = 0

    df = df[df["ano_cofiex"] > 0]

    if ano_range and not df.empty:
        lo, hi = ano_range[0], ano_range[1]
        df = df[df["ano_cofiex"].between(lo, hi)]

    return df


def _normalize_text_col(df: pd.DataFrame, col: str, default: str = "Não informado") -> pd.DataFrame:
    if col in df.columns:
        df[col] = (
            df[col]
            .astype("string")
            .fillna(default)
            .replace(["<NA>", "nan", "None", ""], default)
        )
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
    if "cd_pleito" in df.columns:
        return df["cd_pleito"].dropna().nunique()
    return len(df)


def fmt_bi(v: float) -> str:
    if v >= 1e9:
        return f"US$ {brazil_vlr(v / 1e9, 1)}bi"
    if v >= 1e6:
        return f"US$ {brazil_vlr(v / 1e6, 0)}M"
    return f"US$ {brazil_vlr(v, 0)}"


def build_filtered_carteira_df(
    df_json,
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
):
    df = prep_carteira_df(df_json, ano_range or [ANO_MIN, ANO_MAX])

    if df is None or df.empty:
        return df

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
    return df


def build_filtered_carteira_ca_df(
    df_json_ca,
    de_fase=None,
    de_tipo_operacao=None,
    nm_proponente=None,
    sg_fonte=None,
    de_esfera=None,
    nm_regiao=None,
    nm_setor=None,
    nm_subsetor=None,
    sys=None,
    nm_limite=None,
    nm_cg=None,
    nm_tecnico=None,
    cd_pleito=None,
):
    if not df_json_ca:
        return pd.DataFrame()

    df = pd.read_json(StringIO(df_json_ca), orient="split")
    if df.empty:
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
        "sys": sys,
        "nm_limite": nm_limite,
        "nm_cg": nm_cg,
        "nm_tecnico": nm_tecnico,
        "cd_pleito": cd_pleito,
    }

    out = df.copy()

    # normaliza os campos usados nos filtros
    for col in filtros.keys():
        if col in out.columns:
            out[col] = (
                out[col]
                .astype("string")
                .fillna("Não informado")
                .replace(["<NA>", "nan", "None", ""], "Não informado")
            )

    for col, valores in filtros.items():
        if col in out.columns and valores:
            out = out[out[col].isin([str(v) for v in valores])]

    return out


# ── CONSTRUTORES DE GRÁFICOS ──────────────────────────────────────────────────


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
    customdata = list(zip(grp["valor_fmt"], grp["qtd_fmt"], grp["proj_fmt"]))

    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            x=grp["ano_cofiex"],
            y=y_vals,
            name="Empréstimo (US$ milhões)" if metrica == "valor" else "Nº Operações",
            showlegend=(metrica == "valor"),
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
            xaxis=merge_dict(XAXIS_DEF, dtick=1),
            yaxis=merge_dict(YAXIS_DEF, tickformat=",.0f", ticksuffix="M"),
            margin=dict(t=30, r=20, b=50, l=20),
            legend=DEFAULT_LEGEND,
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
            xaxis=merge_dict(XAXIS_DEF, dtick=1),
            yaxis=merge_dict(YAXIS_DEF, tickformat=",.0f"),
            margin=dict(t=30, r=20, b=50, l=20),
            legend=DEFAULT_LEGEND,
            showlegend=False,
        )

    return fig


def chart_setor(df: pd.DataFrame, metrica: str) -> go.Figure:
    if "nm_setor" not in df.columns or df.empty:
        return EMPTY_FIG

    agg_dict = {
        "val": ("vl_financiamento_dolar", "sum"),
        "qtd": ("nm_setor", "size"),
        "proj": ("cd_pleito", pd.Series.nunique) if "cd_pleito" in df.columns else ("nm_setor", "size"),
    }

    grp = df.groupby("nm_setor").agg(**agg_dict).reset_index().rename(columns={"nm_setor": "setor"})
    if grp.empty:
        return EMPTY_FIG

    grp = grp.nlargest(10, "val" if metrica == "valor" else "qtd").sort_values(
        "val" if metrica == "valor" else "qtd", ascending=False
    )

    grp["val_mi"] = grp["val"] / 1e6
    grp["val_fmt"] = grp["val_mi"].apply(lambda x: f"{brazil_vlr(x, 0)}M")
    grp["qtd_fmt"] = grp["qtd"].apply(fmt_int_br)
    grp["proj_fmt"] = grp["proj"].apply(fmt_int_br)

    xval = grp["val_mi"] if metrica == "valor" else grp["qtd"]

    fig = go.Figure(
        go.Bar(
            x=xval,
            y=grp["setor"],
            orientation="h",
            marker_color=[COLOR_SEQUENCE[i % len(COLOR_SEQUENCE)] for i in range(len(grp))],
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
    )
    return fig


def chart_fonte(df: pd.DataFrame, metrica: str) -> go.Figure:
    if "sg_fonte_resumo" not in df.columns or df.empty:
        return EMPTY_FIG

    grp = (
        df.groupby("sg_fonte_resumo")
        .agg(
            val=("vl_financiamento_dolar", "sum") if metrica == "valor" else ("sg_fonte_resumo", "size"),
            valor_fin=("vl_financiamento_dolar", "sum"),
            qtd=("sg_fonte_resumo", "size"),
            proj=("cd_pleito", pd.Series.nunique) if "cd_pleito" in df.columns else ("sg_fonte_resumo", "size"),
        )
        .reset_index()
        .rename(columns={"sg_fonte_resumo": "fonte"})
    )

    if grp.empty:
        return EMPTY_FIG

    grp = grp.sort_values("val", ascending=False).copy()
    grp["valor_fmt"] = (grp["valor_fin"] / 1e6).apply(lambda x: f"{brazil_vlr(x, 0)}M")
    grp["qtd_fmt"] = grp["qtd"].apply(fmt_int_br)
    grp["proj_fmt"] = grp["proj"].apply(fmt_int_br)

    fig = go.Figure(
        go.Pie(
            labels=grp["fonte"],
            values=grp["val"],
            hole=0.62,
            marker=dict(
                colors=[COLOR_SEQUENCE[i % len(COLOR_SEQUENCE)] for i in range(len(grp))],
                line=dict(color="white", width=2),
            ),
            textinfo="label+percent",
            hovertext=[
                (
                    f"Valor de financiamento: US$ {valor}<br>"
                    f"Quantidade de operações: {qtd}<br>"
                    f"Quantidade de projetos: {proj}"
                )
                for valor, qtd, proj in zip(grp["valor_fmt"], grp["qtd_fmt"], grp["proj_fmt"])
            ],
            hovertemplate=(
                "<b>%{label}</b><br>"
                "%{hovertext}<br>"
                "Percentual: %{percent}<extra></extra>"
            ),
        )
    )

    apply_layout(
        fig,
        margin=dict(t=20, r=130, b=20, l=20),
        legend=dict(orientation="v", x=1.02, y=0.5, font=dict(size=11)),
    )
    return fig


def chart_fase_percentual(df: pd.DataFrame, metrica: str) -> go.Figure:
    if "de_fase" not in df.columns or df.empty:
        return EMPTY_FIG

    grp = (
        df.groupby("de_fase")
        .agg(
            val=("vl_financiamento_dolar", "sum") if metrica == "valor" else ("de_fase", "size"),
            valor_fin=("vl_financiamento_dolar", "sum"),
            qtd=("de_fase", "size"),
            proj=("cd_pleito", pd.Series.nunique) if "cd_pleito" in df.columns else ("de_fase", "size"),
        )
        .reset_index()
        .rename(columns={"de_fase": "fase"})
    )

    if grp.empty:
        return EMPTY_FIG

    grp = grp.sort_values("val", ascending=False).copy()
    grp["valor_fmt"] = (grp["valor_fin"] / 1e6).apply(lambda x: f"{brazil_vlr(x, 0)}M")
    grp["qtd_fmt"] = grp["qtd"].apply(fmt_int_br)
    grp["proj_fmt"] = grp["proj"].apply(fmt_int_br)

    fig = go.Figure(
        go.Pie(
            labels=grp["fase"],
            values=grp["val"],
            hole=0.62,
            marker=dict(
                colors=[COLOR_SEQUENCE[i % len(COLOR_SEQUENCE)] for i in range(len(grp))],
                line=dict(color="white", width=2),
            ),
            textinfo="label+percent",
            hovertext=[
                (
                    f"Valor de financiamento: US$ {valor}<br>"
                    f"Quantidade de operações: {qtd}<br>"
                    f"Quantidade de projetos: {proj}"
                )
                for valor, qtd, proj in zip(grp["valor_fmt"], grp["qtd_fmt"], grp["proj_fmt"])
            ],
            hovertemplate=(
                "<b>%{label}</b><br>"
                "%{hovertext}<br>"
                "Percentual: %{percent}<extra></extra>"
            ),
        )
    )

    apply_layout(
        fig,
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

    grp = (
        df.groupby("de_fase")
        .agg(
            valor=("vl_financiamento_dolar", "sum"),
            qtd=("de_fase", "size"),
            proj=("cd_pleito", pd.Series.nunique) if "cd_pleito" in df.columns else ("de_fase", "size"),
        )
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

    fig = go.Figure(
        go.Bar(
            x=grp["fase_label"],
            y=y_bar,
            marker_color=[COLOR_SEQUENCE[i % len(COLOR_SEQUENCE)] for i in range(len(grp))],
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
        yaxis=merge_dict(YAXIS_DEF, tickformat=",.0f", ticksuffix="M" if metrica == "valor" else ""),
        margin=dict(t=20, r=20, b=110, l=60),
    )
    return fig


def chart_esfera(df: pd.DataFrame, metrica: str) -> go.Figure:
    if "de_esfera" not in df.columns or df.empty:
        return EMPTY_FIG

    grp = (
        df.groupby("de_esfera")
        .agg(
            valor=("vl_financiamento_dolar", "sum"),
            qtd=("de_esfera", "size"),
            qtd_projetos=("cd_pleito", pd.Series.nunique) if "cd_pleito" in df.columns else ("de_esfera", "size"),
        )
        .reset_index()
        .rename(columns={"de_esfera": "esfera"})
    )

    if grp.empty:
        return EMPTY_FIG

    grp = grp.sort_values("valor" if metrica == "valor" else "qtd", ascending=False)
    yval = grp["valor"] / 1e6 if metrica == "valor" else grp["qtd"]

    grp["valor_fmt"] = (grp["valor"] / 1e6).apply(lambda x: f"{brazil_vlr(x, 0)}M")
    grp["qtd_fmt"] = grp["qtd"].apply(fmt_int_br)
    grp["qtd_projetos_fmt"] = grp["qtd_projetos"].apply(fmt_int_br)

    fig = go.Figure(
        go.Bar(
            x=grp["esfera"],
            y=yval,
            width=0.45,
            marker_color=[COLOR_SEQUENCE[i % len(COLOR_SEQUENCE)] for i in range(len(grp))],
            customdata=list(zip(grp["esfera"], grp["valor_fmt"], grp["qtd_fmt"], grp["qtd_projetos_fmt"])),
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Valor de financiamento: US$ %{customdata[1]}<br>"
                "Quantidade de operações: %{customdata[2]}<br>"
                "Quantidade de projetos: %{customdata[3]}<extra></extra>"
            ),
        )
    )

    apply_layout(
        fig,
        yaxis=merge_dict(YAXIS_DEF, tickformat=",.0f", ticksuffix="M" if metrica == "valor" else ""),
        margin=dict(t=20, r=20, b=50, l=60),
    )
    return fig


def chart_regiao(df: pd.DataFrame, metrica: str) -> go.Figure:
    if "nm_regiao" not in df.columns or df.empty:
        return EMPTY_FIG

    grp = (
        df.groupby("nm_regiao")
        .agg(
            valor=("vl_financiamento_dolar", "sum"),
            qtd=("nm_regiao", "size"),
            qtd_projetos=("cd_pleito", pd.Series.nunique) if "cd_pleito" in df.columns else ("nm_regiao", "size"),
        )
        .reset_index()
        .rename(columns={"nm_regiao": "regiao"})
    )

    if grp.empty:
        return EMPTY_FIG

    grp = grp.sort_values("valor" if metrica == "valor" else "qtd", ascending=False)
    yval = grp["valor"] / 1e6 if metrica == "valor" else grp["qtd"]

    grp["valor_fmt"] = (grp["valor"] / 1e6).apply(lambda x: f"{brazil_vlr(x, 0)}M")
    grp["qtd_fmt"] = grp["qtd"].apply(fmt_int_br)
    grp["qtd_projetos_fmt"] = grp["qtd_projetos"].apply(fmt_int_br)

    fig = go.Figure(
        go.Bar(
            x=grp["regiao"],
            y=yval,
            marker_color=[COLOR_SEQUENCE[i % len(COLOR_SEQUENCE)] for i in range(len(grp))],
            customdata=list(zip(grp["regiao"], grp["valor_fmt"], grp["qtd_fmt"], grp["qtd_projetos_fmt"])),
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Valor de financiamento: US$ %{customdata[1]}<br>"
                "Quantidade de operações: %{customdata[2]}<br>"
                "Quantidade de projetos: %{customdata[3]}<extra></extra>"
            ),
        )
    )

    apply_layout(
        fig,
        yaxis=merge_dict(YAXIS_DEF, tickformat=",.0f", ticksuffix="M" if metrica == "valor" else ""),
        margin=dict(t=20, r=20, b=50, l=60),
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

    grp = (
        df_plot.groupby("sg_uf")
        .agg(
            valor=("vl_financiamento_dolar", "sum"),
            qtd=("sg_uf", "size"),
            qtd_projetos=("cd_pleito", pd.Series.nunique) if "cd_pleito" in df_plot.columns else ("sg_uf", "size"),
        )
        .reset_index()
        .rename(columns={"sg_uf": "uf"})
    )

    if grp.empty:
        return EMPTY_FIG

    grp = grp.nlargest(15, "valor" if metrica == "valor" else "qtd").sort_values(
        "valor" if metrica == "valor" else "qtd", ascending=False
    )

    xval = grp["valor"] / 1e6 if metrica == "valor" else grp["qtd"]
    grp["valor_fmt"] = (grp["valor"] / 1e6).apply(lambda x: f"{brazil_vlr(x, 0)}M")
    grp["qtd_fmt"] = grp["qtd"].apply(fmt_int_br)
    grp["qtd_projetos_fmt"] = grp["qtd_projetos"].apply(fmt_int_br)

    fig = go.Figure(
        go.Bar(
            x=xval,
            y=grp["uf"],
            orientation="h",
            marker_color=[COLOR_SEQUENCE[i % len(COLOR_SEQUENCE)] for i in range(len(grp))],
            customdata=list(zip(grp["uf"], grp["valor_fmt"], grp["qtd_fmt"], grp["qtd_projetos_fmt"])),
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Valor de financiamento: US$ %{customdata[1]}<br>"
                "Quantidade de operações: %{customdata[2]}<br>"
                "Quantidade de projetos: %{customdata[3]}<extra></extra>"
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
    )
    return fig


@cache.memoize(timeout=86400)
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

    df_map["cd_uf"] = pd.to_numeric(df_map["cd_uf"], errors="coerce")
    df_map = df_map[df_map["cd_uf"].notna()]

    if df_map.empty:
        return EMPTY_FIG

    grp = df_map.groupby("cd_uf", as_index=False).agg(
        valor_financiamento=("vl_financiamento_dolar", "sum"),
        qtd_operacoes=("cd_uf", "size"),
        qtd_projetos=("cd_pleito", pd.Series.nunique) if "cd_pleito" in df_map.columns else ("cd_uf", "size"),
    )
    grp["val"] = grp["valor_financiamento"] if metrica == "valor" else grp["qtd_operacoes"]

    zmin = float(grp["val"].min()) if not grp.empty else 0
    zmax = float(grp["val"].max()) if not grp.empty else 0

    if metrica == "valor":
        colorbar = dict(
            title=dict(text="US$ milhões", side="right", font=dict(size=12)),
            thickness=14,
            len=0.75,
            y=0.5,
            tickmode="array",
            tickvals=[zmin, zmax],
            ticktext=[f"{brazil_vlr(zmin / 1e6, 0)}M", f"{brazil_vlr(zmax / 1e6, 0)}M"],
        )
    else:
        colorbar = dict(
            title=dict(text="Operações", side="right", font=dict(size=12)),
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

    grp["valor_financiamento_fmt"] = grp["valor_financiamento"].apply(lambda x: f"US$ {brazil_vlr(x / 1e6, 0)}M")
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
                "Quantidade de projetos: %{customdata[3]}<extra></extra>"
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
                    style={"marginBottom": "10px"},
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
                        "display": "flex",
                        "justifyContent": "space-between",
                        "alignItems": "center",
                        "gap": "16px",
                        "flexWrap": "wrap",
                        "padding": "10px 14px",
                        "borderRadius": "14px",
                        "background": "linear-gradient(135deg, rgba(201,168,76,0.12), rgba(29,60,105,0.08))",
                        "border": "1px solid rgba(201,168,76,0.35)",
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
                style={"marginBottom": "10px"},
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
                    "padding": "18px",
                    "borderRadius": "16px",
                    "background": "linear-gradient(135deg, rgba(201,168,76,0.16), rgba(29,60,105,0.08))",
                    "border": "1px solid rgba(201,168,76,0.45)",
                    "display": "flex",
                    "justifyContent": "space-between",
                    "alignItems": "center",
                    "gap": "16px",
                    "flexWrap": "wrap",
                    "boxShadow": "0 10px 30px rgba(29,60,105,0.08)",
                },
                children=[
                    html.Div(
                        [
                            html.Div("Acesso protegido", className="section-title"),
                            html.Div(
                                "A autenticação é obrigatória antes de exibir qualquer conteúdo da aplicação.",
                                className="section-subtitle",
                            ),
                        ]
                    ),
                    html.A("Entrar com Microsoft", href="/login", className="btn btn-primary"),
                ],
            ),
        ],
    )


def unauthenticated_page_layout() -> html.Div:
    return html.Div(className="page-wrap fade-in", children=[auth_status_card()])


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
                    html.P("Carregue, filtre e exporte os dados da planilha compartilhada.", className="page-subtitle"),
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
    return html.Div(
        className="page-wrap fade-in",
        children=[
            auth_status_card(),
            html.Div(
                className="page-header",
                children=[
                    html.H1("Exploração Livre", className="page-title"),
                    html.P("Interface drag-and-drop para análises personalizadas sobre a base carregada.", className="page-subtitle"),
                ],
            ),
            html.Div(id="bi-lazy-container"),
        ],
    )


# ── APP LAYOUT ────────────────────────────────────────────────────────────────

app.layout = html.Div(
    className="app-shell",
    children=[
        dcc.Location(id="url"),
        dcc.Store(id="global-df-json"),
        dcc.Store(id="global-filename"),
        dcc.Store(id="global-selected-columns"),
        dcc.Store(id="global-load-status", data="idle"),
        dcc.Store(id="global-df-json-ca"),
        dcc.Store(id="global-filename-ca"),
        dcc.Store(id="global-load-status-ca", data="idle"),
        dcc.Store(id="carteira-ca-selected-columns"),
        dcc.Store(id="carteira-ca-loaded", data=False),
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
                        html.Div(
                            [
                                html.Div("Controles paralelos", className="sidebar-brand-title"),
                            ],
                            style={
                                "marginTop": "14px",
                                "paddingTop": "14px",
                                "borderTop": "1px solid rgba(148,163,184,0.25)",
                            },
                        ),
                        dcc.Link(
                            [html.Span(className="nav-dot nav-dot-violet"), "[CGs] Carteira Ativa"],
                            href="/carteira-cgs",
                            className="sidebar-link",
                            id="nav-carteira-cgs",
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
    flow = msal_app.initiate_auth_code_flow(scopes=SCOPES, redirect_uri=build_redirect_uri())
    session["auth_flow"] = flow
    return redirect(flow["auth_uri"])


@server.route("/auth/callback")
def auth_callback():
    flow = session.get("auth_flow")
    if not flow:
        return redirect("/")

    msal_app = build_msal_app()
    result = msal_app.acquire_token_by_auth_code_flow(flow, dict(request.args))
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
    logout_url = f"{build_authority()}/oauth2/v2.0/logout?post_logout_redirect_uri={build_logout_redirect_uri()}"
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
    Output("nav-carteira-cgs", "className"),
    Input("url", "pathname"),
)
def update_nav_classes(pathname):
    base = "sidebar-link"
    return (
        base + (" active" if pathname == "/" else ""),
        base + (" active" if pathname == "/painel" else ""),
        base + (" active" if pathname == "/bi" else ""),
        base + (" active" if pathname == "/carteira-cgs" else ""),
    )


@callback(
    Output("page-content", "children"),
    Input("url", "pathname"),
    Input("global-df-json-ca", "data"),
    State("global-df-json", "data"),
    State("global-selected-columns", "data"),
    State("global-filename", "data"),
)
def render_page(pathname, df_json_ca, df_json, selected_columns, filename):
    if not is_authenticated():
        return unauthenticated_page_layout()

    if pathname == "/painel":
        return painel_page_layout(df_json)
    if pathname == "/bi":
        return bi_page_layout(df_json, selected_columns, filename)
    if pathname == "/carteira-cgs":
        return carteira_analistas_page_layout(
            df_json_ca=df_json_ca,
            auth_component=auth_status_card(),
        )
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
    Output("bi-lazy-container", "children"),
    Input("url", "pathname"),
    State("global-df-json", "data"),
    State("global-selected-columns", "data"),
    State("global-filename", "data"),
)
def render_bi_lazy_content(pathname, df_json, selected, filename):
    if pathname != "/bi":
        raise PreventUpdate

    if not df_json:
        return _empty_state("Base não carregada", "Acesse a página de Dados e carregue a base para habilitar a Exploração Livre.")

    df = filter_df_by_columns(df_json, selected)
    if df.empty:
        return _empty_state("Nenhuma coluna selecionada", "Volte para Dados e selecione pelo menos uma coluna.")

    orig_df = pd.read_json(StringIO(df_json), orient="split")
    rows, cols = len(df), len(df.columns)
    orig_cols = len(orig_df.columns)
    completeness = (df.notna().sum().sum() / (rows * cols) * 100) if rows * cols > 0 else 100.0

    pyg_html = pyg.to_html(df, appearance="light")

    return [
        html.Div(
            className="metrics-grid",
            children=[
                metric_card("Registros", fmt_int_br(rows), "linhas na base", ACCENT),
                metric_card("Colunas ativas", fmt_int_br(cols), f"de {fmt_int_br(orig_cols)}", BLUE),
                metric_card("Preenchimento", brazil_per(completeness, 1), "células não nulas", TEAL),
                metric_card("Arquivo ativo", filename or "base carregada", "", ROSE),
            ],
        ),
        glass_card(
            *[
                section_head("PyGWalker", "Carregado somente ao abrir esta aba, utilizando toda a base disponível."),
                dcc.Loading(type="circle", color=ACCENT, children=[html.Iframe(srcDoc=pyg_html, className="bi-iframe")]),
            ]
        ),
    ]


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
            [html.Div("Status", className="status-title"), html.Pre("É necessário entrar com a Microsoft antes de carregar a base.", className="status-message")],
            "status-box status-error",
            [],
            [],
            "error",
            [],
        )

    try:
        df_json = get_prepared_base_json(force_refresh=False)
        df = pd.read_json(StringIO(df_json), orient="split")

        options = [{"label": c, "value": c} for c in df.columns]
        selected = list(df.columns)
        filename = Path(SHAREPOINT_FILE_PATH).name or "arquivo_compartilhado.xlsx"

        msg = (
            f"Arquivo carregado com sucesso.\n"
            f"Linhas: {fmt_int_br(len(df))}  ·  Colunas: {fmt_int_br(len(df.columns))}\n"
        )

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
            [html.Div("Status", className="status-title"), html.Pre(msg, className="status-message")],
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
            [html.Div("Status", className="status-title"), html.Pre(f"Erro ao carregar arquivo:\n\n{exc}", className="status-message")],
            "status-box status-error",
            [],
            [],
            "error",
            [],
        )


@callback(
    Output("global-df-json-ca", "data"),
    Output("global-filename-ca", "data"),
    Output("status-box-ca", "children"),
    Output("status-box-ca", "className"),
    Output("global-load-status-ca", "data"),
    Output("summary-cards-ca", "children"),
    Output("preview-section-ca", "style"),
    Output("carteira-ca-loaded", "data"),
    Output("carteira-ca-column-selector", "options"),
    Output("carteira-ca-column-selector", "value"),
    Input("btn-load-ca", "n_clicks"),
    prevent_initial_call=True,
)
def load_shared_file_ca(n_clicks):
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
            "error",
            [],
            {"display": "none"},
            False,
            [],
            [],
        )

    try:
        df_json = get_prepared_base_json_ca(force_refresh=False)
        df = pd.read_json(StringIO(df_json), orient="split")

        filename = Path(SHAREPOINT_FILE_PATH_CA).name or "Carteira Ativa - CGs.xlsx"

        total_registros = len(df)
        total_colunas = len(df.columns)

        col_cd_pleito = next((c for c in ["cd_pleito", "CD_PLEITO"] if c in df.columns), None)
        n_pleitos_distintos = (
            df[col_cd_pleito].astype("string").dropna().nunique()
            if col_cd_pleito
            else 0
        )

        msg = (
            f"Arquivo carregado com sucesso.\n"
            f"Linhas: {fmt_int_br(total_registros)}  ·  "
            f"Pleitos distintos: {fmt_int_br(n_pleitos_distintos)}  ·  "
            f"Colunas: {fmt_int_br(total_colunas)}"
        )

        summary = [
            metric_card("Operações", fmt_int_br(total_registros), "linhas na base", ACCENT),
            metric_card("Projetos", fmt_int_br(n_pleitos_distintos), "pleitos distintos", BLUE),
            metric_card("Colunas", fmt_int_br(total_colunas), "variáveis disponíveis", TEAL),
            metric_card("Arquivo", filename, "base ativa", ROSE),
        ]

        col_options = [{"label": str(c), "value": str(c)} for c in df.columns]
        default_cols = list(df.columns)

        return (
            df_json,
            filename,
            [
                html.Div("Status", className="status-title"),
                html.Pre(msg, className="status-message"),
            ],
            "status-box status-success",
            "success",
            summary,
            {"display": "block"},
            True,
            col_options,
            default_cols,
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
            "error",
            [],
            {"display": "none"},
            False,
            [],
            [],
        )


@callback(
    Output("carteira-ca-select-de_fase", "value", allow_duplicate=True),
    Output("carteira-ca-select-de_tipo_operacao", "value", allow_duplicate=True),
    Output("carteira-ca-select-nm_proponente", "value", allow_duplicate=True),
    Output("carteira-ca-select-sg_fonte", "value", allow_duplicate=True),
    Output("carteira-ca-select-de_esfera", "value", allow_duplicate=True),
    Output("carteira-ca-select-nm_regiao", "value", allow_duplicate=True),
    Output("carteira-ca-select-nm_setor", "value", allow_duplicate=True),
    Output("carteira-ca-select-nm_subsetor", "value", allow_duplicate=True),
    Output("carteira-ca-select-sys", "value", allow_duplicate=True),
    Output("carteira-ca-select-nm_limite", "value", allow_duplicate=True),
    Output("carteira-ca-select-nm_cg", "value", allow_duplicate=True),
    Output("carteira-ca-select-nm_tecnico", "value", allow_duplicate=True),
    Output("carteira-ca-select-cd_pleito", "value", allow_duplicate=True),
    Input("carteira-ca-btn-clear-selections", "n_clicks", allow_optional=True),
    prevent_initial_call=True,
)
def clear_carteira_ca_selections(n_clicks):
    if not n_clicks:
        raise PreventUpdate
    return [], [], [], [], [], [], [], [], [], [], [], [], []


@callback(
    Output("summary-cards-ca", "children", allow_duplicate=True),
    Output("status-box-ca", "children", allow_duplicate=True),
    Output("status-box-ca", "className", allow_duplicate=True),
    Output("preview-table-ca-dados", "data"),
    Output("preview-table-ca-dados", "columns"),
    Output("preview-table-ca", "data"),
    Output("preview-table-ca", "columns"),
    Input("global-df-json-ca", "data"),
    Input("carteira-ca-select-de_fase", "value", allow_optional=True),
    Input("carteira-ca-select-de_tipo_operacao", "value", allow_optional=True),
    Input("carteira-ca-select-nm_proponente", "value", allow_optional=True),
    Input("carteira-ca-select-sg_fonte", "value", allow_optional=True),
    Input("carteira-ca-select-de_esfera", "value", allow_optional=True),
    Input("carteira-ca-select-nm_regiao", "value", allow_optional=True),
    Input("carteira-ca-select-nm_setor", "value", allow_optional=True),
    Input("carteira-ca-select-nm_subsetor", "value", allow_optional=True),
    Input("carteira-ca-select-sys", "value", allow_optional=True),
    Input("carteira-ca-select-nm_limite", "value", allow_optional=True),
    Input("carteira-ca-select-nm_cg", "value", allow_optional=True),
    Input("carteira-ca-select-nm_tecnico", "value", allow_optional=True),
    Input("carteira-ca-select-cd_pleito", "value", allow_optional=True),
    Input("carteira-ca-column-selector", "value", allow_optional=True),
    State("global-filename-ca", "data"),
    State("carteira-ca-loaded", "data"),
    prevent_initial_call=True,
)
def update_carteira_ca_views(
    df_json_ca,
    de_fase,
    de_tipo_operacao,
    nm_proponente,
    sg_fonte,
    de_esfera,
    nm_regiao,
    nm_setor,
    nm_subsetor,
    sys,
    nm_limite,
    nm_cg,
    nm_tecnico,
    cd_pleito,
    selected_columns,
    filename_ca,
    carteira_ca_loaded,
):
    if not carteira_ca_loaded or not df_json_ca:
        raise PreventUpdate

    df_full = pd.read_json(StringIO(df_json_ca), orient="split")
    df_filtrado = build_filtered_carteira_ca_df(
        df_json_ca,
        de_fase=de_fase,
        de_tipo_operacao=de_tipo_operacao,
        nm_proponente=nm_proponente,
        sg_fonte=sg_fonte,
        de_esfera=de_esfera,
        nm_regiao=nm_regiao,
        nm_setor=nm_setor,
        nm_subsetor=nm_subsetor,
        sys=sys,
        nm_limite=nm_limite,
        nm_cg=nm_cg,
        nm_tecnico=nm_tecnico,
        cd_pleito=cd_pleito,
    )

    total_registros = len(df_filtrado)
    total_colunas = len(df_full.columns)

    col_cd_pleito = next((c for c in ["cd_pleito", "CD_PLEITO"] if c in df_filtrado.columns), None)
    n_pleitos_distintos = (
        df_filtrado[col_cd_pleito].astype("string").dropna().nunique()
        if col_cd_pleito
        else 0
    )

    summary = [
        metric_card("Operações", fmt_int_br(total_registros), "linhas filtradas", ACCENT),
        metric_card("Projetos", fmt_int_br(n_pleitos_distintos), "pleitos distintos", BLUE),
        metric_card("Colunas", fmt_int_br(total_colunas), "variáveis disponíveis", TEAL),
        metric_card("Arquivo", filename_ca or "Carteira Ativa - CGs.xlsx", "base ativa", ROSE),
    ]

    msg = (
        f"Filtros aplicados com sucesso.\n"
        f"Linhas filtradas: {fmt_int_br(total_registros)}  ·  "
        f"Pleitos distintos: {fmt_int_br(n_pleitos_distintos)}  ·  "
        f"Colunas: {fmt_int_br(total_colunas)}"
    )

    preview_dados = df_filtrado.head(200).where(pd.notnull(df_filtrado.head(200)), None)

    if selected_columns:
        valid_cols = [c for c in selected_columns if c in df_filtrado.columns]
        df_tabela = df_filtrado[valid_cols] if valid_cols else df_filtrado.iloc[:, 0:0]
    else:
        df_tabela = df_filtrado.copy()

    df_tabela = df_tabela.where(pd.notnull(df_tabela), None)

    return (
        summary,
        [
            html.Div("Status", className="status-title"),
            html.Pre(msg, className="status-message"),
        ],
        "status-box status-success",
        preview_dados.to_dict("records"),
        [{"name": str(c), "id": str(c)} for c in preview_dados.columns],
        df_tabela.to_dict("records"),
        [{"name": str(c), "id": str(c)} for c in df_tabela.columns],
    )


@callback(
    Output("carteira-ca-selected-columns", "data"),
    Input("carteira-ca-column-selector", "value"),
    prevent_initial_call=True,
)
def sync_carteira_ca_columns(selected):
    return selected or []


@callback(
    Output("carteira-ca-column-selector", "value", allow_duplicate=True),
    Input("carteira-ca-btn-select-all", "n_clicks"),
    State("carteira-ca-column-selector", "options"),
    prevent_initial_call=True,
)
def select_all_carteira_ca_columns(n_clicks, options):
    if not n_clicks or not options:
        raise PreventUpdate
    return [o["value"] for o in options]


@callback(
    Output("carteira-ca-column-selector", "value", allow_duplicate=True),
    Input("carteira-ca-btn-clear-columns", "n_clicks"),
    prevent_initial_call=True,
)
def clear_carteira_ca_columns(n_clicks):
    if not n_clicks:
        raise PreventUpdate
    return []


@callback(
    Output("carteira-ca-selected-count", "children"),
    Input("carteira-ca-column-selector", "options"),
    Input("carteira-ca-column-selector", "value"),
)
def update_carteira_ca_selected_count(options, selected):
    total = len(options or [])
    sel = len(selected or [])
    return f"{fmt_int_br(sel)} de {fmt_int_br(total)} selecionadas" if total else ""


@callback(
    Output("carteira-ca-download-excel", "data"),
    Input("carteira-ca-btn-export", "n_clicks", allow_optional=True),
    State("carteira-ca-select-de_fase", "value", allow_optional=True),
    State("carteira-ca-select-de_tipo_operacao", "value", allow_optional=True),
    State("carteira-ca-select-nm_proponente", "value", allow_optional=True),
    State("carteira-ca-select-sg_fonte", "value", allow_optional=True),
    State("carteira-ca-select-de_esfera", "value", allow_optional=True),
    State("carteira-ca-select-nm_regiao", "value", allow_optional=True),
    State("carteira-ca-select-nm_setor", "value", allow_optional=True),
    State("carteira-ca-select-nm_subsetor", "value", allow_optional=True),
    State("carteira-ca-select-sys", "value", allow_optional=True),
    State("carteira-ca-select-nm_limite", "value", allow_optional=True),
    State("carteira-ca-select-nm_cg", "value", allow_optional=True),
    State("carteira-ca-select-nm_tecnico", "value", allow_optional=True),
    State("carteira-ca-select-cd_pleito", "value", allow_optional=True),
    State("carteira-ca-column-selector", "value", allow_optional=True),
    State("global-df-json-ca", "data"),
    State("carteira-ca-loaded", "data"),
    prevent_initial_call=True,
)
def export_carteira_ca_excel(
    n_clicks,
    de_fase,
    de_tipo_operacao,
    nm_proponente,
    sg_fonte,
    de_esfera,
    nm_regiao,
    nm_setor,
    nm_subsetor,
    sys,
    nm_limite,
    nm_cg,
    nm_tecnico,
    cd_pleito,
    selected_columns,
    df_json_ca,
    carteira_ca_loaded,
):
    if not n_clicks or not carteira_ca_loaded or not df_json_ca:
        raise PreventUpdate

    df = build_filtered_carteira_ca_df(
        df_json_ca,
        de_fase=de_fase,
        de_tipo_operacao=de_tipo_operacao,
        nm_proponente=nm_proponente,
        sg_fonte=sg_fonte,
        de_esfera=de_esfera,
        nm_regiao=nm_regiao,
        nm_setor=nm_setor,
        nm_subsetor=nm_subsetor,
        sys=sys,
        nm_limite=nm_limite,
        nm_cg=nm_cg,
        nm_tecnico=nm_tecnico,
        cd_pleito=cd_pleito,
    )

    if df.empty:
        raise PreventUpdate

    if selected_columns:
        valid_cols = [c for c in selected_columns if c in df.columns]
        df = df[valid_cols] if valid_cols else df.iloc[:, 0:0]

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="carteira_cgs")
    buf.seek(0)

    return dcc.send_bytes(buf.getvalue(), "carteira_ativa_cgs_filtrada.xlsx")

# ── CALLBACKS DA ABA ACOMPANHAMENTO ([CGs] CARTEIRA ATIVA) ───────────────────

@callback(
    Output("ca-acomp-kpi-tecnicos", "children"),
    Output("ca-acomp-kpi-ops", "children"),
    Output("ca-acomp-kpi-valor", "children"),
    Output("ca-acomp-kpi-cgs", "children"),
    Input("carteira-ca-tabs", "value"),
    Input("global-df-json-ca", "data"),
    *(Input(f"carteira-ca-select-{col}", "value", allow_optional=True) for col, _ in SELECT_FIELDS_CA),
    prevent_initial_call=True,
)
def update_acomp_kpis(tab, df_json_ca, *filter_values):
    # Só roda quando o usuário estiver na aba correta.
    if tab != "acompanhamento" or not df_json_ca:
        raise PreventUpdate

    df = _apply_ca_filters(df_json_ca, filter_values)

    if df.empty:
        empty = kpi_block("—", "—", "sem dados")
        return empty, empty, empty, empty

    n_tecnicos = df["nm_tecnico"].nunique() if "nm_tecnico" in df.columns else 0
    n_ops = len(df)
    val_fin = df["vl_financiamento_dolar"].sum() if "vl_financiamento_dolar" in df.columns else 0
    n_cgs = df["nm_cg"].nunique() if "nm_cg" in df.columns else 0

    return (
        kpi_block("Técnicos", fmt_int_br(n_tecnicos), "com operações na carteira"),
        kpi_block("Total de Operações", fmt_int_br(n_ops), "no filtro atual"),
        kpi_block("Financiamento Total", fmt_bi(val_fin), "soma do portfólio"),
        kpi_block("Coordenações Gerais", fmt_int_br(n_cgs), "CGs representadas"),
    )


@callback(
    Output("ca-acomp-fig-tecnico", "figure"),
    Input("carteira-ca-tabs", "value"),
    Input("global-df-json-ca", "data"),
    *(Input(f"carteira-ca-select-{col}", "value", allow_optional=True) for col, _ in SELECT_FIELDS_CA),
    prevent_initial_call=True,
)
def update_acomp_fig_tecnico(tab, df_json_ca, *filter_values):
    if tab != "acompanhamento" or not df_json_ca:
        raise PreventUpdate

    df = _apply_ca_filters(df_json_ca, filter_values)

    if df.empty or "nm_tecnico" not in df.columns or "vl_financiamento_dolar" not in df.columns:
        return EMPTY_FIG

    agg_dict = {
        "val": ("vl_financiamento_dolar", "sum"),
        "qtd": ("nm_tecnico", "size"),
        "proj": ("cd_pleito", pd.Series.nunique) if "cd_pleito" in df.columns else ("nm_tecnico", "size"),
    }

    grp = (
        df.groupby("nm_tecnico")
        .agg(**agg_dict)
        .reset_index()
        .rename(columns={"nm_tecnico": "tecnico"})
    )

    if grp.empty:
        return EMPTY_FIG

    grp = grp.nlargest(15, "val").sort_values("val", ascending=False)

    grp["val_mi"] = grp["val"] / 1e6
    grp["val_fmt"] = grp["val_mi"].apply(lambda x: f"{brazil_vlr(x, 0)}M")
    grp["qtd_fmt"] = grp["qtd"].apply(fmt_int_br)
    grp["proj_fmt"] = grp["proj"].apply(fmt_int_br)

    fig = go.Figure(
        go.Bar(
            x=grp["val_mi"],
            y=grp["tecnico"],
            orientation="h",
            marker_color=[COLOR_SEQUENCE[i % len(COLOR_SEQUENCE)] for i in range(len(grp))],
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
            ticksuffix="M",
        ),
        yaxis=merge_dict(YAXIS_DEF, showgrid=False, automargin=True, autorange="reversed"),
        margin=dict(t=20, r=80, b=40, l=190),
    )
    return fig


@callback(
    Output("ca-acomp-fig-cg", "figure"),
    Input("carteira-ca-tabs", "value"),
    Input("global-df-json-ca", "data"),
    *(Input(f"carteira-ca-select-{col}", "value", allow_optional=True) for col, _ in SELECT_FIELDS_CA),
    prevent_initial_call=True,
)
def update_acomp_fig_cg(tab, df_json_ca, *filter_values):
    if tab != "acompanhamento" or not df_json_ca:
        raise PreventUpdate

    df = _apply_ca_filters(df_json_ca, filter_values)

    if df.empty or "nm_cg" not in df.columns or "vl_financiamento_dolar" not in df.columns:
        return EMPTY_FIG

    grp = (
        df.groupby("nm_cg")
        .agg(
            val=("vl_financiamento_dolar", "sum"),
            valor_fin=("vl_financiamento_dolar", "sum"),
            qtd=("nm_cg", "size"),
            proj=("cd_pleito", pd.Series.nunique) if "cd_pleito" in df.columns else ("nm_cg", "size"),
        )
        .reset_index()
        .rename(columns={"nm_cg": "cg"})
    )

    if grp.empty:
        return EMPTY_FIG

    grp = grp.sort_values("val", ascending=False).copy()
    grp["valor_fmt"] = (grp["valor_fin"] / 1e6).apply(lambda x: f"{brazil_vlr(x, 0)}M")
    grp["qtd_fmt"] = grp["qtd"].apply(fmt_int_br)
    grp["proj_fmt"] = grp["proj"].apply(fmt_int_br)

    customdata = grp[["valor_fmt", "qtd_fmt", "proj_fmt"]].values

    fig = go.Figure(
        go.Pie(
            labels=grp["cg"],
            values=grp["val"],
            hole=0.62,
            marker=dict(
                colors=[COLOR_SEQUENCE[i % len(COLOR_SEQUENCE)] for i in range(len(grp))],
                line=dict(color="white", width=2),
            ),
            textinfo="label+percent",
            hovertext=[
                (
                    f"Valor de financiamento: US$ {valor}<br>"
                    f"Quantidade de operações: {qtd}<br>"
                    f"Quantidade de projetos: {proj}"
                )
                for valor, qtd, proj in zip(grp["valor_fmt"], grp["qtd_fmt"], grp["proj_fmt"])
            ],
            hovertemplate="<b>%{label}</b><br>%{hovertext}<br>Percentual: %{percent}<extra></extra>",
        )
    )
    
    apply_layout(
        fig,
        margin=dict(t=20, r=130, b=20, l=20),
        legend=dict(orientation="v", x=1.02, y=0.5, font=dict(size=11)),
    )
    return fig

@callback(
    Output("ca-acomp-fig-fase-tecnico", "figure"),
    Input("carteira-ca-tabs", "value"),
    Input("global-df-json-ca", "data"),
    *(Input(f"carteira-ca-select-{col}", "value", allow_optional=True) for col, _ in SELECT_FIELDS_CA),
    prevent_initial_call=True,
)
def update_acomp_fig_fase_tecnico(tab, df_json_ca, *filter_values):
    if tab != "acompanhamento" or not df_json_ca:
        raise PreventUpdate

    df = _apply_ca_filters(df_json_ca, filter_values)

    if df.empty or "nm_tecnico" not in df.columns or "de_fase" not in df.columns:
        return EMPTY_FIG

    # Top técnicos por valor total em carteira
    if "vl_financiamento_dolar" in df.columns:
        top_tecnicos = (
            df.groupby("nm_tecnico", dropna=False)["vl_financiamento_dolar"]
            .sum()
            .nlargest(20)
            .index.tolist()
        )
    else:
        top_tecnicos = (
            df.groupby("nm_tecnico", dropna=False)
            .size()
            .nlargest(20)
            .index.tolist()
        )

    df_top = df[df["nm_tecnico"].isin(top_tecnicos)].copy()

    # Agregado total por técnico
    total_por_tecnico = (
        df_top.groupby("nm_tecnico", dropna=False)
        .agg(
            qtd_total=("de_fase", "size"),
            val_total=("vl_financiamento_dolar", "sum") if "vl_financiamento_dolar" in df_top.columns else ("de_fase", "size"),
        )
        .reset_index()
    )
    total_por_tecnico["qtd_total_fmt"] = total_por_tecnico["qtd_total"].apply(fmt_int_br)
    total_por_tecnico["val_total_fmt"] = (total_por_tecnico["val_total"] / 1e6).apply(
        lambda x: f"{brazil_vlr(x, 0)}M"
    )

    pivot = (
        df_top.groupby(["nm_tecnico", "de_fase"], dropna=False)
        .agg(
            qtd=("de_fase", "size"),
            val=("vl_financiamento_dolar", "sum") if "vl_financiamento_dolar" in df_top.columns else ("de_fase", "size"),
        )
        .reset_index()
        .merge(
            total_por_tecnico[["nm_tecnico", "qtd_total_fmt", "val_total_fmt"]],
            on="nm_tecnico",
            how="left",
        )
    )

    if pivot.empty:
        return EMPTY_FIG

    def primeiro_ultimo_nome(txt: str) -> str:
        partes = str(txt).strip().split()
        if not partes:
            return str(txt)
        if len(partes) == 1:
            return partes[0]
        return f"{partes[0]}<br>{partes[-1]}"

    # Ordem dos técnicos por maior valor total
    if "vl_financiamento_dolar" in df_top.columns:
        ordem_tecnicos = (
            df_top.groupby("nm_tecnico", dropna=False)["vl_financiamento_dolar"]
            .sum()
            .sort_values(ascending=False)
            .index.tolist()
        )
    else:
        ordem_tecnicos = (
            df_top.groupby("nm_tecnico", dropna=False)
            .size()
            .sort_values(ascending=False)
            .index.tolist()
        )

    pivot["nm_tecnico_label"] = pivot["nm_tecnico"].apply(primeiro_ultimo_nome)
    pivot["val_fmt"] = (pivot["val"] / 1e6).apply(lambda x: f"{brazil_vlr(x, 0)}M")

    ordem_fases_desejada = [
        "Aprovado",
        "Em negociação",
        "Negociação concluída",
        "Em execução",
        "Repagamento",
        "Arquivado",
        "Finalizado",
    ]

    fases_existentes = pivot["de_fase"].astype(str).tolist()
    fases_unicas_existentes = list(dict.fromkeys(fases_existentes))

    fases_unicas = [f for f in ordem_fases_desejada if f in fases_unicas_existentes]
    fases_unicas += [f for f in fases_unicas_existentes if f not in fases_unicas]

    fig = go.Figure()
    for i, fase in enumerate(fases_unicas):
        sub = pivot[pivot["de_fase"] == fase].copy()
        sub["nm_tecnico"] = pd.Categorical(sub["nm_tecnico"], categories=ordem_tecnicos, ordered=True)
        sub = sub.sort_values("nm_tecnico")

        fig.add_trace(
            go.Bar(
                name=fase,
                x=sub["nm_tecnico_label"],
                y=sub["qtd"],
                marker_color=COLOR_SEQUENCE[i % len(COLOR_SEQUENCE)],
                marker_line_width=0,
                customdata=list(
                    zip(
                        sub["nm_tecnico"].astype(str),
                        sub["de_fase"],
                        sub["qtd"].apply(fmt_int_br),
                        sub["val_fmt"],
                        sub["qtd_total_fmt"],
                        sub["val_total_fmt"],
                    )
                ),
                hovertemplate=(
                    "<b>Técnico:</b> %{customdata[0]}<br>"
                    "<b>Fase:</b> %{customdata[1]}<br>"
                    "<b>Quantidade de operações na fase:</b> %{customdata[2]}<br>"
                    "<b>Valor de financiamento na fase:</b> US$ %{customdata[3]}<br>"
                    "<b>Quantidade total de operações do técnico:</b> %{customdata[4]}<br>"
                    "<b>Valor total em carteira do técnico:</b> US$ %{customdata[5]}"
                    "<extra></extra>"
                ),
            )
        )

    apply_layout(
        fig,
        xaxis=merge_dict(
            XAXIS_DEF,
            tickangle=0,
            automargin=True,
        ),
        yaxis=YAXIS_DEF,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            font=dict(size=10),
        ),
        margin=dict(t=30, r=20, b=110, l=50),
        barmode="stack",
    )
    return fig

@callback(
    Output("ca-acomp-fig-setor-cg", "figure"),
    Input("carteira-ca-tabs", "value"),
    Input("global-df-json-ca", "data"),
    *(Input(f"carteira-ca-select-{col}", "value", allow_optional=True) for col, _ in SELECT_FIELDS_CA),
    prevent_initial_call=True,
)
def update_acomp_fig_setor_cg(tab, df_json_ca, *filter_values):
    if tab != "acompanhamento" or not df_json_ca:
        raise PreventUpdate

    df = _apply_ca_filters(df_json_ca, filter_values)

    if df.empty or "nm_cg" not in df.columns or "nm_setor" not in df.columns:
        return EMPTY_FIG

    grp = (
        df.groupby(["nm_cg", "nm_setor"], dropna=False)
        .agg(
            qtd=("nm_setor", "size"),
            val=("vl_financiamento_dolar", "sum") if "vl_financiamento_dolar" in df.columns else ("nm_setor", "size"),
            proj=("cd_pleito", pd.Series.nunique) if "cd_pleito" in df.columns else ("nm_setor", "size"),
        )
        .reset_index()
    )

    if grp.empty:
        return EMPTY_FIG

    # Ordena as CGs pelo maior valor total financiado
    ordem_cg = (
        grp.groupby("nm_cg", dropna=False)["val"]
        .sum()
        .sort_values(ascending=False)
        .index.tolist()
    )

    # Ordena os setores pelo maior valor total financiado
    setores_unicos = (
        grp.groupby("nm_setor", dropna=False)["val"]
        .sum()
        .sort_values(ascending=False)
        .index.tolist()
    )

    fig = go.Figure()
    for i, setor in enumerate(setores_unicos):
        sub = grp[grp["nm_setor"] == setor].copy()
        sub["nm_cg"] = pd.Categorical(sub["nm_cg"], categories=ordem_cg, ordered=True)
        sub = sub.sort_values("nm_cg")

        sub["val_mi"] = sub["val"] / 1e6
        sub["val_fmt"] = sub["val_mi"].apply(lambda x: f"{brazil_vlr(x, 0)}M")
        sub["qtd_fmt"] = sub["qtd"].apply(fmt_int_br)
        sub["proj_fmt"] = sub["proj"].apply(fmt_int_br)

        fig.add_trace(
            go.Bar(
                name=setor,
                x=sub["val_mi"],
                y=sub["nm_cg"].astype(str),
                orientation="h",
                marker_color=COLOR_SEQUENCE[i % len(COLOR_SEQUENCE)],
                marker_line_width=0,
                customdata=list(
                    zip(
                        sub["nm_cg"].astype(str),
                        sub["nm_setor"],
                        sub["val_fmt"],
                        sub["qtd_fmt"],
                        sub["proj_fmt"],
                    )
                ),
                hovertemplate=(
                    "<b>%{customdata[0]}</b><br>"
                    "Setor: %{customdata[1]}<br>"
                    "Valor de financiamento: US$ %{customdata[2]}<br>"
                    "Quantidade de operações: %{customdata[3]}<br>"
                    "Quantidade de projetos: %{customdata[4]}"
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
            ticksuffix="M",
            title=None,
            showticklabels=False,
        ),
        yaxis=merge_dict(
            YAXIS_DEF,
            showgrid=False,
            automargin=True,
            autorange="reversed",
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            font=dict(size=10),
        ),
        margin=dict(t=30, r=20, b=40, l=120),
        barmode="stack",
    )
    return fig

# ── CALLBACKS DA ABA VALIDADE DA RESOLUÇÃO ([CGs] CARTEIRA ATIVA) ───────────

@callback(
    Output("ca-res-kpi-vencidas", "children"),
    Output("ca-res-kpi-30dias", "children"),
    Output("ca-res-kpi-90dias", "children"),
    Output("ca-res-kpi-180dias", "children"),
    Output("ca-res-kpi-ok", "children"),
    Input("carteira-ca-tabs", "value"),
    Input("ca-res-horizonte", "value"),
    Input("global-df-json-ca", "data"),
    *(Input(f"carteira-ca-select-{col}", "value", allow_optional=True) for col, _ in SELECT_FIELDS_CA),
    prevent_initial_call=True,
)
def update_res_kpis(tab, horizonte, df_json_ca, *filter_values):
    if tab != "resolucao" or not df_json_ca:
        raise PreventUpdate

    df = _apply_ca_filters(df_json_ca, filter_values)

    if "de_fase" in df.columns:
        df = df[df["de_fase"].fillna("").astype(str).str.strip().str.lower().eq("aprovado")]

    if df.empty or "dt_validade_recomendacao" not in df.columns:
        empty = kpi_block("—", "—", "sem dados")
        return empty, empty, empty, empty, empty

    df = df.copy()
    df["dt_validade_recomendacao"] = pd.to_datetime(
        df["dt_validade_recomendacao"], errors="coerce"
    )
    df = df[df["dt_validade_recomendacao"].notna()]

    if df.empty:
        empty = kpi_block("—", "—", "sem dados")
        return empty, empty, empty, empty, empty

    hoje = pd.Timestamp(date.today())
    df["dias"] = (df["dt_validade_recomendacao"] - hoje).dt.days

    n_vencidas = int((df["dias"] < 0).sum())
    n_30 = int(((df["dias"] >= 0) & (df["dias"] < 30)).sum())
    n_90 = int(((df["dias"] >= 0) & (df["dias"] < 90)).sum())
    n_180 = int(((df["dias"] >= 0) & (df["dias"] < 180)).sum())
    n_ok = int((df["dias"] >= 180).sum())

    return (
        kpi_block("Vencidas", fmt_int_br(n_vencidas), "resolução vencida"),
        kpi_block("Vencem em 30 dias", fmt_int_br(n_30), "atenção imediata"),
        kpi_block("Vencem em 90 dias", fmt_int_br(n_90), "atenção moderada"),
        kpi_block("Vencem em 180 dias", fmt_int_br(n_180), "monitorar"),
        kpi_block("Dentro do prazo", fmt_int_br(n_ok), "validade > 180 dias"),
    )


@callback(
    Output("ca-res-fig-timeline", "figure"),
    Input("carteira-ca-tabs", "value"),
    Input("ca-res-horizonte", "value"),
    Input("global-df-json-ca", "data"),
    *(Input(f"carteira-ca-select-{col}", "value", allow_optional=True) for col, _ in SELECT_FIELDS_CA),
    prevent_initial_call=True,
)

def update_res_timeline(tab, horizonte, df_json_ca, *filter_values):
    if tab != "resolucao" or not df_json_ca:
        raise PreventUpdate

    df = _prep_resolucao_df(df_json_ca, filter_values, horizonte)

    if df.empty:
        return EMPTY_FIG

    df = df.copy()
    df["dt_validade_recomendacao"] = pd.to_datetime(df["dt_validade_recomendacao"], errors="coerce")
    df = df[df["dt_validade_recomendacao"].notna()]

    if df.empty:
        return EMPTY_FIG

    # primeiro dia do mês de cada validade
    df["mes_ref"] = df["dt_validade_recomendacao"].dt.to_period("M").dt.to_timestamp()

    grp = df.groupby("mes_ref").size().reset_index(name="qtd").sort_values("mes_ref")
    if grp.empty:
        return EMPTY_FIG

    # preenche meses sem ocorrência para a linha do tempo ficar contínua
    mes_inicio = grp["mes_ref"].min()
    mes_fim = grp["mes_ref"].max()
    faixa_meses = pd.date_range(start=mes_inicio, end=mes_fim, freq="MS")

    grp = (
        pd.DataFrame({"mes_ref": faixa_meses})
        .merge(grp, on="mes_ref", how="left")
        .fillna({"qtd": 0})
    )
    grp["qtd"] = grp["qtd"].astype(int)

    meses_pt = {
        1: "jan", 2: "fev", 3: "mar", 4: "abr", 5: "mai", 6: "jun",
        7: "jul", 8: "ago", 9: "set", 10: "out", 11: "nov", 12: "dez",
    }

    grp["ano"] = grp["mes_ref"].dt.year.astype(str)
    grp["mes_nome"] = grp["mes_ref"].dt.month.map(meses_pt)
    grp["mes_label"] = grp["mes_ref"].apply(lambda d: f"{meses_pt[d.month]}<br>{d.year}")

    mes_atual = pd.Timestamp(date.today()).replace(day=1)
    cores = [ROSE if d <= mes_atual else ACCENT for d in grp["mes_ref"]]

    fig = go.Figure(
        go.Bar(
            x=grp["mes_label"],
            y=grp["qtd"],
            marker_color=cores,
            marker_line_width=0,
            customdata=grp[["ano", "mes_nome"]].to_numpy(),
            hovertemplate=(
                "Ano: %{customdata[0]}<br>"
                "Mês: %{customdata[1]}<br>"
                "Quantidade de operações: %{y}"
                "<extra></extra>"
            ),
        )
    )

    apply_layout(
        fig,
        xaxis=merge_dict(
            XAXIS_DEF,
            type="category",
            tickangle=0,
            automargin=True,
        ),
        yaxis=merge_dict(YAXIS_DEF, dtick=1, title=None),
        margin=dict(t=20, r=20, b=80, l=50),
    )
    return fig


@callback(
    Output("ca-res-table", "data"),
    Output("ca-res-table", "columns"),
    Output("ca-res-selected-count", "children"),
    Input("carteira-ca-tabs", "value"),
    Input("ca-res-horizonte", "value"),
    Input("ca-res-column-selector", "value"),
    Input("global-df-json-ca", "data"),
    *(Input(f"carteira-ca-select-{col}", "value", allow_optional=True) for col, _ in SELECT_FIELDS_CA),
    prevent_initial_call=True,
)
def update_res_table(tab, horizonte, selected_cols, df_json_ca, *filter_values):
    if tab != "resolucao" or not df_json_ca:
        raise PreventUpdate

    df = _prep_resolucao_df(df_json_ca, filter_values, horizonte)

    if df.empty:
        return [], [], "0 de 0 colunas"

    selected_cols = list(selected_cols or [])

    # Garante que a coluna calculada fique disponível na tabela.
    if "dias_para_vencer" not in selected_cols:
        selected_cols.append("dias_para_vencer")

    valid_cols = [c for c in selected_cols if c in df.columns]
    df_out = df[valid_cols].copy()

    for col in df_out.columns:
        if str(col).startswith("dt_"):
            df_out[col] = pd.to_datetime(df_out[col], errors="coerce").dt.strftime("%d/%m/%Y")
            df_out[col] = df_out[col].where(df_out[col].notna(), None)

    df_out = df_out.where(pd.notnull(df_out), None)

    columns = [{"name": str(c), "id": str(c)} for c in df_out.columns]
    count_text = f"{fmt_int_br(len(valid_cols))} colunas exibidas"

    return df_out.to_dict("records"), columns, count_text


@callback(
    Output("ca-res-download-excel", "data"),
    Input("ca-res-btn-export", "n_clicks"),
    State("ca-res-horizonte", "value"),
    State("ca-res-column-selector", "value"),
    State("global-df-json-ca", "data"),
    *(State(f"carteira-ca-select-{col}", "value", allow_optional=True) for col, _ in SELECT_FIELDS_CA),
    prevent_initial_call=True,
)
def export_res_excel(n_clicks, horizonte, selected_cols, df_json_ca, *filter_values):
    if not n_clicks or not df_json_ca:
        raise PreventUpdate

    df = _prep_resolucao_df(df_json_ca, filter_values, horizonte)

    if df.empty:
        raise PreventUpdate

    selected_cols = list(selected_cols or [])
    if "dias_para_vencer" not in selected_cols:
        selected_cols.append("dias_para_vencer")

    valid_cols = [c for c in selected_cols if c in df.columns]
    df_out = df[valid_cols].copy()

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df_out.to_excel(w, index=False, sheet_name="validade_resolucao")
    buf.seek(0)

    return dcc.send_bytes(buf.getvalue(), "validade_resolucao.xlsx")

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
def update_carteira_main(
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

    df = build_filtered_carteira_df(
        df_json,
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
    )

    if df is None or df.empty:
        empty_kpi = kpi_block("—", "—", "sem dados")
        return (EMPTY_FIG,) * 8 + (empty_kpi,) * 5

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
        kpi_block("Total de Projetos", fmt_int_br(n_proj), "projetos distintos"),
        kpi_block("Total de Operações", fmt_int_br(n_ops), "operações exibidas"),
        kpi_block("Financiamento Total", fmt_bi(val_fin), "soma em dólares"),
        kpi_block("Contrapartida Total", fmt_bi(val_contra), "soma em dólares"),
        kpi_block("Ano com Mais Aprovações", ano_rec, f"{fmt_int_br(ano_rec_n)} operações"),
    )


@callback(
    Output("carteira-fig-mapa-uf", "figure"),
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
def update_carteira_map(
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

    df = build_filtered_carteira_df(
        df_json,
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
    )

    if df is None or df.empty:
        return EMPTY_FIG

    return chart_mapa_uf(df, metrica)


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

    df = build_filtered_carteira_df(
        df_json,
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
    )

    if df is None or df.empty:
        return [], []

    if selected_columns:
        valid_cols = [c for c in selected_columns if c in df.columns]
        df = df[valid_cols] if valid_cols else df.iloc[:, 0:0]

    df_out = df.copy().where(pd.notnull(df), None)
    return df_out.to_dict("records"), [{"name": str(c), "id": str(c)} for c in df_out.columns]


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

    df = build_filtered_carteira_df(
        df_json,
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
    app.run(host="localhost", port=8050, debug=True, use_reloader=False)