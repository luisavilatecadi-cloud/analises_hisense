"""
CENTRAL DE ANÁLISE DE TRANSFERÊNCIAS DE ESTOQUE
------------------------------------------------
Regra de negócio:
  - A chave da movimentação é a coluna CODIGO.
  - Cada CODIGO possui linhas de SAIDA (origem) e ENTRADA (destino).
  - Tipo Estoque origem <> destino  -> TROCA DE TIPO DE ESTOQUE
  - Tipo Estoque igual, local diferente -> TRANSFERÊNCIA DE ENDEREÇO
  - Só um lado (sem par)            -> INCOMPLETO (auditoria)
  - Tipo e local iguais             -> SEM ALTERAÇÃO (auditoria)

Execução:  streamlit run app_transferencias_estoque.py
Base: upload manual (.xlsx / .csv), processada na hora. Nenhum dado é gravado.
Login opcional: só é exigido se existir [credenciais] no .streamlit/secrets.toml
"""

import hmac
from io import BytesIO
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# =========================================================
# 1. CONFIGURAÇÃO E ESTILO (padrão TECADI)
# =========================================================
st.set_page_config(page_title="TECADI - Transferências de Estoque", page_icon="🔄", layout="wide")

AZUL_ESCURO = "#133A68"
AZUL_TECADI = "#1D569B"
AZUL_CLARO_TECADI = "#009FE3"
CINZA_FUNDO = "#F8FAFC"
LARANJA = "#FF7F0E"
VERMELHO = "#DC2626"
VERDE = "#16A34A"
CINZA = "#CBD5E0"

CLS_TROCA = "Troca de Tipo de Estoque"
CLS_ENDER = "Transferência de Endereço"
CLS_INCOMP = "Incompleto (sem par)"
CLS_SEMALT = "Sem alteração"

COR_CLASSE = {
    CLS_TROCA: AZUL_ESCURO,
    CLS_ENDER: AZUL_CLARO_TECADI,
    CLS_INCOMP: VERMELHO,
    CLS_SEMALT: CINZA,
}


def formatar_br(valor):
    try:
        return f"{valor:,.0f}".replace(",", ".")
    except Exception:
        return str(valor)


st.markdown(f"""
    <style>
    .stApp {{ background-color: #FFFFFF; }}
    .dash-header {{ border-left: 10px solid {AZUL_TECADI}; padding-left: 20px; margin-bottom: 20px; }}
    .dash-title {{ color: {AZUL_ESCURO}; font-size: 38px !important; font-weight: 800 !important; }}
    .dash-sub {{ color: #64748B; font-size: 14px; margin-top: -6px; }}

    [data-testid="stSidebar"] {{ background: linear-gradient(180deg, {AZUL_ESCURO} 0%, {AZUL_TECADI} 100%); }}
    .sidebar-label {{ color: white !important; font-size: 0.8rem !important; font-weight: 700 !important; margin-top: 20px; margin-bottom: 10px; display: block; }}
    [data-testid="stSidebar"] label p {{ color: white !important; }}
    [data-testid="stSidebar"] .stCaption, [data-testid="stSidebar"] small {{ color: rgba(255,255,255,0.75) !important; }}

    div[data-testid="stMetric"] {{
        background-color: {CINZA_FUNDO} !important;
        border: 1px solid #E2E8F0 !important;
        padding: 15px !important;
        border-radius: 12px !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05) !important;
    }}
    div[data-testid="stMetric"]:hover {{ transform: translateY(-5px) !important; border-color: {AZUL_CLARO_TECADI} !important; }}
    div[data-testid="stMetric"] label p {{ color: {AZUL_ESCURO} !important; font-size: 16px !important; font-weight: 600 !important; }}
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {{ color: {AZUL_TECADI} !important; font-size: 28px !important; font-weight: 800 !important; }}

    div[data-baseweb="tooltip"] {{ display: none !important; visibility: hidden !important; }}
    div[data-baseweb="slider"] > div > div {{ background: rgba(255, 255, 255, 0.2) !important; }}
    div[data-baseweb="slider"] > div > div > div {{ background: #FFFFFF !important; }}
    div[role="slider"] {{ background-color: {AZUL_CLARO_TECADI} !important; border: 2px solid white !important; }}
    div[data-baseweb="slider"] div {{ color: white !important; font-weight: 600 !important; }}
    </style>
""", unsafe_allow_html=True)

# =========================================================
# 2. LOGIN OPCIONAL (só se houver [credenciais] em secrets)
# =========================================================
def _credenciais():
    try:
        return dict(st.secrets.get("credenciais", {}))
    except Exception:
        return {}


_cred = _credenciais()
if _cred:
    if not st.session_state.get("autenticado", False):
        st.title("🔒 Login")
        u = st.text_input("Usuário")
        s = st.text_input("Senha", type="password")
        if st.button("Entrar"):
            d = _cred.get(u)
            if d and hmac.compare_digest(str(s), str(d.get("senha", ""))):
                st.session_state.autenticado = True
                st.session_state.usuario_logado = u
                st.rerun()
            else:
                st.error("Usuário ou senha inválidos.")
        st.stop()

# =========================================================
# 3. MOTOR DE DADOS
# =========================================================
COLS_OBRIG = ["Codigo", "Data", "Usuario", "Tipo Mov", "Produto", "Quantidade", "Localizacao", "Tipo Estoque"]
COLS_OPC = ["Hora", "Filial", "Local", "Serial", "Lote", "Observacao", "Cod.Cliente", "Loja"]


def _normaliza_colunas(df):
    """Remove espaços e iguala nomes ignorando caixa/acentos básicos."""
    def chave(c):
        c = str(c).strip().lower()
        for a, b in (("ç", "c"), ("ã", "a"), ("á", "a"), ("é", "e"), ("í", "i"), ("ó", "o"), ("ú", "u"), ("â", "a"), ("ê", "e")):
            c = c.replace(a, b)
        return c
    alvo = {chave(c): c for c in COLS_OBRIG + COLS_OPC}
    df.columns = [alvo.get(chave(c), str(c).strip()) for c in df.columns]
    return df


@st.cache_data(show_spinner="Processando base...")
def carregar(arquivo_bytes: bytes, nome: str):
    """Lê a base e consolida 1 linha por CODIGO. Retorna (df_det, df_cod, erro)."""
    try:
        if nome.lower().endswith(".csv"):
            df = pd.read_csv(BytesIO(arquivo_bytes), sep=None, engine="python")
        else:
            df = pd.read_excel(BytesIO(arquivo_bytes))
        df = _normaliza_colunas(df)

        faltando = [c for c in COLS_OBRIG if c not in df.columns]
        if faltando:
            return None, None, f"Colunas obrigatórias ausentes: {', '.join(faltando)}"

        for c in COLS_OPC:
            if c not in df.columns:
                df[c] = np.nan

        df = df.dropna(subset=["Codigo"]).copy()
        df["Data"] = pd.to_datetime(df["Data"], errors="coerce", dayfirst=True)
        df = df.dropna(subset=["Data"])
        hora = pd.to_timedelta(df["Hora"].astype(str), errors="coerce").fillna(pd.Timedelta(0))
        df["Timestamp"] = df["Data"].dt.normalize() + hora
        df["Quantidade"] = pd.to_numeric(df["Quantidade"], errors="coerce").fillna(0)
        df["Tipo Estoque"] = pd.to_numeric(df["Tipo Estoque"], errors="coerce")
        df["Tipo Mov"] = df["Tipo Mov"].astype(str).str.strip().str.upper()
        df["Localizacao"] = df["Localizacao"].astype(str).str.strip()
        df["Produto"] = df["Produto"].astype(str).str.strip()
        df["Usuario"] = df["Usuario"].astype(str).str.strip()
        df["Papel"] = np.select(
            [df["Tipo Mov"].str.contains("SAIDA|SAÍDA", regex=True), df["Tipo Mov"].str.contains("ENTRADA")],
            ["ORIGEM", "DESTINO"], default="OUTRO")
        df = df.sort_values(["Codigo", "Timestamp"]).reset_index(drop=True)

        # ---- Consolidação por CODIGO -------------------------------------
        base = df.groupby("Codigo").agg(
            Timestamp=("Timestamp", "min"),
            Usuario=("Usuario", "first"),
            Produto=("Produto", "first"),
            Serial=("Serial", "first"),
            Observacao=("Observacao", "first"),
            Filial=("Filial", "first"),
            Linhas=("Codigo", "size"),
            Produtos_Distintos=("Produto", "nunique"),
        )
        org = df[df["Papel"] == "ORIGEM"].groupby("Codigo").agg(
            Tipo_Origem=("Tipo Estoque", "first"), Local_Origem=("Localizacao", "first"), Qtd_Saida=("Quantidade", "sum"))
        dst = df[df["Papel"] == "DESTINO"].groupby("Codigo").agg(
            Tipo_Destino=("Tipo Estoque", "first"), Local_Destino=("Localizacao", "first"), Qtd_Entrada=("Quantidade", "sum"))
        cod = base.join(org).join(dst).reset_index()
        cod[["Qtd_Saida", "Qtd_Entrada"]] = cod[["Qtd_Saida", "Qtd_Entrada"]].fillna(0)
        cod["Qtd_Mov"] = np.maximum(cod["Qtd_Saida"], cod["Qtd_Entrada"])

        tem_par = cod["Tipo_Origem"].notna() & cod["Tipo_Destino"].notna()
        cod["Classificacao"] = np.select(
            [~tem_par,
             cod["Tipo_Origem"] != cod["Tipo_Destino"],
             cod["Local_Origem"] != cod["Local_Destino"]],
            [CLS_INCOMP, CLS_TROCA, CLS_ENDER], default=CLS_SEMALT)
        cod["Divergencia_Qtd"] = tem_par & (cod["Qtd_Saida"] != cod["Qtd_Entrada"])
        cod["Data"] = cod["Timestamp"].dt.normalize()
        cod["Hora_Dia"] = cod["Timestamp"].dt.hour
        return df, cod, None
    except Exception as e:  # noqa
        return None, None, str(e)


def rotulo_tipo(v, mapa):
    if pd.isna(v):
        return "—"
    v = int(v)
    return f"{v} - {mapa[v]}" if v in mapa else str(v)


def parse_mapa(txt):
    mapa = {}
    for linha in (txt or "").splitlines():
        if "=" in linha:
            k, v = linha.split("=", 1)
            if k.strip().isdigit():
                mapa[int(k.strip())] = v.strip()
    return mapa


def to_excel(abas: dict) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for nome, d in abas.items():
            d.to_excel(w, sheet_name=nome[:31], index=False)
            ws = w.sheets[nome[:31]]
            for i, col in enumerate(d.columns, start=1):
                mx = d[col].astype(str).fillna("").str.len().max() if len(d) else 0
                larg = min(40, max(len(str(col)), int(mx) if pd.notna(mx) else 0) + 2)
                ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = larg
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
    return buf.getvalue()


def estilo(fig, h=None):
    fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="white", margin=dict(t=40, b=30, l=0, r=0))
    if h:
        fig.update_layout(height=h)
    return fig


# =========================================================
# 4. UPLOAD
# =========================================================
st.markdown(
    '<div class="dash-header"><div class="dash-title">Central de Análise de Transferências de Estoque</div>'
    '<div class="dash-sub">Mapeamento de trocas de tipo de estoque por CODIGO de movimentação</div></div>',
    unsafe_allow_html=True)

with st.sidebar:
    st.image("https://tecadi.com.br/wp-content/uploads/2024/01/LOGO-HORIZONTAL_BRANCA_p.png.webp", width=200)
    st.markdown("---")
    st.markdown('<span class="sidebar-label">📂 BASE DE MOVIMENTAÇÕES</span>', unsafe_allow_html=True)
    arq = st.file_uploader("Base", type=["xlsx", "csv"], label_visibility="collapsed")

if arq is None:
    st.info("📂 Carregue a base de movimentações (.xlsx ou .csv) na barra lateral para iniciar a análise.")
    st.markdown(
        "**Colunas obrigatórias:** `Codigo`, `Data`, `Usuario`, `Tipo Mov`, `Produto`, `Quantidade`, `Localizacao`, `Tipo Estoque`  \n"
        "**Opcionais:** `Hora`, `Filial`, `Local`, `Serial`, `Lote`, `Observacao`, `Cod.Cliente`, `Loja`")
    st.stop()

df_det_base, df_cod_base, erro = carregar(arq.getvalue(), arq.name)
if erro:
    st.error(f"❌ Não foi possível processar a base: {erro}")
    st.stop()
if df_cod_base.empty:
    st.warning("A base não possui registros válidos.")
    st.stop()

# =========================================================
# 5. FILTROS (SIDEBAR)
# =========================================================
with st.sidebar:
    with st.expander("🏷️ Dicionário de Tipos de Estoque"):
        st.caption("Um por linha: código=descrição (ex.: 26=Avaria)")
        mapa = parse_mapa(st.text_area("Mapa", value="", height=110, label_visibility="collapsed", key="mapa_tipos"))

    cod_all = df_cod_base.copy()
    cod_all["Origem"] = cod_all["Tipo_Origem"].apply(lambda v: rotulo_tipo(v, mapa))
    cod_all["Destino"] = cod_all["Tipo_Destino"].apply(lambda v: rotulo_tipo(v, mapa))
    cod_all["Fluxo"] = cod_all["Origem"] + "  →  " + cod_all["Destino"]

    st.markdown('<span class="sidebar-label">📅 INTERVALO DE TEMPO</span>', unsafe_allow_html=True)
    min_d, max_d = cod_all["Data"].min().date(), cod_all["Data"].max().date()
    if min_d < max_d:
        date_range = st.slider("Período", min_value=min_d, max_value=max_d, value=(min_d, max_d),
                               format="DD/MM/YYYY", label_visibility="collapsed")
    else:
        date_range = (min_d, max_d)
        st.caption(f"Data única: {min_d.strftime('%d/%m/%Y')}")

    st.markdown('<span class="sidebar-label"> FILTRO: CLASSIFICAÇÃO</span>', unsafe_allow_html=True)
    cls_filt = st.multiselect("Classificação", options=[CLS_TROCA, CLS_ENDER, CLS_INCOMP, CLS_SEMALT],
                              label_visibility="collapsed")
    st.markdown('<span class="sidebar-label"> FILTRO: TIPO ORIGEM</span>', unsafe_allow_html=True)
    org_filt = st.multiselect("Origem", options=sorted(cod_all["Origem"].unique()), label_visibility="collapsed")
    st.markdown('<span class="sidebar-label"> FILTRO: TIPO DESTINO</span>', unsafe_allow_html=True)
    dst_filt = st.multiselect("Destino", options=sorted(cod_all["Destino"].unique()), label_visibility="collapsed")
    st.markdown('<span class="sidebar-label"> FILTRO: OPERADOR</span>', unsafe_allow_html=True)
    user_filt = st.multiselect("Operador", options=sorted(cod_all["Usuario"].unique()), label_visibility="collapsed")
    st.markdown('<span class="sidebar-label"> FILTRO: PRODUTO</span>', unsafe_allow_html=True)
    prod_filt = st.multiselect("Produto", options=sorted(cod_all["Produto"].unique()), label_visibility="collapsed")
    st.markdown('<span class="sidebar-label"> FILTRO: OBSERVAÇÃO</span>', unsafe_allow_html=True)
    obs_opts = sorted(cod_all["Observacao"].dropna().astype(str).unique())
    obs_filt = st.multiselect("Observação", options=obs_opts, label_visibility="collapsed")

df_c = cod_all[(cod_all["Data"].dt.date >= date_range[0]) & (cod_all["Data"].dt.date <= date_range[1])]
if cls_filt:  df_c = df_c[df_c["Classificacao"].isin(cls_filt)]
if org_filt:  df_c = df_c[df_c["Origem"].isin(org_filt)]
if dst_filt:  df_c = df_c[df_c["Destino"].isin(dst_filt)]
if user_filt: df_c = df_c[df_c["Usuario"].isin(user_filt)]
if prod_filt: df_c = df_c[df_c["Produto"].isin(prod_filt)]
if obs_filt:  df_c = df_c[df_c["Observacao"].astype(str).isin(obs_filt)]

if df_c.empty:
    st.warning("Nenhum registro para os filtros selecionados.")
    st.stop()

df_d = df_det_base[df_det_base["Codigo"].isin(df_c["Codigo"])].copy()
df_t = df_c[df_c["Classificacao"] == CLS_TROCA]
df_par = df_c[df_c["Classificacao"] != CLS_INCOMP]
dias_ativos = max(1, df_c["Data"].nunique())

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 Visão Geral", "🔄 Trocas de Tipo", "📍 Endereços", "📦 Produtos e Seriais", "🛡️ Auditoria", "🔍 Detalhamento"])

# ---------------------------------------------------------
# ABA 1: VISÃO GERAL
# ---------------------------------------------------------
with tab1:
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Movimentações (CODIGO)", formatar_br(len(df_c)))
    m2.metric("Trocas de Tipo", formatar_br(len(df_t)))
    m3.metric("Transf. de Endereço", formatar_br((df_c["Classificacao"] == CLS_ENDER).sum()))
    m4.metric("Qtd em Troca de Tipo", formatar_br(df_t["Qtd_Mov"].sum()))
    m5.metric("Média Trocas / Dia", formatar_br(len(df_t) / dias_ativos))

    n_inc = (df_c["Classificacao"] == CLS_INCOMP).sum()
    n_div = int(df_c["Divergencia_Qtd"].sum())
    n_sem = (df_c["Classificacao"] == CLS_SEMALT).sum()
    if n_inc or n_div or n_sem:
        st.warning(f"⚠️ Alertas de auditoria: **{n_inc}** incompletos (sem par) · **{n_div}** com divergência de quantidade · "
                   f"**{n_sem}** sem alteração real. Veja a aba 🛡️ Auditoria.")

    st.divider()
    c_dir, c_esq = st.columns([2, 1])
    with c_dir:
        st.subheader("📈 Evolução Diária por Classificação")
        ev = df_c.groupby([df_c["Data"].dt.date.rename("Dia"), "Classificacao"]).size().reset_index(name="Qtd")
        fig = px.bar(ev, x="Dia", y="Qtd", color="Classificacao", barmode="stack", text="Qtd",
                     color_discrete_map=COR_CLASSE)
        fig.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, title=None),
                          xaxis=dict(tickformat="%d/%m", showgrid=False), yaxis_title="Movimentações")
        st.plotly_chart(estilo(fig), use_container_width=True)
    with c_esq:
        st.subheader("🏆 Ranking Operadores")
        ru = df_c.groupby("Usuario").size().reset_index(name="Mov").sort_values("Mov", ascending=False).head(10)
        fig_u = px.bar(ru, x="Mov", y="Usuario", orientation="h", text="Mov", color_discrete_sequence=[AZUL_ESCURO])
        fig_u.update_layout(yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(estilo(fig_u), use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("🕒 Movimentações por Hora do Dia")
        hh = df_c.groupby("Hora_Dia").size().reset_index(name="Qtd")
        fig_h = px.bar(hh, x="Hora_Dia", y="Qtd", text="Qtd", color_discrete_sequence=[AZUL_CLARO_TECADI])
        fig_h.update_layout(xaxis=dict(dtick=1, title="Hora"), yaxis_title="Movimentações")
        st.plotly_chart(estilo(fig_h), use_container_width=True)
    with c2:
        st.subheader("🧩 Composição por Classificação")
        cp = df_c.groupby("Classificacao").size().reset_index(name="Qtd")
        fig_p = px.pie(cp, names="Classificacao", values="Qtd", hole=0.55, color="Classificacao",
                       color_discrete_map=COR_CLASSE)
        st.plotly_chart(estilo(fig_p), use_container_width=True)

# ---------------------------------------------------------
# ABA 2: TROCAS DE TIPO DE ESTOQUE
# ---------------------------------------------------------
with tab2:
    if df_t.empty:
        st.info("Nenhuma troca de tipo de estoque no filtro atual. "
                "Neste recorte as movimentações mantêm o mesmo Tipo Estoque (veja Endereços).")
    else:
        t1, t2, t3 = st.columns(3)
        t1.metric("Trocas de Tipo", formatar_br(len(df_t)))
        t2.metric("Fluxos Distintos", formatar_br(df_t["Fluxo"].nunique()))
        t3.metric("Qtd Total Movimentada", formatar_br(df_t["Qtd_Mov"].sum()))
        st.divider()

        medida = st.radio("Medida", ["Nº de movimentações", "Quantidade (un.)"], horizontal=True, key="medida_matriz")
        col_val, agg = ("Codigo", "count") if medida.startswith("Nº") else ("Qtd_Mov", "sum")

        cA, cB = st.columns([1, 1])
        with cA:
            st.subheader("🔥 Matriz Origem × Destino")
            mat = df_t.pivot_table(index="Origem", columns="Destino", values=col_val, aggfunc=agg, fill_value=0)
            fig_m = go.Figure(go.Heatmap(
                z=mat.values, x=mat.columns, y=mat.index, text=mat.values, texttemplate="%{text:,.0f}",
                colorscale=[[0, "#EAF2FB"], [1, AZUL_ESCURO]], showscale=False))
            fig_m.update_layout(xaxis_title="Destino", yaxis_title="Origem", yaxis=dict(autorange="reversed"))
            st.plotly_chart(estilo(fig_m, 420), use_container_width=True)
        with cB:
            st.subheader("🏆 Top Fluxos de Troca")
            fl = df_t.groupby("Fluxo").agg(Movimentacoes=("Codigo", "size"), Quantidade=("Qtd_Mov", "sum")).reset_index()
            fl = fl.sort_values("Movimentacoes" if agg == "count" else "Quantidade", ascending=False).head(15)
            fig_f = px.bar(fl, x="Movimentacoes" if agg == "count" else "Quantidade", y="Fluxo", orientation="h",
                           text="Movimentacoes" if agg == "count" else "Quantidade", color_discrete_sequence=[AZUL_TECADI])
            fig_f.update_layout(yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(estilo(fig_f, 420), use_container_width=True)

        st.divider()
        st.subheader("📈 Evolução Diária por Fluxo")
        top_fluxos = df_t["Fluxo"].value_counts().head(8).index
        evf = df_t[df_t["Fluxo"].isin(top_fluxos)].groupby([df_t["Data"].dt.date.rename("Dia"), "Fluxo"]).size().reset_index(name="Qtd")
        fig_ef = px.bar(evf, x="Dia", y="Qtd", color="Fluxo", barmode="stack")
        fig_ef.update_layout(xaxis=dict(tickformat="%d/%m"), legend=dict(orientation="h", y=-0.25, title=None))
        st.plotly_chart(estilo(fig_ef), use_container_width=True)

        st.subheader("👤 Operador × Fluxo (Top 10 operadores)")
        top_users = df_t["Usuario"].value_counts().head(10).index
        uf = df_t[df_t["Usuario"].isin(top_users)].groupby(["Usuario", "Fluxo"]).size().reset_index(name="Qtd")
        fig_uf = px.bar(uf, x="Qtd", y="Usuario", color="Fluxo", orientation="h", barmode="stack")
        fig_uf.update_layout(yaxis={"categoryorder": "total ascending"}, legend=dict(orientation="h", y=-0.25, title=None))
        st.plotly_chart(estilo(fig_uf), use_container_width=True)

# ---------------------------------------------------------
# ABA 3: ENDEREÇOS
# ---------------------------------------------------------
with tab3:
    st.subheader("📍 Locais de Origem e Destino das Movimentações")
    st.caption("Considera movimentações com par completo (Troca de Tipo, Transf. de Endereço e Sem alteração).")
    if df_par.empty:
        st.info("Sem movimentações com par completo no filtro atual.")
    else:
        e1, e2 = st.columns(2)
        with e1:
            st.markdown("**Top 20 Locais de Origem (SAÍDA)**")
            lo = df_par.groupby("Local_Origem").size().reset_index(name="Qtd").sort_values("Qtd", ascending=False).head(20)
            fo = px.bar(lo, x="Qtd", y="Local_Origem", orientation="h", text="Qtd", color_discrete_sequence=[AZUL_ESCURO])
            fo.update_layout(yaxis={"categoryorder": "total ascending", "tickmode": "linear", "dtick": 1},
                             yaxis_title=None)
            st.plotly_chart(estilo(fo, max(420, 26 * len(lo))), use_container_width=True)
        with e2:
            st.markdown("**Top 20 Locais de Destino (ENTRADA)**")
            ld = df_par.groupby("Local_Destino").size().reset_index(name="Qtd").sort_values("Qtd", ascending=False).head(20)
            fd = px.bar(ld, x="Qtd", y="Local_Destino", orientation="h", text="Qtd", color_discrete_sequence=[AZUL_CLARO_TECADI])
            fd.update_layout(yaxis={"categoryorder": "total ascending", "tickmode": "linear", "dtick": 1},
                             yaxis_title=None)
            st.plotly_chart(estilo(fd, max(420, 26 * len(ld))), use_container_width=True)

        st.divider()
        st.subheader("🔀 Top Rotas (Local Origem → Local Destino)")
        rt = df_par.assign(Rota=df_par["Local_Origem"].astype(str) + "  →  " + df_par["Local_Destino"].astype(str))
        rt = rt.groupby("Rota").agg(Movimentacoes=("Codigo", "size"), Quantidade=("Qtd_Mov", "sum")).reset_index()
        rt = rt.sort_values("Movimentacoes", ascending=False).head(30)
        st.dataframe(rt, use_container_width=True, hide_index=True)

# ---------------------------------------------------------
# ABA 4: PRODUTOS E SERIAIS
# ---------------------------------------------------------
with tab4:
    p1, p2 = st.columns(2)
    with p1:
        st.subheader("📦 Top 15 Produtos Movimentados")
        pr = df_c.groupby("Produto").agg(Movimentacoes=("Codigo", "size"), Quantidade=("Qtd_Mov", "sum")).reset_index()
        pr = pr.sort_values("Movimentacoes", ascending=False).head(15)
        fp = px.bar(pr, x="Movimentacoes", y="Produto", orientation="h", text="Movimentacoes", color_discrete_sequence=[LARANJA])
        fp.update_layout(yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(estilo(fp), use_container_width=True)
    with p2:
        st.subheader("🔄 Top 15 Produtos com Troca de Tipo")
        if df_t.empty:
            st.info("Sem trocas de tipo no filtro atual.")
        else:
            pt = df_t.groupby("Produto").agg(Trocas=("Codigo", "size")).reset_index().sort_values("Trocas", ascending=False).head(15)
            fpt = px.bar(pt, x="Trocas", y="Produto", orientation="h", text="Trocas", color_discrete_sequence=[AZUL_ESCURO])
            fpt.update_layout(yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(estilo(fpt), use_container_width=True)

    st.divider()
    st.subheader("🔁 Seriais Reincidentes (mais de 1 movimentação no período)")
    sr = df_c.dropna(subset=["Serial"]).copy()
    sr["Serial"] = sr["Serial"].astype(str)
    if sr.empty:
        st.info("Base sem coluna Serial preenchida.")
    else:
        reinc = sr.groupby("Serial").agg(
            Produto=("Produto", "first"), Movimentacoes=("Codigo", "size"),
            Trocas_de_Tipo=("Classificacao", lambda s: (s == CLS_TROCA).sum()),
            Primeira=("Timestamp", "min"), Ultima=("Timestamp", "max")).reset_index()
        reinc = reinc[reinc["Movimentacoes"] > 1].sort_values(["Trocas_de_Tipo", "Movimentacoes"], ascending=False)
        if reinc.empty:
            st.success("✅ Nenhum serial com mais de uma movimentação.")
        else:
            st.caption(f"{len(reinc)} seriais com movimentação repetida — possível ida e volta ou retrabalho.")
            st.dataframe(reinc, use_container_width=True, hide_index=True)

# ---------------------------------------------------------
# ABA 5: AUDITORIA
# ---------------------------------------------------------
with tab5:
    st.subheader("🛡️ Inconsistências e Conformidade")
    aud = df_c.copy()
    motivos = pd.Series("", index=aud.index)
    motivos += np.where(aud["Classificacao"] == CLS_INCOMP, "Sem par ENTRADA/SAÍDA; ", "")
    motivos += np.where(aud["Divergencia_Qtd"], "Divergência de quantidade; ", "")
    motivos += np.where(aud["Classificacao"] == CLS_SEMALT, "Mesmo tipo e mesmo local; ", "")
    motivos += np.where(aud["Produtos_Distintos"] > 1, "Mais de um produto no CODIGO; ", "")
    aud["Motivo"] = motivos.str.strip("; ").str.strip()
    aud = aud[aud["Motivo"] != ""]

    a1, a2, a3, a4 = st.columns(4)
    a1.metric("Incompletos", formatar_br((aud["Classificacao"] == CLS_INCOMP).sum()))
    a2.metric("Divergência Qtd", formatar_br(aud["Divergencia_Qtd"].sum()))
    a3.metric("Sem Alteração", formatar_br((aud["Classificacao"] == CLS_SEMALT).sum()))
    a4.metric("Multi-Produto", formatar_br((aud["Produtos_Distintos"] > 1).sum()))

    if aud.empty:
        st.success("✅ Nenhuma inconsistência encontrada no filtro atual.")
    else:
        cols_aud = ["Codigo", "Timestamp", "Usuario", "Produto", "Serial", "Origem", "Local_Origem", "Destino",
                    "Local_Destino", "Qtd_Saida", "Qtd_Entrada", "Motivo"]
        st.dataframe(aud[cols_aud], use_container_width=True, hide_index=True)
        st.download_button("📥 Baixar Auditoria (Excel)", data=to_excel({"Auditoria": aud[cols_aud]}),
                           file_name=f"auditoria_transferencias_{datetime.now():%Y%m%d_%H%M}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ---------------------------------------------------------
# ABA 6: DETALHAMENTO
# ---------------------------------------------------------
with tab6:
    st.subheader("🔍 Movimentações Consolidadas por CODIGO")
    f1, f2, f3 = st.columns(3)
    with f1: q_cod = st.text_input("Filtrar CODIGO", key="q_cod")
    with f2: q_prod = st.text_input("Filtrar Produto / Serial", key="q_prod")
    with f3: q_loc = st.text_input("Filtrar Local (origem/destino)", key="q_loc")

    cols_cod = ["Codigo", "Timestamp", "Usuario", "Classificacao", "Produto", "Serial", "Origem", "Destino",
                "Local_Origem", "Local_Destino", "Qtd_Saida", "Qtd_Entrada", "Observacao"]
    det = df_c[cols_cod].copy()
    if q_cod:  det = det[det["Codigo"].astype(str).str.contains(q_cod, case=False)]
    if q_prod: det = det[det["Produto"].str.contains(q_prod, case=False) | det["Serial"].astype(str).str.contains(q_prod, case=False)]
    if q_loc:  det = det[det["Local_Origem"].astype(str).str.contains(q_loc, case=False) | det["Local_Destino"].astype(str).str.contains(q_loc, case=False)]
    st.caption(f"{formatar_br(len(det))} movimentações")
    st.dataframe(det, use_container_width=True, hide_index=True)

    d1, d2 = st.columns(2)
    with d1:
        st.download_button("📥 Baixar Consolidado (Excel)", data=to_excel({"Consolidado": det}),
                           file_name=f"transferencias_consolidado_{datetime.now():%Y%m%d_%H%M}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with d2:
        abas = {"Consolidado": df_c[cols_cod], "Linhas_Originais": df_d.drop(columns=["Papel"])}
        if not df_t.empty:
            abas["Trocas_de_Tipo"] = df_t[cols_cod]
        st.download_button("📥 Baixar Pacote Completo (Excel)", data=to_excel(abas),
                           file_name=f"transferencias_completo_{datetime.now():%Y%m%d_%H%M}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with st.expander("Ver linhas originais (ENTRADA/SAÍDA)"):
        st.dataframe(df_d.drop(columns=["Papel"]), use_container_width=True, hide_index=True)
