import hmac
import io
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Acuracidade de Estoque", page_icon="📦", layout="wide")

# ----------------------------------------------------------------------------
# Configuração
# ----------------------------------------------------------------------------
DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)
ARQ_DADOS = DATA_DIR / "ultimo_inventario.csv"
ARQ_META = DATA_DIR / "ultimo_inventario_meta.txt"
ARQ_LOG = DATA_DIR / "auditoria.csv"
MAX_TENTATIVAS = 5

ALIASES = {
    "endereco": ["endereco", "local", "posicao", "localizacao", "address"],
    "sku": ["sku", "codigo", "cod", "item", "produto", "material"],
    "descricao": ["descricao", "desc", "nome", "description"],
    "qtd_sistema": ["qtd_sistema", "saldo_sistema", "saldo", "estoque", "qtd_wms", "sistema"],
    "qtd_contada": ["qtd_contada", "contagem", "contado", "qtd_fisica", "fisico", "contada"],
    "custo_unit": ["custo_unit", "custo", "preco", "valor_unit"],
    "area": ["area", "setor", "zona", "rua"],
}
OBRIGATORIOS = ["endereco", "sku", "qtd_sistema", "qtd_contada"]
OPCIONAIS = ["descricao", "custo_unit", "area"]


# ----------------------------------------------------------------------------
# Auditoria
# ----------------------------------------------------------------------------
def registrar(evento: str, detalhe: str = "") -> None:
    linha = pd.DataFrame(
        [{
            "data_hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "perfil": st.session_state.get("perfil", "-"),
            "evento": evento,
            "detalhe": detalhe,
        }]
    )
    linha.to_csv(
        ARQ_LOG, mode="a", header=not ARQ_LOG.exists(), index=False, encoding="utf-8-sig"
    )


# ----------------------------------------------------------------------------
# Login
# ----------------------------------------------------------------------------
def check_login() -> str:
    if st.session_state.get("autenticado"):
        return st.session_state["perfil"]

    try:
        credenciais = dict(st.secrets["credenciais"])
    except Exception:
        st.error("Secrets não configurado. Defina [credenciais.<perfil>] com 'senha'.")
        st.stop()

    st.title("🔒 Acesso restrito")
    tentativas = st.session_state.get("tentativas", 0)
    if tentativas >= MAX_TENTATIVAS:
        st.error("Número máximo de tentativas excedido. Recarregue a página.")
        st.stop()

    perfil = st.selectbox("Perfil", list(credenciais.keys()))
    senha = st.text_input("Senha", type="password")

    if st.button("Entrar", type="primary"):
        esperada = str(credenciais[perfil]["senha"])
        if hmac.compare_digest(senha.encode(), esperada.encode()):
            st.session_state.update(autenticado=True, perfil=perfil, tentativas=0)
            registrar("LOGIN_OK")
            st.rerun()
        else:
            st.session_state["tentativas"] = tentativas + 1
            st.session_state["perfil"] = perfil
            registrar("LOGIN_FALHA", f"tentativa {tentativas + 1}")
            st.session_state.pop("perfil", None)
            st.error("Senha incorreta.")
    st.stop()


perfil = check_login()
PODE_EDITAR = perfil == "acuracidade"


# ----------------------------------------------------------------------------
# Dados
# ----------------------------------------------------------------------------
def norm(txt: str) -> str:
    txt = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode()
    return txt.strip().lower().replace(" ", "_")


def detectar_mapa(colunas: list[str]) -> dict:
    mapa = {}
    normalizadas = {norm(c): c for c in colunas}
    for campo, apelidos in ALIASES.items():
        for ap in apelidos:
            if ap in normalizadas:
                mapa[campo] = normalizadas[ap]
                break
    return mapa


def ler_arquivo(arquivo) -> pd.DataFrame:
    nome = arquivo.name.lower()
    if nome.endswith((".xlsx", ".xlsm", ".xls")):
        return pd.read_excel(arquivo)
    return pd.read_csv(arquivo, sep=None, engine="python", encoding="utf-8-sig")


def para_numero(serie: pd.Series) -> pd.Series:
    if serie.dtype == object:
        def conv(v):
            s = str(v).strip()
            if "," in s:
                s = s.replace(".", "").replace(",", ".")
            return s
        serie = serie.map(conv)
    return pd.to_numeric(serie, errors="coerce").fillna(0)


def preparar(df: pd.DataFrame, mapa: dict) -> pd.DataFrame:
    out = pd.DataFrame()
    out["endereco"] = df[mapa["endereco"]].astype(str).str.strip()
    out["sku"] = df[mapa["sku"]].astype(str).str.strip()
    out["descricao"] = (
        df[mapa["descricao"]].astype(str).str.strip() if mapa.get("descricao") else "N/D"
    )
    out["area"] = (
        df[mapa["area"]].astype(str).str.strip() if mapa.get("area") else "N/D"
    )
    out["qtd_sistema"] = para_numero(df[mapa["qtd_sistema"]])
    out["qtd_contada"] = para_numero(df[mapa["qtd_contada"]])
    out["custo_unit"] = (
        para_numero(df[mapa["custo_unit"]]) if mapa.get("custo_unit") else float("nan")
    )
    out["divergencia"] = out["qtd_contada"] - out["qtd_sistema"]
    out["div_abs"] = out["divergencia"].abs()
    out["status"] = out["divergencia"].map(
        lambda d: "OK" if d == 0 else ("Sobra" if d > 0 else "Falta")
    )
    out["valor_div"] = out["divergencia"] * out["custo_unit"]
    return out


def carregar_ultimo() -> pd.DataFrame | None:
    if ARQ_DADOS.exists():
        return pd.read_csv(ARQ_DADOS, encoding="utf-8-sig")
    return None


def calcular_kpis(df: pd.DataFrame) -> dict:
    total = len(df)
    ok = int((df["divergencia"] == 0).sum())
    soma_sis = df["qtd_sistema"].sum()
    acur_qtd = max(0.0, 1 - df["div_abs"].sum() / soma_sis) if soma_sis > 0 else float("nan")
    return {
        "total": total,
        "ok": ok,
        "acur_itens": ok / total if total else float("nan"),
        "acur_qtd": acur_qtd,
        "sobras": df.loc[df["divergencia"] > 0, "divergencia"].sum(),
        "faltas": df.loc[df["divergencia"] < 0, "divergencia"].sum(),
        "liquida": df["divergencia"].sum(),
        "valor_liq": df["valor_div"].sum(skipna=True) if df["custo_unit"].notna().any() else None,
    }


def exportar_excel(df: pd.DataFrame, kpis: dict) -> bytes:
    buf = io.BytesIO()
    resumo = pd.DataFrame(
        {
            "Indicador": [
                "Itens contados", "Itens sem divergência", "Acuracidade por item",
                "Acuracidade por quantidade", "Sobras (un)", "Faltas (un)", "Divergência líquida (un)",
            ],
            "Valor": [
                kpis["total"], kpis["ok"], round(kpis["acur_itens"], 4),
                round(kpis["acur_qtd"], 4), kpis["sobras"], kpis["faltas"], kpis["liquida"],
            ],
        }
    )
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        resumo.to_excel(xw, "Resumo", index=False)
        df[df["divergencia"] != 0].sort_values("div_abs", ascending=False).to_excel(
            xw, "Divergencias", index=False
        )
        df.to_excel(xw, "Base", index=False)
    return buf.getvalue()


# ----------------------------------------------------------------------------
# Sidebar
# ----------------------------------------------------------------------------
with st.sidebar:
    st.markdown(f"**Perfil:** {perfil}")
    meta = st.slider("Meta de acuracidade (%)", 90.0, 100.0, 98.0, 0.5)
    if st.button("Sair"):
        registrar("LOGOUT")
        for k in ["autenticado", "perfil"]:
            st.session_state.pop(k, None)
        st.rerun()

st.title("📦 Acuracidade de Estoque")

# ----------------------------------------------------------------------------
# Abas por perfil
# ----------------------------------------------------------------------------
nomes_abas = ["📊 Dashboard", "🔎 Divergências"]
if PODE_EDITAR:
    nomes_abas = ["⬆️ Importar inventário"] + nomes_abas + ["🧾 Auditoria"]
abas = dict(zip(nomes_abas, st.tabs(nomes_abas)))

# ---- Importação (somente perfil acuracidade) ----
if PODE_EDITAR:
    with abas["⬆️ Importar inventário"]:
        st.caption("Colunas obrigatórias: endereço, SKU, qtd sistema e qtd contada. Opcionais: descrição, custo unitário, área.")
        arquivo = st.file_uploader("Arquivo (.xlsx ou .csv)", type=["xlsx", "xlsm", "xls", "csv"])
        if arquivo:
            try:
                bruto = ler_arquivo(arquivo)
            except Exception as e:
                st.error(f"Não foi possível ler o arquivo: {e}")
                st.stop()

            auto = detectar_mapa(list(bruto.columns))
            st.markdown("**Mapeamento de colunas** (ajuste se necessário)")
            cols = st.columns(len(OBRIGATORIOS) + len(OPCIONAIS))
            mapa = {}
            opcoes = ["—"] + list(bruto.columns)
            for c, campo in zip(cols, OBRIGATORIOS + OPCIONAIS):
                padrao = opcoes.index(auto[campo]) if campo in auto else 0
                esc = c.selectbox(campo, opcoes, index=padrao, key=f"map_{campo}")
                if esc != "—":
                    mapa[campo] = esc

            faltando = [c for c in OBRIGATORIOS if c not in mapa]
            if faltando:
                st.warning(f"Mapeie as colunas obrigatórias: {', '.join(faltando)}")
            else:
                prep = preparar(bruto, mapa)
                st.dataframe(prep.head(20), use_container_width=True)
                dup = prep.duplicated(["endereco", "sku"]).sum()
                if dup:
                    st.warning(f"{dup} linhas duplicadas (endereço + SKU). Verifique antes de publicar.")
                if st.button("Publicar inventário", type="primary"):
                    prep.to_csv(ARQ_DADOS, index=False, encoding="utf-8-sig")
                    ARQ_META.write_text(
                        f"{arquivo.name} | {datetime.now():%d/%m/%Y %H:%M} | {perfil}",
                        encoding="utf-8",
                    )
                    registrar("UPLOAD", f"{arquivo.name} ({len(prep)} linhas)")
                    st.success("Inventário publicado.")
                    st.rerun()

# ---- Dados base para as demais abas ----
base = carregar_ultimo()
if base is None:
    for nome in nomes_abas:
        if nome not in ("⬆️ Importar inventário", "🧾 Auditoria"):
            with abas[nome]:
                st.info("Nenhum inventário publicado ainda.")
else:
    if ARQ_META.exists():
        st.caption(f"Fonte: {ARQ_META.read_text(encoding='utf-8')}")

    with st.sidebar:
        st.markdown("---")
        areas = sorted(base["area"].dropna().unique())
        f_area = st.multiselect("Área", areas, default=areas)
        f_status = st.multiselect("Status", ["OK", "Sobra", "Falta"], default=["OK", "Sobra", "Falta"])

    df = base[base["area"].isin(f_area) & base["status"].isin(f_status)]
    if df.empty:
        with abas["📊 Dashboard"]:
            st.warning("Nenhum registro com os filtros atuais.")
        st.stop()

    k = calcular_kpis(df)

    # ---- Dashboard ----
    with abas["📊 Dashboard"]:
        c = st.columns(5)
        c[0].metric("Itens contados", f"{k['total']:,}".replace(",", "."))
        c[1].metric(
            "Acuracidade (itens)", f"{k['acur_itens']:.2%}",
            delta=f"{(k['acur_itens'] * 100 - meta):+.2f} p.p. vs meta",
        )
        c[2].metric("Acuracidade (qtd)", f"{k['acur_qtd']:.2%}")
        c[3].metric("Sobras / Faltas (un)", f"+{k['sobras']:,.0f} / {k['faltas']:,.0f}")
        c[4].metric(
            "Impacto líquido (R$)",
            f"{k['valor_liq']:,.2f}" if k["valor_liq"] is not None else "N/D",
        )

        g1, g2 = st.columns(2)
        dist = df["status"].value_counts().reset_index()
        dist.columns = ["status", "itens"]
        g1.plotly_chart(
            px.pie(dist, names="status", values="itens", hole=0.5, title="Distribuição por status",
                   color="status", color_discrete_map={"OK": "#2e7d32", "Sobra": "#f9a825", "Falta": "#c62828"}),
            use_container_width=True,
        )
        por_area = (
            df.groupby("area")["divergencia"].apply(lambda s: (s == 0).mean() * 100).reset_index()
        )
        por_area.columns = ["area", "acuracidade"]
        fig = px.bar(por_area.sort_values("acuracidade"), x="acuracidade", y="area",
                     orientation="h", title="Acuracidade por área (%)", text_auto=".1f")
        fig.add_vline(x=meta, line_dash="dash", annotation_text="Meta")
        g2.plotly_chart(fig, use_container_width=True)

        top = df.sort_values("div_abs", ascending=False).head(10)
        st.plotly_chart(
            px.bar(top, x="sku", y="divergencia", color="status", title="Top 10 divergências (un)",
                   color_discrete_map={"Sobra": "#f9a825", "Falta": "#c62828", "OK": "#2e7d32"}),
            use_container_width=True,
        )

    # ---- Divergências ----
    with abas["🔎 Divergências"]:
        div = df[df["divergencia"] != 0].sort_values("div_abs", ascending=False)
        st.write(f"**{len(div)}** itens com divergência para recontagem/tratativa.")
        st.dataframe(
            div[["endereco", "sku", "descricao", "area", "qtd_sistema", "qtd_contada", "divergencia", "status", "valor_div"]],
            use_container_width=True, hide_index=True,
        )
        if st.download_button(
            "⬇️ Exportar relatório (Excel)", data=exportar_excel(df, k),
            file_name=f"acuracidade_{datetime.now():%Y%m%d_%H%M}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ):
            registrar("EXPORT", f"{len(df)} linhas")

# ---- Auditoria (somente perfil acuracidade) ----
if PODE_EDITAR:
    with abas["🧾 Auditoria"]:
        if ARQ_LOG.exists():
            log = pd.read_csv(ARQ_LOG, encoding="utf-8-sig").sort_values("data_hora", ascending=False)
            st.dataframe(log, use_container_width=True, hide_index=True)
            st.download_button("⬇️ Baixar log", log.to_csv(index=False).encode("utf-8-sig"), "auditoria.csv")
        else:
            st.info("Sem eventos registrados.")
