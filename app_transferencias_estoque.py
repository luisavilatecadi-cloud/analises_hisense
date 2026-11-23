# ============================================================
# TROCA DE ESTOQUE - identificação de alterações de Tipo Estoque
# Regra: mesmo CODIGO com Tipo Estoque diferente (SAIDA + ENTRADA)
#        = 1 alteração. A quantidade é contada UMA vez só.
# ============================================================
import io
import unicodedata

import pandas as pd
import streamlit as st

TOP_N = 30

# ------------------------------------------------------------
# PARTE 1 - Leitura e normalização
# ------------------------------------------------------------
ALIASES = {
    "filial": "Filial", "codigo": "Codigo", "data": "Data", "hora": "Hora",
    "usuario": "Usuario", "tipo mov": "Tipo Mov", "produto": "Produto",
    "quantidade": "Quantidade", "tipo estoque": "Tipo Estoque",
    "lote": "Lote", "observacao": "Observacao",
}
REQUIRED = ["Codigo", "Tipo Mov", "Produto", "Quantidade", "Tipo Estoque"]


def _norm(txt):
    t = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode()
    t = t.lower().replace(".", " ").replace("_", " ")
    return " ".join(t.split())


def _to_num(s):
    s = s.astype(str).str.strip().str.replace(" ", "", regex=False)
    com_virgula = s.str.contains(",", regex=False)
    s = s.where(
        ~com_virgula,
        s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
    )
    return pd.to_numeric(s, errors="coerce").fillna(0).abs()


@st.cache_data(show_spinner=False)
def carregar(conteudo: bytes, nome: str) -> pd.DataFrame:
    bio = io.BytesIO(conteudo)
    if nome.lower().endswith(".csv"):
        df = pd.read_csv(bio, sep=None, engine="python", dtype=str)
    else:
        df = pd.read_excel(bio, dtype=str)
    df.columns = [ALIASES.get(_norm(c), str(c).strip()) for c in df.columns]
    faltando = [c for c in REQUIRED if c not in df.columns]
    if faltando:
        raise ValueError("Colunas obrigatórias ausentes: " + ", ".join(faltando))
    for c in ["Codigo", "Produto", "Tipo Estoque", "Tipo Mov"]:
        df[c] = df[c].fillna("").astype(str).str.strip()
    if "Filial" in df.columns:
        df["Filial"] = df["Filial"].fillna("").astype(str).str.strip()
    df["_qtd"] = _to_num(df["Quantidade"])
    mov = df["Tipo Mov"].map(_norm).str.upper()
    df["_mov"] = mov.map(lambda m: "SAIDA" if m.startswith("S") else ("ENTRADA" if m.startswith("E") else ""))
    return df


# ------------------------------------------------------------
# PARTE 2 - Regra de negócio (identificar trocas)
# ------------------------------------------------------------
def montar_trocas(df: pd.DataFrame) -> pd.DataFrame:
    """1 linha por CODIGO de alteração. Quantidade = SAIDA (contada 1 vez)."""
    chaves = [c for c in ["Filial", "Codigo"] if c in df.columns]
    base = df[df["_mov"].isin(["SAIDA", "ENTRADA"])]
    n_est = base.groupby(chaves)["Tipo Estoque"].transform("nunique")
    alt = base[n_est > 1]
    if alt.empty:
        return pd.DataFrame()

    extras = {c: (c, "first") for c in ["Data", "Hora", "Usuario"] if c in alt.columns}
    agg = (
        alt.groupby(chaves + ["_mov"])
        .agg(
            Produto=("Produto", "first"),
            Qtd=("_qtd", "sum"),
            Estoque=("Tipo Estoque", lambda x: " + ".join(sorted(set(x)))),
            **extras,
        )
        .reset_index()
    )
    sai = agg[agg["_mov"] == "SAIDA"].drop(columns="_mov")
    sai = sai.rename(columns={"Qtd": "Qtd Saida", "Estoque": "Estoque Origem"})
    ent = agg[agg["_mov"] == "ENTRADA"][chaves + ["Qtd", "Estoque"]]
    ent = ent.rename(columns={"Qtd": "Qtd Entrada", "Estoque": "Estoque Destino"})

    t = sai.merge(ent, on=chaves, how="inner")
    t = t[t["Estoque Origem"] != t["Estoque Destino"]].copy()
    # ENTRADA + SAIDA = mesma peça: conta só uma vez (usa a SAIDA)
    t["Quantidade"] = t["Qtd Saida"]
    t["Divergencia"] = (t["Qtd Saida"] - t["Qtd Entrada"]).abs() > 1e-9
    return t.reset_index(drop=True)


# ------------------------------------------------------------
# PARTE 3 - Rankings (TOP 30)
# ------------------------------------------------------------
def top(t: pd.DataFrame, por, n: int = TOP_N) -> pd.DataFrame:
    g = (
        t.groupby(por, dropna=False)
        .agg(Quantidade=("Quantidade", "sum"), Trocas=("Codigo", "nunique"))
        .reset_index()
        .sort_values(["Quantidade", "Trocas"], ascending=False)
        .head(n)
        .reset_index(drop=True)
    )
    g.index = range(1, len(g) + 1)
    return g


def para_excel(abas: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for nome, d in abas.items():
            d.to_excel(w, sheet_name=nome[:31], index=False)
    return buf.getvalue()


# ------------------------------------------------------------
# PARTE 4 - Interface
# ------------------------------------------------------------
def fmt(x):
    return f"{x:,.0f}".replace(",", ".")


def mostrar_top(titulo, tabela, total_geral):
    st.subheader(titulo)
    soma = tabela["Quantidade"].sum()
    c1, c2, c3 = st.columns(3)
    c1.metric(f"Soma do TOP {len(tabela)}", fmt(soma))
    c2.metric("Total geral de peças", fmt(total_geral))
    c3.metric("% do total", f"{(soma / total_geral * 100) if total_geral else 0:.1f}%")
    st.dataframe(tabela, use_container_width=True)


def main():
    st.set_page_config(page_title="Troca de Estoque", layout="wide")
    st.title("🔄 Troca de Estoque")
    st.caption(
        "Alteração = mesmo CODIGO com Tipo Estoque diferente (SAIDA + ENTRADA). "
        "A quantidade é contada uma única vez por alteração."
    )

    arq = st.file_uploader("Planilha de movimentações", type=["xlsx", "xls", "csv"])
    if not arq:
        st.info("Envie a planilha para identificar as trocas.")
        return
    try:
        df = carregar(arq.getvalue(), arq.name)
    except Exception as e:
        st.error(f"Erro ao ler a planilha: {e}")
        return

    trocas = montar_trocas(df)
    if trocas.empty:
        st.warning("Nenhuma alteração de estoque encontrada.")
        return

    # ---- filtros
    with st.sidebar:
        st.header("Filtros")
        if "Filial" in trocas.columns:
            filiais = sorted(trocas["Filial"].unique())
            sel = st.multiselect("Filial", filiais, default=filiais)
            trocas = trocas[trocas["Filial"].isin(sel)]
        busca = st.text_input("Produto contém")
        if busca:
            trocas = trocas[trocas["Produto"].str.contains(busca, case=False, na=False)]
    if trocas.empty:
        st.warning("Nenhum registro com os filtros atuais.")
        return

    total = trocas["Quantidade"].sum()
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Alterações (códigos)", fmt(len(trocas)))
    k2.metric("Peças alteradas", fmt(total))
    k3.metric("Produtos distintos", fmt(trocas["Produto"].nunique()))
    k4.metric("Divergências Saída x Entrada", fmt(trocas["Divergencia"].sum()))

    # ---- rankings
    chave = ["Produto", "Estoque Origem", "Estoque Destino"]
    top_combo = top(trocas, chave)
    top_prod = top(trocas, ["Produto"])
    top_rota = top(trocas, ["Estoque Origem", "Estoque Destino"])

    aba1, aba2, aba3, aba4 = st.tabs(
        ["Produto + Origem + Destino", "Top Produtos", "Top Rotas", "Detalhar"]
    )
    with aba1:
        mostrar_top(f"TOP {TOP_N} Produto + Estoque Origem + Estoque Destino", top_combo, total)
    with aba2:
        mostrar_top(f"TOP {TOP_N} Produtos", top_prod, total)
    with aba3:
        mostrar_top(f"TOP {TOP_N} Rotas (Origem → Destino)", top_rota, total)

    # ---- detalhe
    with aba4:
        st.subheader("Detalhar registros completos")
        rotulos = [
            f"{r['Produto']} | {r['Estoque Origem']} → {r['Estoque Destino']} | {fmt(r['Quantidade'])} pç"
            for _, r in top_combo.iterrows()
        ]
        modo = st.radio("Detalhar por", ["Combinação do TOP", "Código"], horizontal=True)
        if modo == "Combinação do TOP":
            i = st.selectbox("Combinação", range(len(rotulos)), format_func=lambda k: rotulos[k])
            linha = top_combo.iloc[i]
            sub = trocas[
                (trocas["Produto"] == linha["Produto"])
                & (trocas["Estoque Origem"] == linha["Estoque Origem"])
                & (trocas["Estoque Destino"] == linha["Estoque Destino"])
            ]
        else:
            cod = st.selectbox("Codigo", sorted(trocas["Codigo"].unique()))
            sub = trocas[trocas["Codigo"] == cod]
        st.dataframe(sub, use_container_width=True)
        codigos = set(sub["Codigo"])
        bruto = df[df["Codigo"].isin(codigos)].drop(columns=["_qtd", "_mov"])
        st.caption("Registros originais (SAIDA e ENTRADA)")
        st.dataframe(bruto, use_container_width=True)

    # ---- exportação
    st.divider()
    st.download_button(
        "⬇️ Baixar Excel",
        data=para_excel({
            "Trocas": trocas,
            "Top Produto+Origem+Destino": top_combo.reset_index(),
            "Top Produtos": top_prod.reset_index(),
            "Top Rotas": top_rota.reset_index(),
        }),
        file_name="troca_estoque.xlsx",
    )


if __name__ == "__main__":
    main()
