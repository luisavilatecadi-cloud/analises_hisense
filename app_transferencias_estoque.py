# ============================================================================
# TROCA DE ESTOQUE - Identificação de alterações de Tipo de Estoque (WMS)
# Regra: mesmo CODIGO com linhas de Tipo Estoque DIFERENTES = alteração de estoque
# ============================================================================
import io
import unicodedata
from datetime import datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Troca de Estoque", page_icon="🔄", layout="wide")

# ----------------------------------------------------------------------------
# PARTE 1 - Configuração e leitura
# ----------------------------------------------------------------------------
ALIASES = {
    "codigo": ["codigo", "cod", "cod_mov", "id_mov", "movimento"],
    "tipo_mov": ["tipo_mov", "tipo_movimento", "movimento_tipo", "tipo_movimentacao"],
    "produto": ["produto", "sku", "item", "material"],
    "quantidade": ["quantidade", "qtd", "qtde", "qtd_mov"],
    "tipo_estoque": ["tipo_estoque", "tp_estoque", "tipo_est", "estoque"],
    "filial": ["filial", "cd", "unidade"],
    "data": ["data", "dt", "data_mov"],
    "hora": ["hora", "hr", "hora_mov"],
    "usuario": ["usuario", "user", "operador"],
    "localizacao": ["localizacao", "endereco", "posicao"],
    "observacao": ["observacao", "obs", "motivo"],
    "serial": ["serial", "numero_serie", "n_serie"],
    "lote": ["lote"],
    "cod_cliente": ["cod_cliente", "cliente", "codigo_cliente"],
    "saldo_bloq": ["saldo_bloq", "saldo_bloq?", "bloqueado"],
}
OBRIGATORIOS = ["codigo", "tipo_mov", "produto", "quantidade", "tipo_estoque"]
OPCIONAIS = [c for c in ALIASES if c not in OBRIGATORIOS]

TIPO_TROCA = "Troca de tipo de estoque"
TIPO_ENDERECO = "Transferência de endereço"


def norm(txt) -> str:
    txt = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode()
    txt = txt.strip().lower().replace(".", "_").replace(" ", "_")
    return txt


def detectar_mapa(colunas: list) -> dict:
    normalizadas = {norm(c): c for c in colunas}
    mapa = {}
    for campo, apelidos in ALIASES.items():
        for ap in apelidos:
            if ap in normalizadas:
                mapa[campo] = normalizadas[ap]
                break
    return mapa


@st.cache_data(show_spinner="Lendo arquivo...")
def ler_arquivo(conteudo: bytes, nome: str) -> pd.DataFrame:
    if nome.lower().endswith((".xlsx", ".xlsm", ".xls")):
        return pd.read_excel(io.BytesIO(conteudo))
    return pd.read_csv(io.BytesIO(conteudo), sep=None, engine="python", encoding="utf-8-sig")


def para_numero(serie: pd.Series) -> pd.Series:
    if serie.dtype == object or str(serie.dtype).startswith("str"):
        def conv(v):
            s = str(v).strip()
            if "," in s:
                s = s.replace(".", "").replace(",", ".")
            return s
        serie = serie.map(conv)
    return pd.to_numeric(serie, errors="coerce").fillna(0)


def limpa_texto(serie: pd.Series) -> pd.Series:
    """Texto limpo; números inteiros guardados como float viram '26' e não '26.0'."""
    def f(v):
        if pd.isna(v):
            return ""
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v).strip()
    return serie.map(f)


def parse_rotulos(txt: str) -> dict:
    """Converte linhas 'codigo=Descrição' em dicionário."""
    rot = {}
    for linha in txt.splitlines():
        if "=" in linha:
            k, v = linha.split("=", 1)
            if k.strip():
                rot[k.strip()] = v.strip()
    return rot


# ----------------------------------------------------------------------------
# PARTE 2 - Preparação e detecção das alterações
# ----------------------------------------------------------------------------
def classifica_mov(v) -> str:
    s = norm(v).upper()
    if s.startswith("ENT"):
        return "ENTRADA"
    if s.startswith("SAI"):
        return "SAIDA"
    return "OUTRO"


def preparar(df: pd.DataFrame, mapa: dict) -> pd.DataFrame:
    """Padroniza as colunas. O índice original do arquivo fica em '_idx'."""
    n = len(df)
    vazio = pd.Series([""] * n, index=df.index)

    def col(campo):
        return limpa_texto(df[mapa[campo]]) if mapa.get(campo) else vazio

    out = pd.DataFrame(index=df.index)
    out["_idx"] = range(n)
    out["filial"] = col("filial")
    out["codigo"] = col("codigo")
    out["tipo_mov"] = df[mapa["tipo_mov"]].map(classifica_mov)
    out["produto"] = col("produto")
    out["quantidade"] = para_numero(df[mapa["quantidade"]])
    out["tipo_estoque"] = col("tipo_estoque")
    out["usuario"] = col("usuario")
    out["localizacao"] = col("localizacao")
    out["observacao"] = col("observacao")
    out["serial"] = col("serial")
    out["lote"] = col("lote")
    out["cod_cliente"] = col("cod_cliente")
    out["saldo_bloq"] = col("saldo_bloq")

    data = pd.to_datetime(df[mapa["data"]], errors="coerce") if mapa.get("data") else pd.NaT
    if mapa.get("hora") and mapa.get("data"):
        hora = df[mapa["hora"]].astype(str).str.strip()
        dh = pd.to_datetime(data.dt.strftime("%Y-%m-%d") + " " + hora, errors="coerce")
        data = dh.fillna(data)
    out["data_hora"] = data
    out["chave"] = out["filial"] + "|" + out["codigo"]
    return out


def marcar_categoria(base: pd.DataFrame) -> pd.DataFrame:
    """Marca cada linha: troca de tipo de estoque, transferência de endereço ou normal."""
    base = base.copy()
    g = base.groupby("chave")
    linhas = g["_idx"].transform("count")
    n_tipos = g["tipo_estoque"].transform("nunique")
    n_locs = g["localizacao"].transform("nunique")
    base["categoria"] = "Normal"
    base.loc[(linhas >= 2) & (n_tipos == 1) & (n_locs > 1), "categoria"] = TIPO_ENDERECO
    base.loc[(linhas >= 2) & (n_tipos > 1), "categoria"] = TIPO_TROCA
    return base


def _juntar(serie: pd.Series) -> str:
    vals = sorted({v for v in serie if v != ""})
    return ", ".join(vals) if vals else "?"


def resumir_alteracoes(base: pd.DataFrame) -> pd.DataFrame:
    """Uma linha por (movimento, produto): de qual estoque para qual, e quanto."""
    alt = base[base["categoria"] != "Normal"].copy()
    if alt.empty:
        return pd.DataFrame()
    alt["qtd_ent"] = alt["quantidade"].where(alt["tipo_mov"] == "ENTRADA", 0)
    alt["qtd_sai"] = alt["quantidade"].where(alt["tipo_mov"] == "SAIDA", 0)
    chaves = ["chave", "produto"]

    res = alt.groupby(chaves, as_index=False).agg(
        categoria=("categoria", "first"),
        filial=("filial", "first"),
        codigo=("codigo", "first"),
        data_hora=("data_hora", "min"),
        usuario=("usuario", "first"),
        observacao=("observacao", "first"),
        qtd_saida=("qtd_sai", "sum"),
        qtd_entrada=("qtd_ent", "sum"),
        linhas=("_idx", "count"),
        seriais=("serial", lambda s: s[s != ""].nunique()),
    )

    def lado(mov, campo, nome):
        sub = alt[alt["tipo_mov"] == mov].groupby(chaves)[campo].agg(_juntar)
        return sub.rename(nome).reset_index()

    for mov, campo, nome in [
        ("SAIDA", "tipo_estoque", "estoque_origem"),
        ("ENTRADA", "tipo_estoque", "estoque_destino"),
        ("SAIDA", "localizacao", "local_origem"),
        ("ENTRADA", "localizacao", "local_destino"),
    ]:
        res = res.merge(lado(mov, campo, nome), on=chaves, how="left")
    for c in ["estoque_origem", "estoque_destino", "local_origem", "local_destino"]:
        res[c] = res[c].fillna("?")

    res["qtd_alterada"] = res["qtd_saida"].where(res["qtd_saida"] > 0, res["qtd_entrada"])
    ok = (res["qtd_saida"] == res["qtd_entrada"]) & (res["qtd_saida"] > 0)
    res["conferencia"] = ok.map({True: "OK", False: "Verificar"})
    return res.sort_values("data_hora", ascending=False).reset_index(drop=True)


# ----------------------------------------------------------------------------
# PARTE 3 - Entrada de dados, mapeamento e filtros
# ----------------------------------------------------------------------------
st.title("🔄 Troca de Estoque")
st.caption(
    "Identifica alterações de estoque: mesmo CÓDIGO com linhas de TIPO DE ESTOQUE diferentes "
    "(saiu de um estoque e entrou em outro)."
)

with st.sidebar:
    st.header("⚙️ Arquivo")
    arquivo = st.file_uploader("Movimentações (.xlsx ou .csv)", type=["xlsx", "xlsm", "xls", "csv"])
    incluir_end = st.checkbox(
        "Incluir transferências de endereço",
        value=False,
        help="Mesmo código e mesmo tipo de estoque, mas endereços diferentes. "
             "Por padrão fica de fora: não é troca de estoque.",
    )
    rotulos_txt = st.text_area(
        "Nomes dos tipos de estoque (opcional)",
        placeholder="26=Disponível\n30=Avaria\n40=Bloqueado",
        help="Uma linha por tipo, no formato código=nome.",
    )

if not arquivo:
    st.info("⬆️ Envie a planilha de movimentações na barra lateral para começar.")
    st.stop()

try:
    bruto = ler_arquivo(arquivo.getvalue(), arquivo.name)
except Exception as e:
    st.error(f"Não foi possível ler o arquivo: {e}")
    st.stop()

auto = detectar_mapa(list(bruto.columns))
mapa = {}
opcoes = ["—"] + list(bruto.columns)
with st.expander("🧩 Mapeamento de colunas (ajuste se necessário)", expanded=any(c not in auto for c in OBRIGATORIOS)):
    campos = OBRIGATORIOS + OPCIONAIS
    cols = st.columns(5)
    for i, campo in enumerate(campos):
        padrao = opcoes.index(auto[campo]) if campo in auto else 0
        esc = cols[i % 5].selectbox(campo, opcoes, index=padrao, key=f"map_{campo}")
        if esc != "—":
            mapa[campo] = esc

faltando = [c for c in OBRIGATORIOS if c not in mapa]
if faltando:
    st.warning(f"Mapeie as colunas obrigatórias: {', '.join(faltando)}")
    st.stop()

rotulos = parse_rotulos(rotulos_txt)


def rot(v) -> str:
    """Mostra '26 - Disponível' quando houver nome cadastrado."""
    partes = [p.strip() for p in str(v).split(",")]
    return ", ".join(f"{p} - {rotulos[p]}" if p in rotulos else p for p in partes)


base = marcar_categoria(preparar(bruto, mapa))
alt_all = resumir_alteracoes(base)

c_info = st.columns(4)
c_info[0].metric("Linhas no arquivo", f"{len(base):,}".replace(",", "."))
c_info[1].metric("Movimentos (códigos)", f"{base['chave'].nunique():,}".replace(",", "."))
c_info[2].metric(
    "Trocas de tipo de estoque",
    f"{(alt_all['categoria'] == TIPO_TROCA).sum() if not alt_all.empty else 0:,}".replace(",", "."),
)
c_info[3].metric(
    "Transferências de endereço",
    f"{(alt_all['categoria'] == TIPO_ENDERECO).sum() if not alt_all.empty else 0:,}".replace(",", "."),
)

if alt_all.empty:
    st.warning("Nenhuma alteração encontrada: não há código repetido com tipo de estoque (ou endereço) diferente.")
    st.stop()

alt = alt_all if incluir_end else alt_all[alt_all["categoria"] == TIPO_TROCA]
if alt.empty:
    st.warning(
        "Nenhuma TROCA de tipo de estoque encontrada neste arquivo. "
        "Existem transferências de endereço: marque a opção na barra lateral para vê-las."
    )
    st.stop()

# Colunas de exibição já com nome do estoque
alt = alt.copy()
alt["de"] = alt["estoque_origem"].map(rot)
alt["para"] = alt["estoque_destino"].map(rot)
alt["fluxo"] = alt["de"] + "  →  " + alt["para"]

with st.sidebar:
    st.markdown("---")
    st.header("🔎 Filtros")
    dmin, dmax = alt["data_hora"].min(), alt["data_hora"].max()
    if pd.notna(dmin) and pd.notna(dmax):
        faixa = st.date_input("Período", (dmin.date(), dmax.date()), min_value=dmin.date(), max_value=dmax.date())
        if isinstance(faixa, (list, tuple)) and len(faixa) == 2:
            ini, fim = pd.Timestamp(faixa[0]), pd.Timestamp(faixa[1]) + pd.Timedelta(days=1)
            alt = alt[(alt["data_hora"] >= ini) & (alt["data_hora"] < fim)]
    f_prod = st.multiselect("Produto", sorted(alt["produto"].unique()))
    f_fluxo = st.multiselect("Fluxo (de → para)", sorted(alt["fluxo"].unique()))
    f_user = st.multiselect("Usuário", sorted(u for u in alt["usuario"].unique() if u))
    f_conf = st.multiselect("Conferência", ["OK", "Verificar"], default=["OK", "Verificar"])
    busca = st.text_input("Buscar código / serial / endereço")

if f_prod:
    alt = alt[alt["produto"].isin(f_prod)]
if f_fluxo:
    alt = alt[alt["fluxo"].isin(f_fluxo)]
if f_user:
    alt = alt[alt["usuario"].isin(f_user)]
alt = alt[alt["conferencia"].isin(f_conf)]
if busca.strip():
    b = busca.strip().lower()
    chaves_busca = base[
        base[["codigo", "serial", "localizacao"]].apply(lambda s: s.str.lower().str.contains(b, regex=False)).any(axis=1)
    ]["chave"].unique()
    alt = alt[alt["chave"].isin(chaves_busca)]

if alt.empty:
    st.warning("Nenhum registro com os filtros atuais.")
    st.stop()


# ----------------------------------------------------------------------------
# PARTE 4 - Exportação e visualizações
# ----------------------------------------------------------------------------
COLS_TABELA = {
    "data_hora": "Data/Hora", "filial": "Filial", "codigo": "Código", "produto": "Produto",
    "qtd_alterada": "Qtd alterada", "de": "De (estoque)", "para": "Para (estoque)",
    "local_origem": "Endereço origem", "local_destino": "Endereço destino",
    "usuario": "Usuário", "observacao": "Observação", "seriais": "Seriais",
    "conferencia": "Conferência", "categoria": "Categoria",
}


def exportar_excel(alt_df: pd.DataFrame, base_df: pd.DataFrame, bruto_df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    resumo = pd.DataFrame({
        "Indicador": ["Alterações", "Quantidade total alterada", "Produtos distintos", "Usuários", "A verificar"],
        "Valor": [
            len(alt_df), alt_df["qtd_alterada"].sum(), alt_df["produto"].nunique(),
            alt_df["usuario"].nunique(), int((alt_df["conferencia"] == "Verificar").sum()),
        ],
    })
    fluxo = (
        alt_df.groupby(["de", "para"], as_index=False)
        .agg(alteracoes=("codigo", "count"), qtd=("qtd_alterada", "sum"), produtos=("produto", "nunique"))
        .sort_values("qtd", ascending=False)
    )
    por_prod = (
        alt_df.groupby("produto", as_index=False)
        .agg(alteracoes=("codigo", "count"), qtd=("qtd_alterada", "sum"))
        .sort_values("qtd", ascending=False)
    )
    chaves = alt_df["chave"].unique()
    registros = bruto_df.loc[base_df[base_df["chave"].isin(chaves)]["_idx"].values]
    tabela = alt_df[list(COLS_TABELA)].rename(columns=COLS_TABELA)
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        resumo.to_excel(xw, sheet_name="Resumo", index=False)
        tabela.to_excel(xw, sheet_name="Alteracoes", index=False)
        fluxo.to_excel(xw, sheet_name="Fluxos", index=False)
        por_prod.to_excel(xw, sheet_name="Por produto", index=False)
        registros.to_excel(xw, sheet_name="Registros completos", index=False)
    return buf.getvalue()


def fmt_int(v) -> str:
    return f"{v:,.0f}".replace(",", ".")


abas = st.tabs(["📊 Resumo", "📋 Alterações", "🔀 Fluxos", "🔍 Detalhar registro", "⬇️ Exportar"])

# ---- Resumo ----
with abas[0]:
    k = st.columns(5)
    k[0].metric("Alterações", fmt_int(len(alt)))
    k[1].metric("Quantidade alterada", fmt_int(alt["qtd_alterada"].sum()))
    k[2].metric("Produtos distintos", fmt_int(alt["produto"].nunique()))
    k[3].metric("Usuários", fmt_int(alt["usuario"].nunique()))
    k[4].metric("A verificar", fmt_int((alt["conferencia"] == "Verificar").sum()),
                help="Quantidade de saída diferente da entrada, ou lado ausente.")

    g1, g2 = st.columns(2)
    top_prod = alt.groupby("produto", as_index=False)["qtd_alterada"].sum().sort_values("qtd_alterada", ascending=False).head(10)
    g1.plotly_chart(
        px.bar(top_prod.sort_values("qtd_alterada"), x="qtd_alterada", y="produto", orientation="h",
               title="Top 10 produtos por quantidade alterada", text_auto=True),
        use_container_width=True,
    )
    top_fluxo = alt.groupby("fluxo", as_index=False)["qtd_alterada"].sum().sort_values("qtd_alterada", ascending=False).head(10)
    g2.plotly_chart(
        px.bar(top_fluxo.sort_values("qtd_alterada"), x="qtd_alterada", y="fluxo", orientation="h",
               title="Fluxos (de → para) por quantidade", text_auto=True),
        use_container_width=True,
    )

    g3, g4 = st.columns(2)
    por_user = alt.groupby("usuario", as_index=False)["qtd_alterada"].sum().sort_values("qtd_alterada", ascending=False).head(10)
    g3.plotly_chart(
        px.bar(por_user, x="usuario", y="qtd_alterada", title="Quantidade alterada por usuário", text_auto=True),
        use_container_width=True,
    )
    if alt["data_hora"].notna().any():
        serie = alt.dropna(subset=["data_hora"]).copy()
        span = (serie["data_hora"].max() - serie["data_hora"].min()).days
        serie["periodo"] = serie["data_hora"].dt.floor("D" if span >= 2 else "h")
        evol = serie.groupby("periodo", as_index=False)["qtd_alterada"].sum()
        g4.plotly_chart(
            px.line(evol, x="periodo", y="qtd_alterada", markers=True,
                    title="Evolução da quantidade alterada"),
            use_container_width=True,
        )

# ---- Alterações ----
with abas[1]:
    st.write(f"**{len(alt)}** alterações encontradas.")
    st.dataframe(
        alt[list(COLS_TABELA)].rename(columns=COLS_TABELA),
        use_container_width=True, hide_index=True,
        column_config={"Data/Hora": st.column_config.DatetimeColumn(format="DD/MM/YYYY HH:mm:ss")},
    )
    st.caption("Para ver todas as linhas originais de um movimento, use a aba 🔍 Detalhar registro.")

# ---- Fluxos ----
with abas[2]:
    fluxo = (
        alt.groupby(["de", "para"], as_index=False)
        .agg(alteracoes=("codigo", "count"), qtd=("qtd_alterada", "sum"), produtos=("produto", "nunique"))
        .sort_values("qtd", ascending=False)
    )
    st.subheader("De qual estoque para qual")
    st.dataframe(
        fluxo.rename(columns={"de": "De", "para": "Para", "alteracoes": "Alterações", "qtd": "Qtd", "produtos": "Produtos"}),
        use_container_width=True, hide_index=True,
    )
    origens = list(pd.unique(fluxo["de"]))
    destinos = [d for d in pd.unique(fluxo["para"])]
    nos = [f"De: {o}" for o in origens] + [f"Para: {d}" for d in destinos]
    fig = go.Figure(go.Sankey(
        node=dict(label=nos, pad=18, thickness=16),
        link=dict(
            source=[origens.index(r) for r in fluxo["de"]],
            target=[len(origens) + destinos.index(r) for r in fluxo["para"]],
            value=fluxo["qtd"].tolist(),
        ),
    ))
    fig.update_layout(title="Fluxo entre tipos de estoque (quantidade)", height=420)
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Produto × fluxo")
    piv = alt.pivot_table(index="produto", columns="fluxo", values="qtd_alterada", aggfunc="sum", fill_value=0)
    st.dataframe(piv, use_container_width=True)

# ---- Detalhar registro ----
with abas[3]:
    opcoes_det = alt.sort_values("data_hora", ascending=False)
    rotulo_det = {
        i: f"{r.codigo} | {r.produto} | {fmt_int(r.qtd_alterada)} un | {r.de} → {r.para}"
        for i, r in opcoes_det.iterrows()
    }
    sel = st.selectbox("Escolha a alteração", list(rotulo_det), format_func=lambda i: rotulo_det[i])
    r = alt.loc[sel]
    d = st.columns(4)
    d[0].metric("Produto", r["produto"])
    d[1].metric("Qtd alterada", fmt_int(r["qtd_alterada"]))
    d[2].metric("De", r["de"])
    d[3].metric("Para", r["para"])
    st.write(
        f"**Usuário:** {r['usuario'] or '-'}  |  **Data/Hora:** {r['data_hora']}  |  "
        f"**Endereço:** {r['local_origem']} → {r['local_destino']}  |  "
        f"**Observação:** {r['observacao'] or '-'}  |  **Conferência:** {r['conferencia']}"
    )
    idx_orig = base[base["chave"] == r["chave"]]["_idx"].values
    st.markdown("**Registro completo (todas as linhas originais do movimento)**")
    st.dataframe(bruto.iloc[idx_orig], use_container_width=True, hide_index=True)

# ---- Exportar ----
with abas[4]:
    st.write("O relatório contém: Resumo, Alterações, Fluxos, Por produto e Registros completos (linhas originais).")
    st.download_button(
        "⬇️ Exportar relatório (Excel)",
        data=exportar_excel(alt, base, bruto),
        file_name=f"troca_estoque_{datetime.now():%Y%m%d_%H%M}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
