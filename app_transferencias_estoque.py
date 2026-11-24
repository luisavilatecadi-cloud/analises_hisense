# ============================================================================
# TROCA DE ESTOQUE - Identificação de alterações de Tipo de Estoque (WMS)
# Regra: mesmo CODIGO + PRODUTO + SERIAL com Tipo Estoque DIFERENTE = troca de estoque
# (serial que não movimentou dentro do mesmo CODIGO fica de fora)
# ============================================================================
import hashlib
import hmac
import io
import unicodedata
from datetime import datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Troca de Estoque", page_icon="🔄", layout="wide")


# ----------------------------------------------------------------------------
# PARTE 0 - Login (usuário e senha em st.secrets)
# Formato esperado no secrets.toml (ou em Settings > Secrets no Streamlit Cloud):
#   [credenciais.acuracidade]
#   senha = "SenhaForte123"
#   [credenciais.geral]
#   senha = "sha256:<hash_hex_da_senha>"   # opcional: senha em hash SHA-256
# O nome da seção (acuracidade, geral) é o USUÁRIO digitado no login.
# ----------------------------------------------------------------------------
def _senha_confere(informada: str, guardada) -> bool:
    guardada = str(guardada).strip()  # str(): evita erro se a senha for só número no TOML
    if guardada.lower().startswith("sha256:"):
        informada = hashlib.sha256(informada.encode()).hexdigest()
        guardada = guardada[7:].strip().lower()
    return hmac.compare_digest(informada.encode(), guardada.encode())


def exigir_login():
    try:
        usuarios = {}
        for k, v in dict(st.secrets["credenciais"]).items():
            senha = v.get("senha") if hasattr(v, "get") else v
            if senha is not None:
                usuarios[str(k).strip().lower()] = senha
    except Exception:
        usuarios = {}
    if not usuarios:
        st.error(
            "Secrets não configurado. Crie seções [credenciais.<usuario>] com `senha = \"...\"` em "
            ".streamlit/secrets.toml (local) ou em Settings > Secrets (Streamlit Cloud)."
        )
        st.stop()

    if st.session_state.get("autenticado"):
        with st.sidebar:
            st.caption(f"👤 {st.session_state.get('usuario_logado', '')}")
            if st.button("Sair"):
                st.session_state.clear()
                st.rerun()
        return

    st.title("🔐 Acesso restrito")
    with st.form("login"):
        u = st.text_input("Usuário").strip().lower()
        p = st.text_input("Senha", type="password")
        entrar = st.form_submit_button("Entrar")
    if entrar:
        if u in usuarios and _senha_confere(p, usuarios[u]):
            st.session_state["autenticado"] = True
            st.session_state["usuario_logado"] = u
            st.rerun()
        st.error("Usuário ou senha inválidos.")
    st.stop()


exigir_login()

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

# DE/PARA oficial dos tipos de estoque (código WMS -> descrição).
# Os códigos são normalizados sem zeros à esquerda (000026 == 26).
DEPARA_ESTOQUE = {
    "000026": "E010 SALES",
    "000027": "E020 SPARE PARTS",
    "000028": "E030 A. SALES",
    "000029": "E050 SAMPLES",
    "000030": "E060 COML. RETURN",
    "000031": "E070 BOX DMG",
    "000032": "E080 PROD DMG",
    "000033": "E130 CERTIFICATION",
    "000034": "E140 E-COMMERCE",
    "000035": "E150 MARKETING",
    "000036": "E160 TECH RETURN",
    "000037": "E170 REWORK",
    "000038": "E200 MISSING WH",
    "000055": "E014 SALES TCD",
    "000056": "E034 A. SALES TCD",
    "000057": "E074 BOX DMG TCD",
    "000058": "E084 PROD DMG TCD",
    "000059": "E162 TECH RETURN TCD",
    "000060": "E172 REWORK TCD",
    "000061": "E202 MISSING WH TCD",
    "000135": "E019",
}


def norm(txt) -> str:
    txt = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode()
    txt = txt.strip().lower().replace(".", "_").replace(" ", "_")
    return txt


def cod_est(v) -> str:
    """Normaliza código de tipo de estoque: '000026' / '26' / 26.0 -> '26'."""
    s = str(v).strip()
    if s.isdigit():
        return s.lstrip("0") or "0"
    return s


DEPARA_NORM = {cod_est(k): v for k, v in DEPARA_ESTOQUE.items()}


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
    """Converte linhas 'codigo=Descrição' em dicionário (código normalizado)."""
    rot = {}
    for linha in txt.splitlines():
        if "=" in linha:
            k, v = linha.split("=", 1)
            if k.strip():
                rot[cod_est(k)] = v.strip()
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
    out["tipo_estoque"] = col("tipo_estoque").map(cod_est)
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
    # Chave do movimento (todas as linhas do código) e chave do ITEM movimentado
    out["chave"] = out["filial"] + "|" + out["codigo"]
    out["chave_item"] = out["chave"] + "|" + out["produto"] + "|" + out["serial"]
    return out


def marcar_categoria(base: pd.DataFrame) -> pd.DataFrame:
    """Marca cada linha: troca de tipo de estoque, transferência de endereço ou normal.

    A validação é por CODIGO + PRODUTO + SERIAL: um serial que não movimentou
    dentro do mesmo código (1 linha só) nunca é marcado como troca.
    """
    base = base.copy()
    g = base.groupby("chave_item")
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
    """Uma linha por (código, produto, serial): de qual estoque para qual, e quanto."""
    alt = base[base["categoria"] != "Normal"].copy()
    if alt.empty:
        return pd.DataFrame()
    alt["qtd_ent"] = alt["quantidade"].where(alt["tipo_mov"] == "ENTRADA", 0)
    alt["qtd_sai"] = alt["quantidade"].where(alt["tipo_mov"] == "SAIDA", 0)
    chaves = ["chave_item"]

    res = alt.groupby(chaves, as_index=False).agg(
        categoria=("categoria", "first"),
        chave=("chave", "first"),
        filial=("filial", "first"),
        codigo=("codigo", "first"),
        produto=("produto", "first"),
        serial=("serial", "first"),
        data_hora=("data_hora", "min"),
        usuario=("usuario", "first"),
        observacao=("observacao", "first"),
        qtd_saida=("qtd_sai", "sum"),
        qtd_entrada=("qtd_ent", "sum"),
        linhas=("_idx", "count"),
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
    "Identifica trocas de estoque: mesmo CÓDIGO + PRODUTO + SERIAL com TIPO DE ESTOQUE diferente "
    "(saiu de um estoque e entrou em outro). Seriais que não movimentaram são ignorados."
)

with st.sidebar:
    st.header("⚙️ Arquivo")
    arquivo = st.file_uploader("Movimentações (.xlsx ou .csv)", type=["xlsx", "xlsm", "xls", "csv"])
    incluir_end = st.checkbox(
        "Incluir transferências de endereço",
        value=False,
        help="Mesmo código + produto + serial e mesmo tipo de estoque, mas endereços diferentes. "
             "Por padrão fica de fora: não é troca de estoque.",
    )
    rotulos_txt = st.text_area(
        "Ajustes no DE/PARA de estoque (opcional)",
        placeholder="26=Disponível\n30=Avaria",
        help="Uma linha por tipo, no formato código=nome. Sobrescreve o DE/PARA padrão do sistema.",
    )
    with st.expander("📖 DE/PARA de tipo de estoque"):
        st.dataframe(
            pd.DataFrame({"DE": list(DEPARA_ESTOQUE), "PARA": list(DEPARA_ESTOQUE.values())}),
            hide_index=True, use_container_width=True,
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

rotulos = {**DEPARA_NORM, **parse_rotulos(rotulos_txt)}


def rot(v) -> str:
    """Mostra '26 - E010 SALES' quando houver nome cadastrado no DE/PARA."""
    partes = [cod_est(p) for p in str(v).split(",")]
    return ", ".join(f"{p} - {rotulos[p]}" if p in rotulos else p for p in partes)


if "serial" not in mapa:
    st.warning(
        "Coluna SERIAL não mapeada: a validação ficará por CÓDIGO + PRODUTO, "
        "e seriais que não movimentaram não serão separados."
    )

base = marcar_categoria(preparar(bruto, mapa))
alt_all = resumir_alteracoes(base)

# Auditoria: tipos de estoque do arquivo que não existem no DE/PARA
sem_depara = sorted({t for t in base["tipo_estoque"].unique() if t and t not in rotulos})
if sem_depara:
    st.warning(
        "Tipos de estoque sem DE/PARA cadastrado: " + ", ".join(sem_depara)
        + ". Inclua em DEPARA_ESTOQUE no fonte ou use o campo da barra lateral."
    )

c_info = st.columns(4)
c_info[0].metric("Linhas no arquivo", f"{len(base):,}".replace(",", "."))
c_info[1].metric("Movimentos (códigos)", f"{base['chave'].nunique():,}".replace(",", "."))
c_info[2].metric(
    "Trocas de tipo de estoque (itens)",
    f"{(alt_all['categoria'] == TIPO_TROCA).sum() if not alt_all.empty else 0:,}".replace(",", "."),
)
c_info[3].metric(
    "Transferências de endereço (itens)",
    f"{(alt_all['categoria'] == TIPO_ENDERECO).sum() if not alt_all.empty else 0:,}".replace(",", "."),
)

if alt_all.empty:
    st.warning(
        "Nenhuma alteração encontrada: não há código + produto + serial repetido "
        "com tipo de estoque (ou endereço) diferente."
    )
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
    itens_busca = base[
        base[["codigo", "serial", "localizacao"]].apply(lambda s: s.str.lower().str.contains(b, regex=False)).any(axis=1)
    ]["chave_item"].unique()
    alt = alt[alt["chave_item"].isin(itens_busca)]

if alt.empty:
    st.warning("Nenhum registro com os filtros atuais.")
    st.stop()


# ----------------------------------------------------------------------------
# PARTE 4 - Exportação e visualizações
# ----------------------------------------------------------------------------
COLS_TABELA = {
    "data_hora": "Data/Hora", "filial": "Filial", "codigo": "Código", "produto": "Produto",
    "serial": "Serial", "qtd_alterada": "Qtd alterada", "de": "De (estoque)", "para": "Para (estoque)",
    "local_origem": "Endereço origem", "local_destino": "Endereço destino",
    "usuario": "Usuário", "observacao": "Observação",
    "conferencia": "Conferência", "categoria": "Categoria",
}

COLS_EXPORT_TROCA = [
    "Produto", "Serial", "Endereço Origem", "Endereço Destino",
    "Estoque Origem", "Estoque Destino", "Usuário", "Horário", "Data",
]


def montar_export_troca(alt_df: pd.DataFrame) -> pd.DataFrame:
    """Layout de exportação: Produto | Serial | End. Origem | End. Destino |
    Estoque Origem | Estoque Destino | Usuário | Horário | Data (respeita os filtros)."""
    dh = pd.to_datetime(alt_df["data_hora"], errors="coerce")
    out = pd.DataFrame({
        "Produto": alt_df["produto"].values,
        "Serial": alt_df["serial"].values,
        "Endereço Origem": alt_df["local_origem"].values,
        "Endereço Destino": alt_df["local_destino"].values,
        "Estoque Origem": alt_df["de"].values,
        "Estoque Destino": alt_df["para"].values,
        "Usuário": alt_df["usuario"].values,
        "Horário": dh.dt.strftime("%H:%M:%S").fillna("").values,
        "Data": dh.dt.date.values,
    })
    return out[COLS_EXPORT_TROCA]


def excel_troca(df_exp: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df_exp.to_excel(xw, sheet_name="Troca de estoque", index=False)
        ws = xw.sheets["Troca de estoque"]
        for i, c in enumerate(df_exp.columns, start=1):
            letra = ws.cell(row=1, column=i).column_letter
            tam = max([len(str(c))] + [len(str(v)) for v in df_exp[c].head(500)])
            ws.column_dimensions[letra].width = min(max(tam + 2, 10), 45)
            if c == "Data":
                for row in range(2, len(df_exp) + 2):
                    ws.cell(row=row, column=i).number_format = "DD/MM/YYYY"
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    return buf.getvalue()


def exportar_excel(alt_df: pd.DataFrame, base_df: pd.DataFrame, bruto_df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    resumo = pd.DataFrame({
        "Indicador": ["Alterações (itens)", "Quantidade total alterada", "Produtos distintos", "Usuários", "A verificar"],
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
    itens = alt_df["chave_item"].unique()
    registros = bruto_df.loc[base_df[base_df["chave_item"].isin(itens)]["_idx"].values]
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


abas = st.tabs(["📊 Resumo", "📋 Alterações", "🔀 Fluxos", "🔍 Detalhar registro", "📤 Exportar troca", "⬇️ Relatório completo"])

# ---- Resumo ----
with abas[0]:
    k = st.columns(5)
    k[0].metric("Alterações (itens)", fmt_int(len(alt)))
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
    st.write(f"**{len(alt)}** alterações encontradas (uma linha por código + produto + serial).")
    st.dataframe(
        alt[list(COLS_TABELA)].rename(columns=COLS_TABELA),
        use_container_width=True, hide_index=True,
        column_config={"Data/Hora": st.column_config.DatetimeColumn(format="DD/MM/YYYY HH:mm:ss")},
    )
    st.caption("Para ver as linhas originais de um item, use a aba 🔍 Detalhar registro.")

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
        i: (f"{r.codigo} | {r.produto} | {('S/N ' + r.serial) if r.serial else 'sem serial'} | "
            f"{fmt_int(r.qtd_alterada)} un | {r.de} → {r.para}")
        for i, r in opcoes_det.iterrows()
    }
    sel = st.selectbox("Escolha a alteração", list(rotulo_det), format_func=lambda i: rotulo_det[i])
    r = alt.loc[sel]
    d = st.columns(5)
    d[0].metric("Produto", r["produto"])
    d[1].metric("Serial", r["serial"] or "-")
    d[2].metric("Qtd alterada", fmt_int(r["qtd_alterada"]))
    d[3].metric("De", r["de"])
    d[4].metric("Para", r["para"])
    st.write(
        f"**Usuário:** {r['usuario'] or '-'}  |  **Data/Hora:** {r['data_hora']}  |  "
        f"**Endereço:** {r['local_origem']} → {r['local_destino']}  |  "
        f"**Observação:** {r['observacao'] or '-'}  |  **Conferência:** {r['conferencia']}"
    )
    todas = st.checkbox("Mostrar todas as linhas do código (inclui seriais que não movimentaram)", value=False)
    if todas:
        idx_orig = base[base["chave"] == r["chave"]]["_idx"].values
        st.markdown("**Registro completo (todas as linhas originais do código)**")
    else:
        idx_orig = base[base["chave_item"] == r["chave_item"]]["_idx"].values
        st.markdown("**Linhas originais do item (código + produto + serial)**")
    st.dataframe(bruto.iloc[idx_orig], use_container_width=True, hide_index=True)

# ---- Exportar troca (layout operacional) ----
with abas[4]:
    exp_troca = montar_export_troca(alt)
    st.write(
        f"**{len(exp_troca)}** registros conforme os filtros atuais. "
        "Colunas: Produto | Serial | Endereço Origem | Endereço Destino | "
        "Estoque Origem | Estoque Destino | Usuário | Horário | Data."
    )
    st.dataframe(exp_troca, use_container_width=True, hide_index=True)
    e1, e2 = st.columns(2)
    carimbo = f"{datetime.now():%Y%m%d_%H%M}"
    e1.download_button(
        "⬇️ Exportar troca de estoque (Excel)",
        data=excel_troca(exp_troca),
        file_name=f"troca_estoque_dados_{carimbo}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    e2.download_button(
        "⬇️ Exportar troca de estoque (CSV)",
        data=exp_troca.to_csv(index=False, sep=";").encode("utf-8-sig"),
        file_name=f"troca_estoque_dados_{carimbo}.csv",
        mime="text/csv",
    )

# ---- Relatório completo ----
with abas[5]:
    st.write("O relatório contém: Resumo, Alterações, Fluxos, Por produto e Registros completos (linhas originais).")
    st.download_button(
        "⬇️ Exportar relatório completo (Excel)",
        data=exportar_excel(alt, base, bruto),
        file_name=f"troca_estoque_{datetime.now():%Y%m%d_%H%M}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
