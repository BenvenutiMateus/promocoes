import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from openpyxl import load_workbook
import difflib

# ================= FUNÇÕES =================

def ler_excel_promocao_com_formulas(file, sheet_name="PROMOÇÃO", header_row=0):
    wb = load_workbook(file, data_only=True)
    ws = wb[sheet_name]
    df = pd.DataFrame(ws.values)

    df.columns = df.iloc[header_row]
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    df = df.loc[:, df.columns.notna()]
    df.columns = df.columns.astype(str)

    return df


def _normalize_merge_key(val):
    s = str(val)
    s = s.replace('.0', '')
    s = s.replace('MLB', '')
    s = s.replace('+T', '')
    s = s.split()[0]  # Pega só a primeira parte antes de espaços
    s = s.strip()
    parts = [p.strip() for p in s.split('-') if p.strip()]
    lower = [p.lower() for p in parts]
    if any('kit' in p for p in lower):
        take = parts[:3]
    else:
        take = parts[:2]
    return '-'.join(take)


def _auto_detect_header_from_excel(file, sheet_name, max_rows=10):
    try:
        tmp = pd.read_excel(file, sheet_name=sheet_name, header=None, nrows=max_rows)
    except Exception:
        return 0

    aliases = [
        'skc', 'sku', 'número do item', 'numero do item', 'numero', 'número', 'item',
        'id do anúncio', 'id do anuncio', 'id', 'ean', 'gtin', 'preço', 'preco', 'valor'
    ]

    for i in range(len(tmp)):
        row = tmp.iloc[i].astype(str).str.lower().tolist()
        matches = sum(any(a in c for a in aliases) for c in row)
        if matches >= 1:
            return i
    return 0


def _suggest_column(columns, aliases):
    cols = list(columns)
    lcols = [c.lower() for c in cols]

    # exact
    for a in aliases:
        if a in lcols:
            return cols[lcols.index(a)]

    # contains
    for idx, c in enumerate(lcols):
        for a in aliases:
            if a in c:
                return cols[idx]

    # fuzzy
    for a in aliases:
        m = difflib.get_close_matches(a, lcols, n=1)
        if m:
            return cols[lcols.index(m[0])]

    return None


def _deduplicate_columns(df):
    counts = {}
    new_columns = []
    renamed = {}

    for col in df.columns:
        base = str(col).strip()
        counts[base] = counts.get(base, 0) + 1
        if counts[base] == 1:
            new_columns.append(base)
            continue

        new_name = f"{base}__{counts[base]}"
        new_columns.append(new_name)
        renamed.setdefault(base, []).append(new_name)

    if renamed:
        df = df.copy()
        df.columns = new_columns

    return df, renamed


# ================= APP =================

st.set_page_config("Gerenciador de Promoções", layout="wide")
st.title("📊 Gerenciador de Promoções por Marketplace")

# ================= SIDEBAR =================

with st.sidebar:
    st.header("⚙️ Configuração")

    arquivo_skus = st.file_uploader("1️⃣ Planilha de SKUs / IDs", type=["xlsx", "xls", "csv"])
    arquivo_precos = st.file_uploader("2️⃣ Base de Preços", type=["xlsx", "xls", "csv"])

    if not arquivo_skus or not arquivo_precos:
        st.info("👆 Envie os dois arquivos")
        st.stop()

    # --- Leitura com seleção de sheet / linha de header ---
    st.sidebar.write("### Selecione sheet e linha de cabeçalho (quando aplicável)")

    # SKUs file
    if arquivo_skus.name.lower().endswith(("xls", "xlsx")):
        xls_skus = pd.ExcelFile(arquivo_skus)
        # Sugerir sheet: prioriza sheet com nome SKU, ID, PRODUTO, ITENS, ou a primeira
        sheet_skus_sug = next((s for s in xls_skus.sheet_names if any(w in s.lower() for w in ["sku", "id", "produto", "itens"])), xls_skus.sheet_names[0])
        idx_sug = xls_skus.sheet_names.index(sheet_skus_sug)
        sheet_skus = st.selectbox("Sheet (SKUs)", xls_skus.sheet_names, index=idx_sug, key="sheet_skus")

        # Sugerir header: busca linha com maior match de aliases
        def _find_best_header(file, sheet, aliases, max_rows=10):
            tmp = pd.read_excel(file, sheet_name=sheet, header=None, nrows=max_rows)
            best = 0
            best_score = 0
            for i in range(len(tmp)):
                row = tmp.iloc[i].astype(str).str.lower().tolist()
                score = sum(any(a in c for a in aliases) for c in row)
                if score > best_score:
                    best = i
                    best_score = score
            return best

        aliases_header = ["sku", "produto", "item", "skc", "número", "numero"]
        header_skus_sug = _find_best_header(arquivo_skus, sheet_skus, aliases_header)
        header_row_skus_1b = st.number_input("Linha do cabeçalho (SKUs) - 1-based, 0=auto", min_value=0, max_value=50, value=header_skus_sug+1, key="header_skus")
        header_skus = header_row_skus_1b
        df_skus = pd.read_excel(arquivo_skus, sheet_name=sheet_skus, header=header_skus)
    else:
        header_row_skus_1b = st.number_input("Linha do cabeçalho (SKUs) - 1-based, 0=auto (CSV)", min_value=0, max_value=50, value=0, key="header_skus_csv")
        if header_row_skus_1b == 0:
            df_skus = pd.read_csv(arquivo_skus)
        else:
            df_skus = pd.read_csv(arquivo_skus, header=header_row_skus_1b)

    # Preços file
    if arquivo_precos.name.lower().endswith(("xls", "xlsx")):
        xls_precos = pd.ExcelFile(arquivo_precos)
        # default to PROMOÇÃO sheet when present (case-insensitive)
        def _find_default_sheet(names):
            for i, s in enumerate(names):
                if s.strip().lower() in ("promoção", "promocaO", "promocao"):
                    return i
                if s.strip().lower() in ("promoção", "promocao"):
                    return i
            return 0

        default_idx = next((i for i, s in enumerate(xls_precos.sheet_names) if s.strip().lower() in ("promoção", "promocao")), 0)
        sheet_precos = st.selectbox("Sheet (Preços)", xls_precos.sheet_names, index=default_idx, key="sheet_precos")
        header_row_precos_1b = st.number_input("Linha do cabeçalho (Preços) - 1-based, 0=auto", min_value=0, max_value=50, value=0, key="header_precos")
        if header_row_precos_1b == 0:
            detected_p = _auto_detect_header_from_excel(arquivo_precos, sheet_precos)
            header_precos = detected_p
            st.sidebar.write(f"Detectado header em linha {detected_p+1} para Preços")
        else:
            header_precos = header_row_precos_1b - 1
        # usamos a função que preserva fórmulas quando necessário
        try:
            df_precos = ler_excel_promocao_com_formulas(arquivo_precos, sheet_name=sheet_precos, header_row=header_precos)
        except Exception:
            df_precos = pd.read_excel(arquivo_precos, sheet_name=sheet_precos, header=header_precos)
    else:
        header_row_precos_1b = st.number_input("Linha do cabeçalho (Preços) - 1-based, 0=auto (CSV)", min_value=0, max_value=50, value=0, key="header_precos_csv")
        if header_row_precos_1b == 0:
            df_precos = pd.read_csv(arquivo_precos)
        else:
            df_precos = pd.read_csv(arquivo_precos, header=header_row_precos_1b - 1)

    # Normaliza nomes de colunas (remove espaços invisíveis)
    df_skus.columns = df_skus.columns.astype(str).str.strip()
    df_precos.columns = df_precos.columns.astype(str).str.strip()

    # Evita erro no Streamlit/PyArrow com nomes de coluna duplicados
    df_skus, skus_renamed = _deduplicate_columns(df_skus)
    df_precos, precos_renamed = _deduplicate_columns(df_precos)

    if skus_renamed:
        st.sidebar.warning("Foram detectadas colunas duplicadas na planilha de SKUs. Ajustamos os nomes automaticamente.")
    if precos_renamed:
        st.sidebar.warning("Foram detectadas colunas duplicadas na base de preços. Ajustamos os nomes automaticamente.")

    # Inicializa variáveis de detecção para evitar NameError
    num_item_col = None
    skc_col = None

    # Limpeza base de preços
    colunas_remover = [
        "descrição", "descricao", "valor a receber", "peso",
        "frete", "taxa", "redução", "reducao", "bruto", "publicação"
    ]

    df_precos = df_precos.loc[:, [c for c in df_precos.columns if not any(r in c.lower() for r in colunas_remover)]]

    # Remove colunas de marketplace do df_skus
    marketplaces = ["mercado", "shopee", "shein", "magalu", "netshoes", "kwai", "tiktok", "mercado livre"]
    df_skus = df_skus.loc[:, [c for c in df_skus.columns if not any(m in c.lower() for m in marketplaces)]]

    st.success("✅ Arquivos carregados")

    st.divider()

    # Primeiro escolha o marketplace (necessário para decisões seguintes)
    marketplace = st.selectbox(
        "Marketplace",
        ["Mercado Livre", "Shopee", "Shein", "Magalu","Netshoes", "Kwai", "TikTok", "Outro"],
        key="marketplace_select"
    )

    # --- Mapeamento de colunas (auto-sugestão básica) ---
    st.sidebar.write("### Mapeamento de colunas (ajuste se necessário)")

    # sugestões para SKUs
    skc_aliases = ["skc"]
    num_item_aliases = ["número do item", "numero do item", "item number", "item no", "numero item"]
    sku_aliases = ["sku", "id do anúncio", "id do anuncio", "id", "seller id"]

    # Recalcula sugestões sempre após leitura do df_skus
    def _auto_suggest(df, aliases):
        # Tenta por nome
        col = _suggest_column(df.columns, aliases)
        if col:
            return col
        # Tenta por conteúdo: procura coluna com valores que parecem SKU (alfanumérico, tamanho típico)
        for c in df.columns:
            vals = df[c].astype(str).str.upper().str.strip().dropna().unique()
            if any(len(v) >= 6 and v.isalnum() for v in vals[:10]):
                return c
        return None

    sugestao_num_item = _auto_suggest(df_skus, num_item_aliases)
    sugestao_skc = _auto_suggest(df_skus, skc_aliases)
    sugestao_sku = _auto_suggest(df_skus, sku_aliases)

    # Mostra SKC / Número do item apenas para Shein
    if marketplace.lower() == "shein":
        col_skc_sel = st.selectbox(
            "Coluna SKC (se existir)",
            ["(nenhuma)"] + list(df_skus.columns),
            index=(1 + list(df_skus.columns).index(sugestao_skc) if sugestao_skc in df_skus.columns else 0),
            key="col_skc_sel"
        )
        col_num_item_sel = st.selectbox(
            "Coluna Número do item (se existir)",
            ["(nenhuma)"] + list(df_skus.columns),
            index=(1 + list(df_skus.columns).index(sugestao_num_item) if sugestao_num_item in df_skus.columns else 0),
            key="col_num_item_sel"
        )
    else:
        col_skc_sel = None
        col_num_item_sel = None

    col_match_skus = st.selectbox(
        "Coluna de match (SKUs)",
        list(df_skus.columns),
        index=(list(df_skus.columns).index(sugestao_sku) if sugestao_sku in df_skus.columns else 0),
        key="col_match_skus_map"
    )

    # sugestões para preços
    price_aliases = ["preço", "preco", "valor", "price"]
    sugestao_price = _suggest_column(df_precos.columns, price_aliases)
    sugestao_precos_match = _suggest_column(df_precos.columns, sku_aliases + num_item_aliases)

    colunas_match_precos = [c for c in df_precos.columns if not any(m in c.lower() for m in marketplaces)]
    col_match_precos = st.selectbox(
        "Coluna de match (Preços)",
        colunas_match_precos,
        index=(colunas_match_precos.index(sugestao_precos_match) if sugestao_precos_match in colunas_match_precos else 0),
        key="col_match_precos_map"
    )
    col_preco = st.selectbox(
        "Coluna de Preço",
        list(df_precos.columns),
        index=(list(df_precos.columns).index(sugestao_price) if sugestao_price in df_precos.columns else 0),
        key="col_preco_map"
    )

    # Normaliza seleção de "(nenhuma)"
    if col_skc_sel == "(nenhuma)":
        col_skc_sel = None

    # Exibe pré-visualizações rápidas para ajudar a confirmar mapeamento
    st.sidebar.write("Preview SKUs:")
    st.sidebar.dataframe(df_skus.head(5))
    st.sidebar.write("Preview Preços:")
    st.sidebar.dataframe(df_precos.head(5))

# ================= PROCESSAMENTO =================

# 🔒 Coluna canônica de ID (NUNCA some)
df_skus["ID_BASE"] = df_skus[col_match_skus]

df_skus["_MERGE_KEY"] = (
    df_skus[col_match_skus]
    .astype(str)
).apply(_normalize_merge_key)



# Ajuste especial para Shein: prefira a coluna "Número do item" como chave de
# agrupamento (isso agrupa variações pelo SKU pai). Também detectamos a
# coluna `SKC` para uso no export quando disponível.
def _find_num_item_col(df):
    for c in df.columns:
        lc = c.lower().strip()
        if ("numero" in lc or "número" in lc) and "item" in lc:
            return c
    return None

def _find_skc_col(df):
    for c in df.columns:
        if c.strip().lower() == "skc" or "skc" in c.lower():
            return c
    return None

num_item_col = _find_num_item_col(df_skus)
skc_col = _find_skc_col(df_skus)

# Se o usuário mapeou explicitamente na sidebar, priorizamos essa seleção
try:
    if 'col_num_item_sel' in locals() and col_num_item_sel and col_num_item_sel in df_skus.columns:
        num_item_col = col_num_item_sel
except Exception:
    pass

try:
    if 'col_skc_sel' in locals() and col_skc_sel and col_skc_sel in df_skus.columns:
        skc_col = col_skc_sel
except Exception:
    pass

# Se estivermos trabalhando com Shein e o arquivo de SKUs contém o
# "Número do item", então usamos essa coluna para gerar o _MERGE_KEY em
# `df_skus` (para procurar no `df_precos`). Caso contrário mantemos o
# comportamento padrão (usar a coluna selecionada em `col_match_skus`).
if marketplace.lower() == "shein" and num_item_col and num_item_col in df_skus.columns:
    df_skus["_MERGE_KEY"] = (
        df_skus[num_item_col]
        .astype(str)
    ).apply(_normalize_merge_key)

df_precos["_MERGE_KEY"] = (
    df_precos[col_match_precos]
    .astype(str)
    .str.replace(r"\\.0$", "", regex=True)
    .str.replace("MLB", "", regex=False)
)

df_precos["_MERGE_KEY"] = df_precos["_MERGE_KEY"].str.split(",")


# Explode IDs múltiplos
df_precos = df_precos.explode("_MERGE_KEY")
df_precos["_MERGE_KEY"] = df_precos["_MERGE_KEY"].astype(str).str.replace(".0","", regex=False).str.strip()
df_precos["_MERGE_KEY"] = df_precos["_MERGE_KEY"].apply(_normalize_merge_key)

# Quando for Shein e houver `SKC`, vamos manter essa coluna para o export
# (o merge usa `Número do item` como chave pai). `skc_col` já foi detectado
# acima; se não existir, nada muda.

# Opção: colapsar linhas idênticas por SKC (útil para Shein)
collapse_skc = False
collapse_post_merge = False
if marketplace.lower() == "shein":
    # Mostrar quais colunas foram detectadas (ajuda debugging)
    st.sidebar.write(f"Detectado Número do item: {num_item_col}")
    st.sidebar.write(f"Detectado SKC: {skc_col}")
    collapse_skc = st.sidebar.checkbox("Colapsar linhas idênticas por SKC", value=True)

if collapse_skc and marketplace.lower() == "shein":
    # Se houver SKC em df_precos, colapsa antes do merge; se houver apenas em
    # df_skus, agendamos o colapso para depois do merge.
    if skc_col and skc_col in df_precos.columns:
        before = len(df_precos)

        def _first_nonnull(s):
            s2 = s.dropna()
            return s2.iloc[0] if not s2.empty else s.iloc[0]

        agg_map = {c: 'first' for c in df_precos.columns}
        if col_preco in df_precos.columns:
            agg_map[col_preco] = _first_nonnull
        agg_map['_MERGE_KEY'] = 'first'

        df_precos = df_precos.groupby(skc_col, as_index=False).agg(agg_map)
        after = len(df_precos)
        removed = before - after
        if removed > 0:
            st.sidebar.info(f"✅ Colapsadas {removed} linhas idênticas por SKC — agora {after} SKC únicos.")
    elif skc_col and skc_col in df_skus.columns:
        collapse_post_merge = True
        st.sidebar.info("ℹ️ SKC detectado apenas em SKUs — colapso será feito após o merge.")
    else:
        st.sidebar.warning("⚠️ SKC não detectado em SKUs nem em Preços — não foi possível colapsar por SKC.")


# Remove colisões (preserva ID_BASE)
colisoes = set(df_skus.columns) & set(df_precos.columns)
colisoes.discard("_MERGE_KEY")
colisoes.discard("ID_BASE")

df_skus_limpo = df_skus.drop(columns=list(colisoes))

# Merge
df_merged = df_skus_limpo.merge(df_precos, on="_MERGE_KEY", how="left")
df_merged.drop(columns="_MERGE_KEY", inplace=True)

# Se foi marcado para colapsar por SKC somente após o merge (quando SKC
# existe apenas no arquivo da Shein), aplicamos o agrupamento aqui.
if 'collapse_post_merge' in globals() and collapse_post_merge:
    if skc_col and skc_col in df_merged.columns:
        before = len(df_merged)

        def _first_nonnull(s):
            s2 = s.dropna()
            return s2.iloc[0] if not s2.empty else s.iloc[0]

        agg_map = {c: 'first' for c in df_merged.columns}
        if col_preco in df_merged.columns:
            agg_map[col_preco] = _first_nonnull

        df_merged = df_merged.groupby(skc_col, as_index=False).agg(agg_map)
        after = len(df_merged)
        removed = before - after
        if removed > 0:
            st.sidebar.info(f"✅ Colapsadas {removed} linhas idênticas por SKC após o merge — agora {after} SKC únicos.")

# ================= TABS =================

tab1, tab2, tab3 = st.tabs(["📋 Dados", "🔗 Match", "⬇️ Download"])

# ---------- TAB 1 ----------
with tab1:
    st.subheader("SKUs")
    st.dataframe(df_skus, use_container_width=True)

    st.subheader("Preços")
    st.dataframe(df_precos, use_container_width=True)

# ---------- TAB 2 ----------
with tab2:
    st.subheader("🔗 Resultado do Match")

    total = len(df_merged)
    matched = df_merged[col_preco].notna().sum()
    nao_matched = df_merged[col_preco].isna().sum()

    c1, c2, c3 = st.columns(3)
    c1.metric("Total SKUs", total)
    c2.metric("Matched", matched)
    c3.metric("Não encontrados", nao_matched)

    st.divider()

    st.write("### 📌 Amostra geral")
    st.dataframe(df_merged, use_container_width=True)

    st.divider()

    df_nao_encontrados = df_merged[df_merged[col_preco].isna()]

    if not df_nao_encontrados.empty:
        st.warning(f"⚠️ {len(df_nao_encontrados)} SKUs não tiveram match")

        if st.checkbox("🔍 Mostrar apenas SKUs não encontrados"):
            st.dataframe(
                df_nao_encontrados[["ID_BASE"]],
                use_container_width=True
            )
    else:
        st.success("🎉 Todos os SKUs tiveram match!")

# ---------- TAB 3 ----------
with tab3:
    df_final = df_merged[df_merged[col_preco].notna()].copy()

    # Trata #REF!, texto, etc
    df_final[col_preco] = pd.to_numeric(df_final[col_preco], errors="coerce")
    df_final[col_preco] = df_final[col_preco].round(2)

    # Para Shein: se detectamos a coluna `SKC`, exportamos `SKC` + preço de campanha
    if marketplace.lower() == "shein" and skc_col and skc_col in df_final.columns:
        df_export = df_final[[skc_col, col_preco]].copy()
        df_export = df_export.rename(columns={
            skc_col: "SKC (obrigatório)",
            col_preco: f"Preço de campanha"
        })
    else:
        df_export = df_final[["ID_BASE", col_preco]].copy()

        df_export = df_export.rename(columns={
            "ID_BASE": "ID",
            col_preco: f"Preço {marketplace}"
        })

    st.info(f"📊 {len(df_export)} registros prontos")

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False)

    st.download_button(
        "📥 Baixar Excel",
        buffer.getvalue(),
        file_name=f"promo_{marketplace}_{datetime.now():%d%m%Y_%H%M}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary"
    )
