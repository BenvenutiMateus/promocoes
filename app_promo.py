import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from openpyxl import load_workbook

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

    # Leitura
    df_skus = pd.read_csv(arquivo_skus) if arquivo_skus.name.endswith("csv") else pd.read_excel(arquivo_skus)
    df_precos = (
        pd.read_csv(arquivo_precos)
        if arquivo_precos.name.endswith("csv")
        else ler_excel_promocao_com_formulas(arquivo_precos)
    )
    # Normaliza nomes de colunas (remove espaços invisíveis)
    df_skus.columns = df_skus.columns.astype(str).str.strip()
    df_precos.columns = df_precos.columns.astype(str).str.strip()

    # Inicializa variáveis de detecção para evitar NameError
    num_item_col = None
    skc_col = None

    # Limpeza base de preços
    colunas_remover = [
        "descrição", "descricao", "valor a receber", "peso",
        "frete", "taxa", "redução", "reducao", "bruto", "publicação"
    ]

    df_precos = df_precos.loc[
        :,
        [c for c in df_precos.columns if not any(r in c.lower() for r in colunas_remover)]
    ]

    # Remove colunas de marketplace do df_skus
    marketplaces = ["mercado", "shopee", "shein", "magalu", "netshoes", "kwai", "tiktok", "mercado livre"]
    df_skus = df_skus.loc[
        :,
        [c for c in df_skus.columns if not any(m in c.lower() for m in marketplaces)]
    ]

    st.success("✅ Arquivos carregados")

    st.divider()

    # Primeiro escolha o marketplace (necessário para decisões seguintes)
    marketplace = st.selectbox(
        "Marketplace",
        ["Mercado Livre", "Shopee", "Shein", "Magalu","Netshoes", "Kwai", "TikTok", "Outro"]
    )

    # Detecta 'Número do item' no arquivo da Shein (se existir, usamos como match)
    num_item_col_detected = next(
        (c for c in df_skus.columns if (("numero" in c.lower() or "número" in c.lower()) and "item" in c.lower())),
        None
    )
    num_item_col = num_item_col_detected
    if num_item_col_detected and marketplace.lower() == "shein":
        st.sidebar.write(f"Detectado Número do item: {num_item_col_detected} — usando como coluna de match")
        col_match_skus = num_item_col_detected
    else:
        col_match_skus = st.selectbox("Coluna de match (SKUs)", df_skus.columns)

    col_match_precos = st.selectbox(
        "Coluna de match (Preços)",
        [c for c in df_precos.columns if c.lower() not in marketplaces]
    )

    col_preco = st.selectbox(
        "Coluna de Preço",
        [c for c in df_precos.columns if marketplace.lower() in c.lower()]
    )

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
