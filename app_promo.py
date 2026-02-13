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


def detectar_coluna_id(df):
    possiveis = [
        "id", "id anuncio", "id do anuncio", "id do anúncio",
        "anuncio", "anúncio", "sku", "codigo", "código"
    ]

    for col in df.columns:
        nome = str(col).lower().strip().replace("_", " ")
        if nome in possiveis:
            return col
    return None


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

    # Limpeza da base de preços
    colunas_remover = [
        "descrição", "descricao", "valor a receber", "peso",
        "frete", "taxa", "redução", "reducao", "bruto", "publicação"
    ]

    df_precos = df_precos.loc[
        :,
        [
            c for c in df_precos.columns
            if not any(r in c.lower() for r in colunas_remover)
        ]
    ]

    # Remove marketplaces do df_skus
    marketplaces = ["mercado", "shopee", "shein", "magalu", "netshoes", "kwai", "tiktok", "mercado livre"]
    df_skus = df_skus.loc[
        :,
        [c for c in df_skus.columns if not any(m in c.lower() for m in marketplaces)]
    ]

    st.success("✅ Arquivos carregados")

    st.divider()
    
    # Match
    col_match_skus = st.selectbox("Coluna de match (SKUs)", df_skus.columns)
    col_match_precos = st.selectbox("Coluna de match (Preços)", [col for col in df_precos.columns if col.lower() not in marketplaces])

    # Marketplace
    marketplace = st.selectbox(
        "Marketplace",
        ["Mercado Livre", "Shopee", "Shein", "Magalu"]
    )

    col_preco = st.selectbox(
        "Coluna de Preço",
        [c for c in df_precos.columns if marketplace.lower() in c.lower()]
    )

# ================= PROCESSAMENTO =================

# Cria chaves temporárias
df_skus["_MERGE_KEY"] = df_skus[col_match_skus].astype(str).str.replace("MLB", "").str.strip()
df_precos["_MERGE_KEY"] = (
    df_precos[col_match_precos]
    .astype(str)
    .str.replace("MLB", "")
    .str.split(",")
)

df_precos = df_precos.explode("_MERGE_KEY")

df_precos["_MERGE_KEY"] = df_precos["_MERGE_KEY"].str.strip()
# Remove colisões
colisoes = set(df_skus.columns) & set(df_precos.columns)
colisoes.discard("_MERGE_KEY")
df_skus_limpo = df_skus.drop(columns=list(colisoes))

# Merge seguro
df_merged = df_skus_limpo.merge(df_precos, on="_MERGE_KEY", how="left")
df_merged.drop(columns="_MERGE_KEY", inplace=True)

# ================= TABS =================

tab1, tab2, tab3 = st.tabs(["📋 Dados", "🔗 Match", "⬇️ Download"])

with tab1:
    st.subheader("SKUs")
    st.dataframe(df_skus.head(10), use_container_width=True)

    st.subheader("Preços")
    st.dataframe(df_precos.head(10), use_container_width=True)

with tab2:
    st.subheader("🔗 Resultado do Match")

    total = len(df_merged)
    matched = df_merged[col_preco].notna().sum()
    nao_matched = df_merged[col_preco].isna().sum()

    col1, col2, col3 = st.columns(3)
    col1.metric("Total SKUs", total)
    col2.metric("Matched", matched)
    col3.metric("Não encontrados", nao_matched)

    st.divider()

    st.write("### 📌 Amostra geral (com match e sem match)")
    st.dataframe(df_merged.head(20), use_container_width=True)

    st.divider()

    # ================= NÃO ENCONTRADOS =================
    df_nao_encontrados = df_merged[df_merged[col_preco].isna()]

    if not df_nao_encontrados.empty:
        st.warning(f"⚠️ {len(df_nao_encontrados)} SKUs não tiveram match")

        if st.checkbox("🔍 Mostrar apenas SKUs não encontrados"):
            st.dataframe(
                df_nao_encontrados[[col_match_skus]],
                use_container_width=True
            )
    else:
        st.success("🎉 Todos os SKUs tiveram match!")

with tab3:
    df_final = df_merged[df_merged[col_preco].notna()].copy()
    df_final[col_preco] = pd.to_numeric(df_final[col_preco], errors="coerce")
    df_final[col_preco] = df_final[col_preco].apply(lambda x: round(x, 2) if pd.notna(x) else x)
    # Mantém apenas ID e preço do marketplace selecionado
    df_export = df_final[[col_match_skus, col_preco]].copy()

    # Renomeia colunas para ficar bonito no arquivo final
    df_export = df_export.rename(columns={
        col_match_skus: "ID",
        col_preco: f"Preço {marketplace}"
    })


    st.info(f"📊 {len(df_final)} registros prontos")

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
