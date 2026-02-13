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

    col_match_skus = st.selectbox("Coluna de match (SKUs)", df_skus.columns)
    col_match_precos = st.selectbox(
        "Coluna de match (Preços)",
        [c for c in df_precos.columns if c.lower() not in marketplaces]
    )

    marketplace = st.selectbox(
        "Marketplace",
        ["Mercado Livre", "Shopee", "Shein", "Magalu"]
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
    .str.replace(r"\.0$", "", regex=True)  # remove .0 no final
    .str.replace("MLB", "", regex=False)
    .str.strip()
)


df_precos["_MERGE_KEY"] = (
    df_precos[col_match_precos]
    .astype(str)
    .str.replace(r"\.0$", "", regex=True)  # remove .0 no final
    .str.replace("MLB", "", regex=False)
    .str.split(",")
)


# Explode IDs múltiplos
df_precos = df_precos.explode("_MERGE_KEY")
df_precos["_MERGE_KEY"] = df_precos["_MERGE_KEY"].replace(".0","", regex=False).str.strip()

df_precos["_MERGE_KEY"] = (
    df_precos["_MERGE_KEY"]
    .astype(str)
    .str.replace(r"\.0$", "", regex=True)
    .str.strip()
)


# Remove colisões (preserva ID_BASE)
colisoes = set(df_skus.columns) & set(df_precos.columns)
colisoes.discard("_MERGE_KEY")
colisoes.discard("ID_BASE")

df_skus_limpo = df_skus.drop(columns=list(colisoes))

# Merge
df_merged = df_skus_limpo.merge(df_precos, on="_MERGE_KEY", how="left")
df_merged.drop(columns="_MERGE_KEY", inplace=True)

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
    st.dataframe(df_merged.head(20), use_container_width=True)

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
