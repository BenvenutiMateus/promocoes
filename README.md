Aplicação desenvolvida em Python + Streamlit para cruzamento automático de SKUs/IDs com bases de preços, focada em operações de e-commerce e marketplaces.

O sistema permite importar planilhas do Excel ou CSV, inclusive com valores provenientes de fórmulas, identificar automaticamente campos de identificação (ID, SKU, código), realizar o match seguro entre bases, visualizar registros encontrados e não encontrados, e exportar apenas os preços finais do marketplace selecionado, prontos para publicação em promoções.

🚀 Principais funcionalidades

Upload de planilhas de SKUs/IDs e base de preços

Leitura correta de arquivos Excel com fórmulas calculadas

Detecção automática de colunas de ID / SKU

Match flexível entre diferentes identificadores

Visualização de registros com e sem match

Tratamento automático de erros de fórmula (#REF!, textos, valores inválidos)

Exportação limpa contendo apenas ID + preço do marketplace escolhido

Interface simples, rápida e intuitiva via Streamlit

🛠 Tecnologias utilizadas

Python

Pandas

Streamlit

OpenPyXL

🎯 Objetivo

Automatizar e reduzir erros no processo de criação de promoções em marketplaces, economizando tempo operacional e garantindo consistência nos preços publicados.