import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Conversor PDF → Excel", layout="centered")

st.title("📄 Conversor de Contrato PDF → Excel (Unificado)")
st.write("""
Esta ferramenta extrai automaticamente todas as tabelas contendo **CÓDIGO | DESCRIÇÃO | VALOR** 
de contratos em PDF e une tudo em uma única tabela Excel pronta para subir no sistema.
""")

uploaded_file = st.file_uploader("Envie o contrato em PDF", type=["pdf"])

if uploaded_file:
    with st.spinner("🔍 Processando o PDF..."):
        tabelas_extraidas = []

        try:
            with pdfplumber.open(uploaded_file) as pdf:
                for pagina in pdf.pages:
                    tabelas = pagina.extract_tables()
                    for tabela in tabelas:
                        df = pd.DataFrame(tabela)
                        # Identifica se a primeira linha contém algo como "CÓDIGO"
                        if df.shape[1] >= 3 and any(df.iloc[0].astype(str).str.contains("CÓDIGO", case=False, na=False)):
                            # Define cabeçalhos
                            df.columns = ["CÓDIGO", "DESCRIÇÃO", "VALOR", *df.columns[3:len(df.columns)]]
                            df = df.drop(0)  # Remove a linha de cabeçalho duplicado
                            tabelas_extraidas.append(df)

            if tabelas_extraidas:
                # Junta todas as tabelas em uma só
                resultado = pd.concat(tabelas_extraidas, ignore_index=True)
                resultado = resultado.iloc[:, :3]  # Mantém apenas as 3 primeiras colunas
                resultado.columns = ["CÓDIGO", "DESCRIÇÃO", "VALOR"]

                # Limpa dados em branco e duplicados
                resultado = resultado.dropna(how="all")
                resultado = resultado[resultado["CÓDIGO"].astype(str).str.strip() != ""]
                resultado = resultado.drop_duplicates()

                # Cria arquivo Excel na memória
                output = BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    resultado.to_excel(writer, index=False, sheet_name="Tabelas Unificadas")

                st.success(f"✅ Extração concluída! {len(resultado)} linhas extraídas.")
                st.dataframe(resultado.head(20))
                st.download_button(
                    label="⬇️ Baixar Excel Unificado",
                    data=output.getvalue(),
                    file_name="tabelas_unificadas.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("⚠️ Nenhuma tabela contendo 'CÓDIGO' foi encontrada neste PDF.")
        except Exception as e:
            st.error(f"❌ Erro ao processar o PDF: {e}")
else:
    st.info("📥 Faça upload de um arquivo PDF para iniciar a conversão.")
