import streamlit as st
import pandas as pd

# Configuração da página
st.set_page_config(page_title="Consolidador de Medicamentos", layout="centered")

st.title("Consolidador de Quantidades por Código Simpas")
st.write("Faça o upload do arquivo Excel (.xls). O sistema irá ler os dados a partir da linha 8, filtrar as colunas necessárias e somar as quantidades com base no Código Simpas.")

# Upload do arquivo (.xls)
arquivo_upado = st.file_uploader("Selecione o arquivo Excel", type=["xls"])

if arquivo_upado is not None:
    try:
        # Lendo o arquivo Excel (.xls) com o cabeçalho na linha 8 (índice 7)
        df = pd.read_excel(
            arquivo_upado, 
            engine='xlrd', 
            header=7, 
            usecols=["Código Simpas", "Medicamento", "Quantidade Encontrada"]
        )
        
        # Limpando possíveis espaços nas nomenclaturas das colunas
        df.columns = df.columns.str.strip()
        
        # Garantindo que a coluna de quantidade seja número (erros/nulos viram 0)
        df['Quantidade Encontrada'] = pd.to_numeric(df['Quantidade Encontrada'], errors='coerce').fillna(0)
        
        # ---------------------------------------------------------
        # MODIFICAÇÃO: Tratamento dos Códigos em Branco
        # Substitui valores nulos (NaN) por "CÓDIGO EM BRANCO"
        df['Código Simpas'] = df['Código Simpas'].fillna('CÓDIGO EM BRANCO')
        
        # Substitui também casos onde o código venha como texto vazio ou apenas espaços
        df['Código Simpas'] = df['Código Simpas'].replace(r'^\s*$', 'CÓDIGO EM BRANCO', regex=True)
        # ---------------------------------------------------------
        
        # Agrupando pelo "Código Simpas" (inclusive os "CÓDIGO EM BRANCO") e somando
        df_consolidado = df.groupby('Código Simpas', as_index=False, dropna=False).agg({
            'Medicamento': 'first', # Pega o primeiro nome de medicamento que aparecer no grupo
            'Quantidade Encontrada': 'sum'
        })
        
        st.success("Arquivo processado com sucesso!")
        
        # Mostrando o dataframe na tela
        st.subheader("Resultado Consolidado")
        st.dataframe(df_consolidado, use_container_width=True)
        
        # Botão para baixar o resultado
        csv = df_consolidado.to_csv(index=False, sep=';').encode('utf-8')
        st.download_button(
            label="Baixar resultado consolidado (CSV)",
            data=csv,
            file_name="simpas_consolidado.csv",
            mime="text/csv"
        )
        
    except ValueError as ve:
        st.error("Erro na leitura das colunas.")
        st.warning(f"Detalhe do erro: {ve}. Verifique se os nomes das colunas na linha 8 estão EXATAMENTE como solicitado ('Código Simpas', 'Medicamento', 'Quantidade Encontrada').")
    except Exception as e:
        st.error(f"Ocorreu um erro inesperado: {e}")