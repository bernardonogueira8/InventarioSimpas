import streamlit as st
import pandas as pd

# Configuração da página
st.set_page_config(page_title="Consolidador de Medicamentos", layout="centered")

st.title("Consolidador de Quantidades por Código Simpas")
st.write("Faça o upload do arquivo Excel. O sistema irá ler os dados a partir da linha 8, filtrar as colunas necessárias e somar as quantidades com base no Código Simpas.")

arquivo_upado = st.file_uploader("Selecione o arquivo Excel", type=["xls", "xlsx"])

if arquivo_upado is not None:
    try:
        # TENTATIVA 1: Tenta ler com 'openpyxl' (Resolve o caso do "ARBOVIROSE.xls" que é um xlsx disfarçado)
        try:
            df = pd.read_excel(
                arquivo_upado, 
                engine='openpyxl', 
                header=7, 
                usecols=["Código Simpas", "Medicamento", "Quantidade Encontrada"]
            )
        except Exception:
            # TENTATIVA 2: Se falhar, reseta a leitura do arquivo e tenta com 'xlrd' (Para arquivos .xls originais)
            arquivo_upado.seek(0)
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
        
        # Tratamento dos Códigos em Branco
        df['Código Simpas'] = df['Código Simpas'].fillna('CÓDIGO EM BRANCO')
        df['Código Simpas'] = df['Código Simpas'].replace(r'^\s*$', 'CÓDIGO EM BRANCO', regex=True)
        
        # Agrupando pelo "Código Simpas" e somando
        df_consolidado = df.groupby('Código Simpas', as_index=False, dropna=False).agg({
            'Medicamento': 'first',
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
        # Se cair aqui, mostrará o erro final para podermos investigar
        st.error(f"Ocorreu um erro inesperado: {e}")