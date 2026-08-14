import streamlit as st
import pandas as pd
import os
from io import BytesIO
import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER

# ---------------------------------------------------------
# FUNÇÕES DE GERAÇÃO DE ARQUIVOS
# ---------------------------------------------------------
def gerar_excel(df):
    output = BytesIO()
    # Usando openpyxl para criar o xlsx formatado
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Consolidado')
        worksheet = writer.sheets['Consolidado']
        
        # Definindo a largura das colunas
        worksheet.column_dimensions['A'].width = 20  # Código Simpas
        worksheet.column_dimensions['B'].width = 60  # Medicamento
        worksheet.column_dimensions['C'].width = 25  # Quantidade
        
        # Aplicando a quebra de texto na coluna de Medicamento e alinhando ao topo
        wrap_format = openpyxl.styles.Alignment(wrap_text=True, vertical='top')
        top_format = openpyxl.styles.Alignment(vertical='top')
        
        for row in worksheet.iter_rows(min_row=2, max_col=3):
            row[0].alignment = top_format
            row[1].alignment = wrap_format
            row[2].alignment = top_format

    return output.getvalue()

def gerar_pdf(df, nome_do_arquivo):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    elementos = []
    
    estilos = getSampleStyleSheet()
    
    # --- NOVO: Adicionando o Cabeçalho Centralizado ---
    estilo_titulo = estilos['Title']
    estilo_titulo.alignment = TA_CENTER # Garante que fique no centro
    estilo_titulo.fontSize = 14
    
    # Cria o parágrafo do título com o nome do arquivo e adiciona um espaço abaixo dele
    titulo_pdf = Paragraph(f"Consolidado: {nome_do_arquivo}", estilo_titulo)
    elementos.append(titulo_pdf)
    elementos.append(Spacer(1, 20)) # Espaço de 20 pontos entre o título e a tabela
    # --------------------------------------------------
    
    estilo_medicamento = estilos['Normal']
    estilo_medicamento.fontSize = 9
    
    dados_tabela = [["Código Simpas", "Medicamento", "Quantidade Encontrada"]]
    
    for _, row in df.iterrows():
        med_paragrafo = Paragraph(str(row['Medicamento']), estilo_medicamento)
        dados_tabela.append([str(row['Código Simpas']), med_paragrafo, str(row['Quantidade Encontrada'])])
    
    larguras_colunas = [100, 335, 120] 
    
    tabela = Table(dados_tabela, colWidths=larguras_colunas)
    tabela.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#4F81BD")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.whitesmoke, colors.white])
    ]))
    
    elementos.append(tabela)
    doc.build(elementos)
    return buffer.getvalue()

# ---------------------------------------------------------
# APLICATIVO STREAMLIT
# ---------------------------------------------------------
st.set_page_config(page_title="Consolidador de Medicamentos", layout="centered")

st.title("Consolidador de Quantidades por Código Simpas")
st.write("Faça o upload do arquivo Excel. O sistema irá ler os dados a partir da linha 8, filtrar as colunas necessárias e somar as quantidades com base no Código Simpas.")

arquivo_upado = st.file_uploader("Selecione o arquivo Excel", type=["xls", "xlsx"])

if arquivo_upado is not None:
    try:
        # --- NOVO: Capturando o nome do arquivo ---
        nome_original = arquivo_upado.name
        # Separa o nome da extensão (ex: 'ARBOVIROSE' e '.xls')
        nome_sem_extensao, _ = os.path.splitext(nome_original)
        # ------------------------------------------

        try:
            df = pd.read_excel(
                arquivo_upado, 
                engine='openpyxl', 
                header=7, 
                usecols=["Código Simpas", "Medicamento", "Quantidade Encontrada"]
            )
        except Exception:
            arquivo_upado.seek(0)
            df = pd.read_excel(
                arquivo_upado, 
                engine='xlrd', 
                header=7, 
                usecols=["Código Simpas", "Medicamento", "Quantidade Encontrada"]
            )

        df.columns = df.columns.str.strip()
        df['Quantidade Encontrada'] = pd.to_numeric(df['Quantidade Encontrada'], errors='coerce').fillna(0)
        
        df['Código Simpas'] = df['Código Simpas'].fillna('CÓDIGO EM BRANCO')
        df['Código Simpas'] = df['Código Simpas'].replace(r'^\s*$', 'CÓDIGO EM BRANCO', regex=True)
        
        df_consolidado = df.groupby('Código Simpas', as_index=False, dropna=False).agg({
            'Medicamento': 'first',
            'Quantidade Encontrada': 'sum'
        })
        
        st.success(f"Arquivo '{nome_original}' processado com sucesso!")
        
        st.subheader("Resultado Consolidado")
        st.dataframe(df_consolidado, use_container_width=True)
        
        st.write("---")
        st.write("### Exportar Resultados")
        
        col1, col2 = st.columns(2)
        
        with col1:
            arquivo_xlsx = gerar_excel(df_consolidado)
            st.download_button(
                label="📥 Baixar em Excel (.xlsx)",
                data=arquivo_xlsx,
                file_name=f"{nome_sem_extensao}_consolidado.xlsx", # Usa o nome capturado
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        with col2:
            # Passamos o nome original para aparecer no título de dentro do PDF
            arquivo_pdf = gerar_pdf(df_consolidado, nome_original)
            st.download_button(
                label="📄 Baixar em PDF (.pdf)",
                data=arquivo_pdf,
                file_name=f"{nome_sem_extensao}_consolidado.pdf", # Usa o nome capturado
                mime="application/pdf"
            )
        
    except ValueError as ve:
        st.error("Erro na leitura das colunas.")
        st.warning(f"Detalhe do erro: {ve}. Verifique se os nomes das colunas na linha 8 estão EXATAMENTE como solicitado ('Código Simpas', 'Medicamento', 'Quantidade Encontrada').")
    except Exception as e:
        st.error(f"Ocorreu um erro inesperado: {e}")