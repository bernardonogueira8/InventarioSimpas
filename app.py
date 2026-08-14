import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# Funções de geração de arquivos
def gerar_excel(df):
    output = BytesIO()
    # Usando openpyxl para criar o xlsx formatado
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Consolidado')
        worksheet = writer.sheets['Consolidado']
        
        # Definindo a largura das colunas
        worksheet.column_dimensions['A'].width = 20  # Código Simpas
        worksheet.column_dimensions['B'].width = 60  # Medicamento (Mais largo)
        worksheet.column_dimensions['C'].width = 25  # Quantidade
        
        # Aplicando a quebra de texto na coluna de Medicamento e alinhando tudo ao topo
        wrap_format = openpyxl.styles.Alignment(wrap_text=True, vertical='top')
        top_format = openpyxl.styles.Alignment(vertical='top')
        
        for row in worksheet.iter_rows(min_row=2, max_col=3):
            row[0].alignment = top_format       # Código
            row[1].alignment = wrap_format      # Medicamento quebrando a linha
            row[2].alignment = top_format       # Quantidade

    return output.getvalue()

def gerar_pdf(df):
    buffer = BytesIO()
    # A4 vertical padrão, definindo as margens
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    elementos = []
    
    # Criando o estilo para o texto da coluna de medicamentos (permite quebrar linha)
    estilos = getSampleStyleSheet()
    estilo_medicamento = estilos['Normal']
    estilo_medicamento.fontSize = 9
    
    # Cabeçalho da tabela
    dados_tabela = [["Código Simpas", "Medicamento", "Quantidade Encontrada"]]
    
    # Inserindo os dados (Transformando o texto do medicamento num Parágrafo para quebrar a linha automaticamente)
    for _, row in df.iterrows():
        med_paragrafo = Paragraph(str(row['Medicamento']), estilo_medicamento)
        dados_tabela.append([str(row['Código Simpas']), med_paragrafo, str(row['Quantidade Encontrada'])])
    
    # Largura total de uma página A4 utilizável é aprox. 555 pontos.
    # Distribuindo a largura para caber exatamente em 1 página
    larguras_colunas = [100, 335, 120] 
    
    tabela = Table(dados_tabela, colWidths=larguras_colunas)
    tabela.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#4F81BD")), # Cor do cabeçalho
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'), # Alinha o texto sempre na parte de cima da célula
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.whitesmoke, colors.white]) # Efeito zebrado nas linhas
    ]))
    
    elementos.append(tabela)
    doc.build(elementos)
    return buffer.getvalue()


# Configuração da página principal
st.set_page_config(page_title="Consolidador de Medicamentos", layout="centered")

st.title("Consolidador de Quantidades por Código Simpas")
st.write("Faça o upload do arquivo Excel. O sistema irá ler os dados a partir da linha 8, filtrar as colunas necessárias e somar as quantidades com base no Código Simpas.")

arquivo_upado = st.file_uploader("Selecione o arquivo Excel", type=["xls", "xlsx"])

if arquivo_upado is not None:
    try:
        # TENTATIVA 1: Ler com openpyxl (xlsx)
        try:
            df = pd.read_excel(
                arquivo_upado, 
                engine='openpyxl', 
                header=7, 
                usecols=["Código Simpas", "Medicamento", "Quantidade Encontrada"]
            )
        except Exception:
            # TENTATIVA 2: Ler com xlrd (xls)
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
        
        st.success("Arquivo processado com sucesso!")
        
        st.subheader("Resultado Consolidado")
        st.dataframe(df_consolidado, use_container_width=True)
        
        st.write("---")
        st.write("### Exportar Resultados")
        
        # Colocando os botões lado a lado
        col1, col2 = st.columns(2)
        
        with col1:
            # Botão XLSX
            arquivo_xlsx = gerar_excel(df_consolidado)
            st.download_button(
                label="📥 Baixar em Excel (.xlsx)",
                data=arquivo_xlsx,
                file_name="simpas_consolidado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        with col2:
            # Botão PDF
            arquivo_pdf = gerar_pdf(df_consolidado)
            st.download_button(
                label="📄 Baixar em PDF (.pdf)",
                data=arquivo_pdf,
                file_name="simpas_consolidado.pdf",
                mime="application/pdf"
            )
        
    except ValueError as ve:
        st.error("Erro na leitura das colunas.")
        st.warning(f"Detalhe do erro: {ve}. Verifique se os nomes das colunas na linha 8 estão EXATAMENTE como solicitado ('Código Simpas', 'Medicamento', 'Quantidade Encontrada').")
    except Exception as e:
        st.error(f"Ocorreu um erro inesperado: {e}")