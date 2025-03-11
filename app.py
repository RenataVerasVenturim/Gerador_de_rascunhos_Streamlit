import os
import pdfquery
import openpyxl
import shutil
import pandas as pd
import streamlit as st
from io import BytesIO

# Diretórios
directory_path = os.getcwd()
pdfs_directory = os.path.join(directory_path, 'pdfs')
drafts_directory = os.path.join(directory_path, 'RascunhosGerados')
excel_filename = "Modelo.xlsx"
excel_path = os.path.join(directory_path, excel_filename)

# Criar diretórios se não existirem
os.makedirs(pdfs_directory, exist_ok=True)
os.makedirs(drafts_directory, exist_ok=True)

# Função para esvaziar pastas
def empty_folder(folder_path):
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if os.path.isfile(file_path):
            os.remove(file_path)

# Definição das coordenadas dos elementos no PDF
coordinates = [
    {'left': 200.0, 'top': 549.52, 'width': 16.68, 'height': 10.0},    # Número de empenho
    {'left': 41.0, 'top': 418.52, 'width': 374.62, 'height': 10.0},    # Fornecedor (nome e CNPJ)
    {'left': 421.0, 'top': 642.52, 'width': 50.02, 'height': 10.0},    # Valor da nota
    {'left': 200.0, 'top': 464.52, 'width': 139.57, 'height': 10.0},   # Número do processo
    {'left': 200.0, 'top': 503.52, 'width': 56.7, 'height': 10.0},     # Fonte de despesa
    {'left': 43.0, 'top': 627.52, 'width': 387.29, 'height': 10.0},    # Natureza da despesa
    {'left': 125.0, 'top': 306.52, 'width': 122.66, 'height': 10.0},   # Modalidade da licitação
    {'left': 122.0, 'top': 503.52, 'width': 33.36, 'height': 10.0},    # PTRES
    {'left': 296.0, 'top': 503.52, 'width': 33.36, 'height': 10.0},    # Nº da natureza da despesa
    {'left': 485.0, 'top': 503.52, 'width': 73.88, 'height': 10.0},    # Plano interno
    {'left': 407.0, 'top': 503.52, 'width': 33.36, 'height': 10.0},    # UGR
]

# Interface no Streamlit
st.title("Extração de Dados de PDFs para Excel")
st.markdown("Envie arquivos PDF para extrair informações e gerar uma planilha.")

# Upload de arquivos
uploaded_files = st.file_uploader("Selecione os PDFs", type="pdf", accept_multiple_files=True)

if uploaded_files:
    # Limpar diretórios antes de processar
    empty_folder(pdfs_directory)
    empty_folder(drafts_directory)

    # Copiar modelo de Excel
    shutil.copy(excel_path, os.path.join(pdfs_directory, "Temp_Consolidado.xlsx"))

    # Abrir o Excel copiado
    copied_workbook = openpyxl.load_workbook(os.path.join(pdfs_directory, "Temp_Consolidado.xlsx"))
    copied_sheet = copied_workbook.active

    start_row = 2
    empenho_number = None

    # Processamento dos PDFs
    for i, pdf_file in enumerate(uploaded_files):
        pdf_path = os.path.join(pdfs_directory, pdf_file.name)
        
        # Salvar PDF temporariamente
        with open(pdf_path, "wb") as f:
            f.write(pdf_file.read())

        pdf = pdfquery.PDFQuery(pdf_path)
        pdf.load()

        for j, coord in enumerate(coordinates):
            target_left = coord['left']
            target_top = coord['top']
            target_width = coord['width']
            target_height = coord['height']

            element = pdf.pq('LTTextLineHorizontal:in_bbox("%s, %s, %s, %s")' % 
                            (target_left, target_top, target_left + target_width, target_top + target_height))
            text = element.text().strip()
            copied_sheet.cell(row=start_row + i, column=j + 1).value = text

            if j == 0:  # Número de empenho
                empenho_number = text

    # Definir o nome do arquivo final
    if empenho_number:
        new_filename = f"Rascunho_inicial-{empenho_number}.xlsx"
    else:
        new_filename = "Consolidado.xlsx"

    final_path = os.path.join(drafts_directory, new_filename)
    copied_workbook.save(final_path)
    copied_workbook.close()

    st.success(f"Planilha gerada: **{new_filename}**")

    # Criar arquivo em memória para download
    with open(final_path, "rb") as file:
        excel_bytes = file.read()
    
    st.download_button(
        label="📥 Baixar Planilha",
        data=excel_bytes,
        file_name=new_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
