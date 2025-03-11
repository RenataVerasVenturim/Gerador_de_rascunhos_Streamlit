from pathlib import Path
import os
import pdfquery
import openpyxl
import shutil
import streamlit as st
from io import BytesIO

def empty_folder(folder_path):
    """Esvazia a pasta especificada."""
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if os.path.isfile(file_path):
            os.remove(file_path)
    print(f"Pasta '{folder_path}' esvaziada.")

def main():
    st.title('Gerador de Rascunhos a partir de PDFs')

    # Nome do arquivo Excel
    excel_filename = "Modelo.xlsx"

    # Carregar arquivo Excel
    excel_path = Path(__file__).parent / excel_filename

    # Carregar arquivos PDF
    uploaded_files = st.file_uploader("Carregar arquivos PDF", type=["pdf"], accept_multiple_files=True)

    if uploaded_files:
        # Definir coordenadas dos elementos desejados
        coordinates = [
            {'left': 200.0, 'top': 549.52, 'width': 16.68, 'height': 10.0},
            {'left': 41.0, 'top': 418.52, 'width': 374.62, 'height': 10.0},
            {'left': 421.0, 'top': 642.52, 'width': 50.02, 'height': 10.0},
            {'left': 200.0, 'top': 464.52, 'width': 139.57, 'height': 10.0},
            {'left': 200.0, 'top': 503.52, 'width': 56.7, 'height': 10.0},
            {'left': 43.0, 'top': 627.52, 'width': 387.29, 'height': 10.0},
            {'left': 125.0, 'top': 306.52, 'width': 122.66, 'height': 10.0},
            {'left': 122.0, 'top': 503.52, 'width': 33.36, 'height': 10.0},
            {'left': 296.0, 'top': 503.52, 'width': 33.36, 'height': 10.0},
            {'left': 485.0, 'top': 503.52, 'width': 73.88, 'height': 10.0},
            {'left': 407.0, 'top': 503.52, 'width': 33.36, 'height': 10.0},
        ]

        # Carregar o Excel
        copied_workbook = openpyxl.load_workbook(excel_path)
        copied_sheet = copied_workbook.active
        st.write('Excel aberto para inserção dos dados')

        start_row = 2
        empenho_number = None

        # Processar os arquivos PDF
        for i, pdf_file in enumerate(uploaded_files):
            pdf_path = BytesIO(pdf_file.getbuffer())  # Lê o arquivo PDF diretamente da memória

            pdf = pdfquery.PDFQuery(pdf_path)
            pdf.load()

            for j, coord in enumerate(coordinates):
                target_left = coord['left']
                target_top = coord['top']
                target_width = coord['width']
                target_height = coord['height']
                
                element = pdf.pq('LTTextLineHorizontal:in_bbox("%s, %s, %s, %s")' % (target_left, target_top, target_left + target_width, target_top + target_height))
                text = element.text().strip()
                copied_sheet.cell(row=start_row + i, column=j + 1).value = text
                
                if j == 0:  # Supondo que o número de empenho está na primeira coordenada
                    empenho_number = text

        # Definir o nome final do arquivo
        if empenho_number:
            new_filename = f"Rascunho inicial-{empenho_number}.xlsx"
        else:
            new_filename = "Consolidado.xlsx"

        # Salvar o arquivo gerado
        final_path = Path(__file__).parent / new_filename
        copied_workbook.save(final_path)
        copied_workbook.close()

        st.write(f"Dados inseridos na planilha: {new_filename}")

        # Exibir link de download
        st.download_button(
            label="Baixar Excel Gerado",
            data=open(final_path, 'rb').read(),
            file_name=new_filename,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

if __name__ == '__main__':
    main()
