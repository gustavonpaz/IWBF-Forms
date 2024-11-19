import os
import pandas as pd
from flask import Flask, request, render_template
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.generic import NameObject, BooleanObject, DictionaryObject

app = Flask(__name__)

# Função para preencher os PDFs
def preencher_pdf(template_path, output_path, field_values):
    pdf_reader = PdfReader(template_path)
    pdf_writer = PdfWriter()

    # Copiar todas as páginas do modelo
    for page in pdf_reader.pages:
        pdf_writer.add_page(page)

    # Preencher os campos do formulário
    pdf_writer.update_page_form_field_values(
        pdf_writer.pages[0],  # Preenche os campos na primeira página
        field_values
    )

    # Configurar "/NeedAppearances" no dicionário do formulário
    if "/AcroForm" not in pdf_writer._root_object:
        pdf_writer._root_object[NameObject("/AcroForm")] = DictionaryObject()
    pdf_writer._root_object[NameObject("/AcroForm")][NameObject("/NeedAppearances")] = BooleanObject(True)

    # Salvar o PDF preenchido
    with open(output_path, "wb") as output_pdf:
        pdf_writer.write(output_pdf)

@app.route("/")
def home():
    return render_template("upload.html")

@app.route("/process", methods=["POST"])
def process():
    # Receber arquivos do formulário
    excel_file = request.files["excel"]
    worksheet_pdf = request.files["worksheet_pdf"]
    assessment_pdf = request.files["assessment_pdf"]

    # Salvar os arquivos temporariamente
    excel_path = "temp_players.xlsx"
    worksheet_path = "temp_worksheet.pdf"
    assessment_path = "temp_assessment.pdf"
    excel_file.save(excel_path)
    worksheet_pdf.save(worksheet_path)
    assessment_pdf.save(assessment_path)

    # Carregar Excel
    planilhas = pd.read_excel(excel_path, sheet_name=None)

    output_base_path = "Generated_PDFs"

    # Processar cada aba
    for aba, dados in planilhas.items():
        aba_path = os.path.join(output_base_path, aba)
        os.makedirs(aba_path, exist_ok=True)

        stages_subfolder_path = os.path.join(aba_path, "Stages 2C and 3")
        os.makedirs(stages_subfolder_path, exist_ok=True)

        for _, row in dados.iterrows():
            field_values_worksheet = {
                "number": str(row["number"]),
                "proposed-class": row["proposed-class"],
                "name": row["name"],
                "country": row["country"],
                "date": str(row["date"]),
                "competition": row["competition"],
            }

            field_values_assessment = {
                "name": row["name"],
                "country": row["country"],
                "dob": row["dob"].strftime("%d-%m-%Y") if isinstance(row["dob"], pd.Timestamp) else row["dob"],
            }

            worksheet_output_path = os.path.join(stages_subfolder_path, f"{row['name']}-Worksheet-Stages-2C-and-3.pdf")
            assessment_output_path = os.path.join(aba_path, f"{row['name']}-Assessment-Form-Stages-2AB.pdf")

            preencher_pdf(worksheet_path, worksheet_output_path, field_values_worksheet)
            preencher_pdf(assessment_path, assessment_output_path, field_values_assessment)

    return f"Processamento concluído! Verifique a pasta '{output_base_path}' para os PDFs gerados."

if __name__ == "__main__":
    app.run(debug=True)
