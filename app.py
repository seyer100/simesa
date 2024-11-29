from flask import Flask, render_template, request, send_file, redirect, url_for
import os
import pandas as pd
from fpdf import FPDF
from io import BytesIO

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
PDF_FOLDER = "pdfs"
TEMPLATES_PDF_FOLDER = "templates_pdf"  # Para la plantilla
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PDF_FOLDER, exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return render_template("index.html", error_message="No se seleccionó ningún archivo.")

    file = request.files["file"]
    if file.filename == "":
        return render_template("index.html", error_message="El archivo no tiene nombre.")

    # Guardar el archivo
    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(file_path)

    # Procesar el Excel y generar PDFs
    try:
        data = pd.read_excel(file_path)
        pdf_files = []

        # Cargar plantilla PDF
        template_path = os.path.join(TEMPLATES_PDF_FOLDER, "plantilla.pdf")

        # Generar un PDF por cada fila
        for index, row in data.iterrows():
            pdf_in_memory = generate_pdf(row, template_path)
            pdf_files.append(pdf_in_memory)

        # Combinar PDFs en memoria
        combined_pdf = combine_pdfs(pdf_files)

        # Enviar el PDF combinado como respuesta para descarga
        return send_file(combined_pdf, as_attachment=True, download_name="combined.pdf", mimetype='application/pdf')

    except Exception as e:
        return render_template("index.html", error_message=f"Error al procesar el archivo: {str(e)}")


def generate_pdf(data, template_path):
    """Genera un PDF en memoria a partir de una plantilla y los datos"""
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    # Aquí se carga la plantilla y se rellena con los datos (modificar según necesidades)
    pdf.image(template_path, x=0, y=0, w=210, h=297)  # Tamaño A4 estándar

    # Llenado de campos (ajustar las posiciones)
    pdf.set_xy(40, 50)
    pdf.cell(0, 10, f"{data['Nombres']} {data['Apellido paterno']} {data['Apellido materno']}", ln=True)

    pdf.set_xy(40, 60)
    pdf.cell(0, 10, f"Sexo: {data['Sexo']}", ln=True)

    # Añadir más campos según sea necesario
    # ...

    # Guardar el PDF en memoria en lugar de en un archivo
    pdf_output = BytesIO()
    pdf.output(pdf_output)
    pdf_output.seek(0)  # Volver al inicio del archivo en memoria
    return pdf_output

def combine_pdfs(pdf_files):
    """Combina múltiples archivos PDF en uno solo, todo en memoria"""
    from PyPDF2 import PdfMerger

    merger = PdfMerger()

    # Añadir los PDFs en memoria
    for pdf in pdf_files:
        merger.append(pdf)

    # Guardar el PDF combinado en memoria
    combined_pdf_output = BytesIO()
    merger.write(combined_pdf_output)
    combined_pdf_output.seek(0)  # Volver al inicio
    return combined_pdf_output


if __name__ == "__main__":
    app.run(debug=True)
