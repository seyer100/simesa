from flask import Flask, render_template, request, send_file
import os
import pandas as pd
from fpdf import FPDF
from PyPDF2 import PdfReader, PdfWriter
from io import BytesIO

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
TEMPLATES_PDF_FOLDER = "templates_pdf"  # Carpeta para la plantilla PDF
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

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

    # Guardar el archivo Excel cargado
    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(file_path)

    try:
        # Leer los datos del archivo Excel
        data = pd.read_excel(file_path)
        pdf_files = []

        # Ruta de la plantilla
        template_path = os.path.join(TEMPLATES_PDF_FOLDER, "plantilla.pdf")

        # Generar un PDF para cada fila
        for _, row in data.iterrows():
            pdf_in_memory = generate_pdf(row, template_path)
            pdf_files.append(pdf_in_memory)

        # Combinar todos los PDFs en un solo archivo en memoria
        combined_pdf = combine_pdfs(pdf_files)

        # Descargar el archivo combinado
        return send_file(combined_pdf, as_attachment=True, download_name="combined.pdf", mimetype='application/pdf')

    except Exception as e:
        return render_template("index.html", error_message=f"Error al procesar el archivo: {str(e)}")


def generate_pdf(data, template_path):
    """
    Genera un PDF en memoria combinando una plantilla y los datos.
    """
    # Cargar la plantilla PDF
    reader = PdfReader(template_path)
    page = reader.pages[0]  # Asumimos una sola página en la plantilla

    # Crear un archivo PDF temporal con los datos
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=10)

    # Llenar los datos en posiciones específicas (ajustar según plantilla)

    # Datos personales
    pdf.set_xy(40, 50)  # Posición para el nombre completo
    pdf.cell(0, 10, f"{data['Nombres']} {data['Apellido paterno']} {data['Apellido materno']}", ln=True)

    # Sexo
    if data['Sexo'].lower() == "femenino":
        pdf.set_xy(50, 80)  # Posición para "Femenino"
        pdf.cell(5, 5, "X", ln=1)
    elif data['Sexo'].lower() == "masculino":
        pdf.set_xy(80, 80)  # Posición para "Masculino"
        pdf.cell(5, 5, "X", ln=1)

    # RFC - Separar por dígitos
    rfc_positions = [(50 + i * 10, 90) for i in range(13)]  # Ajustar las posiciones
    for i, char in enumerate(data['Registro federal de contribuyente']):
        pdf.set_xy(*rfc_positions[i])
        pdf.cell(5, 5, char, ln=1)

    # Domicilio
    pdf.set_xy(40, 100)
    pdf.cell(0, 10, data['Domicilio'], ln=True)

    # Edad
    pdf.set_xy(40, 110)
    pdf.cell(0, 10, str(data['Edad']), ln=True)

    # Municipio/Alcaldía
    pdf.set_xy(40, 120)
    pdf.cell(0, 10, data['Municipio/alcaldía'], ln=True)

    # Estado
    pdf.set_xy(40, 130)
    pdf.cell(0, 10, data['Estado'], ln=True)

    # Código postal - Separar por dígitos
    cp_positions = [(50 + i * 10, 140) for i in range(5)]
    for i, char in enumerate(data['Código postal']):
        pdf.set_xy(*cp_positions[i])
        pdf.cell(5, 5, char, ln=1)

    # Datos laborales
    pdf.set_xy(40, 150)
    pdf.cell(0, 10, f"Fecha de ingreso: {data['Fecha de ingreso a la institución']}", ln=True)

    pdf.set_xy(40, 160)
    pdf.cell(0, 10, f"Puesto: {data['Denominación de puesto']}", ln=True)

    pdf.set_xy(40, 170)
    pdf.cell(0, 10, f"Servicio: {data['Servicio asignado']}", ln=True)

    pdf.set_xy(40, 180)
    pdf.cell(0, 10, f"Fecha de ingreso a OPD: {data['Fecha de ingreso a OPD']}", ln=True)

    pdf.set_xy(40, 190)
    pdf.cell(0, 10, f"Función Real: {data['Función real']}", ln=True)

    # Responsable de afiliación
    pdf.set_xy(40, 200)
    pdf.cell(0, 10, f"Responsable: {data['Responsable de afiliación']}", ln=True)

    # Guardar el PDF con datos en memoria
    data_pdf = BytesIO()
    pdf.output(data_pdf)
    data_pdf.seek(0)

    # Superponer el contenido del nuevo PDF sobre la plantilla
    writer = PdfWriter()
    data_reader = PdfReader(data_pdf)
    template_page = page
    data_page = data_reader.pages[0]
    template_page.merge_page(data_page)

    # Crear un PDF combinado en memoria
    combined_pdf = BytesIO()
    writer.add_page(template_page)
    writer.write(combined_pdf)
    combined_pdf.seek(0)
    return combined_pdf


def combine_pdfs(pdf_files):
    """
    Combina múltiples PDFs en un solo archivo en memoria.
    """
    writer = PdfWriter()

    for pdf in pdf_files:
        reader = PdfReader(pdf)
        for page in reader.pages:
            writer.add_page(page)

    # Guardar el archivo combinado en memoria
    combined_pdf = BytesIO()
    writer.write(combined_pdf)
    combined_pdf.seek(0)
    return combined_pdf


if __name__ == "__main__":
    app.run(debug=True)
