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
        # Leer los datos del archivo Excel y limpiar nombres de columnas
        data = pd.read_excel(file_path)
        data.columns = data.columns.str.strip()  # Eliminar espacios en los nombres de las columnas

        pdf_files = []
        template_path = os.path.join(TEMPLATES_PDF_FOLDER, "plantilla.pdf")

        # Generar un PDF para cada fila
        for index, row in data.iterrows():
            if pd.isna(row["Nombres"]):  # Validar si "Nombres" está vacío
                print(f"Fila {index} tiene un valor nulo en 'Nombres'. Saltando...")
                continue
            pdf_in_memory = generate_pdf(row, template_path)
            pdf_files.append(pdf_in_memory)

        # Combinar todos los PDFs en un solo archivo en memoria
        combined_pdf = combine_pdfs(pdf_files)

        # Descargar el archivo combinado
        return send_file(combined_pdf, as_attachment=True, download_name="combined.pdf", mimetype='application/pdf')

    except Exception as e:
        import traceback
        error_details = traceback.format_exc()
        print("Error detallado:", error_details)
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

    # Llenar los datos en posiciones específicas
    pdf.set_xy(40, 50)
    pdf.cell(0, 10, f"{data['Nombres']} {data['Apellido paterno']} {data['Apellido materno']}", ln=True)

    if data['Sexo'].lower() == "femenino":
        pdf.set_xy(50, 80)
        pdf.cell(5, 5, "X", ln=1)
    elif data['Sexo'].lower() == "masculino":
        pdf.set_xy(80, 80)
        pdf.cell(5, 5, "X", ln=1)

    rfc = str(data["Registro federal de contribuyente"]).zfill(13)
    rfc_positions = [(50 + i * 10, 90) for i in range(13)]
    for i, char in enumerate(rfc):
        pdf.set_xy(*rfc_positions[i])
        pdf.cell(5, 5, char, ln=1)

    pdf.set_xy(40, 100)
    pdf.cell(0, 10, data['Domicilio'], ln=True)

    pdf.set_xy(40, 110)
    pdf.cell(0, 10, str(data['Edad']), ln=True)

    pdf.set_xy(40, 120)
    pdf.cell(0, 10, data['Municipio/alcaldía'], ln=True)

    pdf.set_xy(40, 130)
    pdf.cell(0, 10, data['Estado'], ln=True)

    cp = str(data["Código postal"]).zfill(5)
    cp_positions = [(50 + i * 10, 140) for i in range(5)]
    for i, char in enumerate(cp):
        pdf.set_xy(*cp_positions[i])
        pdf.cell(5, 5, char, ln=1)

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

    pdf.set_xy(40, 200)
    pdf.cell(0, 10, f"Responsable: {data['Responsable de afiliación']}", ln=True)

    # Guardar el PDF como string y escribirlo en BytesIO
    pdf_content = pdf.output(dest='S').encode('latin1')  # El contenido del PDF como bytes
    data_pdf = BytesIO(pdf_content)
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
        # Asegurarse de que el archivo en memoria esté al inicio
        pdf.seek(0)  
        reader = PdfReader(pdf)  # Leer desde BytesIO
        for page in reader.pages:
            writer.add_page(page)

    # Guardar el archivo combinado en memoria
    combined_pdf = BytesIO()
    writer.write(combined_pdf)
    combined_pdf.seek(0)  # Volver al inicio del archivo en memoria
    return combined_pdf


if __name__ == "__main__":
    app.run(debug=True)
