from flask import Flask, render_template, request, send_file
import os
import pandas as pd
from fpdf import FPDF
from PyPDF2 import PdfReader, PdfWriter
from io import BytesIO
from datetime import datetime  # Importación añadida para manejar objetos datetime

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

        # Convertir la fecha al formato DD/MM/AAAA usando pandas
        if 'Fecha de ingreso a la institución' in data.columns:
            data['Fecha de ingreso a la institución'] = pd.to_datetime(
                data['Fecha de ingreso a la institución'], 
                errors='coerce'  # Maneja errores convirtiendo valores no válidos a NaT
            ).dt.strftime("%d/%m/%Y")  # Formato DD/MM/AAAA

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
    
    pdf.set_xy(120, 52.5)  # Coordenada para Fecha afiliacion
    pdf.cell(0, 10, f"{data['Fecha afiliacion']}", ln=False)

    pdf.set_xy(50, 90)  # Coordenada para Unidad administrativa
    pdf.cell(0, 10, f"{data['Unidad operativa']}", ln=False)
    
    pdf.set_xy(33, 157)  # Coordenada para Apellido paterno
    pdf.cell(0, 10, f"{data['Apellido paterno']}", ln=False)

    pdf.set_xy(83, 157)  # Coordenada para Apellido materno
    pdf.cell(0, 10, f"{data['Apellido materno']}", ln=False)

    pdf.set_xy(127, 157)  # Coordenada para Nombres
    pdf.cell(0, 10, f"{data['Nombres']}", ln=True)

    if data['Sexo'].lower() == "femenino":
        pdf.set_xy(42, 172)  # Coordenadas para "Femenino"
        pdf.cell(5, 5, "X", ln=1)
    elif data['Sexo'].lower() == "masculino":
        pdf.set_xy(58, 172)  # Coordenadas para "Masculino"
        pdf.cell(5, 5, "X", ln=1)

    rfc = str(data["Registro federal de contribuyente"]).zfill(13)
    rfc_positions = [(67 + i * 6, 172) for i in range(13)]
    for i, char in enumerate(rfc):
        pdf.set_xy(*rfc_positions[i])
        pdf.cell(5, 5, char, ln=1)

    if data['Nacionalidad'].lower() == "mexicano":
        pdf.set_xy(157, 168)  # Coordenadas para "Mexicano"
        pdf.cell(5, 5, "X", ln=1)
    elif data['Nacionalidad'].lower() == "extranjero":
        pdf.set_xy(175, 168)  # Coordenadas para "extranjero"
        pdf.cell(5, 5, "X", ln=1)

    pdf.set_xy(30, 181)
    pdf.cell(0, 10, data['Domicilio'], ln=True)

    pdf.set_xy(170, 181)
    pdf.cell(0, 10, str(data['Edad']), ln=True)
    
    pdf.set_xy(30, 192)  # Coordenadas para Municipio/Alcaldía
    pdf.cell(0, 10, data['Municipio/alcaldía'], ln=True)

    # Estado
    pdf.set_xy(98, 192)  # Coordenadas para Estado
    pdf.cell(0, 10, data['Estado'], ln=True)

    # Código Postal
    cp = str(data["Código postal"]).zfill(5)  # Asegurarse de que tenga 5 caracteres, rellenando con ceros si es necesario
    cp_positions = [(158 + i * 6, 194.4) for i in range(5)]  # Coordenadas iniciales y espaciado para cada dígito
    for i, char in enumerate(cp):
        pdf.set_xy(cp_positions[i][0], cp_positions[i][1])  # Misma 'y' para todos, ajustando 'x'
        pdf.cell(5, 5, char, ln=False)  # ln=False evita que salte de línea después de cada dígito

    # Fecha de ingreso
    fecha_ingreso = data.get('Fecha de ingreso a la institución', '')
    if pd.notna(fecha_ingreso) and isinstance(fecha_ingreso, (datetime, pd.Timestamp)):
        fecha_ingreso = fecha_ingreso.strftime("%d/%m/%Y")
    elif isinstance(fecha_ingreso, str) and '/' in fecha_ingreso:
        # Si ya es un string en formato DD/MM/AAAA
        pass
    else:
        fecha_ingreso = ""  # Campo vacío

    if fecha_ingreso:
        dia, mes, anio = fecha_ingreso.split('/')  # Asume formato válido
        pdf.set_xy(29, 215)
        pdf.cell(10, 10, dia, ln=False)
        pdf.set_xy(42, 215)
        pdf.cell(10, 10, mes, ln=False)
        pdf.set_xy(53, 215)
        pdf.cell(10, 10, anio, ln=True)

    
    pdf.set_xy(71, 215)
    pdf.cell(0, 10, f"{data['Denominación de puesto']}", ln=True)

    pdf.set_xy(134, 211)
    pdf.cell(0, 10, f"{data['Servicio asignado']}", ln=True)

    # Fecha de ingreso a OPD
    fecha_opd = data.get('Fecha de ingreso a OPD', '')
    if pd.notna(fecha_opd) and isinstance(fecha_opd, (datetime, pd.Timestamp)):
        fecha_opd = fecha_opd.strftime("%d/%m/%Y")
    elif isinstance(fecha_opd, str) and '/' in fecha_opd:
        # Si ya es un string en formato DD/MM/AAAA
        pass
    else:
        fecha_opd = ""  # Campo vacío

    if fecha_opd:
        dia_opd, mes_opd, anio_opd = fecha_opd.split('/')
        pdf.set_xy(29, 240)
        pdf.cell(10, 10, dia_opd, ln=False)
        pdf.set_xy(42, 240)
        pdf.cell(10, 10, mes_opd, ln=False)
        pdf.set_xy(53, 240)
        pdf.cell(10, 10, anio_opd, ln=True)


    # Manejo de Función real
    funcion_real = data.get('Función real', '')
    if pd.notna(funcion_real) and funcion_real:
        pdf.set_xy(134, 233)
        pdf.cell(0, 10, funcion_real, ln=True)

    pdf.set_xy(115, 263)
    pdf.cell(0, 10, f"{data['Responsable de afiliación']}", ln=True)

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
