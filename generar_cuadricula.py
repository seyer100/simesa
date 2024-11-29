from PyPDF2 import PdfReader, PdfWriter
from fpdf import FPDF

class GridPDF(FPDF):
    def __init__(self, template_path):
        super().__init__()
        self.template_path = template_path
        self.page_width = 216  # Ancho de la página Carta (mm)
        self.page_height = 279  # Alto de la página Carta (mm)

    def add_grid(self, output_pdf):
        # Crear una página con la cuadrícula
        self.add_page()
        self.set_font("Arial", size=6)
        self.set_text_color(255, 0, 0)

        # Dibujar líneas horizontales (cada 10 mm)
        for y in range(0, self.page_height, 10):  # Ajusta el paso según lo necesites
            self.line(0, y, self.page_width, y)  # Línea horizontal
            self.set_xy(0, y)
            self.cell(0, 5, f"Y={y}", ln=1)

        # Dibujar líneas verticales (cada 10 mm)
        for x in range(0, self.page_width, 10):  # Ajusta el paso según lo necesites
            self.line(x, 0, x, self.page_height)  # Línea vertical
            self.set_xy(x, 0)
            self.cell(5, 5, f"X={x}")

        # Guardar la cuadrícula en un archivo temporal
        self.output(output_pdf)

# Ruta de la plantilla y archivo de salida
template_path = "templates_pdf/plantilla.pdf"  # Ruta a tu plantilla
output_pdf = "plantilla_con_cuadricula.pdf"

# Crear PDF con cuadrícula
grid_pdf = GridPDF(template_path)
grid_pdf.add_grid(output_pdf)

# Ahora combinar la cuadrícula con la plantilla
reader = PdfReader(template_path)
writer = PdfWriter()

# Leer la plantilla
template_page = reader.pages[0]

# Leer la cuadrícula generada
grid_reader = PdfReader(output_pdf)
grid_page = grid_reader.pages[0]

# Fusionar las páginas
template_page.merge_page(grid_page)

# Escribir el resultado final
with open("final_plantilla_con_cuadricula.pdf", "wb") as f:
    writer.add_page(template_page)
    writer.write(f)

print("PDF con cuadrícula generado correctamente.")


