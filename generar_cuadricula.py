from fpdf import FPDF

class GridPDF(FPDF):
    def __init__(self, template_path):
        super().__init__()
        self.template_path = template_path

    def add_grid(self):
        self.add_page()
        self.image(self.template_path, x=0, y=0, w=210, h=297)  # Tamaño A4 estándar
        self.set_font("Arial", size=6)
        self.set_text_color(255, 0, 0)

        # Dibujar líneas horizontales
        for y in range(0, 297, 10):  # Ajusta el paso según lo necesites
            self.line(0, y, 210, y)  # Línea horizontal
            self.set_xy(0, y)
            self.cell(0, 5, f"Y={y}", ln=1)

        # Dibujar líneas verticales
        for x in range(0, 210, 10):  # Ajusta el paso según lo necesites
            self.line(x, 0, x, 297)  # Línea vertical
            self.set_xy(x, 0)
            self.cell(5, 5, f"X={x}")

# Crear un PDF con la cuadrícula
template_path = "templates_pdf/plantilla.pdf"  # Ruta a tu plantilla
grid_pdf = GridPDF(template_path)
grid_pdf.add_grid()
grid_pdf.output("plantilla_con_cuadricula.pdf")
