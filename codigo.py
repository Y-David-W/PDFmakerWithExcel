import os
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from pdfrw import PdfReader, PdfWriter, PageMerge

# Archivos
pdf_base = "plantilla.pdf"
excel_file = "nombres.xlsx"
output_folder = "PDFs_Alumnos"

os.makedirs(output_folder, exist_ok=True)

# Mira el excel
df = pd.read_excel(excel_file)

def agregar_datos_a_pdf(row, output_path):
    temp_pdf = os.path.join(output_folder, "temp.pdf")

    # Crear PDF temporal con los datos
    c = canvas.Canvas(temp_pdf, pagesize=letter)
    c.setFont("Helvetica-Bold", 10)

    # Coordenadas en la que va a escribir cada datos
    c.drawString(160, 670, str(row["Alumno"]))
    c.drawString(160, 650, str(row["Ciclo"]))
    c.drawString(160, 633, str(row["Curso académico"]))
    c.drawString(430, 633, str(row["Grupo"]))
    c.drawString(160, 620, str(row["Año escolar"]))
    c.drawString(160, 605, str(row["Módulo profesional"]))

    c.save()

    # Escribe con el PDF base
    base_pdf = PdfReader(pdf_base)
    overlay_pdf = PdfReader(temp_pdf)

    base_page = base_pdf.pages[0]
    overlay_page = overlay_pdf.pages[0]

    PageMerge(base_page).add(overlay_page).render()

    PdfWriter(output_path, trailer=base_pdf).write()

# Generar PDF con bucle for hasta que termine
for _, row in df.iterrows():
    nombre_limpio = str(row["Alumno"]).replace(" ", "_")
    output_path = os.path.join(output_folder, f"Anexo1_{nombre_limpio}.pdf")

    agregar_datos_a_pdf(row, output_path)

print(f"Proceso completado, se generaron {len(df)} PDFs.")
