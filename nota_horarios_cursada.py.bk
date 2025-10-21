import pandas as pd
from datetime import datetime
import fitz  # PyMuPDF
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm

# Cargar datos desde Google Sheets (publicado)
def cargar_datos_desde_sheets():
    url = "https://docs.google.com/spreadsheets/d/1qx1AHGXUkS1MnVlQerKQFDiJ4HHvti0zszDcBOFoJts"
    archivo_id = url.split("/d/")[1].split("/")[0]
    hoja_inscripciones = f"https://docs.google.com/spreadsheets/d/{archivo_id}/gviz/tq?tqx=out:csv&sheet=Inscripciones"
    inscripciones = pd.read_csv(hoja_inscripciones)
    return inscripciones

inscripciones = cargar_datos_desde_sheets()
print(inscripciones.columns.tolist())

# Formatear tabla de materias
def formatear_tabla(materias):
    bloques = ["Materia | Días y horarios | Comisión", "-" * 60]
    for _, row in materias.iterrows():
        nombre_materia = row["materia"]
        horarios = row["Horarios"]
        comision = row["Comisión"]
        bloques.append(f"{nombre_materia} | {horarios} | {comision}")
    return bloques

# Crear PDF directamente

def generar_certificado(dni, incluir_sello=True):
    inscripciones = cargar_datos_desde_sheets()
    datos_estudiante = inscripciones[inscripciones["DNI"] == dni]

    if datos_estudiante.empty:
        print("DNI no encontrado.")
        return

    nombre = datos_estudiante.iloc[0]["Nombre"]
    carrera = datos_estudiante.iloc[0]["Carrera"]
    cuatrimestre = datos_estudiante.iloc[0]["Cuatrimestre"]
    tabla_materias = formatear_tabla(datos_estudiante)

    fecha = datetime.today().strftime('%d de %B de %Y')
    nombre_pdf = f"Certificado_{dni}.pdf"

    c = canvas.Canvas(nombre_pdf, pagesize=A4)
    c.setFont("Helvetica", 11)
    c.drawString(2.5*cm, 27*cm, f"San Martín, {fecha}")

    texto = f"""
Por medio de la presente se deja constancia que {nombre}, DNI: {dni}, se inscribió a las materias del {cuatrimestre} correspondiente a la carrera de {carrera}.
Se extiende el presente certificado a la persona solicitante para ser presentado ante quien corresponda.
"""
    textobject = c.beginText(2.5*cm, 25.5*cm)
    for linea in texto.strip().split("\n"):
        textobject.textLine(linea)
    c.drawText(textobject)

    # Materias
    textobject = c.beginText(2.5*cm, 22*cm)
    for bloque in tabla_materias:
        for linea in bloque.strip().split("\n"):
            textobject.textLine(linea)
        textobject.textLine("")
    c.drawText(textobject)

    c.drawString(2.5*cm, 3*cm, "Saludos cordiales.")
    c.save()

    if incluir_sello:
        print("(Sello digital aún no implementado en esta versión)")

    print(f"Certificado generado: {nombre_pdf}")

generar_certificado(42564981)