
import pandas as pd
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from io import StringIO
import csv
import locale
import os
from babel.dates import format_date

# Establecer configuración regional para meses en español
'''try:
    locale.setlocale(locale.LC_TIME, 'es_AR.utf8')
except:
    try:
        locale.setlocale(locale.LC_TIME, 'es_ES.utf8')
    except:
        locale.setlocale(locale.LC_TIME, '')'''


def cargar_datos_desde_sheets():
    url = "https://docs.google.com/spreadsheets/d/1qx1AHGXUkS1MnVlQerKQFDiJ4HHvti0zszDcBOFoJts"
    archivo_id = url.split("/d/")[1].split("/")[0]
    hoja_inscripciones = f"https://docs.google.com/spreadsheets/d/{archivo_id}/gviz/tq?tqx=out:csv&sheet=Inscripciones"

    # 👇 Acá forzamos el tipo str para Comisión
    inscripciones = inscripciones = pd.read_csv(hoja_inscripciones, quoting=csv.QUOTE_ALL, dtype=str)

    inscripciones.columns = inscripciones.columns.str.strip()
    return inscripciones

def generar_certificado(dni, incluir_sello=True):
    inscripciones = cargar_datos_desde_sheets()
    datos_estudiante = inscripciones[inscripciones["DNI"] == dni]

    if datos_estudiante.empty:
        print("DNI no encontrado.")
        return

    nombre = datos_estudiante.iloc[0]["Nombre"]
    carrera = datos_estudiante.iloc[0]["Carrera"]
    cuatrimestre = datos_estudiante.iloc[0]["Cuatrimestre"]

    fecha_dt = datetime.today()
    fecha = format_date(fecha_dt, format="d 'de' MMMM 'de' y", locale='es')
    nombre_pdf = f"Certificado_{dni}.pdf"

    doc = SimpleDocTemplate(nombre_pdf, pagesize=A4,
                            rightMargin=2.5*cm, leftMargin=2.5*cm,
                            topMargin=2.5*cm, bottomMargin=2.5*cm)
    elementos = []
    estilos = getSampleStyleSheet()

    # Estilos personalizados
    estilo_justificado = ParagraphStyle(
        'Justificado',
        parent=estilos["Normal"],
        alignment=4,
        leading=16,
    )
    estilo_negrita = ParagraphStyle(
        'Negrita',
        parent=estilos["Normal"],
        fontName="Helvetica-Bold"
    )
    estilo_derecha = ParagraphStyle(
        'Derecha',
        parent=estilos["Normal"],
        alignment=2
    )

    # Logo UNSAM
    if os.path.exists("static/logo_unsam.png"):
        elementos.append(Image("static/logo_unsam.png", width=4 * cm, height=2 * cm, hAlign="LEFT"))
        elementos.append(Spacer(1, 12))

    elementos.append(Paragraph(f"San Martín, {fecha}", estilo_derecha))
    elementos.append(Spacer(1, 12))

    texto = f"""
    Por medio de la presente se deja constancia que <b>{nombre}</b>, DNI: <b>{dni}</b>, se inscribió a las materias del <b>{cuatrimestre}</b> correspondiente a la carrera de <b>{carrera}</b>.
    Se extiende el presente certificado a la persona solicitante para ser presentado ante quien corresponda.
    """
    elementos.append(Paragraph(texto.strip(), estilo_justificado))
    elementos.append(Spacer(1, 24))

    # Armar tabla
    columnas = ["materia", "comision", "Horarios"]
    for col in columnas:
        if col not in datos_estudiante.columns:
            print(f"Columna faltante: {col}")
            return

    data = [["Materia", "Días y horarios"]]
    for _, row in datos_estudiante.iterrows():

        comision_valida = str(row["comision"]).strip()
        print(comision_valida)
        materia_texto = f"{row['materia'].strip()} ({comision_valida})"


        materia = Paragraph(materia_texto, estilo_negrita)
        horarios = Paragraph(str(row["Horarios"]).replace("\n", "<br/>"))
        data.append([materia, horarios])

    tabla = Table(data, colWidths=[9*cm, 7*cm])
    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 10),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    elementos.append(tabla)
    elementos.append(Spacer(1, 24))
    elementos.append(Paragraph("Saludos cordiales.", estilos["Normal"]))

    if incluir_sello and os.path.exists("sello.png"):
        elementos.append(Spacer(1, 36))
        elementos.append(Image("sello.png", width=5.5*cm, height=6.2*cm, hAlign="RIGHT"))

    doc.build(elementos)
    print(datos_estudiante[["materia", "comision"]])
    print(f"Certificado generado: {nombre_pdf}")

# Ejecutar para probar
#generar_certificado(str(95848914), incluir_sello=True)
