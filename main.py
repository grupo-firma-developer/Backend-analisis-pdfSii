from fastapi import FastAPI, Response, UploadFile, File
import shutil
import os
import pdfplumber
import re
import pandas as pd
from pathlib import Path
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from reportlab.pdfgen import canvas
import pdfkit
from fpdf import FPDF
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape  



app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:4200"], 
    allow_credentials=True,
    allow_methods=["*"],  
    allow_headers=["*"],  
)

# Directorios
BASE_DIR = Path(__file__).resolve().parent
PDF_SIN_ANALIZAR = BASE_DIR / "pdf_sin_analizar"
PDF_ANALIZADOS = BASE_DIR / "pdf_analizados"
EXCELS_GENERADOS = BASE_DIR / "excels_generados"
PDF_GENERADOS = BASE_DIR / "pdfs_generados"
# Crear directorios si no existen
for folder in [PDF_SIN_ANALIZAR, PDF_ANALIZADOS, EXCELS_GENERADOS, PDF_GENERADOS]:
    folder.mkdir(parents=True, exist_ok=True)

# Expresiones regulares
#pattern_periodo = re.compile(r"PERIODO\s+(\d{2}\s\d{2}\s/\s\d{4})")
pattern_periodo = re.compile(r"PERIODO\s+\d{2}\s(\d{2}\s/\s\d{4})")
pattern_142 = re.compile(r"VENTAS Y/O SERV. EXENTOS O NO\s[^\d\-−]*([\-−]?[\d,.]+)")
pattern_537 = re.compile(r"TOTAL CRÉDITOS\s[^\d\-−]*([\-−]?[\d,.]+)")
pattern_538 = re.compile(r"TOTAL DÉBITOS\s[^\d\-−]*([\-−]?[\d,.]+)")

@app.post("/subir_pdf/")
async def subir_pdf(file: UploadFile = File(...)):
    file_path = PDF_SIN_ANALIZAR / file.filename
    with file_path.open("wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    
    return {"mensaje": "Archivo subido exitosamente", "nombre": file.filename}


def limpiar_numero(valor):
    if valor == "N/A":
        return 0
    return int(valor.replace('\u2212', '-').replace(',', '').replace('.', ''))

@app.get("/procesar_pdf/{filename}")
def procesar_pdf(filename: str):
    pdf_path = PDF_SIN_ANALIZAR / filename
    if not pdf_path.exists():
        return {"error": "Archivo no encontrado"}

    datos = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                match_periodo = pattern_periodo.search(text)
                periodo = match_periodo.group(1) if match_periodo else "Sin Periodo"
                
                valores_142 = pattern_142.findall(text)
                valores_537 = pattern_537.findall(text)
                match_538 = pattern_538.search(text)
                valores_538 = [match_538.group(1)] if match_538 else []
                
                max_length = max(len(valores_142), len(valores_537), len(valores_538))
                valores_142 += ["N/A"] * (max_length - len(valores_142))
                valores_537 += ["N/A"] * (max_length - len(valores_537))
                valores_538 += ["N/A"] * (max_length - len(valores_538))
                
                for i in range(max_length):
                    try:
                        #val_142 = limpiar_numero(valores_142[i]) if valores_142[i] != "N/A" else 0
                        #val_537 = limpiar_numero(valores_537[i]) if valores_537[i] != "N/A" else 0
                        #val_538 = limpiar_numero(valores_538[i]) if valores_538[i] != "N/A" else 0

                        val_142 = limpiar_numero(valores_142[i])
                        val_537 = limpiar_numero(valores_537[i])
                        val_538 = limpiar_numero(valores_538[i])


                        ventas_netas = (val_538 / 0.19) + val_142
                        compras_netas = val_537 / 0.19
                        ventas_netas_m = ventas_netas / 1000
                        compras_netas_m = compras_netas / 1000
                        margen = ventas_netas - compras_netas
                    except ValueError:
                        ventas_netas = compras_netas = ventas_netas_m = compras_netas_m = margen = "Error"
                    
                    datos.append({
                        "PERIODO": periodo,
                        "538": valores_538[i],
                        "537": valores_537[i],
                        "142": valores_142[i],
                        "Ventas Netas": ventas_netas,
                        "Compras Netas": compras_netas,
                        "Ventas Netas M$": ventas_netas_m,
                        "Compras Netas M$": compras_netas_m,
                        "Margen": margen
                    })
    
    df = pd.DataFrame(datos)

    df.replace(["N/A", "NaN", None, pd.NA], 0, inplace=True)

    columnas_a_formatear = ["Ventas Netas", "Compras Netas", "Ventas Netas M$", "Compras Netas M$", "Margen"]

    for col in columnas_a_formatear:
        if col in df.columns and pd.api.types.is_numeric_dtype(df[col]):
            df[col] = df[col].apply(lambda x: f"$ {int(x):,}".replace(",", ".") if pd.notna(x) else x)

    excel_filename = filename.replace(".pdf", ".xlsx")
    excel_path = EXCELS_GENERADOS / excel_filename
    df.to_excel(excel_path, sheet_name="Datos", index=False)
    
    shutil.move(pdf_path, PDF_ANALIZADOS / filename)
    
    return {"mensaje": "Procesamiento exitoso", "archivo_excel": excel_filename}

@app.get("/descargar_excel/{filename}")
def descargar_excel(filename: str):
    excel_path = EXCELS_GENERADOS / filename
    if not excel_path.exists():
        return {"error": "Archivo no encontrado"}
    
    with open(excel_path, "rb") as f:
        excel_data = f.read()

    return Response(
        content=excel_data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.get("/descargar_pdf/{filename}")
def descargar_pdf(filename: str):
    excel_path = EXCELS_GENERADOS / filename
    if not excel_path.exists():
        return {"error": "Archivo Excel no encontrado"}

    pdf_filename = filename.replace(".xlsx", ".pdf")
    pdf_path =  PDF_GENERADOS / pdf_filename

    try:
        df = pd.read_excel(excel_path)

        df.replace(["N/A", "NaN", None, pd.NA], 0, inplace=True)

        columnas_a_redondear = [
            "Ventas Netas",
            "Compras Netas",
            "Ventas Netas M$",
            "Compras Netas M$",
            "Margen"
        ]

        for col in columnas_a_redondear:
            if col in df.columns and pd.api.types.is_numeric_dtype(df[col]):
                df[col] = df[col].apply(lambda x: f"${int(x):,}".replace(",",".") if pd.notna(x) else x)

        pdf_path_str = str(pdf_path)

        # Usar orientación horizontal (landscape)
        doc = SimpleDocTemplate(pdf_path_str, pagesize=landscape(letter))

        table_data = [df.columns.tolist()] + df.values.tolist()

        # Ajustar el ancho de las columnas (puedes ajustar estos valores)
        col_widths = [85] * len(df.columns)  # Ancho fijo para cada columna

        table = Table(table_data, colWidths=col_widths) #asignamos el ancho de las columnas

        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 0), (-1, -1), 6), #Reducimos el tamaño de la fuente
        ]))

        elements = [table]
        doc.build(elements)

        with open(pdf_path_str, "rb") as f:
            pdf_data = f.read()

        return Response(
            content=pdf_data,
            media_type="application/pdf",
            headers={"Content-Disposition": f"attachment; filename={pdf_filename}"}
        )

    except Exception as e:
        return {"error": f"Error procesando el PDF: {str(e)}"}

