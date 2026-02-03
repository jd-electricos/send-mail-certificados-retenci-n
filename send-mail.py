import json
import time
import win32com.client as win32
import os
import random
import pandas as pd
import unicodedata
from datetime import datetime

# ------------------------------
# Rutas base
# ------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_DIR = os.path.join(BASE_DIR, "data")
HTML_DIR = os.path.join(BASE_DIR, "html")

data_path = os.path.join(DATA_DIR, "data.json")
subject_path = os.path.join(DATA_DIR, "subject.json")

# Excel proveedores
EXCEL_PATH = r"C:\Users\andres.lopez.JDELECTRICOS\Documents\codes\send-mails-certificados-retenciones\providers\datos-proveedores.xlsx"

# Carpeta PDFs
PDF_DIR = r"C:\Users\andres.lopez.JDELECTRICOS\Documents\codes\send-mails-certificados-retenciones\pdf"

# ------------------------------
# Función para normalizar texto
# ------------------------------
def normalizar(texto):
    texto = unicodedata.normalize('NFKD', texto)
    texto = texto.encode('ascii', 'ignore').decode('ascii')
    return texto.replace(" ", "").upper()

# ------------------------------
# Cargar data JSON
# ------------------------------
with open(data_path, "r", encoding="utf-8") as file:
    data = json.load(file)

asesor = data["asesor"]
correo_asesor = data["correoAsesor"]

# ------------------------------
# Cargar subjects
# ------------------------------
with open(subject_path, "r", encoding="utf-8") as file:
    subjects = json.load(file)["subjects"]

# ------------------------------
# Cargar plantillas HTML
# ------------------------------
html_templates = []
for filename in sorted(os.listdir(HTML_DIR)):
    if filename.endswith(".html"):
        with open(os.path.join(HTML_DIR, filename), "r", encoding="utf-8") as file:
            html_templates.append(file.read())

# ------------------------------
# Leer proveedores desde Excel
# ------------------------------
df = pd.read_excel(EXCEL_PATH)

# Normalizar encabezados
df.columns = df.columns.astype(str).str.strip().str.lower()

proveedores = []
for _, row in df.iterrows():
    proveedor = row.get("proveedor")
    correo = row.get("correo")

    if pd.notna(proveedor):
        proveedores.append({
            "proveedor": str(proveedor).strip(),
            "correo": str(correo).strip() if pd.notna(correo) else ""
        })

# ------------------------------
# Outlook
# ------------------------------
outlook = win32.Dispatch('Outlook.Application')

# ------------------------------
# Variables de control
# ------------------------------
novedades = []
correos_enviados = 0

# ------------------------------
# Envío de correos
# ------------------------------
for proveedor in proveedores:
    nombre_proveedor = proveedor["proveedor"]
    correo_proveedor = proveedor["correo"]

    # ------------------------------
    # Validar correo
    # ------------------------------
    if not correo_proveedor or "@" not in correo_proveedor:
        novedades.append(f"{nombre_proveedor}: no tiene correo válido.")
        continue

    nombre_normalizado = normalizar(nombre_proveedor)

    # ------------------------------
    # Buscar PDF correspondiente
    # ------------------------------
    pdfs_encontrados = []

    for archivo in os.listdir(PDF_DIR):
        if (
            archivo.upper().endswith(".PDF")
            and archivo.upper().startswith(nombre_normalizado)
        ):
            pdfs_encontrados.append(os.path.join(PDF_DIR, archivo))

    if not pdfs_encontrados:
        novedades.append(f"{nombre_proveedor}: no tiene documentos PDF para enviar.")
        continue

    # ------------------------------
    # Crear y enviar correo
    # ------------------------------
    html_template = random.choice(html_templates)
    subject = random.choice(subjects)

    html_content = (
        html_template
        .replace("{{ cliente }}", nombre_proveedor)
        .replace("{{ asesor }}", asesor)
        .replace("{{ correoAsesor }}", correo_asesor)
    )

    mail = outlook.CreateItem(0)
    mail.To = correo_proveedor
    mail.Subject = subject
    mail.HTMLBody = html_content
    # Adjuntar TODOS los PDFs
    for pdf in pdfs_encontrados:
        mail.Attachments.Add(pdf)

    mail.Send()
    correos_enviados += 1

    print(f"✔ Enviado a: {nombre_proveedor} -> {correo_proveedor}")
    print(f"📎 PDFs adjuntos: {len(pdfs_encontrados)}")

    if len(pdfs_encontrados) > 1:
        novedades.append(
            f"{nombre_proveedor}: se enviaron {len(pdfs_encontrados)} documentos PDF."
        )

    # ------------------------------
    # Límite horario 5:30 PM
    # ------------------------------
    hora_actual = datetime.now()
    hora_limite = hora_actual.replace(hour=17, minute=30, second=0, microsecond=0)

    if hora_actual > hora_limite:
        novedades.append("Proceso detenido por límite horario (5:30 PM).")
        break

    # Espera para pruebas
    wait_time = random.uniform(3, 6)
    print(f"⏱ Esperando {wait_time:.1f} segundos...\n")
    time.sleep(wait_time)

# ------------------------------
# Enviar correo de reporte FINAL
# ------------------------------
reporte = []
reporte.append("REPORTE ENVÍO CORREOS RETENCIONES\n")
reporte.append(f"Correos enviados correctamente: {correos_enviados}\n")

if novedades:
    reporte.append("NOVEDADES DETECTADAS:")
    for n in novedades:
        reporte.append(f"- {n}")
else:
    reporte.append("No se presentaron novedades.")

reporte.append("\nEl proceso de envío ha finalizado correctamente.")

reporte_texto = "\n".join(reporte)

mail_reporte = outlook.CreateItem(0)
mail_reporte.To = "abicdev26@gmail.com"
mail_reporte.Subject = "reporte envio correo de retenciones"
mail_reporte.Body = reporte_texto
mail_reporte.Send()

print("📧 Correo de reporte enviado a abicdev26@gmail.com")
print("\n🎉 Proceso finalizado.")
