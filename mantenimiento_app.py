import streamlit as st
import pandas as pd
from datetime import date
import os
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

# ===============================
# CONFIGURACIÓN
# ===============================
DATA_FILE = "data/equipos.xlsx"
EXPORT_FOLDER = "exports"

# Crear carpetas si no existen
os.makedirs("data", exist_ok=True)
os.makedirs(EXPORT_FOLDER, exist_ok=True)

# ===============================
# Cargar y guardar datos
# ===============================
def load_data():
    if os.path.exists(DATA_FILE):
        return pd.read_excel(DATA_FILE)
    else:
        df = pd.DataFrame(columns=[
            "Tipo", "Departamento", "Sucursal", "Responsable", "Posicion",
            "Nombre de Equipo", "Correo", "Fecha de Mantenimiento", "Hora"
        ])
        df.to_excel(DATA_FILE, index=False)
        return df

def save_data(df):
    df.to_excel(DATA_FILE, index=False)

# ===============================
# Generar PDF formateado
# ===============================
def export_pdf(df, filename, rango_semana):
    df = df.copy()

    # Formatear columnas
    if "Fecha de Mantenimiento" in df.columns:
        df["Fecha de Mantenimiento"] = pd.to_datetime(df["Fecha de Mantenimiento"], errors="coerce").dt.strftime("%d-%b-%Y")
    if "Hora" in df.columns:
        df["Hora"] = pd.to_datetime(df["Hora"], errors="coerce").dt.strftime("%I:%M %p")

    # Crear documento PDF
    doc = SimpleDocTemplate(filename, pagesize=landscape(letter), leftMargin=25, rightMargin=25, topMargin=25, bottomMargin=25)
    styles = getSampleStyleSheet()
    elements = []

    # Título principal
    title = Paragraph("Calendario de Mantenimiento Preventivo", styles['Title'])
    subtitle = Paragraph(rango_semana, styles['Normal'])
    elements.append(title)
    elements.append(subtitle)
    elements.append(Spacer(1, 12))

    # Ajustar anchos de columnas dinámicamente
    col_widths = [max(80, min(150, len(str(col)) * 7)) for col in df.columns]

    # Crear tabla
    data = [list(df.columns)] + df.values.tolist()
    table = Table(data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.gray),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.whitesmoke, colors.lightgrey]),
    ]))
    elements.append(table)
    doc.build(elements)

# ===============================
# STREAMLIT APP
# ===============================
st.set_page_config(page_title="Mantenimiento Preventivo", layout="wide")
st.title("🛠️ Mantenimiento Preventivo de Computadoras")

df = load_data()

# ----- FORMULARIO PARA AGREGAR EQUIPO -----
st.subheader("Agregar nuevo equipo")
with st.form("agregar_equipo"):
    col1, col2, col3 = st.columns(3)
    tipo = col1.text_input("Tipo")
    depto = col1.text_input("Departamento")
    sucursal = col1.text_input("Sucursal")
    responsable = col2.text_input("Responsable")
    posicion = col2.text_input("Posición")
    nombre = col2.text_input("Nombre de Equipo")
    correo = col3.text_input("Correo")
    fecha_mantenimiento = col3.date_input("Fecha de Mantenimiento", value=date.today())
    hora = col3.time_input("Hora de mantenimiento")
    submitted = st.form_submit_button("Agregar equipo")

    if submitted and nombre:
        new_row = {
            "Tipo": tipo,
            "Departamento": depto,
            "Sucursal": sucursal,
            "Responsable": responsable,
            "Posicion": posicion,
            "Nombre de Equipo": nombre,
            "Correo": correo,
            "Fecha de Mantenimiento": fecha_mantenimiento,
            "Hora": str(hora)
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        save_data(df)
        st.success(f"✅ Equipo {nombre} agregado correctamente!")

# ----- LIMPIAR TIPOS DE DATOS -----
for col in df.columns:
    if "Fecha" in col:
        df[col] = pd.to_datetime(df[col], errors="coerce").dt.date
    elif "Hora" in col:
        df[col] = df[col].astype(str)
    else:
        df[col] = df[col].astype(str).fillna("")

# ----- TABLA EDITABLE -----
st.subheader("Lista de equipos")
edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")

if not edited_df.equals(df):
    save_data(edited_df)
    st.success("✅ Cambios guardados correctamente.")
    df = edited_df

# ===============================
# EXPORTACIÓN
# ===============================
st.subheader("Exportar calendario")

if not df.empty:
    # --- Formatear rango de fechas personalizado ---
    min_date = pd.to_datetime(df["Fecha de Mantenimiento"], errors="coerce").min()
    max_date = pd.to_datetime(df["Fecha de Mantenimiento"], errors="coerce").max()
    rango_semana = f"Semana del ({min_date.strftime('%d–%b')} al {max_date.strftime('%d–%b %Y')})"

    st.write(f"📅 Rango actual: **{rango_semana}**")

    # --- Formatear antes de exportar ---
    df_export = df.copy()
    df_export["Fecha de Mantenimiento"] = pd.to_datetime(df_export["Fecha de Mantenimiento"], errors="coerce").dt.strftime("%d-%b-%Y")
    df_export["Hora"] = pd.to_datetime(df_export["Hora"], errors="coerce").dt.strftime("%I:%M %p")

    # --- Exportar Excel ---
    excel_filename = os.path.join(EXPORT_FOLDER, "mantenimiento_semana.xlsx")
    df_export.to_excel(excel_filename, index=False)
    with open(excel_filename, "rb") as file:
        st.download_button("📥 Descargar Excel", file, file_name="mantenimiento_semana.xlsx")

    # --- Exportar PDF ---
    pdf_filename = os.path.join(EXPORT_FOLDER, "mantenimiento_semana.pdf")
    export_pdf(df_export, pdf_filename, rango_semana)
    with open(pdf_filename, "rb") as file:
        st.download_button("📥 Descargar PDF", file, file_name="mantenimiento_semana.pdf")
else:
    st.info("Aún no hay equipos registrados para exportar.")
