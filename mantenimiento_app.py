import streamlit as st
import pandas as pd
from datetime import date
from dateutil.relativedelta import relativedelta
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
MAINTENANCE_INTERVAL_MONTHS = 3  # mantenimiento cada 3 meses

# Crear carpetas si no existen
os.makedirs("data", exist_ok=True)
os.makedirs(EXPORT_FOLDER, exist_ok=True)

# ===============================
# Cargar datos
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
# Generar PDF (ajustado y formateado)
# ===============================
def export_pdf(df, filename):
    from reportlab.lib.pagesizes import landscape, letter
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet

    df = df.copy()
    # Formatear columnas
    if "Fecha de Mantenimiento" in df.columns:
        df["Fecha de Mantenimiento"] = pd.to_datetime(df["Fecha de Mantenimiento"], errors="coerce").dt.strftime("%d-%b-%Y")
    if "Hora" in df.columns:
        try:
            df["Hora"] = pd.to_datetime(df["Hora"], errors="coerce").dt.strftime("%I:%M %p")
        except:
            pass

    # Documento PDF
    doc = SimpleDocTemplate(filename, pagesize=landscape(letter), leftMargin=25, rightMargin=25, topMargin=25, bottomMargin=25)
    styles = getSampleStyleSheet()
    elements = []

    # Título
    title = Paragraph("Calendario de Mantenimiento Preventivo", styles['Title'])
    elements.append(title)
    elements.append(Spacer(1, 12))

    # Ajustar anchos de columna dinámicamente
    col_widths = [max(90, min(150, len(str(col)) * 7)) for col in df.columns]

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
    df["Fecha de Mantenimiento"] = pd.to_datetime(df["Fecha de Mantenimiento"], errors="coerce")
    df["Semana"] = df["Fecha de Mantenimiento"].dt.isocalendar().week
    df["Año"] = df["Fecha de Mantenimiento"].dt.year

    # Crear rango personalizado: Semana del (20–26 Oct 2025)
    semanas = df.groupby(["Año", "Semana"])["Fecha de Mantenimiento"].agg(["min", "max"]).reset_index()
    semanas["Rango"] = semanas.apply(
        lambda x: f"Semana del ({x['min'].strftime('%d–%b')} al {x['max'].strftime('%d–%b %Y')})", axis=1
    )

    selected = st.selectbox("Seleccionar semana a exportar", semanas["Rango"])

    # Filtrar por semana seleccionada
    semana_sel = semanas[semanas["Rango"] == selected].iloc[0]
    año_sel, sem_sel = semana_sel["Año"], semana_sel["Semana"]
    df_week = df[(df["Año"] == año_sel) & (df["Semana"] == sem_sel)].copy()

    # --- Formatear antes de exportar ---
    df_week["Fecha de Mantenimiento"] = df_week["Fecha de Mantenimiento"].dt.strftime("%d-%b-%Y")
    df_week["Hora"] = pd.to_datetime(df_week["Hora"], errors="coerce").dt.strftime("%I:%M %p")

    # --- Exportar Excel ---
    excel_filename = os.path.join(EXPORT_FOLDER, f"mantenimiento_semana_{año_sel}_{sem_sel}.xlsx")
    df_week.to_excel(excel_filename, index=False)
    with open(excel_filename, "rb") as file:
        st.download_button("📥 Descargar Excel de la semana seleccionada", file, file_name=f"mantenimiento_{selected}.xlsx")

    # --- Exportar PDF ---
    pdf_filename = os.path.join(EXPORT_FOLDER, f"mantenimiento_semana_{año_sel}_{sem_sel}.pdf")
    export_pdf(df_week, pdf_filename)
    with open(pdf_filename, "rb") as file:
        st.download_button("📥 Descargar PDF de la semana seleccionada", file, file_name=f"mantenimiento_{selected}.pdf")
else:
    st.info("Aún no hay equipos registrados para exportar.")
