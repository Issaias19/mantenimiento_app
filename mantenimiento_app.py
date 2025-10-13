import streamlit as st
import pandas as pd
from datetime import date, timedelta
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
        df = pd.DataFrame(columns=["Tipo","Departamento","Sucursal","Responsable",
                                   "Posicion","Nombre de Equipo","Correo",
                                   "Fecha de Mantenimiento","Hora"])
        df.to_excel(DATA_FILE, index=False)
        return df

def save_data(df):
    df.to_excel(DATA_FILE, index=False)

# ===============================
# Generar PDF
# ===============================
def export_pdf(df, filename):
    doc = SimpleDocTemplate(filename, pagesize=landscape(letter))
    styles = getSampleStyleSheet()
    elements = []

    title = Paragraph("Calendario de Mantenimiento Preventivo", styles['Title'])
    elements.append(title)
    elements.append(Spacer(1,12))

    data = [list(df.columns)] + df.values.tolist()
    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,0),colors.gray),
        ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),
        ('ALIGN',(0,0),(-1,-1),'CENTER'),
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
        ('GRID',(0,0),(-1,-1),0.5,colors.black),
        ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.whitesmoke,colors.lightgrey])
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
        next_date = fecha_mantenimiento + relativedelta(months=MAINTENANCE_INTERVAL_MONTHS)
        new_row = {
            "Tipo": tipo, "Departamento": depto, "Sucursal": sucursal,
            "Responsable": responsable, "Posicion": posicion, "Nombre de Equipo": nombre,
            "Correo": correo, "Fecha de Mantenimiento": fecha_mantenimiento, "Hora": hora
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        save_data(df)
        st.success(f"✅ Equipo {nombre} agregado correctamente!")

# ----- TABLA DE EQUIPOS -----
st.subheader("Lista de equipos")
st.dataframe(df, use_container_width=True)

# ----- EXPORTACIÓN -----
st.subheader("Exportar calendario")
# Agrupar por semana
df_export = df.copy()
df_export["Semana"] = df_export["Fecha de Mantenimiento"].apply(lambda x: x.isocalendar()[1])
weeks = df_export["Semana"].unique()
selected_week = st.selectbox("Seleccionar semana para exportar", sorted(weeks))

df_week = df_export[df_export["Semana"]==selected_week]

# Exportar Excel
excel_filename = os.path.join(EXPORT_FOLDER,f"mantenimiento_semana_{selected_week}.xlsx")
df_week.to_excel(excel_filename, index=False)
st.download_button("📥 Descargar Excel de la semana seleccionada", excel_filename)

# Exportar PDF
pdf_filename = os.path.join(EXPORT_FOLDER,f"mantenimiento_semana_{selected_week}.pdf")
export_pdf(df_week, pdf_filename)
st.download_button("📥 Descargar PDF de la semana seleccionada", pdf_filename)
