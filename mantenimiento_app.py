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
# FUNCIONES DE DATOS
# ===============================
def load_data():
    if os.path.exists(DATA_FILE):
        return pd.read_excel(DATA_FILE)
    else:
        df = pd.DataFrame(columns=[
            "Tipo","Departamento","Sucursal","Responsable",
            "Posicion","Nombre de Equipo","Correo",
            "Fecha de Mantenimiento","Hora"
        ])
        df.to_excel(DATA_FILE, index=False)
        return df

def save_data(df):
    df.to_excel(DATA_FILE, index=False)

# ===============================
# FUNCIÓN PARA EXPORTAR A PDF
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
        new_row = {
            "Tipo": tipo, "Departamento": depto, "Sucursal": sucursal,
            "Responsable": responsable, "Posicion": posicion, "Nombre de Equipo": nombre,
            "Correo": correo, "Fecha de Mantenimiento": fecha_mantenimiento, "Hora": hora
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        save_data(df)
        st.success(f"✅ Equipo '{nombre}' agregado correctamente!")

# ----- TABLA EDITABLE -----
st.subheader("Lista de equipos (editable)")
edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")

# Guardar si se modifica algo
if not edited_df.equals(df):
    save_data(edited_df)
    st.success("💾 Cambios guardados correctamente.")
    df = edited_df

# ----- EXPORTACIÓN -----
st.subheader("Exportar calendario")

# Calcular rango de semanas disponibles
df["Fecha de Mantenimiento"] = pd.to_datetime(df["Fecha de Mantenimiento"], errors='coerce')
df = df.dropna(subset=["Fecha de Mantenimiento"])
df["Semana"] = df["Fecha de Mantenimiento"].apply(lambda x: x.isocalendar()[1])
df["Año"] = df["Fecha de Mantenimiento"].dt.year

# Obtener rangos de fechas por semana
semana_rangos = {}
for _, grupo in df.groupby(["Año", "Semana"]):
    semana = grupo["Semana"].iloc[0]
    año = grupo["Año"].iloc[0]
    inicio_semana = grupo["Fecha de Mantenimiento"].min().strftime("%d %b")
    fin_semana = grupo["Fecha de Mantenimiento"].max().strftime("%d %b")
    rango = f"Semana {semana} ({inicio_semana} - {fin_semana}) {año}"
    semana_rangos[(año, semana)] = rango

# Selector con rangos legibles
if semana_rangos:
    selected_label = st.selectbox("📅 Selecciona la semana para exportar", list(semana_rangos.values()))
    selected_key = list(semana_rangos.keys())[list(semana_rangos.values()).index(selected_label)]
    año, semana = selected_key
    df_week = df[(df["Año"] == año) & (df["Semana"] == semana)]

    # Exportar Excel
    excel_filename = os.path.join(EXPORT_FOLDER, f"mantenimiento_{año}_semana_{semana}.xlsx")
    df_week.to_excel(excel_filename, index=False)
    with open(excel_filename, "rb") as f:
        st.download_button("📥 Descargar Excel", f, file_name=f"mantenimiento_{año}_semana_{semana}.xlsx")

    # Exportar PDF
    pdf_filename = os.path.join(EXPORT_FOLDER, f"mantenimiento_{año}_semana_{semana}.pdf")
    export_pdf(df_week, pdf_filename)
    with open(pdf_filename, "rb") as f:
        st.download_button("📥 Descargar PDF", f, file_name=f"mantenimiento_{año}_semana_{semana}.pdf")
else:
    st.info("No hay semanas registradas aún para exportar.")
