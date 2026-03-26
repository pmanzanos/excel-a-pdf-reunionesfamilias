import streamlit as st
import pandas as pd
from fpdf import FPDF
import datetime

class PDFReunion(FPDF):
    def header(self):
        try:
            self.image('encabezado.png', 10, 8, 190)
            self.ln(33) 
        except:
            self.ln(10)
        
        self.set_font('helvetica', 'B', 16)
        self.cell(0, 10, 'REUNIÓN PRESENCIAL CON FAMILIAS', 0, 1, 'C')
        self.ln(5)

    def seccion(self, titulo):
        self.set_font('helvetica', 'B', 11)
        self.set_fill_color(240, 240, 240)
        self.cell(0, 8, f" {titulo}", 0, 1, 'L', fill=True)
        self.ln(2)

    def campo(self, etiqueta, valor):
        self.set_font('helvetica', 'B', 10)
        self.write(6, f"{etiqueta}: ")
        self.set_font('helvetica', '', 10)
        if pd.isna(valor) or str(valor).strip() in ["nan", "#VALUE!", ""]:
            val_str = "---"
        elif isinstance(valor, (datetime.datetime, pd.Timestamp)):
            val_str = valor.strftime('%d/%m/%Y %H:%M')
        else:
            val_str = str(valor)
        self.multi_cell(0, 6, val_str.encode('latin-1', 'replace').decode('latin-1'))
        self.ln(1)

    def dibujar_firmas_paralelo(self, docente, jefe):
        self.ln(10)
        y_pos = self.get_y()
        self.rect(10, y_pos, 4, 4)
        self.set_font('helvetica', 'B', 8)
        self.text(11, y_pos + 3.2, "X")
        self.set_xy(16, y_pos - 1)
        self.set_font('helvetica', '', 10)
        self.cell(85, 6, "Conforme del Docente / Ed. Social", 0, 0)
        self.rect(105, y_pos, 4, 4)
        self.text(106, y_pos + 3.2, "X")
        self.set_xy(111, y_pos - 1)
        self.cell(85, 6, "Conforme de la Jefatura de Estudios", 0, 1)
        self.ln(15)
        y_nombres = self.get_y()
        self.set_xy(10, y_nombres); self.set_font('helvetica', 'B', 10)
        self.cell(90, 5, "V.º B.º El Docente / Ed. Social", 0, 1, 'L')
        self.set_font('helvetica', 'I', 9)
        self.cell(90, 5, f"Fdo: {docente}".encode('latin-1', 'replace').decode('latin-1'), 0, 0, 'L')
        self.set_xy(105, y_nombres); self.set_font('helvetica', 'B', 10)
        self.cell(90, 5, "V.º B.º Jefatura de Estudios", 0, 1, 'L')
        self.set_font('helvetica', 'I', 9); self.set_x(105)
        self.cell(90, 5, f"Fdo: {jefe}".encode('latin-1', 'replace').decode('latin-1'), 0, 0, 'L')

def generar_pdf_reunion(datos_fila, nombre_jefatura):
    pdf = PDFReunion()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.seccion("DATOS DE LA REUNIÓN")
    
    id_val = datos_fila.get('ID_REDONDEADA', '---')
    peticion = datos_fila.get('La reunión se produce a petición de...', '---')
    pdf.campo("ID", f"{id_val}. La reunión se produce a petición de: {peticion}")
    
    pdf.campo("Fecha y hora de la comunicación", datos_fila.get('Fecha y hora de la comunicación', '---'))
    pdf.campo("ALUMN@/O", datos_fila.get('Nombre del alumno/a', '---'))
    pdf.campo("CURSO", datos_fila.get('Curso', '---'))
    
    docente_nombre = datos_fila.get('Docente que convoca la reunión', '---')
    pdf.campo("DOCENTE RESPONSABLE", docente_nombre)
    pdf.dibujar_firmas_paralelo(docente_nombre, nombre_jefatura)
    return pdf.output()

# --- INTERFAZ ---
st.set_page_config(page_title="Actas de Reunión", page_icon="🤝")
st.title("🤝 Generador de Actas de Reunión")

archivo = st.file_uploader("Sube el archivo Excel", type=['xlsx'])

if archivo:
    try:
        df = pd.read_excel(archivo, sheet_name='Informe')
        
        # BUSCADOR AUTOMÁTICO DE LA COLUMNA DE ID (por si no se llama 'NUMERO')
        columna_id = None
        for col in df.columns:
            if 'NUMERO' in col.upper() or 'MARCA' in col.upper() or col == df.columns[0]:
                columna_id = col
                break

        def extraer_id_redondeada(valor):
            try:
                valor_redondeado = round(float(valor), 4)
                return str(f"{valor_redondeado:.4f}").split('.')[1]
            except: return None

        df['ID_REDONDEADA'] = df[columna_id].apply(extraer_id_redondeada)
        
        # Buscamos el nombre del alumno (intentamos varias opciones comunes)
        col_alumno = next((c for c in df.columns if 'ALUMNO' in c.upper() or 'NOMBRE' in c.upper()), df.columns[1])
        
        df['ETIQUETA'] = df['ID_REDONDEADA'].astype(str) + " - " + df[col_alumno].astype(str)
        st.success(f"✅ Hoja 'Informe' cargada. Usando columna '{columna_id}' para generar IDs.")

        opciones = ["Selecciona un alumno..."] + sorted(df['ETIQUETA'].dropna().tolist())
        seleccion = st.selectbox("Introduce la ID:", opciones)

        if seleccion != "Selecciona un alumno...":
            id_real = seleccion.split(" - ")[0]
            fila = df[df['ID_REDONDEADA'] == id_real].iloc[0]
            
            # Intentamos obtener Jefatura de la otra hoja
            try:
                df_p = pd.read_excel(archivo, sheet_name='PARTE', header=None)
                jefe = df_p.iloc[48, 3] if not pd.isna(df_p.iloc[48, 3]) else "Jefatura de Estudios"
            except: jefe = "Jefatura de Estudios"

            if st.button("🚀 Generar Acta de Reunión"):
                pdf_bytes = generar_pdf_reunion(fila, jefe)
                st.download_button(label="⬇️ Descargar Acta PDF", data=bytes(pdf_bytes), file_name=f"Acta_{id_real}.pdf", mime="application/pdf")

    except Exception as e:
        st.error(f"⚠️ Error: {e}")
