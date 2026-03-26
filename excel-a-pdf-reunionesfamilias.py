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

    def campo(self, etiqueta, valor, estilo_valor=''):
        self.set_font('helvetica', 'B', 10)
        self.write(6, f"{etiqueta}: ")
        self.set_font('helvetica', estilo_valor, 10)
        val_str = str(valor) if pd.notna(valor) and str(valor).strip() not in ["nan", "None", ""] else "---"
        self.multi_cell(0, 6, val_str.encode('latin-1', 'replace').decode('latin-1'))
        self.ln(1)

    def dibujar_firmas_paralelo(self, docente, jefe):
        self.ln(10)
        y_pos = self.get_y()
        # Casillas marcadas con X
        for x in [10, 105]:
            self.rect(x, y_pos, 4, 4)
            self.set_font('helvetica', 'B', 8)
            self.text(x+1, y_pos + 3.2, "X")
        
        self.set_font('helvetica', '', 10)
        self.set_xy(16, y_pos - 1); self.cell(85, 6, "Conforme del Docente / Ed. Social", 0, 0)
        self.set_xy(111, y_pos - 1); self.cell(85, 6, "Conforme de la Jefatura de Estudios", 0, 1)
        
        self.ln(15)
        y_nombres = self.get_y()
        self.set_xy(10, y_nombres); self.set_font('helvetica', 'B', 10)
        self.cell(90, 5, "V.º B.º El Docente / Ed. Social", 0, 1, 'L')
        self.set_font('helvetica', 'I', 9); self.cell(90, 5, f"Fdo: {docente}".encode('latin-1', 'replace').decode('latin-1'), 0, 0, 'L')
        
        self.set_xy(105, y_nombres); self.set_font('helvetica', 'B', 10)
        self.cell(90, 5, "V.º B.º Jefatura de Estudios", 0, 1, 'L')
        self.set_font('helvetica', 'I', 9); self.set_x(105); self.cell(90, 5, f"Fdo: {jefe}".encode('latin-1', 'replace').decode('latin-1'), 0, 0, 'L')

def generar_pdf_reunion(fila, nombre_jefatura):
    pdf = PDFReunion()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    
    # DATOS GENERALES
    pdf.seccion("DATOS DE LA REUNIÓN")
    id_val = fila['ID_REDONDEADA']
    peticion = fila.iloc[3] # Columna D (Índice 3)
    pdf.campo("ID", f"{id_val}. La reunión se produce a petición de: {peticion}")
    
    fecha_hora = f"{fila.iloc[4]} a las {fila.iloc[5]}" # E (4) y F (5)
    pdf.campo("Fecha y hora de la comunicación", fecha_hora)
    pdf.campo("ALUMN@/O Y CURSO", fila.iloc[7]) # Columna H (7)
    pdf.campo("FAMILIAR/ES PRESENTES", fila.iloc[9]) # Columna J (9)
    pdf.campo("OTROS PRESENTES", fila.iloc[8]) # Columna I (8)
    
    # CONTENIDO
    pdf.ln(2); pdf.seccion("DESARROLLO DE LA REUNIÓN")
    pdf.campo("ASUNTO A TRATAR", fila.iloc[10], 'B') # Columna K (10)
    pdf.campo("DESCRIPCIÓN DE LO TRATADO", fila.iloc[12]) # Columna M (12)
    pdf.campo("TRÁMITE A SEGUIR", fila.iloc[11]) # Columna L (11)

    pdf.dibujar_firmas_paralelo("Docente responsable", nombre_jefatura)
    
    return pdf.output()

# --- INTERFAZ ---
st.set_page_config(page_title="Generador Actas", page_icon="🤝")
st.title("🤝 Actas de Reunión (Pestaña RPTS)")

archivo = st.file_uploader("Sube el Excel de Respuestas", type=['xlsx'])

if archivo:
    try:
        # CAMBIO CLAVE: Ahora leemos de 'RPTS'
        df = pd.read_excel(archivo, sheet_name='RPTS')
        
        def extraer_id(valor):
            try:
                v = round(float(valor), 4)
                return str(f"{v:.4f}").split('.')[1]
            except: return "0000"

        # ID basada en Columna C (índice 2)
        df['ID_REDONDEADA'] = df.iloc[:, 2].apply(extraer_id)
        # Etiqueta basada en Columna H (índice 7)
        df['ETIQUETA'] = df['ID_REDONDEADA'].astype(str) + " - " + df.iloc[:, 7].astype(str)
        
        st.success(f"✅ Cargados {len(df)} registros de la hoja 'RPTS'.")

        opciones = ["Selecciona un alumno..."] + sorted(df['ETIQUETA'].dropna().tolist())
        seleccion = st.selectbox("Busca por ID o Nombre:", opciones)

        if seleccion != "Selecciona un alumno...":
            id_buscada = seleccion.split(" - ")[0]
            fila_sel = df[df['ID_REDONDEADA'] == id_buscada].iloc[0]
            
            # Intentar sacar Jefatura de hoja PARTE
            try:
                df_p = pd.read_excel(archivo, sheet_name='PARTE', header=None)
                jefe = df_p.iloc[48, 3] if not pd.isna(df_p.iloc[48, 3]) else "Jefatura de Estudios"
            except: jefe = "Jefatura de Estudios"

            if st.button("🚀 Generar Acta PDF"):
                pdf_bytes = generar_pdf_reunion(fila_sel, jefe)
                st.download_button(label="⬇️ Descargar Acta", data=bytes(pdf_bytes), file_name=f"Acta_{id_buscada}.pdf", mime="application/pdf")

    except Exception as e:
        st.error(f"⚠️ Error: {e}. Revisa que la pestaña se llame 'RPTS'.")
else:
    st.info("👋 Sube el archivo para comenzar.")
