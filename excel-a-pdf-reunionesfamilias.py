import streamlit as st
import pandas as pd
from fpdf import FPDF
import datetime

class PDFReunion(FPDF):
    def header(self):
        try:
            # 1. Logo
            self.image('encabezado.png', 10, 8, 190)
            self.ln(33) 
        except:
            self.ln(10)
        
        # 2. TÍTULO DE REUNIÓN
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
            val_str = valor.strftime('%d/%m/%Y')
        else:
            val_str = str(valor)
        self.multi_cell(0, 6, val_str.encode('latin-1', 'replace').decode('latin-1'))
        self.ln(1)

    def dibujar_firmas_paralelo(self, docente, jefe):
        self.ln(10)
        y_pos = self.get_y()
        
        # Bloque Izquierdo (Marcado con X)
        self.rect(10, y_pos, 4, 4)
        self.set_font('helvetica', 'B', 8)
        self.text(11, y_pos + 3.2, "X")
        self.set_xy(16, y_pos - 1)
        self.set_font('helvetica', '', 10)
        self.cell(85, 6, "Conforme del Docente / Ed. Social", 0, 0)
        
        # Bloque Derecho (Marcado con X)
        self.rect(105, y_pos, 4, 4)
        self.set_font('helvetica', 'B', 8)
        self.text(106, y_pos + 3.2, "X")
        self.set_xy(111, y_pos - 1)
        self.set_font('helvetica', '', 10)
        self.cell(85, 6, "Conforme de la Jefatura de Estudios", 0, 1)
        
        self.ln(15)
        y_nombres = self.get_y()
        self.set_xy(10, y_nombres)
        self.set_font('helvetica', 'B', 10)
        self.cell(90, 5, "V.º B.º El Docente / Ed. Social", 0, 1, 'L')
        self.set_font('helvetica', 'I', 9)
        self.cell(90, 5, f"Fdo: {docente}".encode('latin-1', 'replace').decode('latin-1'), 0, 0, 'L')
        
        self.set_xy(105, y_nombres)
        self.set_font('helvetica', 'B', 10)
        self.cell(90, 5, "V.º B.º Jefatura de Estudios", 0, 1, 'L')
        self.set_font('helvetica', 'I', 9)
        self.set_x(105)
        self.cell(90, 5, f"Fdo: {jefe}".encode('latin-1', 'replace').decode('latin-1'), 0, 0, 'L')

def generar_pdf_reunion(datos_fila, nombre_jefatura):
    pdf = PDFReunion()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    
    pdf.seccion("DATOS DE LA REUNIÓN")
    # Campo 1 modificado: ID + Texto solicitado
    id_val = datos_fila.get('ID_REDONDEADA', '---')
    pdf.campo("ID", f"{id_val}. La reunión se produce a petición de...")
    
    # Campo 2 modificado: Fecha y hora
    pdf.campo("Fecha y hora de la comunicación", datos_fila.get('FECHA DEL INCIDENTE', '---'))
    
    # Datos de identificación
    pdf.campo("ALUMN@/O", datos_fila.get('ALUMNO OBJETO DEL PARTE', '---'))
    pdf.campo("CURSO / GRUPO / TUTOR", datos_fila.get('CURSO / GRUPO / TUTOR', '---'))
    
    docente_nombre = datos_fila.get('DOCENTE / ED. SOCIAL QUE IMPONE EL PARTE', '---')
    pdf.campo("DOCENTE RESPONSABLE", docente_nombre)
    
    # Espacio para firmas con casillas marcadas
    pdf.dibujar_firmas_paralelo(docente_nombre, nombre_jefatura)
    
    return pdf.output()

# --- INTERFAZ STREAMLIT ---
st.set_page_config(page_title="Informes de Reunión", page_icon="🤝")
st.title("🤝 Reunión Presencial con Familias")

archivo = st.file_uploader("Sube el archivo Excel", type=['xlsx'])

if archivo:
    try:
        df = pd.read_excel(archivo, sheet_name='RPTS')
        # Capturamos el nombre de jefatura desde la hoja PARTE
        df_parte = pd.read_excel(archivo, sheet_name='PARTE', header=None)
        nombre_jefatura = df_parte.iloc[48, 3] if not pd.isna(df_parte.iloc[48, 3]) else "Jefatura de Estudios"

        def extraer_id_redondeada(valor):
            try:
                valor_redondeado = round(float(valor), 4)
                return str(f"{valor_redondeado:.4f}").split('.')[1]
            except:
                return None

        # Procesamos IDs y creamos la etiqueta para el buscador
        df['ID_REDONDEADA'] = df['NUMERO'].apply(extraer_id_redondeada)
        df['ETIQUETA'] = df['ID_REDONDEADA'].astype(str) + " - " + df['ALUMNO OBJETO DEL PARTE'].astype(str)
        
        st.success("✅ Archivo cargado correctamente.")

        # Desplegable de selección
        opciones = ["Selecciona un alumno..."] + sorted(df['ETIQUETA'].dropna().tolist())
        seleccion = st.selectbox("Introduce la ID:", opciones)

        if seleccion != "Selecciona un alumno...":
            id_real = seleccion.split(" - ")[0]
            fila = df[df['ID_REDONDEADA'] == id_real].iloc[0]
            
            st.info(f"📋 Registro seleccionado: {fila['ALUMNO OBJETO DEL PARTE']}")
            
            if st.button("🚀 Generar Informe de Reunión"):
                pdf_bytes = generar_pdf_reunion(fila, nombre_jefatura)
                st.download_button(
                    label="⬇️ Descargar Informe PDF",
                    data=bytes(pdf_bytes),
                    file_name=f"Reunion_{id_real}.pdf",
                    mime="application/pdf"
                )

    except Exception as e:
        st.error(f"⚠️ Error al procesar el archivo: {e}")
else:
    st.info("👋 Por favor, sube el archivo Excel para empezar.")
