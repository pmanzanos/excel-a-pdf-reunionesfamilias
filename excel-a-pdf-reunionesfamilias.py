import streamlit as st
import pandas as pd
from fpdf import FPDF
import re

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
        val_str = str(valor) if pd.notna(valor) and str(valor).strip() not in ["nan", "None", "", "#VALUE!"] else "---"
        self.multi_cell(0, 6, val_str.encode('latin-1', 'replace').decode('latin-1'))
        self.ln(1)

    def dibujar_firmas_paralelo(self, docente, jefe):
        self.ln(10)
        y_pos = self.get_y()
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

def limpiar_nombre_y_curso(texto):
    if pd.isna(texto): return "---", "---"
    t = str(texto)
    match = re.search(r'(\d|Diver)', t, re.IGNORECASE)
    if match:
        indice = match.start()
        nombre = t[:indice].strip().rstrip(',')
        curso = t[indice:].strip()
        return nombre, curso
    return t, "---"

def generar_pdf_reunion(fila, nombre_jefatura):
    pdf = PDFReunion()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    
    pdf.seccion("DATOS DE LA REUNIÓN")
    id_val = fila['ID_GENERADA']
    
    pdf.campo("ID", f"{id_val}. La reunión se produce a petición de: {fila.iloc[3]}") # D
    pdf.campo("Fecha y hora", f"{fila.iloc[4]} a las {fila.iloc[5]}") # E y F
    
    solo_nombre, solo_curso = limpiar_nombre_y_curso(fila.iloc[7]) # H
    
    pdf.campo("ALUMN@/O", solo_nombre) 
    pdf.campo("CURSO", solo_curso)
    pdf.campo("FAMILIAR/ES PRESENTES", fila.iloc[9]) # J
    pdf.campo("OTROS PRESENTES", fila.iloc[8]) # I
    
    pdf.ln(2); pdf.seccion("DESARROLLO DE LA REUNIÓN")
    pdf.campo("ASUNTO A TRATAR", fila.iloc[10], 'B') # K
    pdf.campo("DESCRIPCIÓN DE LO TRATADO", fila.iloc[12]) # M
    pdf.campo("TRÁMITE A SEGUIR", fila.iloc[11]) # L

    pdf.dibujar_firmas_paralelo("Docente responsable", nombre_jefatura)
    return pdf.output()

# --- INTERFAZ STREAMLIT ---
st.set_page_config(page_title="Generador Actas", page_icon="🤝")
st.title("🤝 Actas de Reunión (ID desde Marca Temporal)")

archivo = st.file_uploader("Sube el Excel de Respuestas", type=['xlsx'])

if archivo:
    try:
        df = pd.read_excel(archivo, sheet_name='RPTS')
        
        # SOLUCIÓN AL REDONDEO: Generar la ID directamente de la Marca Temporal (Columna C / Índice 2)
        # Esto replica lo que hace Google Sheets internamente pero de forma controlada en Python
        def generar_id_desde_fecha(fecha_val):
            try:
                # Convertimos la fecha a número decimal (formato Excel)
                # 25569 es el offset de días entre Unix y Excel
                if isinstance(fecha_val, str):
                    fecha_val = pd.to_datetime(fecha_val)
                
                excel_date = (fecha_val - pd.Timestamp("1899-12-30")).total_seconds() / 86400.0
                # Redondeamos a 5 decimales (que es lo que suele mostrar Google Sheets)
                num_redondeado = round(excel_date, 5)
                # Extraemos los 4 dígitos decimales finales como hace tu fórmula
                id_str = str(f"{num_redondeado:.5f}").split(".")[1][:4]
                return id_str
            except:
                return "ERR"

        # Aplicamos a la Columna C (índice 2) que es la "Marca temporal" original
        df['ID_GENERADA'] = df.iloc[:, 2].apply(generar_id_desde_fecha)
        
        # Limpieza de nombre en columna H (índice 7)
        df['NOMBRE_LIMPIO'] = df.iloc[:, 7].apply(lambda x: limpiar_nombre_y_curso(x)[0])
        df['ETIQUETA'] = df['ID_GENERADA'].astype(str) + " - " + df['NOMBRE_LIMPIO']
        
        st.success(f"✅ Cargados {len(df)} registros. IDs sincronizadas con éxito.")

        df_validos = df[df['ID_GENERADA'] != "ERR"]
        opciones = ["Selecciona un alumno..."] + sorted(df_validos['ETIQUETA'].tolist())
        
        seleccion = st.selectbox("Busca por ID o Nombre:", opciones)

        if seleccion != "Selecciona un alumno...":
            id_buscada = seleccion.split(" - ")[0]
            fila_sel = df[df['ID_GENERADA'] == id_buscada].iloc[0]
            
            try:
                df_p = pd.read_excel(archivo, sheet_name='PARTE', header=None)
                jefe = df_p.iloc[48, 3] if not pd.isna(df_p.iloc[48, 3]) else "Jefatura de Estudios"
            except: jefe = "Jefatura de Estudios"

            if st.button("🚀 Generar Acta PDF"):
                pdf_bytes = generar_pdf_reunion(fila_sel, jefe)
                st.download_button(label="⬇️ Descargar Acta", data=bytes(pdf_bytes), file_name=f"Acta_{id_buscada}.pdf", mime="application/pdf")

    except Exception as e:
        st.error(f"⚠️ Error: {e}")
else:
    st.info("👋 Sube el archivo para comenzar.")
