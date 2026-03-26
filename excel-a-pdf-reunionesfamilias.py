import streamlit as st
import pandas as pd
from fpdf import FPDF
import re
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
        val_str = str(valor) if pd.notna(valor) and str(valor).strip() not in ["nan", "None", "", "#VALUE!"] else "---"
        self.multi_cell(0, 6, val_str.encode('latin-1', 'replace').decode('latin-1'))
        self.ln(1)

    def dibujar_bloque_firmas(self, lista_firmas, fecha_texto):
        self.ln(10)
        self.set_font('helvetica', 'I', 10)
        self.cell(0, 10, f"En Hoyos, a {fecha_texto}", 0, 1, 'L')
        self.ln(5)
        
        self.set_font('helvetica', '', 10)
        for cargo, nombre in lista_firmas:
            # Formato solicitado: Fdo. Nombre (Cargo)
            # Limpiamos posibles "Fdo." duplicados
            limpio_nombre = str(nombre).replace("Fdo.", "").replace("Fdo:", "").strip()
            texto_firma = f"Fdo. {limpio_nombre} ({cargo})"
            
            self.multi_cell(0, 7, texto_firma.encode('latin-1', 'replace').decode('latin-1'))
            self.ln(1)

def formatear_fecha_espanol(fecha_obj):
    try:
        if isinstance(fecha_obj, str): fecha_obj = pd.to_datetime(fecha_obj)
        meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", 
                  "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
        return f"{fecha_obj.day} de {meses[fecha_obj.month - 1]} de {fecha_obj.year}"
    except: return "--- de --- de ---"

def limpiar_nombre_y_curso(texto):
    if pd.isna(texto): return "---", "---"
    t = str(texto)
    match = re.search(r'(\d|Diver)', t, re.IGNORECASE)
    if match:
        idx = match.start()
        return t[:idx].strip().rstrip(','), t[idx:].strip()
    return t, "---"

def generar_pdf_reunion(fila_rpts, lista_firmas):
    pdf = PDFReunion()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    
    fecha_com = formatear_fecha_espanol(fila_rpts.iloc[4])
    
    pdf.seccion("DATOS DE LA REUNIÓN")
    pdf.campo("ID", fila_rpts['ID_GENERADA']) 
    pdf.campo("LA REUNIÓN SE PRODUCE A PETICIÓN DE", fila_rpts.iloc[3])
    pdf.campo("FECHA Y HORA DE LA COMUNICACIÓN", f"{fecha_com} a las {fila_rpts.iloc[5]}")
    
    nom, cur = limpiar_nombre_y_curso(fila_rpts.iloc[7])
    pdf.campo("ALUMN@/O", nom) 
    pdf.campo("CURSO", cur)
    pdf.campo("FAMILIAR/ES PRESENTES", fila_rpts.iloc[9])
    pdf.campo("OTROS PRESENTES", fila_rpts.iloc[8])
    
    pdf.ln(2); pdf.seccion("DESARROLLO DE LA REUNIÓN")
    pdf.campo("ASUNTO A TRATAR", fila_rpts.iloc[10], 'B')
    pdf.campo("DESCRIPCIÓN DE LO TRATADO", fila_rpts.iloc[12])
    pdf.campo("TRÁMITE A SEGUIR", fila_rpts.iloc[11])

    pdf.dibujar_bloque_firmas(lista_firmas, fecha_com)
    return pdf.output()

# --- STREAMLIT ---
st.set_page_config(page_title="Generador Actas", page_icon="🤝")
st.title("🤝 Generador de Actas (Filtro de Firmas)")

archivo = st.file_uploader("Sube el Excel", type=['xlsx'])

if archivo:
    try:
        df_rpts = pd.read_excel(archivo, sheet_name='RPTS')
        df_inf = pd.read_excel(archivo, sheet_name='Informe', header=None)

        def gen_id(f):
            try:
                if isinstance(f, str): f = pd.to_datetime(f)
                val = (f - pd.Timestamp("1899-12-30")).total_seconds() / 86400.0
                return str(f"{round(val, 5):.5f}").split(".")[1][:4]
            except: return "ERR"

        df_rpts['ID_GENERADA'] = df_rpts.iloc[:, 2].apply(gen_id)
        df_rpts['ETIQUETA'] = df_rpts['ID_GENERADA'] + " - " + df_rpts.iloc[:, 7].apply(lambda x: limpiar_nombre_y_curso(x)[0])
        
        opc = ["Selecciona..."] + sorted(df_rpts[df_rpts['ID_GENERADA'] != "ERR"]['ETIQUETA'].tolist())
        sel = st.selectbox("Acta:", opc)

        if sel != "Selecciona...":
            id_b = sel.split(" - ")[0]
            fila_sel = df_rpts[df_rpts['ID_GENERADA'] == id_b].iloc[0]
            
            # --- ESCANEO DE FIRMAS AMPLIADO ---
            firmas_encontradas = []
            # Escaneamos un rango más amplio (hasta la fila 80 de Excel para no dejarnos a nadie)
            for r en range(48, 80): 
                try:
                    # Columna A (Cargo) y Columna C (Nombre)
                    cargo_raw = df_inf.iloc[r, 0]
                    nombre_raw = df_inf.iloc[r, 2]
                    
                    if pd.notna(nombre_raw) and str(nombre_raw).strip() not in ["", "0", "nan"]:
                        cargo = str(cargo_raw).strip().rstrip(":") if pd.notna(cargo_raw) else "Representante"
                        nombre = str(nombre_raw).strip()
                        firmas_encontradas.append((cargo, nombre))
                except: continue

            if st.button("🚀 Generar PDF"):
                pdf_bytes = generar_pdf_reunion(fila_sel, firmas_encontradas)
                st.download_button("⬇️ Descargar", data=bytes(pdf_bytes), file_name=f"Acta_{id_b}.pdf")

    except Exception as e:
        st.error(f"Error: {e}")
