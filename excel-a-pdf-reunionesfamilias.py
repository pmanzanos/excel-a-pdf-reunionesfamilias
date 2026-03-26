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
        # Lugar y fecha formateada
        self.set_font('helvetica', 'I', 10)
        self.cell(0, 10, f"En Hoyos, a {fecha_texto}", 0, 1, 'L')
        self.ln(5)
        
        # Cargos dinámicos
        for cargo, nombre in lista_firmas:
            self.set_font('helvetica', 'B', 10)
            self.write(7, f"{cargo}: ")
            self.set_font('helvetica', '', 10)
            self.multi_cell(0, 7, f"Fdo. {nombre}".encode('latin-1', 'replace').decode('latin-1'))
            self.ln(2)

def formatear_fecha_espanol(fecha_obj):
    try:
        if isinstance(fecha_obj, str):
            fecha_obj = pd.to_datetime(fecha_obj)
        meses = [
            "enero", "febrero", "marzo", "abril", "mayo", "junio",
            "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
        ]
        dia = fecha_obj.day
        mes = meses[fecha_obj.month - 1]
        anio = fecha_obj.year
        return f"{dia} de {mes} de {anio}"
    except:
        return "--- de --- de ---"

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

def generar_pdf_reunion(fila_rpts, lista_firmas):
    pdf = PDFReunion()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    
    # Formatear la fecha de la reunión (columna E -> índice 4)
    fecha_formateada = formatear_fecha_espanol(fila_rpts.iloc[4])
    
    pdf.seccion("DATOS DE LA REUNIÓN")
    pdf.campo("ID", fila_rpts['ID_GENERADA']) 
    pdf.campo("LA REUNIÓN SE PRODUCE A PETICIÓN DE", fila_rpts.iloc[3])
    pdf.campo("FECHA Y HORA DE LA COMUNICACIÓN", f"{fecha_formateada} a las {fila_rpts.iloc[5]}")
    
    solo_nombre, solo_curso = limpiar_nombre_y_curso(fila_rpts.iloc[7])
    pdf.campo("ALUMN@/O", solo_nombre) 
    pdf.campo("CURSO", solo_curso)
    pdf.campo("FAMILIAR/ES PRESENTES", fila_rpts.iloc[9])
    pdf.campo("OTROS PRESENTES", fila_rpts.iloc[8])
    
    pdf.ln(2); pdf.seccion("DESARROLLO DE LA REUNIÓN")
    pdf.campo("ASUNTO A TRATAR", fila_rpts.iloc[10], 'B')
    pdf.campo("DESCRIPCIÓN DE LO TRATADO", fila_rpts.iloc[12])
    pdf.campo("TRÁMITE A SEGUIR", fila_rpts.iloc[11])

    # Bloque de firmas con la fecha formateada
    pdf.dibujar_bloque_firmas(lista_firmas, fecha_formateada)
    return pdf.output()

# --- INTERFAZ ---
st.set_page_config(page_title="Generador Actas", page_icon="🤝")
st.title("🤝 Generador de Actas")

archivo = st.file_uploader("Sube el Excel con pestañas RPTS e Informe", type=['xlsx'])

if archivo:
    try:
        df_rpts = pd.read_excel(archivo, sheet_name='RPTS')
        
        # Lógica de ID desde Marca Temporal (Columna C)
        def generar_id(fecha_val):
            try:
                if isinstance(fecha_val, str): fecha_val = pd.to_datetime(fecha_val)
                excel_date = (fecha_val - pd.Timestamp("1899-12-30")).total_seconds() / 86400.0
                return str(f"{round(excel_date, 5):.5f}").split(".")[1][:4]
            except: return "ERR"

        df_rpts['ID_GENERADA'] = df_rpts.iloc[:, 2].apply(generar_id)
        df_rpts['NOMBRE_LIMPIO'] = df_rpts.iloc[:, 7].apply(lambda x: limpiar_nombre_y_curso(x)[0])
        df_rpts['ETIQUETA'] = df_rpts['ID_GENERADA'].astype(str) + " - " + df_rpts['NOMBRE_LIMPIO']
        
        opciones = ["Selecciona un acta..."] + sorted(df_rpts[df_rpts['ID_GENERADA'] != "ERR"]['ETIQUETA'].tolist())
        seleccion = st.selectbox("Busca al alumno:", opciones)

        if seleccion != "Selecciona un acta...":
            id_buscada = seleccion.split(" - ")[0]
            fila_sel = df_rpts[df_rpts['ID_GENERADA'] == id_buscada].iloc[0]
            
            # Leer firmas de pestaña 'Informe' (Fila 49 en Excel es 48 en Python)
            df_inf = pd.read_excel(archivo, sheet_name='Informe', header=None)
            firmas = []
            for i in range(48, 65): # Rango de firmas
                try:
                    cargo = df_inf.iloc[i, 0] # Columna A
                    nombre = df_inf.iloc[i, 2] # Columna C
                    if pd.notna(cargo) and pd.notna(nombre) and str(nombre).strip() not in ["", "nan", "0"]:
                        firmas.append((str(cargo).strip(": "), str(nombre).strip()))
                except: continue

            if st.button("🚀 Crear PDF"):
                pdf_bytes = generar_pdf_reunion(fila_sel, firmas)
                st.download_button("⬇️ Descargar Acta", data=bytes(pdf_bytes), file_name=f"Acta_{id_buscada}.pdf")

    except Exception as e:
        st.error(f"⚠️ Revisa las pestañas: {e}")
