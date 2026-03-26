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
        val_str = str(valor) if pd.notna(valor) and str(valor).strip() not in ["nan", "None", "", "#VALUE!", "0"] else "---"
        self.multi_cell(0, 6, val_str.encode('latin-1', 'replace').decode('latin-1'))
        self.ln(1)

    def dibujar_bloque_firmas(self, lista_firmas, fecha_texto):
        self.ln(10)
        self.set_font('helvetica', 'I', 10)
        self.cell(0, 10, f"En Hoyos, a {fecha_texto}", 0, 1, 'L')
        self.ln(5)
        
        self.set_font('helvetica', '', 10)
        for cargo, nombre in lista_firmas:
            # Limpiamos el nombre por si ya trae "Fdo."
            n_limpio = str(nombre).replace("Fdo.", "").replace("Fdo:", "").strip()
            # Formato: Fdo. Nombre (Cargo)
            texto_firma = f"Fdo. {n_limpio} ({cargo})"
            self.multi_cell(0, 7, texto_firma.encode('latin-1', 'replace').decode('latin-1'))
            self.ln(1)

def formatear_fecha_espanol(fecha_obj):
    try:
        if isinstance(fecha_obj, str): fecha_obj = pd.to_datetime(fecha_obj)
        meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
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

# --- INTERFAZ ---
st.set_page_config(page_title="Generador Actas", page_icon="🤝")
st.title("🤝 Generador de Actas (Lectura de Informe)")

archivo = st.file_uploader("Sube el Excel (Asegúrate de que la pestaña Informe esté actualizada)", type=['xlsx'])

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
        
        opciones = ["Selecciona el alumno..."] + sorted(df_rpts[df_rpts['ID_GENERADA'] != "ERR"]['ETIQUETA'].tolist())
        seleccion = st.selectbox("Busca el registro:", opciones)

        if seleccion != "Selecciona el alumno...":
            id_b = seleccion.split(" - ")[0]
            fila = df_rpts[df_rpts['ID_GENERADA'] == id_b].iloc[0]
            
            # --- MAPEO DE FIRMAS (Pestaña Informe) ---
            # Diccionario con Fila Excel (convertida a índice Python -1) y su Cargo
            mapeo = {
                48: "Director",
                54: "Jefatura de Estudios",
                60: "Secretario/a",
                66: "Padre, madre, tutor o tutora legal",
                72: "Orientador/a",
                78: "Educador/a social",
                84: "Tutor/a del grupo",
                90: "Docente/s"
            }

            firmas_validas = []
            for fila_idx, cargo_txt in mapeo.items():
                try:
                    nombre = df_inf.iloc[fila_idx, 2] # Columna C
                    if pd.notna(nombre) and str(nombre).strip() not in ["", "0", "nan"]:
                        firmas_validas.append((cargo_txt, nombre))
                except: continue

            if st.button("🚀 Crear PDF con Firmas de Informe"):
                pdf = PDFReunion()
                pdf.set_auto_page_break(auto=True, margin=15)
                pdf.add_page()
                
                fecha_f = formatear_fecha_espanol(fila.iloc[4])
                
                pdf.seccion("DATOS DE LA REUNIÓN")
                pdf.campo("ID", fila['ID_GENERADA'])
                pdf.campo("LA REUNIÓN SE PRODUCE A PETICIÓN DE", fila.iloc[3])
                pdf.campo("FECHA Y HORA DE LA COMUNICACIÓN", f"{fecha_f} a las {fila.iloc[5]}")
                
                nom, cur = limpiar_nombre_y_curso(fila.iloc[7])
                pdf.campo("ALUMN@/O", nom)
                pdf.campo("CURSO", cur)
                
                pdf.seccion("DESARROLLO DE LA REUNIÓN")
                pdf.campo("ASUNTO A TRATAR", fila.iloc[10], 'B')
                pdf.campo("DESCRIPCIÓN DE LO TRATADO", fila.iloc[12])
                pdf.campo("TRÁMITE A SEGUIR", fila.iloc[11])
                
                pdf.dibujar_bloque_firmas(firmas_validas, fecha_f)
                
                st.download_button("⬇️ Descargar Acta", data=bytes(pdf.output()), file_name=f"Acta_{id_b}.pdf")

    except Exception as e:
        st.error(f"Error: {e}")
