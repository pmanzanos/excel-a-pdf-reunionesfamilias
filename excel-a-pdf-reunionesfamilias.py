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
        # Lugar y fecha formateada: En Hoyos, a X de mes de año
        self.set_font('helvetica', 'I', 10)
        self.cell(0, 10, f"En Hoyos, a {fecha_texto}", 0, 1, 'L')
        self.ln(5)
        
        # Formato: Fdo. Nombre (Cargo)
        self.set_font('helvetica', '', 10)
        for cargo, nombre in lista_firmas:
            # Aseguramos que el nombre lleve el prefijo Fdo. si no lo tiene
            fdo_nombre = f"Fdo. {nombre}" if not str(nombre).lower().startswith("fdo") else nombre
            texto_firma = f"{fdo_nombre} ({cargo})"
            
            self.multi_cell(0, 7, texto_firma.encode('latin-1', 'replace').decode('latin-1'))
            self.ln(1)

def formatear_fecha_espanol(fecha_obj):
    try:
        if isinstance(fecha_obj, str):
            fecha_obj = pd.to_datetime(fecha_obj)
        meses = [
            "enero", "febrero", "marzo", "abril", "mayo", "junio",
            "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
        ]
        return f"{fecha_obj.day} de {meses[fecha_obj.month - 1]} de {fecha_obj.year}"
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
    
    fecha_com = formatear_fecha_espanol(fila_rpts.iloc[4])
    
    pdf.seccion("DATOS DE LA REUNIÓN")
    pdf.campo("ID", fila_rpts['ID_GENERADA']) 
    pdf.campo("LA REUNIÓN SE PRODUCE A PETICIÓN DE", fila_rpts.iloc[3])
    pdf.campo("FECHA Y HORA DE LA COMUNICACIÓN", f"{fecha_com} a las {fila_rpts.iloc[5]}")
    
    solo_nombre, solo_curso = limpiar_nombre_y_curso(fila_rpts.iloc[7])
    pdf.campo("ALUMN@/O", solo_nombre) 
    pdf.campo("CURSO", solo_curso)
    pdf.campo("FAMILIAR/ES PRESENTES", fila_rpts.iloc[9])
    pdf.campo("OTROS PRESENTES", fila_rpts.iloc[8])
    
    pdf.ln(2); pdf.seccion("DESARROLLO DE LA REUNIÓN")
    pdf.campo("ASUNTO A TRATAR", fila_rpts.iloc[10], 'B')
    pdf.campo("DESCRIPCIÓN DE LO TRATADO", fila_rpts.iloc[12])
    pdf.campo("TRÁMITE A SEGUIR", fila_rpts.iloc[11])

    # Bloque de firmas con el nuevo formato solicitado
    pdf.dibujar_bloque_firmas(lista_firmas, fecha_com)
    return pdf.output()

# --- INTERFAZ STREAMLIT ---
st.set_page_config(page_title="Generador Actas", page_icon="🤝")
st.title("🤝 Generador de Actas Oficiales")

archivo = st.file_uploader("Sube el archivo Excel", type=['xlsx'])

if archivo:
    try:
        # Pestaña 1 para datos: RPTS
        df_rpts = pd.read_excel(archivo, sheet_name='RPTS')
        
        # Pestaña 2 para firmas: Informe
        df_inf = pd.read_excel(archivo, sheet_name='Informe', header=None)

        # Re-generar ID para el buscador
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
        seleccion = st.selectbox("Selecciona la reunión:", opciones)

        if seleccion != "Selecciona un acta...":
            id_buscada = seleccion.split(" - ")[0]
            fila_sel = df_rpts[df_rpts['ID_GENERADA'] == id_buscada].iloc[0]
            
            # Obtener firmas de la pestaña 'Informe' desde fila 49 (índice 48)
            firmas_finales = []
            for i in range(48, 65):
                try:
                    cargo = df_inf.iloc[i, 0]  # Columna A
                    nombre = df_inf.iloc[i, 2] # Columna C
                    if pd.notna(cargo) and pd.notna(nombre) and str(nombre).strip() not in ["", "nan", "0"]:
                        # Quitamos dos puntos si existen al final del cargo
                        cargo_limpio = str(cargo).strip().rstrip(":")
                        firmas_finales.append((cargo_limpio, str(nombre).strip()))
                except:
                    continue

            if st.button("🚀 Crear y Descargar PDF"):
                pdf_bytes = generar_pdf_reunion(fila_sel, firmas_finales)
                st.download_button(
                    label="⬇️ Descargar Acta PDF",
                    data=bytes(pdf_bytes),
                    file_name=f"Acta_{id_buscada}.pdf",
                    mime="application/pdf"
                )

    except Exception as e:
        st.error(f"⚠️ Error al procesar: {e}")
