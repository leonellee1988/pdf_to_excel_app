import streamlit as st
import pdfplumber
import pandas as pd
import numpy as np
import re
import time
import zipfile
from io import BytesIO

# Columnas deseadas
columnas_deseadas = [
    "Empresa", "Autorización", "Serie", "Numero-DTE", "Fecha Emisión",
    "#No.", "Cantidad", "Descripcion", "P. Unitario con IVA (Q)",
    "Descuentos (Q)", "Total (Q)", "Impuestos", "Monto IVA"
]

def encontrar_columna_similar(columnas, patron_keywords):
    for col in columnas:
        col_normalizada = col.lower()
        if any(keyword in col_normalizada for keyword in patron_keywords):
            return col
    return None

def limpiar_campos_numericos(df, columnas):
    for col in columnas:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(r'[^\d.,]', '', regex=True)
            df[col] = df[col].str.replace(',', '', regex=False)
            df[col] = pd.to_numeric(df[col], errors='coerce')
    return df

def procesar_pdfs_desde_zip(zip_file):
    tabla_final = []
    archivos_procesados = 0
    inicio = time.time()

    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        pdf_names = [name for name in zip_ref.namelist() if name.lower().endswith('.pdf')]

        for archivo_nombre in pdf_names:
            archivos_procesados += 1
            with zip_ref.open(archivo_nombre) as archivo_pdf:
                with pdfplumber.open(BytesIO(archivo_pdf.read())) as pdf:
                    texto = pdf.pages[0].extract_text()

                    try:
                        lineas = texto.split('\n')
                        empresa = "Empresa no identificada"
                        for i, linea in enumerate(lineas):
                            if "Nit Emisor:" in linea and i > 0:
                                empresa_linea = lineas[i - 1].strip()
                                empresa = re.sub(r"NÚMERO DE AUTORIZACIÓN:.*", "", empresa_linea, flags=re.IGNORECASE).strip()
                                break

                        autorizacion = re.search(r"([A-Z0-9]{8}-[A-Z0-9\-]{27})", texto).group(1)
                        serie = re.search(r"Serie:\s+([A-Z0-9]+)", texto).group(1)
                        numero_dte_regex = re.search(r"Número de DTE:\s+(\d+)", texto)
                        numero_dte = numero_dte_regex.group(1) if numero_dte_regex else "No encontrado"
                        match_fecha = re.search(r"(\d{2}-[a-zA-Z]{3}-\d{4} \d{2}:\d{2}:\d{2})", texto)
                        fecha_emision = match_fecha.group(1) if match_fecha else ""
                    except Exception as e:
                        st.warning(f"⚠️ Error extrayendo encabezado de {archivo_nombre}: {e}")
                        continue

                    for pagina in pdf.pages:
                        tablas = pagina.extract_tables()
                        for tabla in tablas:
                            if tabla and len(tabla) > 1:
                                encabezado = tabla[0]
                                encabezado_limpio = []
                                for i, col in enumerate(encabezado):
                                    col = f"Col_{i}" if not col else re.sub(r'\s+', ' ', col.strip())
                                    encabezado_limpio.append(col)

                                encabezado_final = [
                                    f"{col}_{i}" if encabezado_limpio.count(col) > 1 else col
                                    for i, col in enumerate(encabezado_limpio)
                                ]

                                df = pd.DataFrame(tabla[1:], columns=encabezado_final)
                                df.replace(r'^\s*$', np.nan, regex=True, inplace=True)
                                df.dropna(how='all', inplace=True)
                                df.dropna(axis=1, how='all', inplace=True)

                                if not any("Cantidad" in col for col in df.columns):
                                    continue

                                df.columns = [col.strip() for col in df.columns]

                                col_precio_unitario = encontrar_columna_similar(df.columns, ["unitario", "valor", "precio"])
                                if col_precio_unitario and "P. Unitario con IVA (Q)" not in df.columns:
                                    df.rename(columns={col_precio_unitario: "P. Unitario con IVA (Q)"}, inplace=True)

                                df = df[~(
                                    df["Cantidad"].isna() &
                                    df["Descripcion"].isna() &
                                    df["P. Unitario con IVA (Q)"].isna()
                                )]

                                df = limpiar_campos_numericos(df, ["P. Unitario con IVA (Q)", "Descuentos (Q)"])

                                if not df.empty:
                                    df.insert(0, "Empresa", empresa)
                                    df.insert(1, "Autorización", autorizacion)
                                    df.insert(2, "Serie", serie)
                                    df.insert(3, "Numero-DTE", numero_dte)
                                    df.insert(4, "Fecha Emisión", fecha_emision)

                                    if "Col_8" in df.columns:
                                        df.rename(columns={"Col_8": "Monto IVA"}, inplace=True)

                                    columnas_presentes = [col for col in columnas_deseadas if col in df.columns]
                                    df = df[columnas_presentes]
                                    tabla_final.append(df)

    tiempo_total = round(time.time() - inicio, 2)

    if tabla_final:
        df_total = pd.concat(tabla_final, ignore_index=True)
        output = BytesIO()
        df_total.to_excel(output, index=False)
        output.seek(0)
        return output, archivos_procesados, tiempo_total
    else:
        return None, archivos_procesados, tiempo_total

# Interfaz Streamlit
st.set_page_config(page_title="Extractor de Facturas PDF", page_icon="📄", layout="centered")

st.title("📄 Extractor de Datos Facturas (ZIP de PDFs a Excel)")
st.markdown("Sube un archivo ZIP que contenga múltiples PDFs de facturas para generar un Excel consolidado.")

archivo_zip = st.file_uploader("📤 Carga tu archivo ZIP con PDFs", type="zip")

if st.button("Procesar ZIP"):
    if archivo_zip:
        excel_file, total_pdfs, duracion = procesar_pdfs_desde_zip(archivo_zip)
        if excel_file:
            st.success(f"✅ ¡Archivo generado con éxito! PDFs procesados: {total_pdfs} | Tiempo: {duracion} segundos")
            st.download_button("📥 Descargar Excel", data=excel_file, file_name="facturas_consolidadas.xlsx")
        else:
            st.warning(f"⚠️ No se encontraron tablas válidas. PDFs evaluados: {total_pdfs} | Tiempo: {duracion} segundos")
    else:
        st.error("❌ Por favor sube un archivo ZIP válido.")

with st.expander("ℹ️ Soporte Técnico"):
    st.markdown("""
    **Creador:** Edwin Leonel Lee Tiño  
    **Mail:** leonellee2016@gmail.com  
    **Teléfono:** 4087-3658  
    **Dirección:** Antigua Guatemala
    """)
