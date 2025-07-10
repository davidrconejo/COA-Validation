import streamlit as st
import pandas as pd
import fitz
import pytesseract
from PIL import Image
from io import BytesIO
from rapidfuzz import process, fuzz
from unidecode import unidecode
import io

st.set_page_config(page_title="Validación de COAs", layout="centered")
st.title("Validación Semanal de COAs (PDF escaneado + Lista Maestra Excel)")

def fuzzy_find(text, choices, threshold=80):
    found = []
    for choice, score, _ in process.extract(unidecode(text), choices, scorer=fuzz.partial_ratio, score_cutoff=threshold):
        found.append(choice)
    return list(set(found))

with st.form("upload_form"):
    st.subheader("1. Sube tu PDF escaneado")
    pdf_file = st.file_uploader("Archivo PDF", type=["pdf"], key="pdf")
    st.subheader("2. Sube la lista maestra (Excel)")
    master_file = st.file_uploader("Archivo Excel", type=["xls", "xlsx"], key="master")
    submit = st.form_submit_button("Procesar y Validar")

if submit and pdf_file and master_file:
    with st.spinner("Procesando archivos..."):
        # Leer lista maestra
        df_maestra = pd.read_excel(master_file)
        df_maestra.dropna(subset=['Fabricante', 'Lugar de Manufactura'], inplace=True)
        df_maestra['Fabricante'] = df_maestra['Fabricante'].apply(lambda x: unidecode(str(x)).lower().strip())
        df_maestra['Lugar de Manufactura'] = df_maestra['Lugar de Manufactura'].apply(lambda x: unidecode(str(x)).lower().strip())
        proveedores = df_maestra['Fabricante'].unique().tolist()
        lugares = df_maestra['Lugar de Manufactura'].unique().tolist()

        # Convertir PDF a imágenes
        pdf_doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
        images = []

    for page in pdf_doc:
        pix = page.get_pixmap(dpi=300)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        images.append(img)
    
        # Analizar cada página
        results = []
        for i, img in enumerate(images):
            text = pytesseract.image_to_string(img)
            text_norm = unidecode(text).lower()
            proveedores_detectados = fuzzy_find(text_norm, proveedores, threshold=80)
            lugares_detectados = fuzzy_find(text_norm, lugares, threshold=80)
            combinaciones_validas = [
                (prov, lugar)
                for prov in proveedores_detectados
                for lugar in lugares_detectados
                if ((df_maestra['Fabricante'] == prov) & (df_maestra['Lugar de Manufactura'] == lugar)).any()
            ]
            results.append({
                'Página': i + 1,
                'Proveedores Detectados': ', '.join(proveedores_detectados) or 'No encontrado',
                'Lugares Detectados': ', '.join(lugares_detectados) or 'No encontrado',
                'Combinacion Válida': len(combinaciones_validas) > 0,
                'Lista Combinaciones Validas': ', '.join([f'{p} + {l}' for p, l in combinaciones_validas]) or 'Ninguna'
            })

        df_resultados = pd.DataFrame(results)
        # Descargar resultados
        output = BytesIO()
        df_resultados.to_excel(output, index=False)
        output.seek(0)
        st.success("¡Verificación completada!")
        st.dataframe(df_resultados)
        st.download_button(
            label="Descargar resultados Excel",
            data=output,
            file_name="Verificacion_Semanal_COAs.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
elif submit:
    st.error("Por favor, sube ambos archivos para continuar.")
