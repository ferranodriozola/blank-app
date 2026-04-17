import streamlit as st
import pandas as pd

# --- CONFIGURACIÓ DE LA PÀGINA ---
st.set_page_config(page_title="Visor de Text Pla", page_icon="📄")

st.title("📄 Dades en format Text")

# ID del teu Google Sheet
SHEET_ID = "1FvNGh_SySwgVFaPHBzAxd6EocWnOCPQ-kUTFe1TIWSE"
URL_CSV = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

if st.button('Convertir dades a text'):
    with st.spinner('Llegint dades...'):
        try:
            # 1. Llegim l'Excel
            df = pd.read_csv(URL_CSV)
            
            # 2. Agafem el rang A1:H10 (Files 0-10, Columnes 0-8)
            df_filtrat = df.iloc[0:10, 0:8]
            
            # 3. Convertim la taula a una cadena de text (String)
            # index=False fa que no surtin els números de fila a l'esquerra
            text_pla = df_filtrat.to_string(index=False)
            
            st.success("Aquí tens el resultat en text pla:")
            
            # 4. Imprimim a la web dins d'un bloc de codi
            st.code(text_pla, language='text')
            
            # Opcional: Si vols un format de text encara més simple i petit:
            # st.text(text_pla)

        except Exception as e:
            st.error(f"Error al processar: {e}")
else:
    st.info("Prem el botó per transformar l'Excel en text.")