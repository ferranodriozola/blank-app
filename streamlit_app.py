import streamlit as st
import pandas as pd

# --- CONFIGURACIÓ DE LA PÀGINA ---
st.set_page_config(page_title="Eina de Dades Multiformat", page_icon="🛠️")

st.title("🛠️ Visor de Dades de l'Equip")
st.write("Tria com vols visualitzar les dades del Google Sheet.")

# 1. Configuració de la font de dades
SHEET_ID = "1FvNGh_SySwgVFaPHBzAxd6EocWnOCPQ-kUTFe1TIWSE"
URL_CSV = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

# 2. Interfície de selecció
opcio = st.radio(
    "Selecciona el format de sortida:",
    ('Taula Interactiva', 'Text Pla (Còpia ràpida)'),
    index=0
)

# Afegim un separador visual
st.divider()

# 3. Botó d'execució
if st.button('Carregar i Convertir'):
    with st.spinner('Processant dades...'):
        try:
            # Llegim les dades
            df = pd.read_csv(URL_CSV)
            
            # Filtrem el rang A1:H10
            df_filtrat = df.iloc[0:10, 0:8]
            
            if opcio == 'Taula Interactiva':
                st.subheader("Visualització en format Taula")
                st.dataframe(df_filtrat, use_container_width=True)
                st.info("💡 Consell: Pots clicar a les capçaleres per ordenar les dades.")
            
            else:
                st.subheader("Visualització en format Text")
                text_pla = df_filtrat.to_string(index=False)
                # Fem servir language='text' perquè no ressalti colors estranys
                st.code(text_pla, language='text')
                st.info("💡 Consell: Usa el botó de la dreta per copiar tot el text al portaretalls.")

        except Exception as e:
            st.error(f"S'ha produït un error en la connexió: {e}")
else:
    st.info("Selecciona un format a dalt i prem el botó per començar.")