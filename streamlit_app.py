import streamlit as st
import pandas as pd
import requests
import certifi
from io import BytesIO
from xml.sax.saxutils import escape

# --- CONFIGURACIÓ DE LA PÀGINA ---
st.set_page_config(page_title="Dades", page_icon="🛠️")

st.title("Visor de dades")

# 1. Configuració de la font de dades
URL_XLSX = f"https://docs.google.com/spreadsheets/d/{st.secrets['SHEET_ID']}/export?format=xlsx"


@st.cache_data
def descarregar_excel(url_xlsx: str) -> bytes:
    resposta = requests.get(url_xlsx, timeout=30, verify=certifi.where())
    resposta.raise_for_status()
    return resposta.content


@st.cache_data
def obtenir_fulls(url_xlsx: str):
    xls = pd.ExcelFile(BytesIO(descarregar_excel(url_xlsx)))
    return xls.sheet_names


@st.cache_data
def llegir_full(url_xlsx: str, full: str):
    return pd.read_excel(BytesIO(descarregar_excel(url_xlsx)), sheet_name=full)


def _text_segura(valor) -> str:
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def _llista_camp(valor) -> list[str]:
    text = _text_segura(valor)
    if not text:
        return []
    return [item.strip() for item in text.split(",") if item.strip()]


def _etiqueta_opcional(tag: str, contingut: str, cert: str = "") -> str:
    atribut_cert = f' cert="{escape(cert)}"' if cert else ""
    return f'   <{tag}{atribut_cert}>{escape(contingut)}</{tag}>'


def construir_person_xml(fila: pd.Series) -> str:
    nom = _text_segura(fila.get('FDHV'))
    xml_id = _text_segura(fila.get('ID'))
    role = _text_segura(fila.get('role'))
    ref = _text_segura(fila.get('Ref VIAF'))

    ocupacions = _llista_camp(fila.get("occupation (més d'una?)"))
    birth = _text_segura(fila.get('birth'))
    death = _text_segura(fila.get('death'))
    cert_birth = _text_segura(fila.get('certainty1'))
    cert_death = _text_segura(fila.get('certainty2'))
    lang_knowledge = _llista_camp(fila.get("lang knowledge (més d'un?)"))
    faith = _text_segura(fila.get('faith'))

    linies = [f'<person xml:id="{escape(xml_id)}">']

    attrs_persname = []
    if role:
        attrs_persname.append(f'role="{escape(role)}"')
    if ref:
        attrs_persname.append(f'ref="{escape(ref)}"')

    attrs_text = f" {' '.join(attrs_persname)}" if attrs_persname else ""
    linies.append(f'   <persName{attrs_text}>{escape(nom)}</persName>')

    for ocupacio in ocupacions:
        linies.append(f'   <occupation>{escape(ocupacio)}</occupation>')

    if birth:
        linies.append(_etiqueta_opcional('birth', birth, cert_birth))
    if death:
        linies.append(_etiqueta_opcional('death', death, cert_death))
    if lang_knowledge:
        linies.append('   <langKnowledge>')
        for idioma in lang_knowledge:
            linies.append(f'      <langKnown tag="***">{escape(idioma)}</langKnown>')
        linies.append('   </langKnowledge>')
    if faith:
        linies.append(f'   <faith>{escape(faith)}</faith>')

    linies.append('</person>')
    return '\n'.join(linies)

fulls_disponibles = obtenir_fulls(URL_XLSX)[:2]

if not fulls_disponibles:
    st.error("No s'han trobat fulls disponibles a l'Excel.")
    st.stop()

# 2. Interfície de selecció
full_seleccionat = st.selectbox("Selecciona el full de l'Excel:", fulls_disponibles)

# Afegim un separador visual
st.divider()

# 3. Botó d'execució
if st.button('Carregar i Convertir'):
    with st.spinner('Processant dades...'):
        try:
            # Llegim el full seleccionat
            df = llegir_full(URL_XLSX, full_seleccionat)
            
            # Mostrem les primeres files i columnes utilitzades per les llistes
            df_filtrat = df.iloc[0:250, 0:11]
            
            st.subheader("Personatges en format XML")

            files_valides = df_filtrat.copy()
            if 'ID' in files_valides.columns:
                files_valides = files_valides[files_valides['ID'].notna()]
            if 'FDHV' in files_valides.columns:
                files_valides = files_valides[files_valides['FDHV'].notna()]

            if files_valides.empty:
                st.warning("No hi ha personatges vàlids al full seleccionat.")
            else:
                for _, fila in files_valides.iterrows():
                    nom = _text_segura(fila.get('FDHV')) or '(sense nom)'
                    xml_id = _text_segura(fila.get('ID')) or '(sense id)'
                    st.markdown(f"**{nom} ({xml_id})**")
                    st.code(construir_person_xml(fila), language='xml')

        except Exception as e:
            st.error(f"S'ha produït un error en la connexió: {e}")
else:
    st.info("Selecciona un full i prem el botó per començar.")
