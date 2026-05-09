import streamlit as st
import pandas as pd
import requests
import certifi
from io import BytesIO
from xml.sax.saxutils import escape

st.set_page_config(page_title="Llista a xml", page_icon="🛠️")

st.title("XLSX a XML")

URL_XLSX = f"https://docs.google.com/spreadsheets/d/{st.secrets['SHEET_ID']}/export?format=xlsx"

PERSON_COLS = {
    'name': 0,           
    'id': 1,             
    'ref_viaf': 2,       
    'ref_2': 3,
    'ref_3': 4,
    'role': 5,           
    'occupation': 6,     
    'birth': 7,          
    'certainty1': 8,     
    'death': 9,          
    'certainty2': 10,     
    'lang_knowledge': 11, 
    'faith': 12,         
}

PLACE_COLS = {
    'name': 0,      
    'id': 1,        
    'country': 2,   
    'ref_maps': 3,  
    'latitude': 4,  
    'longitude': 5, 
}


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
    nom = _text_segura(fila.iloc[PERSON_COLS['name']])
    xml_id = _text_segura(fila.iloc[PERSON_COLS['id']])
    role = _text_segura(fila.iloc[PERSON_COLS['role']])
    ref = _text_segura(fila.iloc[PERSON_COLS['ref_viaf']])
    ref_2 = _text_segura(fila.iloc[PERSON_COLS['ref_2']])
    ref_3 = _text_segura(fila.iloc[PERSON_COLS['ref_3']])

    ocupacions = _llista_camp(fila.iloc[PERSON_COLS['occupation']])
    birth = _text_segura(fila.iloc[PERSON_COLS['birth']])
    death = _text_segura(fila.iloc[PERSON_COLS['death']])
    cert_birth = _text_segura(fila.iloc[PERSON_COLS['certainty1']])
    cert_death = _text_segura(fila.iloc[PERSON_COLS['certainty2']])
    lang_knowledge = _llista_camp(fila.iloc[PERSON_COLS['lang_knowledge']])
    faith = _text_segura(fila.iloc[PERSON_COLS['faith']])

    linies = [f'<!-- {escape(nom)} -->\n<person xml:id="{escape(xml_id)}">']

    attrs_persname = []
    if role:
        attrs_persname.append(f'role="{escape(role)}"')
    if ref:
        attrs_persname.append(f'ref="{escape(ref)}"')
    if ref_2:
        attrs_persname.append(f'ref="{escape(ref_2)}"')
    if ref_3:
        attrs_persname.append(f'ref="{escape(ref_3)}"')

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

    linies.append('</person>\n')
    return '\n'.join(linies)

def construir_place_xml(fila: pd.Series) -> str:
    nom = _text_segura(fila.iloc[PLACE_COLS['name']])
    xml_id = _text_segura(fila.iloc[PLACE_COLS['id']]) or nom
    ref = _text_segura(fila.iloc[PLACE_COLS['ref_maps']])
    pais = _text_segura(fila.iloc[PLACE_COLS['country']])
    latitud = _text_segura(fila.iloc[PLACE_COLS['latitude']])
    longitud = _text_segura(fila.iloc[PLACE_COLS['longitude']])

    linies = [f'<!-- {escape(nom)} -->\n<place xml:id="{escape(xml_id)}">']

    attrs_placename = []
    if ref:
        attrs_placename.append(f'ref="{escape(ref)}"')

    attrs_text = f" {' '.join(attrs_placename)}" if attrs_placename else ""
    linies.append(f'   <placeName{attrs_text}>{escape(nom)}</placeName>')

    if pais:
        linies.append(f'   <country>{escape(pais)}</country>')

    if latitud or longitud:
        linies.append('   <location>')
        geo_text = f"{latitud}, {longitud}" if (latitud and longitud) else (latitud or longitud)
        linies.append(f'      <geo>{escape(geo_text)}</geo>')
        linies.append('   </location>')

    linies.append('</place>\n')
    return '\n'.join(linies)

fulls_disponibles = obtenir_fulls(URL_XLSX)[:2]

if not fulls_disponibles:
    st.error("No s'han trobat fulls disponibles a l'Excel.")
    st.stop()

full_seleccionat = st.selectbox("Selecciona el full de l'Excel:", fulls_disponibles)

if st.button('Forçar recàrrega'):
    descarregar_excel.clear()
    obtenir_fulls.clear()
    llegir_full.clear()
    if hasattr(st, 'rerun'):
        st.rerun()
    else:
        st.stop()

with st.spinner('Processant dades...'):
    try:
        df = llegir_full(URL_XLSX, full_seleccionat)

        df_filtrat = df.iloc[0:1000, 0:12]

        if full_seleccionat == 'listPerson':
            st.subheader("Personatges en format XML")

            files_valides = df_filtrat.copy()
            files_valides = files_valides[files_valides.iloc[:, PERSON_COLS['id']].notna()]
            files_valides = files_valides[files_valides.iloc[:, PERSON_COLS['name']].notna()]

            if files_valides.empty:
                st.warning("No hi ha personatges vàlids al full seleccionat.")
            else:
                for _, fila in files_valides.iterrows():
                    nom = _text_segura(fila.iloc[PERSON_COLS['name']]) or '(sense nom)'
                    xml_id = _text_segura(fila.iloc[PERSON_COLS['id']]) or '(sense id)'
                    st.markdown(f"**{nom} ({xml_id})**")
                    st.code(construir_person_xml(fila), language='xml')

        elif full_seleccionat == 'listPlace':
            st.subheader("Llocs en format XML")

            files_valides = df_filtrat.copy()
            files_valides = files_valides[files_valides.iloc[:, PLACE_COLS['id']].notna()]
            files_valides = files_valides[files_valides.iloc[:, PLACE_COLS['name']].notna()]

            if files_valides.empty:
                st.warning("No hi ha llocs vàlids al full seleccionat.")
            else:
                for _, fila in files_valides.iterrows():
                    nom = _text_segura(fila.iloc[PLACE_COLS['name']]) or '(sense nom)'
                    xml_id = _text_segura(fila.iloc[PLACE_COLS['id']]) or nom
                    st.markdown(f"**{nom} ({xml_id})**")
                    st.code(construir_place_xml(fila), language='xml')

        else:
            st.warning("Full no reconegut. Prova amb listPerson o listPlace.")

    except Exception as e:
        st.error(f"S'ha produït un error en la connexió: {e}")