import io
import re
import json
import base64
import requests
import streamlit as st
import pandas as pd
import gspread

from google.oauth2.service_account import Credentials as ServiceAccountCredentials
from google.oauth2.credentials import Credentials as OAuthCredentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

from datetime import date, datetime
from io import BytesIO

# Librerías para generación de Word
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# Librerías para generación de PDF
from reportlab.lib.pagesizes import letter
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image as RLImage,
    PageBreak
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader


# =========================================================
# CONFIGURACIÓN DE LA PÁGINA Y CONSTANTES
# =========================================================

st.set_page_config(
    page_title="Generar Informe Visual",
    page_icon="📄",
    layout="wide"
)


# =========================================================
# AUTENTICACIÓN GOOGLE
# =========================================================

if not st.user.is_logged_in:

    st.title("🔐 Autorización de Google")

    st.write(
        "Para utilizar este módulo debes iniciar sesión "
        "con la cuenta de Google que contiene tu Google Drive personal."
    )

    if st.button(
        "🔑 Iniciar sesión con Google",
        type="primary"
    ):
        st.login()

    st.stop()


# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================

URL_LOGO_GITHUB = (
    "https://raw.githubusercontent.com/death-87/"
    "app-inspecciones/main/logo.png"
)

SPREADSHEET_ID = (
    "1eJpQXWqe4AyyrFm_6wlnfzm-KYSGPeTtX_EWCIJYE1I"
)

DRIVE_FOLDER_ID = st.secrets.get(
    "DRIVE_FOLDER_ID",
    "1a2b3c4d5e6f7g8h9i_REEMPLAZAR_POR_TU_ID"
)


LISTA_INSPECTORES = [
    "Juan Navarrete",
    "Jorge Hernandez",
    "Arlem Sarmiento",
    "Harold Castillo",
    "Miguel Chirinos"
]


LISTA_PLANTAS = [
    "A0AEX", "A0ALQ", "A0BUT", "A0CCK", "A0CCR", "A0CKR",
    "A0HDG", "A0HDT", "A0HCK", "A0ISO", "A0LAB", "A0MHC",
    "A0NHT", "A0SAR", "A0SHP", "A0SWS", "AACID", "AAMAR",
    "AAMIN", "AAMPL", "AANTO", "AAREF", "AASER", "AALQU",
    "ADESO", "ADEV1", "ADEV2", "ADIPE", "AE501", "ALNHT",
    "ALPG1", "ALPG2", "ALPG3", "AMACO", "AMDEA", "AMRX1",
    "AMRX2", "AMRX3", "AMRX4", "AMVPR", "AOLEO", "APBMP",
    "APBTQ", "APCAR", "APFEN", "APRCO", "ARPLU", "AREFO",
    "AREMO", "ARILE", "ASAIC", "ASOLV", "ASPLI", "ASRCO",
    "ASUEL", "ASVAQ", "ASVAP", "ASWS2", "ASYBR", "ASEFL",
    "ASEFQ", "ATOP1", "ATOP2", "ATRAG", "AURA1", "AURA2",
    "AURA3", "AVAC1", "AVAC2", "ACOKE"
]


# =========================================================
# FUNCIONES AUXILIARES WORD
# =========================================================

def set_cell_background(cell, fill_hex):

    tcPr = cell._tc.get_or_add_tcPr()

    shd = parse_xml(
        f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'
    )

    tcPr.append(shd)


# =========================================================
# CONEXIÓN A GOOGLE SHEETS
# =========================================================

@st.cache_resource
def obtener_credenciales():

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    creds_dict = dict(
        st.secrets["connections"]["gsheets"]
    )

    if "private_key" in creds_dict:

        creds_dict["private_key"] = (
            creds_dict["private_key"]
            .replace("\\n", "\n")
        )

    return ServiceAccountCredentials.from_service_account_info(
        creds_dict,
        scopes=scopes
    )


@st.cache_resource
def conectar_google_sheets():

    credentials = obtener_credenciales()

    client = gspread.authorize(credentials)

    return client.open_by_key(SPREADSHEET_ID)


# =========================================================
# CONEXIÓN A GOOGLE DRIVE PERSONAL
# =========================================================

def conectar_google_drive():

    if not st.user.is_logged_in:

        st.error(
            "Debes iniciar sesión con Google para acceder "
            "a tu Google Drive."
        )

        return None

    try:

        access_token = st.user.tokens.get("access")

        if not access_token:

            st.error(
                "No se recibió el access token de Google."
            )

            return None

        credentials = OAuthCredentials(
            token=access_token
        )

        drive_service = build(
            "drive",
            "v3",
            credentials=credentials
        )

        return drive_service

    except Exception as e:

        st.error(
            "Error conectando con Google Drive mediante OAuth: "
            f"{e}"
        )

        return None


# =========================================================
# FUNCIONES DE ALMACENAMIENTO EN GOOGLE DRIVE
# =========================================================

def subir_imagen_a_drive(
    nombre_archivo,
    img_bytes
):

    try:

        drive_service = conectar_google_drive()

        if drive_service is None:
            return None

        file_metadata = {
            "name": nombre_archivo,
            "parents": [DRIVE_FOLDER_ID]
        }

        media = MediaIoBaseUpload(
            BytesIO(img_bytes),
            mimetype="image/jpeg",
            resumable=True
        )

        archivo_creado = drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id"
        ).execute()

        file_id = archivo_creado.get("id")

        if not file_id:

            st.error(
                f"Google Drive no devolvió ID para "
                f"'{nombre_archivo}'."
            )

            return None

        drive_service.permissions().create(
            fileId=file_id,
            body={
                "role": "reader",
                "type": "anyone"
            }
        ).execute()

        return (
            f"https://lh3.googleusercontent.com/d/{file_id}"
        )

    except Exception as e:

        st.error(
            f"Error al subir imagen '{nombre_archivo}' "
            f"a Google Drive: {e}"
        )

        return None


# =========================================================
# DESCARGA DE IMÁGENES DESDE URL
# =========================================================

@st.cache_data(ttl=3600)
def obtener_bytes_imagen(url):

    try:

        response = requests.get(
            url,
            timeout=10
        )

        if response.status_code == 200:

            return response.content

    except Exception:

        pass

    return None


# =========================================================
# LECTURA / ESCRITURA EN HISTORIAL
# =========================================================

@st.cache_data(ttl=300)
def cargar_base_equipos():

    try:

        client = conectar_google_sheets()

        ws = client.worksheet(
            "BASE EQUIPOS"
        )

        filas = ws.get_all_values()

        if len(filas) <= 1:

            return pd.DataFrame(
                columns=[
                    "UNIDAD",
                    "TAG",
                    "DESCRIPCION",
                    "ACA"
                ]
            )

        datos = []

        for f in filas[1:]:

            unidad = (
                f[2].strip()
                if len(f) > 2
                else ""
            )

            tag = (
                f[3].strip()
                if len(f) > 3
                else ""
            )

            descripcion = (
                f[4].strip()
                if len(f) > 4
                else ""
            )

            aca_val = (
                f[7].strip()
                if len(f) > 7
                else ""
            )

            if tag:

                datos.append(
                    {
                        "UNIDAD": unidad,
                        "TAG": tag,
                        "DESCRIPCION": descripcion,
                        "ACA": aca_val
                    }
                )

        return pd.DataFrame(datos)

    except Exception:

        return pd.DataFrame(
            columns=[
                "UNIDAD",
                "TAG",
                "DESCRIPCION",
                "ACA"
            ]
        )


def obtener_o_crear_hoja_historial():

    client = conectar_google_sheets()

    try:

        ws = client.worksheet(
            "HISTORIAL_INFORMES"
        )

    except Exception:

        ws = client.add_worksheet(
            title="HISTORIAL_INFORMES",
            rows="1000",
            cols="20"
        )

        ws.append_row(
            [
                "num_informe",
                "ot",
                "fecha",
                "unidad",
                "tag",
                "descripcion",
                "aca",
                "motivo",
                "alcance",
                "secciones_json",
                "inspector",
                "fotos_json"
            ]
        )

    return ws


# =========================================================
# GUARDAR INFORME
# =========================================================

def guardar_resguardo_informe(
    datos_encabezado,
    secciones_dinamicas,
    imagenes_procesadas,
    inspector_firma
):
    """
    Sube fotos a Google Drive personal y guarda
    los enlaces y metadata en Google Sheets.
    """

    try:
        ws = obtener_o_crear_hoja_historial()

        num_inf = datos_encabezado["num_informe"].strip()

        if not num_inf:
            return (
                False,
                "Debe ingresar un N.º DE INFORME para poder resguardar."
            )

        # =================================================
        # SUBIR FOTOGRAFÍAS A GOOGLE DRIVE
        # =================================================

        fotos_guardadas = []

        for idx, (img_bytes, pie) in enumerate(
            imagenes_procesadas,
            start=1
        ):
            nombre_foto = f"{num_inf}_foto_{idx}.jpg"

            url_drive = subir_imagen_a_drive(
                nombre_foto,
                img_bytes
            )

            fotos_guardadas.append(
                {
                    "url": url_drive if url_drive else "",
                    "pie": pie
                }
            )

        # =================================================
        # PREPARAR INFORMACIÓN PARA GOOGLE SHEETS
        # =================================================

        secciones_serializables = {
            str(k): v
            for k, v in secciones_dinamicas.items()
        }

        secciones_json = json.dumps(
            secciones_serializables,
            ensure_ascii=False
        )

        fotos_json = json.dumps(
            fotos_guardadas,
            ensure_ascii=False
        )

        fila_nueva = [
            num_inf,
            str(datos_encabezado["ot"]),
            datos_encabezado["fecha"].strftime("%Y-%m-%d"),
            str(datos_encabezado["unidad"]),
            str(datos_encabezado["tag"]),
            str(datos_encabezado["descripcion"]),
            str(datos_encabezado["aca"]),
            str(datos_encabezado["motivo"]),
            str(datos_encabezado["alcance"]),
            secciones_json,
            str(inspector_firma),
            fotos_json
        ]

        ws.append_row(
            fila_nueva,
            value_input_option="USER_ENTERED"
        )

        return (
            True,
            f"Informe {num_inf} guardado correctamente."
        )

    except Exception as e:
        return (
            False,
            f"Error al guardar el informe: {e}"
        )
