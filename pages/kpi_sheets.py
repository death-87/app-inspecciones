"""Almacenamiento de indicadores en una pestaña independiente."""
import hashlib
import json
from datetime import datetime, timezone

SPREADSHEET_ID = "1eJpQXWqe4AyyrFm_6wlnfzm-KYSGPeTtX_EWCIJYE1I"
HOJA = "Indicadores_Inspeccion"
CAMPOS = ["Fecha", "Semana", "Inspector", "Equipo_TAG", "Horas", "Estado", "Hallazgos"]
ENCABEZADOS = ["ID", *CAMPOS, "Usuario", "Creado_UTC"]


def conectar(secretos):
    import gspread
    from google.oauth2.service_account import Credentials
    config = dict(secretos)
    config["private_key"] = config["private_key"].replace("\\n", "\n")
    credenciales = Credentials.from_service_account_info(config, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    return gspread.authorize(credenciales).open_by_key(SPREADSHEET_ID)


def obtener_hoja(libro, crear=False):
    from gspread.exceptions import WorksheetNotFound
    try:
        return libro.worksheet(HOJA)
    except WorksheetNotFound:
        if not crear:
            return None
        hoja = libro.add_worksheet(title=HOJA, rows=1000, cols=len(ENCABEZADOS))
        hoja.append_row(ENCABEZADOS, value_input_option="RAW")
        return hoja


def leer_hoja(hoja):
    valores = hoja.get_all_values()
    if not valores or valores[0] != ENCABEZADOS:
        raise ValueError(f"La pestaña {HOJA} no tiene los encabezados esperados.")
    resultado = []
    for numero, fila in enumerate(valores[1:], start=2):
        if not any(fila):
            continue
        if len(fila) > len(ENCABEZADOS):
            raise ValueError(f"La fila {numero} contiene columnas inesperadas.")
        fila = fila + [""] * (len(ENCABEZADOS) - len(fila))
        resultado.append(dict(zip(ENCABEZADOS, fila)))
    return resultado


def cargar(libro):
    hoja = obtener_hoja(libro)
    return leer_hoja(hoja) if hoja is not None else []


def guardar(libro, fila, usuario):
    if usuario != "jnavarrete":
        raise PermissionError("Usuario no autorizado.")
    normalizada = {c: fila[c] for c in CAMPOS}
    normalizada["Equipo_TAG"] = str(normalizada["Equipo_TAG"]).strip().upper()
    normalizada["Horas"] = float(normalizada["Horas"])
    normalizada["Hallazgos"] = int(normalizada["Hallazgos"])
    normalizada["Semana"] = int(normalizada["Semana"])
    contenido = json.dumps(normalizada, ensure_ascii=False, sort_keys=True)
    registro_id = hashlib.sha256(contenido.encode("utf-8")).hexdigest()
    hoja = obtener_hoja(libro, crear=True)
    if any(r["ID"] == registro_id for r in leer_hoja(hoja)):
        return False
    valores = [registro_id, *[normalizada[c] for c in CAMPOS], usuario, datetime.now(timezone.utc).isoformat()]
    try:
        hoja.append_row(valores, value_input_option="RAW")
    except Exception:
        # Comprobar si Google guardó antes de interrumpirse la respuesta.
        try:
            confirmado = any(r["ID"] == registro_id for r in leer_hoja(hoja))
        except Exception:
            confirmado = False
        if not confirmado:
            raise
    return True
