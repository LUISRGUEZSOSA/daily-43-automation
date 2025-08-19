# -*- coding: utf-8 -*-
# config.py
import os
from dotenv import load_dotenv

load_dotenv()  # carga .env si existe

def _parse_tiendas(s: str | None):
    if not s: 
        return None  # equivalente a “todas”
    s = s.strip()
    if not s:
        return None
    try:
        return [int(x) for x in s.split(",") if x.strip() != ""]
    except Exception:
        return None

# --- Credenciales/API ---
BASE = os.getenv("TOUCH_BASE", "https://touchm.net/interfacestouchm/public-api/v1/TE_Interfaces")
USER = os.getenv("TOUCH_USER", "")
PASSWORD = os.getenv("TOUCH_PASSWORD", "")
TIMEOUT = int(os.getenv("TOUCH_TIMEOUT", "45"))

# --- Tiendas ---
TIENDAS = _parse_tiendas(os.getenv("TOUCH_TIENDAS"))

# --- Plantilla y salidas ---
TEMPLATE_XLSX = os.getenv("DAILY_TEMPLATE", "Daily plantilla 2025.xlsx")
OUTPUT_DIR = os.getenv("DAILY_OUTPUT_DIR", ".")

# --- Hoja y cabecera destino en Excel ---
TARGET_SHEET = "BBDDcoste"

CSV_COLUMNS = [
    "IDTRANS","NSERIE","SERIE","NUMTIKET","NUMBARRA","NNUMBARRA","FECHA","JORNADA",
    "CREDITO","NCREDITO","NUMCLIE","PUNTOVENTA","NPUNTOVENTA","NUMCUEN","NNUMCUEN",
    "SERVICIO","NSERVICIO","ALMACEN","NALMACEN","CABIMPORTE","CABDESCUENTO","CABNETO",
    "CAMARERO","NCAMARERO","MACROGRUPO","NMACROGRUPO","GRUPO","NGRUPO","FAMILIA","NFAMILIA",
    "TIPOPRODUCTO","NTIPOPRODUCTO","PRODUCTO","NPRODUCTO","CANTIDAD","PRECIO","IVA",
    "IMPORTE","IMPORTESINIVA","DESCUENTO","IMPORTEDESCUENTO","IMPORTESINIVADESCUENTO",
    "ANULADA","FORMATO","NFORMATO","TIENDA","ESTABLECIMIENTO","CTACONTABLE","CECO",
    "NIFCLIENTE","NOMBRECLIENTE","VENCIMIENTO","PROMOCION","COMENSALES","COSTE",
    "OBSERVACIONES","Turno","Denominacion 2","Factura","Motivo"
]
