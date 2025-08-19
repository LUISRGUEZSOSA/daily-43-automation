#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import argparse
from datetime import date

import pandas as pd
import gspread
from gspread_dataframe import set_with_dataframe
from google.oauth2.service_account import Credentials

try:
    # Carga .env si existe (opcional)
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

# --- Config por defecto ---
DEFAULT_WORKSHEET = "Daily"

def load_service_account(creds_json_path: str):
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(creds_json_path, scopes=scopes)
    client = gspread.authorize(creds)
    return client

def read_daily_from_xlsx(xlsx_path: str, sheet_name: str = DEFAULT_WORKSHEET) -> pd.DataFrame:
    # Lee la hoja 'Daily'; pandas toma el valor visible en celdas merged
    df = pd.read_excel(xlsx_path, sheet_name=sheet_name, engine="openpyxl", dtype=object)
    return df.fillna("")

def clear_worksheet(ws):
    ws.clear()

def push_dataframe(df: pd.DataFrame, ws):
    # RAW por defecto en gspread_dataframe
    set_with_dataframe(ws, df, include_index=False, include_column_header=True, resize=True)

def main():
    ap = argparse.ArgumentParser(
        description="Sube la hoja 'Daily' de un XLSX a un Google Sheet (pestaña Daily)."
    )
    ap.add_argument("--creds", help="Ruta al JSON de Service Account (o GOOGLE_SA_JSON en .env)")
    ap.add_argument("--sheet-id", help="Spreadsheet ID destino (o GOOGLE_SHEET_ID en .env)")
    ap.add_argument("--xlsx", help="Ruta al XLSX con la hoja 'Daily'. Por defecto: Daily_YYYY-MM-DD.xlsx en cwd")
    ap.add_argument("--worksheet", default=DEFAULT_WORKSHEET, help="Nombre de la pestaña destino (por defecto: Daily)")
    ap.add_argument("--source-sheet", default=DEFAULT_WORKSHEET, help="Nombre de la hoja en el XLSX origen (por defecto: Daily)")
    args = ap.parse_args()

    # Fallbacks desde .env
    if not args.sheet_id:
        env_id = os.getenv("GOOGLE_SHEET_ID")
        if not env_id:
            raise SystemExit("--sheet-id requerido o variable GOOGLE_SHEET_ID en .env")
        args.sheet_id = env_id

    if not args.creds:
        env_path = os.getenv("GOOGLE_SA_JSON")
        if not env_path:
            raise SystemExit("--creds requerido o variable GOOGLE_SA_JSON en .env")
        args.creds = env_path

    if not args.xlsx:
        hoy = date.today().isoformat()
        args.xlsx = os.path.join(os.getcwd(), f"Daily_{hoy}.xlsx")

    if not os.path.exists(args.xlsx):
        raise SystemExit(f"No encuentro el fichero: {args.xlsx}")

    print(f"→ Cargando XLSX: {args.xlsx} (hoja origen: {args.source_sheet})")
    df = read_daily_from_xlsx(args.xlsx, sheet_name=args.source_sheet)
    print(f"→ Filas: {len(df)} · Columnas: {len(df.columns)}")

    gc = load_service_account(args.creds)
    sh = gc.open_by_key(args.sheet_id)

    # Crea la pestaña si no existe
    try:
        ws = sh.worksheet(args.worksheet)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(
            title=args.worksheet,
            rows=str(max(1000, len(df) + 10)),
            cols=str(max(26, len(df.columns) + 5)),
        )

    print(f"→ Limpiando pestaña destino: {args.worksheet}")
    clear_worksheet(ws)

    print("→ Subiendo datos…")
    push_dataframe(df, ws)

    print("✅ Listo. Google Sheet actualizado.")

if __name__ == "__main__":
    main()
