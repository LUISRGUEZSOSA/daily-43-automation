#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import argparse
from datetime import date
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

def build_drive(creds_json: str):
    scopes = ["https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file(creds_json, scopes=scopes)
    return build("drive", "v3", credentials=creds)

def upload_excel(drive, filepath: str, dest_name: str = None, folder_id: str = None, replace: bool = False):
    if not dest_name:
        dest_name = os.path.basename(filepath)

    file_metadata = {"name": dest_name}
    if folder_id:
        file_metadata["parents"] = [folder_id]

    media = MediaFileUpload(
        filepath,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True,
    )

    # Reemplazo opcional por nombre en la carpeta (si existe)
    if replace:
        safe_name = dest_name.replace("'", "\\'")
        q = (
            f"name = '{safe_name}' "
            f"and mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' "
            f"and trashed = false"
        )
        if folder_id:
            q += f" and '{folder_id}' in parents"
        res = drive.files().list(
            q=q,
            fields="files(id,name)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        files = res.get("files", [])
        if files:
            file_id = files[0]["id"]
            updated = drive.files().update(
                fileId=file_id,
                media_body=media,
                fields="id, webViewLink, webContentLink",
                supportsAllDrives=True,
            ).execute()
            return updated

    created = drive.files().create(
        body=file_metadata,
        media_body=media,
        fields="id, webViewLink, webContentLink",
        supportsAllDrives=True,
    ).execute()
    return created

def main():
    ap = argparse.ArgumentParser(description="Sube un Excel a Google Drive (sin convertir).")
    ap.add_argument("--creds", required=True, help="Ruta al JSON de Service Account")
    ap.add_argument("--xlsx", help="Ruta al XLSX; por defecto usa Daily_YYYY-MM-DD.xlsx en cwd")
    ap.add_argument("--name", help="Nombre destino en Drive (por defecto: basename del XLSX)")
    ap.add_argument("--folder-id", help="ID de carpeta de Drive donde subir (opcional)")
    ap.add_argument("--replace", action="store_true", help="Si existe un archivo con el mismo nombre, lo reemplaza")
    args = ap.parse_args()

    if not args.xlsx:
        args.xlsx = os.path.join(os.getcwd(), f"Daily_{date.today().isoformat()}.xlsx")
    if not os.path.exists(args.xlsx):
        raise SystemExit(f"No encuentro el fichero: {args.xlsx}")

    drive = build_drive(args.creds)
    res = upload_excel(drive, args.xlsx, dest_name=args.name, folder_id=args.folder_id, replace=args.replace)
    print("✅ Subido a Drive:")
    print("  ID:          ", res.get("id"))
    print("  Ver online:  ", res.get("webViewLink"))
    print("  Descargar:   ", res.get("webContentLink"))

if __name__ == "__main__":
    
    if not args.creds:
        env_path = os.getenv("GOOGLE_SA_JSON")
        if env_path:
            args.creds = env_path
        else:
            raise SystemExit("--creds requerido o GOOGLE_SA_JSON en .env")

    if not args.name:
        # opcional: nombre por defecto
        args.name = os.path.basename(args.xlsx)
    # folder-id puede venir por CLI o por env (si lo quieres):
    if not args.folder_id:
        args.folder_id = os.getenv("GDRIVE_FOLDER_ID")
        
    main()
