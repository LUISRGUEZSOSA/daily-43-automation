#!/usr/bin/env python3
# -*- coding: utf-8 -*-
#build_daily_today.py
import os, shutil
from datetime import date
from config import TEMPLATE_XLSX, OUTPUT_DIR, TARGET_SHEET
from excel_writer import overwrite_non_formula_cells_with_csv

def main():
    hoy = date.today()
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    in_csv = os.path.join(OUTPUT_DIR, f"ventas_{hoy.isoformat()}.csv")
    if not os.path.exists(in_csv):
        raise SystemExit(f"No existe el CSV de hoy: {in_csv}. Ejecuta primero fetch_today.py")

    out_xlsx = os.path.join(OUTPUT_DIR, f"Daily_{hoy.isoformat()}.xlsx")
    shutil.copy2(TEMPLATE_XLSX, out_xlsx)
    print(f"Plantilla copiada a: {out_xlsx}")

    # 🔑 Solo tocamos celdas SIN fórmula en BBDDcoste
    overwrite_non_formula_cells_with_csv(out_xlsx, TARGET_SHEET, in_csv, backup=False)

    print(f"✅ Daily del día generado: {out_xlsx}")

if __name__ == "__main__":
    main()
