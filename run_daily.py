#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
run_daily.py
-------------
Pipeline diario (sin romper fórmulas):

  1) fetch_today.main() -> genera ventas_YYYY-MM-DD.csv del día actual (o DAILY_DATE)
  2) build_daily_today.main() -> copia plantilla y escribe BBDDcoste (solo celdas sin fórmula)

Para testear otro día:
  DAILY_DATE=2025-02-28 python3 run_daily.py
"""

import importlib
import sys

def main():
    print("== Paso 1/2: obtener CSV del día ==")
    importlib.import_module("fetch_today").main()

    print("\n== Paso 2/2: construir Daily (copiar plantilla + BBDDcoste) ==")
    importlib.import_module("build_daily_today").main()

if __name__ == "__main__":
    try:
        main()
        print("\n✅ Proceso diario completado (cabecera intacta; Excel recalculará al abrir).")
        sys.exit(0)
    except SystemExit:
        raise
    except Exception as e:
        print(f"\n❌ Error: {e}")
        sys.exit(1)
