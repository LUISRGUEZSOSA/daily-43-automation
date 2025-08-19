#!/usr/bin/env python3
# -*- coding: utf-8 -*-
#fetch_today.py

import csv, base64, json, os
from datetime import date, datetime, timedelta
from urllib import request, error
from typing import Dict, Any, List, Tuple, Optional

from config import BASE, USER, PASSWORD, TIMEOUT, TIENDAS, OUTPUT_DIR, CSV_COLUMNS

# ---------- fecha objetivo ----------
def get_target_date() -> date:
    s = os.environ.get("DAILY_DATE")
    if s:
        return date.fromisoformat(s)
    return date.today()

# ---------- HTTP ----------
def http_post_json(url: str, payload: dict) -> Any:
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = request.Request(url, data=data, method="POST")
    req.add_header("Content-Type", "application/json")
    req.add_header("Accept", "application/json")
    token = base64.b64encode(f"{USER}:{PASSWORD}".encode()).decode()
    req.add_header("Authorization", f"Basic {token}")
    with request.urlopen(req, timeout=TIMEOUT) as r:
        raw = r.read().decode("utf-8", errors="replace")

    try:
        outer = json.loads(raw)
        if isinstance(outer, str):
            return json.loads(outer)
        return outer
    except json.JSONDecodeError:
        if raw.startswith('"') and raw.endswith('"'):
            raw = raw[1:-1].encode("utf-8").decode("unicode_escape")
            return json.loads(raw)
        raise

# ---------- Utils ----------
def iso_to_dt(iso_str: str) -> datetime:
    s = (iso_str or "").replace("Z", "")
    for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M"):
        try:
            return datetime.strptime(s[:len(fmt)], fmt)
        except Exception:
            pass
    try:
        return datetime.fromisoformat(s)
    except Exception:
        return datetime.strptime(s.split("T")[0], "%Y-%m-%d")

def fmt_fecha(dt: datetime) -> str:
    return f"{dt.day}/{dt.month}/{dt.year} {dt.hour:02d}:{dt.minute:02d}"

def fmt_jornada(dt: datetime) -> str:
    return f"{dt.day}/{dt.month}/{dt.year}"

def pick_first_key(d: dict, *candidates: str):
    for k in candidates:
        if k in d and d[k] not in (None, ""):
            return d[k]
    norm = {k.replace(" ", "").lower(): v for k, v in d.items()}
    for k in candidates:
        nk = k.replace(" ", "").lower()
        if nk in norm and norm[nk] not in (None, ""):
            return norm[nk]
    return None

def to_float(x) -> Optional[float]:
    if x in (None, "", "NaN"): return None
    try: return float(x)
    except Exception:
        try: return float(str(x).replace(",", "."))
        except Exception: return None

# ---------- Rango de fechas ----------
def daterange(start: date, end: date):
    cur = start
    while cur <= end:
        yield cur
        cur += timedelta(days=1)

def month_bounds(d: date):
    start = date(d.year, d.month, 1)
    end = date(d.year + (d.month==12), 1 if d.month==12 else d.month+1, 1) - timedelta(days=1)
    return start, end

# ---------- Endpoints ----------
def get_tiendas(dref: date) -> Dict[int, Dict[str, str]]:
    payload = {"TouchExpress_IF": {"Tienda": 1, "Fecha": dref.isoformat()}}
    resp = http_post_json(f"{BASE}/MPTiendas", payload)
    tiendas = resp.get("TouchExpress_IF", {}).get("Tiendas", [])
    out = {}
    for t in tiendas:
        try:
            code = int(t.get("codigo"))
        except Exception:
            continue
        out[code] = {
            "nombre": t.get("nombre") or "",
            "grupo": t.get("grupo") or "",
            "social": t.get("social") or "",
            "nif": t.get("nif") or "",
        }
    return out

def get_ventas_dia(tienda: int, d: date) -> List[dict]:
    payload = {"TouchExpress_IF": {"Tienda": tienda, "Fecha": d.isoformat()}}
    resp = http_post_json(f"{BASE}/MPVentasMesa", payload)
    te = resp.get("TouchExpress_IF", {}) if isinstance(resp, dict) else {}
    docs = te.get("Documentos") or []
    return docs if isinstance(docs, list) else []

def get_compras_dia(tienda: int, d: date) -> List[dict]:
    payload = {"TouchExpress_IF": {"Tienda": tienda, "Fecha": d.isoformat()}}
    resp = http_post_json(f"{BASE}/MPCompras", payload)
    te = resp.get("TouchExpress_IF", {}) if isinstance(resp, dict) else {}
    docs = te.get("Documentos") or []
    return docs if isinstance(docs, list) else []

# ---------- Cost index (del 1 del mes actual hasta el día objetivo) ----------
def build_cost_index(tienda_ids: List[int], target_day: date):
    """
    Construye índice de costes desde el día 1 del mes ACTUAL hasta target_day
    """
    hoy = date.today()
    start = date(hoy.year, hoy.month, 1)  # 1º del mes ACTUAL
    
    idx: Dict[int, Dict[str, Tuple[datetime, float]]] = {tid: {} for tid in tienda_ids}
    
    print(f"Construyendo índice de costes desde {start} hasta {target_day}")
    
    for d in daterange(start, target_day):  # 1º del mes actual .. día objetivo
        for tienda in tienda_ids:
            try:
                docs = get_compras_dia(tienda, d)
                print(f"  Compras {d} tienda {tienda}: {len(docs)} documentos")
            except Exception as e:
                print(f"  Error compras {d} tienda {tienda}: {e}")
                continue
            
            for doc in docs:
                fecha_iso = pick_first_key(doc, "fecha", "Fecha", "FechaReg")
                if not fecha_iso: 
                    continue
                dt = iso_to_dt(fecha_iso)
                for p in (doc.get("productos") or []):
                    ref = p.get("referencia")
                    if ref in (None, ""): 
                        continue
                    ref_str = str(ref)
                    cant = to_float(pick_first_key(p, "can tad", "cantidad"))
                    imp  = to_float(p.get("importe"))
                    if not cant or cant == 0 or imp is None:
                        continue
                    unit = imp / cant
                    prev = idx[tienda].get(ref_str)
                    if (prev is None) or (dt >= prev[0]):
                        idx[tienda][ref_str] = (dt, unit)
    
    return idx

# ---------- Mapper ----------
def make_rows_from_doc(doc, tienda_id, tienda_info, cost_index_for_tienda):
    filas: List[dict] = []
    fecha_iso = pick_first_key(doc, "fecha", "Fecha", "FechaReg")
    if not fecha_iso:
        return filas
    dt = iso_to_dt(fecha_iso)

    serie = pick_first_key(doc, "serie", "Serie")
    numtiket = pick_first_key(doc, "num ket", "num tket", "numtiket", "num")
    seccion = doc.get("seccion") or {}
    servicio = doc.get("servicio") or {}
    cliente  = doc.get("cliente")  or {}
    totales  = doc.get("totales")  or {}

    cab_total = to_float(totales.get("total"))
    cab_base  = to_float(totales.get("baseImponible"))

    for p in (doc.get("productos") or []):
        ref = p.get("referencia")
        ref_str = str(ref) if ref not in (None, "") else ""
        desc = p.get("descripcion") or ""
        grupo = p.get("grupo") or ""

        cantidad = to_float(pick_first_key(p, "can tad", "cantidad"))
        precio   = to_float(p.get("precio"))
        iva      = to_float(p.get("iva"))
        descuento= to_float(p.get("descuento"))
        importe  = to_float(p.get("importe"))

        base = round(importe / (1.0 + iva/100.0), 6) if (importe is not None and iva is not None) else ""
        imp_desc = round(importe - descuento, 6) if (importe is not None and descuento is not None) else ""
        base_desc = round(imp_desc / (1.0 + iva/100.0), 6) if (imp_desc != "" and iva is not None) else ""

        coste = ""
        if ref_str and ref_str in cost_index_for_tienda:
            coste = round(cost_index_for_tienda[ref_str][1], 6)

        fila = {
            "IDTRANS":"", "NSERIE":"", "SERIE":serie or "", "NUMTIKET":numtiket or "",
            "NUMBARRA":seccion.get("codigo") or "", "NNUMBARRA":seccion.get("nombre") or "",
            "FECHA":fmt_fecha(dt), "JORNADA":fmt_jornada(dt),
            "CREDITO":"","NCREDITO":"", "NUMCLIE":cliente.get("codigo") or "",
            "PUNTOVENTA":"","NPUNTOVENTA":"", "NUMCUEN":"","NNUMCUEN":"",
            "SERVICIO":servicio.get("codigo") or "", "NSERVICIO":servicio.get("nombre") or "",
            "ALMACEN":"","NALMACEN":"", "CABIMPORTE":cab_total if cab_total is not None else "",
            "CABDESCUENTO":"", "CABNETO":cab_base if cab_base is not None else "",
            "CAMARERO":"","NCAMARERO":"","MACROGRUPO":"","NMACROGRUPO":"",
            "GRUPO":"","NGRUPO":grupo or "", "FAMILIA":"","NFAMILIA":"",
            "TIPOPRODUCTO":"","NTIPOPRODUCTO":"", "PRODUCTO":ref_str, "NPRODUCTO":desc,
            "CANTIDAD":cantidad if cantidad is not None else "", "PRECIO":precio if precio is not None else "",
            "IVA":iva if iva is not None else "", "IMPORTE":importe if importe is not None else "",
            "IMPORTESINIVA":base, "DESCUENTO":descuento if descuento is not None else "",
            "IMPORTEDESCUENTO":imp_desc, "IMPORTESINIVADESCUENTO":base_desc,
            "ANULADA":"","FORMATO":"","NFORMATO":"",
            "TIENDA":tienda_id, "ESTABLECIMIENTO":tienda_info.get("nombre") or "",
            "CTACONTABLE":"", "CECO":"", "NIFCLIENTE":cliente.get("nif") or "",
            "NOMBRECLIENTE":cliente.get("nombre") or "", "VENCIMIENTO":"","PROMOCION":"",
            "COMENSALES":"", "COSTE":coste, "OBSERVACIONES":"",
            "Turno":"", "Denominacion 2":"", "Factura":"", "Motivo":""
        }
        # asegurar todas las columnas
        for col in CSV_COLUMNS:
            fila.setdefault(col, "")
        filas.append(fila)
    return filas

# ---------- MAIN CORREGIDO ----------
def main():
    target_day = get_target_date()
    
    # 🔑 CAMBIO CRÍTICO: Definir el rango de fechas para ventas
    hoy = date.today()
    start_date = date(hoy.year, hoy.month, 1)  # 1º del mes actual
    end_date = target_day  # día objetivo
    
    print(f"📅 Obteniendo ventas desde {start_date} hasta {end_date}")
    print(f"📅 Día objetivo para el archivo: {target_day.isoformat()}")

    # Tiendas
    tiendas = get_tiendas(target_day)
    if not tiendas:
        raise SystemExit("No se han podido obtener tiendas.")
    tienda_ids = sorted(tiendas.keys()) if TIENDAS is None else TIENDAS
    print("Tiendas:", tienda_ids)

    # Índice de costes desde el 1 del mes hasta el día objetivo
    print("Construyendo índice de costes (MPCompras)…")
    cost_index = build_cost_index(tienda_ids, target_day)

    # 🔑 CAMBIO PRINCIPAL: Ventas de TODO EL RANGO, no solo un día
    rows: List[dict] = []
    
    for current_date in daterange(start_date, end_date):
        print(f"\n📊 Procesando ventas del {current_date}")
        
        for tid in tienda_ids:
            try:
                docs = get_ventas_dia(tid, current_date)
                print(f"  Tienda {tid}: {len(docs)} documentos")
                
                for doc in docs:
                    rows.extend(make_rows_from_doc(doc, tid, tiendas.get(tid, {}), cost_index.get(tid, {})))
                    
            except error.HTTPError as e:
                print(f"  [{current_date}] HTTPError tienda {tid}: {e.code}")
                continue
            except Exception as e:
                print(f"  [{current_date}] Error tienda {tid}: {e}")
                continue

    print(f"\n📈 TOTAL de filas generadas: {len(rows)}")
    
    # Estadísticas por día
    if rows:
        from collections import Counter
        fechas_count = Counter(row['JORNADA'] for row in rows)
        print("\n📊 Distribución por fecha:")
        for fecha, count in sorted(fechas_count.items()):
            print(f"  {fecha}: {count} líneas")
    
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    out_csv = os.path.join(OUTPUT_DIR, f"ventas_{target_day.isoformat()}.csv")

    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_COLUMNS, extrasaction="ignore")
        writer.writeheader()
        if rows:
            writer.writerows(rows)

    print(f"\n✅ CSV generado: {out_csv}")
    print(f"📁 Contiene ventas desde {start_date} hasta {end_date}")

if __name__ == "__main__":
    main()