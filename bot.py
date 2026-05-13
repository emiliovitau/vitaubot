import schedule
import time
import json
import requests
import glob
import os
import pandas as pd
from datetime import datetime, date

SLACK_TOKEN = "xoxb-624701405537-10984934095984-otO0CIVU9npB5kt7WJaBJBSu"
CANAL_MARKETPLACES = "C040H7TV6JG"
CANAL_HELP_MKT = "C05KVEGS2TG"
CANAL_ALERTAS = "C05UH4BFCC8"
METABASE_URL = "https://analytics.vitau.mx"
METABASE_USER = "fatima@vitau.mx"
METABASE_PASS = "9MC69C2L7iFHndu"
METABASE_CARD = 2203
BASE_DIR = r"C:\Users\emigt\OneDrive\Documents\VIATU BOT"
ESTADO_FILE = os.path.join(BASE_DIR, "snapshot.json")
CARPETA_ML  = os.path.join(BASE_DIR, "archivos")
MENCIONES_MTY = "<@U08341SD2A2> <@U06GNMKAL2Z> <@U0467HCL5AP>"
MENCIONES_CDMX = "<@U08QT8JRRG9> <@U0467HCL5AP>"
MENCIONES_DEMSA = "<@U055FRNF1CP> <@U0467HCL5AP>"
MENCIONES_CS = "<@U02MB4L0NDC> <@U0467HCL5AP> <@U0404PD44AZ>"
MENCIONES_ALERTAS = "<@U0404PD44AZ>"
STATUS_ES = {"in_process":"🟡 En proceso","approved":"🟠 Aprobado","bought":"🟣 Comprado","ready":"🟢 Listo","delivered":"🔵 Entregado","shipped":"🚚 Enviado"}
COLECTAS = {"Sucursal Monterrey":{"ventana":"15:00-17:00","limite":"10:00 hs (lunes 08:00)"},"Sucursal Cd. Mx.":{"ventana":"14:35-16:35","limite":"07:35 hs"},"DEMSA":{"ventana":"09:15-11:15","limite":"02:15 hs"}}
CEDIS_CONFIG = [("Sucursal Monterrey",MENCIONES_MTY),("Sucursal Cd. Mx.",MENCIONES_CDMX),("DEMSA",MENCIONES_DEMSA)]
DEPOSITO_A_CEDIS = {"Monterrey Avenida General Pabl":"Sucursal Monterrey","Iztapalapa Espana":"Sucursal Cd. Mx.","Iztapalapa España":"Sucursal Cd. Mx.","DEMSA GDL":"DEMSA"}

def slack_post(canal,texto):
    r = requests.post("https://slack.com/api/chat.postMessage",headers={"Authorization":f"Bearer {SLACK_TOKEN}"},json={"channel":canal,"text":texto,"mrkdwn":True})
    print(f"Slack {canal}: {r.json().get('ok')}")

# FIX: ML exporta filas de metadata antes del header real — detectarlo dinámicamente
def _find_ml_header_row(path, sheet="Ventas MX", max_scan=15):
    raw = pd.read_excel(path, sheet_name=sheet, header=None, nrows=max_scan)
    for i, row in raw.iterrows():
        if any(str(v).strip() == "# de venta" for v in row):
            return i
    raise ValueError(f"No se encontró el header de ML en las primeras {max_scan} filas de {path}")

def get_ml_orders():
    os.makedirs(CARPETA_ML, exist_ok=True)
    archivos = glob.glob(os.path.join(CARPETA_ML, "*.xlsx"))
    if not archivos:
        print("No hay archivo ML")
        return []
    archivo = max(archivos, key=os.path.getmtime)
    print(f"Leyendo ML: {os.path.basename(archivo)}")
    try:
        df = pd.read_excel(archivo, sheet_name="Ventas MX", header=_find_ml_header_row(archivo))
        df["# de venta"] = df["# de venta"].astype(str).str.strip()
        df["SKU"] = df["SKU"].astype(str).str.strip()
        dep_col = "Depósito" if "Depósito" in df.columns else "Deposito"
        df[dep_col] = df[dep_col].astype(str).str.strip()
        df["Forma de entrega"] = df["Forma de entrega"].astype(str).str.strip()
        df = df[df["Forma de entrega"] != "Mercado Envíos Full"]
        df = df[df["Estado"] == "Etiqueta lista para imprimir"]
        df["cedis_ml"] = df[dep_col].map(DEPOSITO_A_CEDIS).fillna("Desconocido")
        df["fecha_ml"] = pd.to_datetime(df["Fecha de venta"], dayfirst=True, errors="coerce")
        orders = df[["# de venta","SKU","cedis_ml","fecha_ml"]].rename(columns={"# de venta":"id","SKU":"sku"}).to_dict("records")
        print(f"ML: {len(orders)} ordenes etiqueta lista (sin FULL)")
        return orders
    except Exception as e:
        print(f"Error leyendo ML: {e}")
        return []

def get_vitau_orders():
    r = requests.post(f"{METABASE_URL}/api/session",json={"username":METABASE_USER,"password":METABASE_PASS})
    token = r.json().get("id")
    h = {"X-Metabase-Session":token}
    r2 = requests.post(f"{METABASE_URL}/api/card/{METABASE_CARD}/query",headers=h,json={})
    data = r2.json().get("data",{})
    cols = [c["name"] for c in data.get("cols",[])]
    return [dict(zip(cols,row)) for row in data.get("rows",[])]

def cruzar(ml_orders, vitau_orders):
    vitau_by_id = {str(o.get("ID externo","")):o for o in vitau_orders}
    vitau_by_ean = {}
    for o in vitau_orders:
        ean = str(o.get("ean_key",""))
        if ean not in vitau_by_ean:
            vitau_by_ean[ean] = []
        vitau_by_ean[ean].append(o)

    result = []
    discrepancias = []

    for ml in ml_orders:
        v = vitau_by_id.get(ml["id"])
        match_method = "ID"

        if not v:
            cands = vitau_by_ean.get(ml.get("sku",""),[])
            if cands:
                ml_dt = ml.get("fecha_ml")
                mejor = None
                mejor_diff = 999999
                for c in cands:
                    try:
                        v_dt = pd.to_datetime(c.get("fecha de creación") or c.get("fecha de creacion",""))
                        if hasattr(v_dt, 'tz_localize'):
                            v_dt = v_dt.tz_localize(None) if v_dt.tzinfo else v_dt
                        if ml_dt and not pd.isna(ml_dt):
                            ml_dt_naive = ml_dt.tz_localize(None) if hasattr(ml_dt,'tzinfo') and ml_dt.tzinfo else ml_dt
                            diff = abs((v_dt - ml_dt_naive).total_seconds()/3600)
                            if diff <= 24 and diff < mejor_diff:
                                mejor_diff = diff
                                mejor = c
                    except: pass
                if mejor:
                    v = mejor
                    match_method = "SKU+fecha"
                elif cands:
                    v = cands[0]
                    match_method = "SKU"

        if v:
            cedis_vitau = v.get("CEDIS","")
            cedis_ml = ml.get("cedis_ml","")
            if cedis_ml in DEPOSITO_A_CEDIS.values() and cedis_vitau != cedis_ml:
                discrepancias.append({
                    "id": ml["id"],
                    "orden": v.get("orden"),
                    "producto": v.get("producto",""),
                    "cedis_ml": cedis_ml,
                    "cedis_vitau": cedis_vitau,
                    "match": match_method
                })
            result.append({**ml,"orden":v.get("orden"),"producto":v.get("producto"),"status":v.get("status"),"CEDIS":cedis_vitau,"match":match_method})
        else:
            result.append({**ml,"orden":None,"producto":"Sin match","status":"sin_match","CEDIS":ml.get("cedis_ml"),"match":"none"})

    return result, discrepancias

def build_msg_despacho(ordenes, cedis, menciones, es_seguimiento=False):
    grp = [o for o in ordenes if o.get("CEDIS")==cedis and o.get("status") not in ["cancelled","delivered","sin_match"]]
    if not grp: return None
    nombre = cedis.replace("Sucursal ","")
    emoji = "🔁" if es_seguimiento else "🚨"
    titulo = f"SIGUEN PENDIENTES - {nombre}" if es_seguimiento else f"DESPACHO HOY - {nombre}"
    intro = f"*{len(grp)} ordenes siguen pendientes desde las 9 AM:*" if es_seguimiento else f"Las siguientes *{len(grp)} ordenes* deben salir hoy:"
    lineas = []
    for o in grp:
        st = STATUS_ES.get(o.get("status",""),o.get("status",""))
        orden_str = f"Orden *{o['orden']}*" if o.get("orden") else f"`{o['id']}`"
        lineas.append(f"🏷️ {orden_str} - {o.get('producto','')} | {st}")
    col = COLECTAS.get(cedis,{})
    hora = datetime.now().strftime("%d/%m/%Y %H:%M")
    col_txt = f"\n\n🚚 *Colecta:* {col.get('ventana','')} | 📦 *Limite:* {col.get('limite','')}" if col else ""
    return f"{emoji} *{titulo}* {emoji}\n{menciones}\n\n{intro}\n\n"+"\n".join(lineas)+col_txt+f"\n\n_Corte: {hora}_"

def build_msg_discrepancias(discrepancias, menciones):
    if not discrepancias: return None
    hora = datetime.now().strftime("%d/%m/%Y %H:%M")
    lineas = []
    for d in discrepancias:
        orden_str = f"Orden *{d['orden']}*" if d.get("orden") else f"`{d['id']}`"
        lineas.append(f"⚠️ {orden_str} — {d.get('producto','')}\n   ML espera: *{d['cedis_ml']}* | Vitau asignó: *{d['cedis_vitau']}*")
    return f"🔴 *DISCREPANCIA DE CEDIS — Mercado Libre* 🔴\n{menciones}\n\n*{len(discrepancias)} ordenes* con CEDIS incorrecto:\n\n"+"\n".join(lineas)+f"\n\n_Verificar y reasignar si es necesario._\n_Corte: {hora}_"

def build_msg_recetas(recetas, menciones):
    if not recetas: return None
    hora = datetime.now().strftime("%d/%m/%Y %H:%M")
    lineas = [f"• `{r['id']}` | SKU: {r['sku']}" for r in recetas]
    return f"💊 *RECETAS PENDIENTES DE REVISIÓN*\n{menciones}\n\nHay *{len(recetas)} ordenes* con receta pendiente:\n\n"+"\n".join(lineas)+f"\n\n_Corte: {hora}_"

def run_corte(es_seguimiento=False):
    tipo = "1PM" if es_seguimiento else "9AM"
    print(f"\nCORTE {tipo} - {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    try:
        vitau_orders = get_vitau_orders()
        print(f"Vitau: {len(vitau_orders)} ordenes")
        ml_orders = get_ml_orders()
        if not ml_orders:
            slack_post(CANAL_ALERTAS, f"⚠️ VitauBot {tipo}: No hay archivo ML en {CARPETA_ML}")
            return
        ordenes, discrepancias = cruzar(ml_orders, vitau_orders)
        sin_match = [o for o in ordenes if o.get("match")=="none"]
        print(f"Match: {len(ordenes)-len(sin_match)}/{len(ordenes)} | Sin match: {len(sin_match)} | Discrepancias: {len(discrepancias)}")
    except Exception as e:
        print(f"Error: {e}")
        slack_post(CANAL_ALERTAS,f"⚠️ Error VitauBot {tipo}: {e}")
        return

    if es_seguimiento:
        try:
            snap = json.load(open(ESTADO_FILE))
            if snap.get("fecha")==str(date.today()):
                ids = set(snap.get("ids",[]))
                ordenes = [o for o in ordenes if str(o.get("id","")) in ids]
                print(f"Filtradas a {len(ordenes)} pendientes desde 9AM")
        except: pass

    for cedis,menciones in CEDIS_CONFIG:
        msg = build_msg_despacho(ordenes,cedis,menciones,es_seguimiento)
        if msg:
            slack_post(CANAL_MARKETPLACES,msg)

    if discrepancias and not es_seguimiento:
        msg_disc = build_msg_discrepancias(discrepancias, MENCIONES_ALERTAS)
        if msg_disc:
            slack_post(CANAL_ALERTAS, msg_disc)
            print(f"Discrepancias enviadas: {len(discrepancias)}")

    if not es_seguimiento:
        try:
            archivo = max(glob.glob(os.path.join(CARPETA_ML, "*.xlsx")), key=os.path.getmtime)
            df_rec = pd.read_excel(archivo, sheet_name="Ventas MX", header=_find_ml_header_row(archivo))
            df_rec["# de venta"] = df_rec["# de venta"].astype(str).str.strip()
            df_rec["SKU"] = df_rec["SKU"].astype(str).str.strip()
            recetas = df_rec[df_rec["Estado"]=="Receta pendiente de revisión"][["# de venta","SKU"]].rename(columns={"# de venta":"id","SKU":"sku"}).to_dict("records")
            if recetas:
                slack_post(CANAL_HELP_MKT, build_msg_recetas(recetas, MENCIONES_CS))
                print(f"Recetas: {len(recetas)}")
        except: pass
        ids = [str(o.get("id","")) for o in ordenes if o.get("status") not in ["cancelled","delivered"]]
        json.dump({"fecha":str(date.today()),"ids":ids},open(ESTADO_FILE,"w"))
        print(f"Snapshot: {len(ids)} ordenes")

for dia in ["monday","tuesday","wednesday","thursday","friday"]:
    getattr(schedule.every(),dia).at("09:00").do(run_corte,es_seguimiento=False)
    getattr(schedule.every(),dia).at("13:00").do(run_corte,es_seguimiento=True)

print("="*55)
print("  VitauBot activo - 9AM y 1PM Lun-Vie")
print(f"  Deposita el Excel de ML en:")
print(f"  {CARPETA_ML}")
print("="*55)
while True:
    schedule.run_pending()
    time.sleep(30)
