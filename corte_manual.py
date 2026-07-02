import json, requests, glob, os, pandas as pd
from datetime import datetime, date

SLACK_TOKEN = "xoxb-624701405537-10984934095984-otO0CIVU9npB5kt7WJaBJBSu"
CANAL = "C040H7TV6JG"
CANAL_ALERTAS = "C05UH4BFCC8"
METABASE_URL = "https://analytics.vitau.mx"
METABASE_USER = "fatima@vitau.mx"
METABASE_PASS = "9MC69C2L7iFHndu"
METABASE_CARD = 2203

# Rutas relativas — funciona tanto en GitHub Actions (Linux) como en Windows
BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
CARPETA_ML   = os.path.join(BASE_DIR, "archivos")
SNAPSHOT_FILE = os.path.join(BASE_DIR, "snapshot.json")

MENCIONES_MTY = "<@U08341SD2A2> <@U06GNMKAL2Z> <@U0467HCL5AP>"
MENCIONES_CDMX = "<@U08QT8JRRG9> <@U0467HCL5AP>"
MENCIONES_DEMSA = "<@U055FRNF1CP> <@U0467HCL5AP>"
MENCIONES_ALERTAS = "<@U0404PD44AZ>"
STATUS_ES = {"in_process":"En proceso","approved":"Aprobado","bought":"Comprado","ready":"Listo","delivered":"Entregado"}
COLECTAS = {"Sucursal Monterrey":{"ventana":"15:00-17:00","limite":"10:00 hs"},"Sucursal Cd. Mx.":{"ventana":"14:35-16:35","limite":"07:35 hs"},"DEMSA":{"ventana":"09:15-11:15","limite":"02:15 hs"}}
DEPOSITO_A_CEDIS = {"Monterrey Avenida General Pabl":"Sucursal Monterrey","Iztapalapa Espana":"Sucursal Cd. Mx.","Iztapalapa España":"Sucursal Cd. Mx.","DEMSA GDL":"DEMSA"}
CEDIS_CONFIG = [("Sucursal Monterrey",MENCIONES_MTY),("Sucursal Cd. Mx.",MENCIONES_CDMX),("DEMSA",MENCIONES_DEMSA)]

def slack_post(canal, texto):
    r = requests.post("https://slack.com/api/chat.postMessage",
                      headers={"Authorization": f"Bearer {SLACK_TOKEN}"},
                      json={"channel": canal, "text": texto, "mrkdwn": True})
    print(f"Slack {canal}: {r.json().get('ok')}")

# Metabase
r = requests.post(f"{METABASE_URL}/api/session", json={"username": METABASE_USER, "password": METABASE_PASS})
token = r.json().get("id")
h = {"X-Metabase-Session": token}
r2 = requests.post(f"{METABASE_URL}/api/card/{METABASE_CARD}/query", headers=h, json={})
data = r2.json().get("data", {})
cols = [c["name"] for c in data.get("cols", [])]
vitau = [dict(zip(cols, row)) for row in data.get("rows", [])]
vitau_by_id = {str(o.get("ID externo", "")): o for o in vitau}
vitau_by_ean = {}
for o in vitau:
    ean = str(o.get("ean_key", ""))
    if ean not in vitau_by_ean: vitau_by_ean[ean] = []
    vitau_by_ean[ean].append(o)
print(f"Vitau: {len(vitau)} ordenes")

# Excel ML
archivos = glob.glob(os.path.join(CARPETA_ML, "*.xlsx"))
if not archivos:
    print(f"ERROR: No hay archivos .xlsx en {CARPETA_ML}")
    slack_post(CANAL_ALERTAS, f"⚠️ VitauBot: No hay archivo ML en la carpeta de Drive")
    exit(1)
archivo = max(archivos, key=os.path.getmtime)
print(f"Leyendo ML: {os.path.basename(archivo)}")

# FIX: ML exporta filas de metadata antes del header real
def _find_ml_header_row(path, sheet="Ventas MX", max_scan=15):
    raw = pd.read_excel(path, sheet_name=sheet, header=None, nrows=max_scan)
    for i, row in raw.iterrows():
        if any(str(v).strip() == "# de venta" for v in row):
            return i
    raise ValueError(f"No se encontró el header de ML en las primeras {max_scan} filas de {path}")

df = pd.read_excel(archivo, sheet_name="Ventas MX", header=_find_ml_header_row(archivo))
print("Columnas:", list(df.columns[:8]))
id_col  = [c for c in df.columns if "venta"   in str(c).lower()][0]
sku_col = [c for c in df.columns if "sku"     in str(c).lower() or "SKU" in str(c)][0]
dep_col = [c for c in df.columns if "dep"     in str(c).lower()][0]
ent_col = [c for c in df.columns if "entrega" in str(c).lower()][0]
est_col = [c for c in df.columns if "estado"  in str(c).lower()][0]
df[id_col]  = df[id_col].astype(str).str.strip()
df[sku_col] = df[sku_col].astype(str).str.strip()
df[dep_col] = df[dep_col].astype(str).str.strip()
df = df[df[ent_col].astype(str).str.strip() != "Mercado Envíos Full"]
df = df[df[est_col].astype(str).str.strip() == "Etiqueta lista para imprimir"]
# Fix: excluir órdenes cuya colecta es mañana — solo queremos las de hoy
desc_col = [c for c in df.columns if "descripci" in str(c).lower()][0]
df = df[~df[desc_col].astype(str).str.contains("mañana", na=False)]
df["cedis_ml"] = df[dep_col].map(DEPOSITO_A_CEDIS).fillna("Desconocido")
print(f"ML: {len(df)} ordenes etiqueta lista (solo hoy)")

hora = datetime.now().strftime("%d/%m/%Y %H:%M")
ids_snapshot = []
discrepancias = []

for cedis, menciones in CEDIS_CONFIG:
    grp = df[df["cedis_ml"] == cedis]
    if grp.empty: continue
    lineas = []
    for _, row in grp.iterrows():
        oid = row[id_col]
        v = vitau_by_id.get(oid, {})
        if not v:
            cands = vitau_by_ean.get(row[sku_col], [])
            if cands: v = cands.pop(0)  # Fix: consumir el candidato para que el siguiente no reutilice el mismo orden
        st = STATUS_ES.get(v.get("status", ""), "Sin match")
        orden_num = v.get("orden")
        orden_str = f"Orden *{orden_num}*" if orden_num else f"`{oid}`"
        prod = v.get("producto", row[sku_col])
        lineas.append(f"🏷️ {orden_str} - {prod} | {st}")
        ids_snapshot.append(oid)
        if v.get("CEDIS") and v.get("CEDIS") != cedis:
            discrepancias.append({"id": oid, "orden": orden_num, "producto": prod, "cedis_ml": cedis, "cedis_vitau": v.get("CEDIS")})
    col = COLECTAS.get(cedis, {})
    msg = (f"🚨 *DESPACHO HOY - {cedis.replace('Sucursal ', '')}* 🚨\n{menciones}\n\n"
           f"Las siguientes *{len(lineas)} ordenes* deben salir hoy:\n\n" +
           "\n".join(lineas) +
           f"\n\n🚚 Colecta: {col.get('ventana', '')} | Limite: {col.get('limite', '')}\n\n_Corte: {hora}_")
    slack_post(CANAL, msg)

if discrepancias:
    lineas_d = [f"⚠️ Orden *{d.get('orden', d['id'])}* — {d['producto']}\n   ML: *{d['cedis_ml']}* | Vitau: *{d['cedis_vitau']}*"
                for d in discrepancias]
    slack_post(CANAL_ALERTAS, f"🔴 *DISCREPANCIA DE CEDIS*\n{MENCIONES_ALERTAS}\n\n" + "\n".join(lineas_d) + f"\n\n_Corte: {hora}_")

json.dump({"fecha": str(date.today()), "ids": ids_snapshot}, open(SNAPSHOT_FILE, "w"))
print(f"Snapshot: {len(ids_snapshot)} ordenes")
