#!/usr/bin/env python3
"""
vitau_buybox_diario.py — Monitor diario del Buy Box de Mercado Libre (Vitau).

Jala la card 4552 (rendimiento-en-ml) DIRECTO de Metabase (igual que VitauBot
con la 2203), calcula las acciones de precio del dia, compara contra el snapshot
de ayer y publica el digest en Slack. Sin archivos, sin Drive.

Uso en GitHub Actions:
  python vitau_buybox_diario.py
Modo prueba con un export local (sin tocar Metabase):
  python vitau_buybox_diario.py --export rendimiento_en_ml_XXXX.xlsx

Variables de entorno (todas como GitHub Secrets, NUNCA en el codigo):
  METABASE_URL    ej. https://analytics.vitau.mx
  METABASE_USER   usuario de Metabase (idealmente cuenta de servicio, no la de Fati)
  METABASE_PASS   contrasena de Metabase
  METABASE_CARD   id de la card (default 4552)
  SLACK_TOKEN     token del bot de Slack (xoxb-...)
  SLACK_CHANNEL   id del canal (ej. C040H7TV6JG)
  SNAPSHOT_FILE   ruta del json de tracking (default snapshot_buybox.json)
"""

import argparse
import datetime as dt
import json
import os
import sys
from decimal import Decimal, ROUND_HALF_UP

import pandas as pd
import requests

GANANDO = "Ganando buy box"
PUEDE_BAJAR = "Puede bajar sin perder margen"

COL_ID = "ID listing ML"
COL_PROD = "Producto"
COL_ESTADO = "Estado buy box"
COL_NUESTRO = "Nuestro precio"
COL_GANADOR = "Precio del ganador"
COL_OBJETIVO = "Precio sugerido para ganar"
COL_BAJAR = "Cuánto debemos bajar ($)"
COL_PISO = "Precio mínimo (con margen)"
COL_PUEDE = "¿Puede ganar buybox?"
REQUERIDAS = [COL_ID, COL_ESTADO, COL_PUEDE, COL_OBJETIVO, COL_BAJAR, COL_PISO]


def money(x):
    if pd.isna(x):
        return None
    return float(Decimal(str(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def cargar_metabase():
    url = os.environ["METABASE_URL"]
    card = os.environ.get("METABASE_CARD", "4552")
    s = requests.post(f"{url}/api/session",
                      json={"username": os.environ["METABASE_USER"],
                            "password": os.environ["METABASE_PASS"]}, timeout=30)
    token = s.json().get("id")
    if not token:
        sys.exit(f"ERROR: no se pudo autenticar en Metabase: {s.text[:200]}")
    q = requests.post(f"{url}/api/card/{card}/query",
                      headers={"X-Metabase-Session": token}, json={}, timeout=120)
    data = q.json().get("data", {})
    cols = [c["name"] for c in data.get("cols", [])]
    rows = data.get("rows", [])
    print(f"Metabase card {card}: {len(rows)} filas")
    return pd.DataFrame(rows, columns=cols)


def cargar_export(path):
    return pd.read_excel(path, dtype={COL_ID: str})


def validar(df):
    faltan = [c for c in REQUERIDAS if c not in df.columns]
    if faltan:
        sys.exit(f"ERROR: faltan columnas {faltan}.\nColumnas recibidas: {list(df.columns)}")
    df[COL_ID] = df[COL_ID].astype(str).str.strip()
    for c in (COL_NUESTRO, COL_GANADOR, COL_OBJETIVO, COL_BAJAR, COL_PISO):
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def calcular_acciones(df):
    no_gana = df[COL_ESTADO] != GANANDO
    puede = df[COL_PUEDE] == PUEDE_BAJAR
    mueve = df[COL_BAJAR].fillna(0) > 0
    seguro = df[COL_OBJETIVO].fillna(-1) >= df[COL_PISO].fillna(1e12)
    acc = df[no_gana & puede & mueve & seguro].sort_values(COL_BAJAR, ascending=False)
    escalar = df[no_gana & puede & mueve & ~seguro]
    return acc, escalar


def tracking(fecha, df):
    path = os.environ.get("SNAPSHOT_FILE", "snapshot_buybox.json")
    ganan_hoy = sorted(df.loc[df[COL_ESTADO] == GANANDO, COL_ID].astype(str))
    perdidos = ganados = None
    if os.path.exists(path):
        prev = json.load(open(path))
        if prev.get("fecha") != fecha:
            ayer = set(prev.get("ganando", []))
            hoy = set(ganan_hoy)
            perdidos, ganados = ayer - hoy, hoy - ayer
    json.dump({"fecha": fecha, "ganando": ganan_hoy}, open(path, "w"))
    return perdidos, ganados, len(ganan_hoy)


def construir_mensaje(fecha, n, acc, escalar, perdidos, ganados, total_gana):
    L = [f"*Buy Box ML — {fecha}*", f"Publicaciones: {n:,} · Ganando: {total_gana}"]
    if perdidos is not None:
        if perdidos:
            L.append(f":small_red_triangle_down: Perdimos el Buy Box en {len(perdidos)} (ayer ganábamos)")
        if ganados:
            L.append(f":white_check_mark: Recuperamos {len(ganados)}")
        if not perdidos and not ganados:
            L.append("Sin cambios de Buy Box vs ayer")
    L.append("")
    if acc.empty:
        L.append(":tada: *Sin acciones de precio hoy.* Todo lo ganable ya está al precio correcto.")
    else:
        L.append(f"*Acciones de precio hoy: {len(acc)}*  (bajar sin perder margen)")
        for _, r in acc.iterrows():
            L.append(f"• {str(r[COL_PROD])[:42]} — ${money(r[COL_NUESTRO]):,.2f} → "
                     f"*${money(r[COL_OBJETIVO]):,.2f}* (piso ${money(r[COL_PISO]):,.2f})")
    if not escalar.empty:
        L.append(f"\n:warning: *{len(escalar)} para escalar a Fati* — ganar implica bajar del margen mínimo.")
    return "\n".join(L)


def enviar_slack(texto):
    token, canal = os.environ.get("SLACK_TOKEN"), os.environ.get("SLACK_CHANNEL")
    if not token or not canal:
        print("[dry-run] SLACK_TOKEN/SLACK_CHANNEL no definidos. Mensaje:\n")
        print(texto)
        return
    r = requests.post("https://slack.com/api/chat.postMessage",
                      headers={"Authorization": f"Bearer {token}"},
                      json={"channel": canal, "text": texto, "mrkdwn": True}, timeout=20)
    print(f"Slack: ok={r.json().get('ok')}")


def escribir_excel(salida, acc, escalar):
    cols = [c for c in (COL_ID, COL_PROD, COL_NUESTRO, COL_GANADOR, COL_OBJETIVO,
                        COL_BAJAR, COL_PISO, COL_ESTADO) if c in acc.columns]
    with pd.ExcelWriter(salida, engine="openpyxl") as xw:
        acc[cols].to_excel(xw, sheet_name="Acciones precio", index=False)
        if not escalar.empty:
            escalar[cols].to_excel(xw, sheet_name="Escalar a Fati", index=False)
    print(f"Excel: {salida}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--export", help="ruta xlsx para modo prueba (sin Metabase)")
    ap.add_argument("--fecha", default=dt.date.today().isoformat())
    args = ap.parse_args()

    df = cargar_export(args.export) if args.export else cargar_metabase()
    df = validar(df)
    acc, escalar = calcular_acciones(df)
    perdidos, ganados, total_gana = tracking(args.fecha, df)
    texto = construir_mensaje(args.fecha, len(df), acc, escalar, perdidos, ganados, total_gana)
    enviar_slack(texto)
    escribir_excel(f"acciones_buybox_{args.fecha}.xlsx", acc, escalar)


if __name__ == "__main__":
    main()
  
