#!/usr/bin/env python3
"""
vitau_buybox_diario.py — Monitor diario de precios / Buy Box de Mercado Libre (Vitau).

Jala la card 4556 (publicaciones-meli-snapshot-mas-reciente) DIRECTO de Metabase.
Fati ya calcula la decision dentro del query: la columna accion_precio dice
OK / Bajar precio / Subir precio, y precio_sugerido es el precio apto por margenes.
El bot solo lee esa decision, resume en Slack (Top por ventas) y deja la lista
completa en un Excel.

Uso en GitHub Actions:
  python vitau_buybox_diario.py
Modo prueba con un export local:
  python vitau_buybox_diario.py --export snapshot.xlsx

Variables de entorno (GitHub Secrets):
  METABASE_URL, METABASE_USER, METABASE_PASS, METABASE_CARD (default 4556)
  SLACK_TOKEN, SLACK_CHANNEL
  SNAPSHOT_FILE (default snapshot_buybox.json)
  TOP_N (cuantas filas mostrar por bloque en Slack, default 12)
"""

import argparse
import datetime as dt
import json
import os
import sys
from decimal import Decimal, ROUND_HALF_UP

import pandas as pd
import requests

ACTIVO = "Activo"
GANANDO = "Ganando"
A_BAJAR = "Bajar precio"
A_SUBIR = "Subir precio"

COL_ID = "item_id"
COL_PROD = "titulo"
COL_ESTADO = "estado"
COL_BB = "buybox_status"
COL_ACTUAL = "precio_actual"
COL_SUGERIDO = "precio_sugerido"
COL_ACCION = "accion_precio"
COL_STOCK = "stock"
COL_VENTAS = "ventas_totales"
COL_URL = "url"
COL_GANADOR = "precio_ganador_competencia"
REQUERIDAS = [COL_ID, COL_PROD, COL_ESTADO, COL_ACCION, COL_ACTUAL, COL_SUGERIDO]


def money(x):
    if pd.isna(x):
        return None
    return float(Decimal(str(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def cargar_metabase():
    url = os.environ["METABASE_URL"]
    card = os.environ.get("METABASE_CARD", "4556")
    s = requests.post(f"{url}/api/session",
                      json={"username": os.environ["METABASE_USER"],
                            "password": os.environ["METABASE_PASS"]}, timeout=30)
    token = s.json().get("id")
    if not token:
        sys.exit(f"ERROR: no se pudo autenticar en Metabase: {s.text[:200]}")
    # /query/json devuelve TODAS las filas (sin el tope de 2000 de /query)
    q = requests.post(f"{url}/api/card/{card}/query/json",
                      headers={"X-Metabase-Session": token}, json={}, timeout=300)
    rows = q.json()
    if not isinstance(rows, list):
        sys.exit(f"ERROR: respuesta inesperada de Metabase: {str(rows)[:200]}")
    print(f"Metabase card {card}: {len(rows)} filas")
    return pd.DataFrame(rows)


def cargar_export(path):
    return pd.read_excel(path, dtype={COL_ID: str})


def validar(df):
    faltan = [c for c in REQUERIDAS if c not in df.columns]
    if faltan:
        sys.exit(f"ERROR: faltan columnas {faltan}.\nColumnas recibidas: {list(df.columns)}")
    df[COL_ID] = df[COL_ID].astype(str).str.strip()
    for c in (COL_ACTUAL, COL_SUGERIDO, COL_VENTAS, COL_STOCK, COL_GANADOR):
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def acciones(df):
    """Solo publicaciones activas; ordenadas por ventas (lo que mas urge primero)."""
    act = df[df[COL_ESTADO] == ACTIVO]
    orden = COL_VENTAS if COL_VENTAS in df.columns else COL_ACTUAL
    bajar = act[act[COL_ACCION] == A_BAJAR].sort_values(orden, ascending=False)
    subir = act[act[COL_ACCION] == A_SUBIR].sort_values(orden, ascending=False)
    return act, bajar, subir


def tracking(fecha, df):
    path = os.environ.get("SNAPSHOT_FILE", "snapshot_buybox.json")
    ganan_hoy = sorted(df.loc[df.get(COL_BB) == GANANDO, COL_ID].astype(str)) if COL_BB in df.columns else []
    perdidos = ganados = None
    if os.path.exists(path):
        prev = json.load(open(path))
        if prev.get("fecha") != fecha:
            ayer, hoy = set(prev.get("ganando", [])), set(ganan_hoy)
            perdidos, ganados = ayer - hoy, hoy - ayer
    json.dump({"fecha": fecha, "ganando": ganan_hoy}, open(path, "w"))
    return perdidos, ganados, len(ganan_hoy)


def fmt_filas(dfx, n):
    L = []
    for _, r in dfx.head(n).iterrows():
        prod = str(r[COL_PROD])[:46]
        a, sug = money(r[COL_ACTUAL]), money(r[COL_SUGERIDO])
        ventas = f" · {int(r[COL_VENTAS])} ventas" if COL_VENTAS in dfx.columns and pd.notna(r[COL_VENTAS]) else ""
        link = f"<{r[COL_URL]}|{prod}>" if COL_URL in dfx.columns and pd.notna(r.get(COL_URL)) else prod
        L.append(f"• {link} — ${a:,.2f} → *${sug:,.2f}*{ventas}")
    return L


def construir_mensaje(fecha, act, bajar, subir, perdidos, ganados, total_gana, top_n):
    L = [f"*Buy Box ML — {fecha}*", f"Publicaciones activas: {len(act):,} · Ganando: {total_gana}"]
    if perdidos is not None:
        if perdidos:
            L.append(f":small_red_triangle_down: Perdimos el Buy Box en {len(perdidos)} (ayer ganábamos)")
        if ganados:
            L.append(f":white_check_mark: Recuperamos {len(ganados)}")
        if not perdidos and not ganados:
            L.append("Sin cambios de Buy Box vs ayer")

    L.append("")
    if bajar.empty and subir.empty:
        L.append(":tada: *Todos los precios están en su valor óptimo.* Sin acciones hoy.")
        return "\n".join(L)

    if not bajar.empty:
        L.append(f":arrow_down: *Bajar precio: {len(bajar)}*  (Top {min(top_n, len(bajar))} por ventas)")
        L += fmt_filas(bajar, top_n)
    if not subir.empty:
        L.append(f"\n:arrow_up: *Subir precio: {len(subir)}*  (margen sin capturar — Top {min(5, len(subir))})")
        L += fmt_filas(subir, 5)
    L.append(f"\n_Lista completa de {len(bajar) + len(subir)} cambios en el Excel (abajo)._")
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


def subir_excel_slack(path, titulo):
    """Sube el Excel al canal (best-effort). Requiere scope files:write en el bot.
    Si falla, el mensaje de texto ya se envió y el Excel queda en los artifacts."""
    token, canal = os.environ.get("SLACK_TOKEN"), os.environ.get("SLACK_CHANNEL")
    if not token or not canal:
        return
    try:
        size = os.path.getsize(path)
        r1 = requests.get("https://slack.com/api/files.getUploadURLExternal",
                          headers={"Authorization": f"Bearer {token}"},
                          params={"filename": os.path.basename(path), "length": size}, timeout=20).json()
        if not r1.get("ok"):
            print(f"Slack upload: {r1.get('error')} (revisa scope files:write)")
            return
        with open(path, "rb") as fh:
            requests.post(r1["upload_url"], files={"file": fh}, timeout=60)
        r3 = requests.post("https://slack.com/api/files.completeUploadExternal",
                           headers={"Authorization": f"Bearer {token}"},
                           json={"files": [{"id": r1["file_id"], "title": titulo}],
                                 "channel_id": canal}, timeout=20).json()
        print(f"Slack archivo: ok={r3.get('ok')}")
    except Exception as e:
        print(f"Slack upload falló (no bloqueante): {e}")


def escribir_excel(salida, bajar, subir):
    cols = [c for c in (COL_ID, COL_PROD, COL_ACTUAL, COL_SUGERIDO, COL_GANADOR,
                        COL_STOCK, COL_VENTAS, COL_BB, COL_URL) if c in bajar.columns]
    with pd.ExcelWriter(salida, engine="openpyxl") as xw:
        (bajar[cols] if not bajar.empty else pd.DataFrame(columns=cols)).to_excel(xw, sheet_name="Bajar precio", index=False)
        if not subir.empty:
            subir[cols].to_excel(xw, sheet_name="Subir precio", index=False)
    print(f"Excel: {salida}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--export", help="ruta xlsx para modo prueba (sin Metabase)")
    ap.add_argument("--fecha", default=dt.date.today().isoformat())
    args = ap.parse_args()

    df = cargar_export(args.export) if args.export else cargar_metabase()
    df = validar(df)
    act, bajar, subir = acciones(df)
    perdidos, ganados, total_gana = tracking(args.fecha, df)
    top_n = int(os.environ.get("TOP_N", "12"))
    texto = construir_mensaje(args.fecha, act, bajar, subir, perdidos, ganados, total_gana, top_n)
    salida = f"acciones_buybox_{args.fecha}.xlsx"
    escribir_excel(salida, bajar, subir)
    enviar_slack(texto)
    if not (bajar.empty and subir.empty):
        subir_excel_slack(salida, f"Acciones de precio {args.fecha}")


if __name__ == "__main__":
    main()
