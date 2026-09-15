# -*- coding: utf-8 -*-
"""
Análisis de prefijos de nombre.

Para cada instalación definida en utils/prefijos.json, verifica que los
registros de su tabla (tabla_psql) comiencen con alguno de los prefijos
indicados. Una tabla puede tener más de un prefijo válido (listados en el
json).

Salidas:
    docs/reporte/analisis_prefijo.xlsx
    docs/reporte/analisis_prefijo_detalle.csv
    docs/historico/analisis_prefijo-<YYYY-MM-DD_HH-MM-SS>.xlsx
    docs/historico/analisis_prefijo_detalle-<YYYY-MM-DD_HH-MM-SS>.csv
"""

import os
import sys
import json
import csv
import warnings
from datetime import datetime
from collections import defaultdict

import pandas as pd
import psycopg2

warnings.filterwarnings(
    "ignore",
    message="pandas only supports SQLAlchemy connectable",
    category=UserWarning,
)

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE_DIR)

import config

REPORTE_DIR = config.REPORTE_DIR
HISTORICO_DIR = config.HISTORICO_DIR
DB_CONN = config.DB_CONN

PREFIJOS_PATH = os.path.join(BASE_DIR, "utils", "prefijos.json")

COLUMNAS_DETALLE = [
    "tabla", "instalacion", "codigo", "id", "nombre", "prefijos_esperados",
    "propietario_id", "propietario",
]


def cargar_prefijos():
    with open(PREFIJOS_PATH, encoding="utf-8") as f:
        return json.load(f)


def cumple_prefijo(nombre, prefijos):
    n = str(nombre or "").strip().upper()
    if not n:
        return False
    for p in prefijos:
        p = str(p or "").strip().upper()
        if p and n.startswith(p):
            return True
    return False


def columnas_de(conn, tabla):
    cur = conn.cursor()
    cur.execute(
        "SELECT column_name FROM information_schema.columns "
        "WHERE table_schema='public' AND table_name=%s",
        (tabla,),
    )
    cols = [r[0] for r in cur.fetchall()]
    cur.close()
    return cols


def analizar(conn):
    empresa_map = {}
    cur = conn.cursor()
    cur.execute("SELECT id, name FROM api_empresa")
    empresa_map = {r[0]: r[1] for r in cur.fetchall()}
    cur.close()

    prefijos = cargar_prefijos()

    findings = []
    resumen = []

    for entry in prefijos:
        tabla = entry.get("tabla_psql", "")
        lista_prefijos = entry.get("prefijo", [])
        instalacion = entry.get("instalacion", "")

        if not tabla or not lista_prefijos:
            continue

        cols = columnas_de(conn, tabla)
        if "name" not in cols or "id" not in cols:
            continue

        has_prop = "propietario_id" in cols
        has_status = "status" in cols

        select = ["id", "name"]
        if has_prop:
            select.append("propietario_id")
        where = "WHERE status='EN_OPERACION'" if has_status else ""

        try:
            df = pd.read_sql(
                f"SELECT {', '.join(select)} FROM {tabla} {where}", conn
            )
        except Exception as e:
            print(f"[WARN] No se pudo leer {tabla}: {e}")
            continue

        if df.empty:
            continue

        df["name"] = df["name"].fillna("").astype(str)

        n_sin = 0
        for _, row in df.iterrows():
            nombre = row["name"]
            pid = row["propietario_id"] if has_prop else None
            if not cumple_prefijo(nombre, lista_prefijos):
                n_sin += 1
                findings.append({
                    "tabla": tabla,
                    "instalacion": instalacion,
                    "codigo": "PREFIJO_NOMBRE",
                    "id": row["id"],
                    "nombre": nombre,
                    "prefijos_esperados": ", ".join(lista_prefijos),
                    "propietario_id": "" if (pid is None or pd.isna(pid)) else str(pid),
                    "propietario": "" if (pid is None or pd.isna(pid)) else str(empresa_map.get(pid, "")),
                })

        resumen.append({
            "tabla": tabla,
            "instalacion": instalacion,
            "prefijos": ", ".join(lista_prefijos),
            "total_revisados": len(df),
            "sin_prefijo": n_sin,
        })

    return findings, resumen


def exportar(findings, resumen):
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    df_findings = pd.DataFrame(findings, columns=COLUMNAS_DETALLE)
    df_resumen = pd.DataFrame(resumen)

    prop = defaultdict(int)
    for f in findings:
        key = (f["propietario_id"], f["propietario"])
        prop[key] += 1
    por_propietario = []
    for (pid, pname), n in sorted(prop.items(), key=lambda kv: -kv[1]):
        por_propietario.append({
            "propietario_id": pid,
            "propietario": pname,
            "sin_prefijo": n,
        })
    df_propietario = pd.DataFrame(por_propietario)

    os.makedirs(REPORTE_DIR, exist_ok=True)
    os.makedirs(HISTORICO_DIR, exist_ok=True)

    base = "analisis_prefijo"
    xlsx_reporte = os.path.join(REPORTE_DIR, f"{base}.xlsx")
    xlsx_historico = os.path.join(HISTORICO_DIR, f"{base}-{ts}{config.sufijo_historico_bd()}.xlsx")
    csv_reporte = os.path.join(REPORTE_DIR, f"{base}_detalle.csv")
    csv_historico = os.path.join(HISTORICO_DIR, f"{base}_detalle-{ts}{config.sufijo_historico_bd()}.csv")

    for ruta in (xlsx_reporte, xlsx_historico):
        with pd.ExcelWriter(ruta, engine="xlsxwriter") as w:
            df_resumen.to_excel(w, sheet_name="Resumen", index=False)
            df_propietario.to_excel(w, sheet_name="Por_Propietario", index=False)
            df_findings.to_excel(w, sheet_name="Detalle", index=False)

    for ruta in (csv_reporte, csv_historico):
        with open(ruta, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(COLUMNAS_DETALLE)
            for item in findings:
                writer.writerow([item[c] for c in COLUMNAS_DETALLE])

    print(f"Total sin prefijo: {len(findings)}")
    print("Resumen (top 15):")
    if not df_resumen.empty:
        print(df_resumen.sort_values("sin_prefijo", ascending=False).head(15).to_string(index=False))
    print(f"Excel reporte: {xlsx_reporte}")
    print(f"CSV reporte:   {csv_reporte}")


def main():
    conn = psycopg2.connect(**DB_CONN)
    try:
        findings, resumen = analizar(conn)
    finally:
        conn.close()

    exportar(findings, resumen)


if __name__ == "__main__":
    main()
