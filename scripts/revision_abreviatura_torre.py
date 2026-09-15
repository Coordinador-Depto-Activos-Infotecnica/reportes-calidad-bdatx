# -*- coding: utf-8 -*-
"""
Revisión de la abreviatura de línea en el nombre de las torres.

El nombre de una torre se construye como:
    {PREFIJO} {LINEA_ABREVIADA} {NOMBRE TORRE}
    p. ej. TOR ARRI-SGOR E79

Se verifica que la abreviatura de línea del nombre coincida con la abreviatura
oficial de la línea (api_torre.linea_id) registrada en
utils/abreviatura_linea.json (campo "nombre_abreviado"). Si la línea no está
en el json, o no tiene nombre_abreviado, el registro se salta.

La comparación normaliza mayúsculas, tildes y espacios alrededor del guion
(" - "), de modo que TOR LAG - GRS E040 se evalúa con la abreviatura LAG-GRS.

Salidas:
    docs/reporte/revision_abreviatura_torre.xlsx
    docs/reporte/revision_abreviatura_torre_detalle.csv
    docs/historico/revision_abreviatura_torre-<YYYY-MM-DD_HH-MM-SS>.xlsx
    docs/historico/revision_abreviatura_torre_detalle-<YYYY-MM-DD_HH-MM-SS>.csv
"""

import os
import re
import sys
import csv
import json
import unicodedata
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

ABREV_JSON_PATH = os.path.join(BASE_DIR, "utils", "abreviatura_linea.json")

TABLA = "api_torre"
INSTALACION = "Torre"
CODIGO = "ABREVIATURA_LINEA_TORRE"

COLUMNAS_DETALLE = [
    "tabla", "instalacion", "codigo", "id", "nombre", "linea_id", "linea_name",
    "abreviatura_encontrada", "abreviatura_esperada", "detalle",
    "propietario_id", "propietario",
]


def _sin_acentos(texto):
    return unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")


def normalizar(texto):
    texto = _sin_acentos(str(texto or "").upper())
    texto = re.sub(r"\s*[-\u2013\u2014]\s*", "-", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def extraer_abreviatura(nombre):
    partes = normalizar(nombre).split(" ")
    if len(partes) < 3:
        return None
    return partes[1]


def cargar_abreviaturas():
    if not os.path.exists(ABREV_JSON_PATH):
        print(
            f"[WARN] No existe {ABREV_JSON_PATH}. Se omite la revisión de "
            "abreviatura de línea (genere el archivo con "
            "utils/scripts/lineas_abreviadas.py)."
        )
        return {}
    with open(ABREV_JSON_PATH, encoding="utf-8") as f:
        registros = json.load(f)
    mapa = {}
    for r in registros:
        lid = r.get("linea_id")
        if lid is None:
            continue
        mapa[lid] = {
            "nombre_abreviado": r.get("nombre_abreviado"),
            "nombre_linea": r.get("nombre_linea") or "",
        }
    return mapa


def empresa_de(pid, empresa_map):
    if pid is None or (isinstance(pid, float) and pd.isna(pid)):
        return "Sin empresa"
    return empresa_map.get(pid) or "Sin empresa"


def analizar(conn):
    cur = conn.cursor()
    cur.execute("SELECT id, name FROM api_empresa")
    empresa_map = {r[0]: r[1] for r in cur.fetchall()}
    cur.execute(
        f"SELECT id, name, linea_id, propietario_id FROM {TABLA} "
        f"WHERE status='EN_OPERACION'"
    )
    torres = cur.fetchall()
    cur.close()

    abreviaturas = cargar_abreviaturas()

    findings = []
    totales = defaultdict(int)
    hallazgos = defaultdict(int)

    for tid, nombre, linea_id, pid in torres:
        info = abreviaturas.get(linea_id)
        if not info or not info.get("nombre_abreviado"):
            continue

        esperada = info["nombre_abreviado"]
        linea_name = info.get("nombre_linea", "")
        pname = empresa_de(pid, empresa_map)
        totales[pname] += 1

        encontrada = extraer_abreviatura(nombre)
        if encontrada is not None and encontrada == normalizar(esperada):
            continue

        hallazgos[pname] += 1
        if encontrada is None:
            detalle = (
                f'No se pudo identificar la abreviatura de línea en el nombre; '
                f'se esperaba "{esperada}" para la línea {linea_name} (id {linea_id}).'
            )
        else:
            detalle = (
                f'La abreviatura de línea "{encontrada}" no coincide con la esperada '
                f'"{esperada}" de la línea {linea_name} (id {linea_id}).'
            )

        findings.append({
            "tabla": TABLA,
            "instalacion": INSTALACION,
            "codigo": CODIGO,
            "id": tid,
            "nombre": "" if nombre is None else str(nombre),
            "linea_id": linea_id,
            "linea_name": linea_name,
            "abreviatura_encontrada": encontrada or "",
            "abreviatura_esperada": esperada,
            "detalle": detalle,
            "propietario_id": "" if pid is None else str(pid),
            "propietario": "" if pname == "Sin empresa" else pname,
        })

    return findings, totales, hallazgos


def exportar(findings, totales, hallazgos):
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    df_findings = pd.DataFrame(findings, columns=COLUMNAS_DETALLE)

    total_evaluados = sum(totales.values())
    df_resumen = pd.DataFrame([{
        "tabla": TABLA,
        "codigo": CODIGO,
        "total_evaluados": total_evaluados,
        "hallazgos": len(findings),
    }])

    empresas = sorted(set(totales) | set(hallazgos))
    df_propietario = pd.DataFrame([{
        "propietario": e,
        "total_evaluados": totales.get(e, 0),
        "hallazgos": hallazgos.get(e, 0),
    } for e in empresas])

    df_empresa = pd.DataFrame([{
        "empresa_name": e,
        "total_registros": totales.get(e, 0),
        "hallazgos": hallazgos.get(e, 0),
    } for e in empresas])

    os.makedirs(REPORTE_DIR, exist_ok=True)
    os.makedirs(HISTORICO_DIR, exist_ok=True)

    base = "revision_abreviatura_torre"
    xlsx_reporte = os.path.join(REPORTE_DIR, f"{base}.xlsx")
    xlsx_historico = os.path.join(HISTORICO_DIR, f"{base}-{ts}{config.sufijo_historico_bd()}.xlsx")
    csv_reporte = os.path.join(REPORTE_DIR, f"{base}_detalle.csv")
    csv_historico = os.path.join(HISTORICO_DIR, f"{base}_detalle-{ts}{config.sufijo_historico_bd()}.csv")

    for ruta in (xlsx_reporte, xlsx_historico):
        with pd.ExcelWriter(ruta, engine="xlsxwriter") as w:
            df_resumen.to_excel(w, sheet_name="Resumen", index=False)
            df_propietario.to_excel(w, sheet_name="Por_Propietario", index=False)
            df_empresa.to_excel(w, sheet_name="analisis_por_empresa", index=False)
            df_findings.to_excel(w, sheet_name="Detalle", index=False)

    for ruta in (csv_reporte, csv_historico):
        with open(ruta, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(COLUMNAS_DETALLE)
            for item in findings:
                writer.writerow([item[c] for c in COLUMNAS_DETALLE])

    print(f"Total torres evaluadas: {total_evaluados}")
    print(f"Total hallazgos ({CODIGO}): {len(findings)}")
    if not df_propietario.empty:
        print(df_propietario.sort_values("hallazgos", ascending=False).head(15).to_string(index=False))
    print(f"Excel reporte: {xlsx_reporte}")
    print(f"CSV reporte:   {csv_reporte}")


def main():
    conn = psycopg2.connect(**DB_CONN)
    try:
        findings, totales, hallazgos = analizar(conn)
    finally:
        conn.close()

    exportar(findings, totales, hallazgos)


if __name__ == "__main__":
    main()
