# -*- coding: utf-8 -*-
"""
Revisión de la abreviatura de línea en el nombre de los vanos.

El nombre de un vano se construye como:
    {PREFIJO} {LINEA_ABREVIADA} {TORRE1}:{TORRE2} C{N}
    p. ej. VN CALAN-LSNA TOR E5:TOR E6 C1

Se verifica que la abreviatura de línea del nombre coincida con la abreviatura
oficial de la línea del vano, registrada en utils/abreviatura_linea.json
(campo "nombre_abreviado"). La línea se obtiene por el relacionamiento
estructural: vano.tramo_id -> tramo.circuito_id -> circuito.linea_id.

Si la línea no está en el json, o no tiene nombre_abreviado, el registro se
salta. Las consideraciones son las mismas que para torres
(scripts/revision_abreviatura_torre.py), incluida la normalización de espacios
alrededor del guion (" - ").

Salidas:
    docs/reporte/revision_abreviatura_vano.xlsx
    docs/reporte/revision_abreviatura_vano_detalle.csv
    docs/historico/revision_abreviatura_vano-<YYYY-MM-DD_HH-MM-SS>.xlsx
    docs/historico/revision_abreviatura_vano_detalle-<YYYY-MM-DD_HH-MM-SS>.csv
"""

import os
import sys
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
sys.path.insert(0, os.path.join(BASE_DIR, "scripts"))

import config
from revision_abreviatura_torre import (
    normalizar,
    extraer_abreviatura,
    cargar_abreviaturas,
    empresa_de,
)

REPORTE_DIR = config.REPORTE_DIR
HISTORICO_DIR = config.HISTORICO_DIR
DB_CONN = config.DB_CONN

TABLA = "api_vano"
INSTALACION = "Vano"
CODIGO = "ABREVIATURA_LINEA_VANO"

COLUMNAS_DETALLE = [
    "tabla", "instalacion", "codigo", "id", "nombre", "tramo_id", "circuito_id",
    "linea_id", "linea_name", "abreviatura_encontrada", "abreviatura_esperada",
    "detalle", "propietario_id", "propietario",
]


def analizar(conn):
    cur = conn.cursor()
    cur.execute("SELECT id, name FROM api_empresa")
    empresa_map = {r[0]: r[1] for r in cur.fetchall()}
    cur.execute(
        "SELECT v.id, v.name, v.tramo_id, v.propietario_id, "
        "       t.circuito_id, ci.linea_id "
        "FROM api_vano v "
        "LEFT JOIN api_tramo t ON v.tramo_id = t.id "
        "LEFT JOIN api_circuito ci ON t.circuito_id = ci.id "
        "WHERE v.status='EN_OPERACION'"
    )
    vanos = cur.fetchall()
    cur.close()

    abreviaturas = cargar_abreviaturas()

    findings = []
    totales = defaultdict(int)
    hallazgos = defaultdict(int)

    for vid, nombre, tramo_id, pid, circuito_id, linea_id in vanos:
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
            "id": vid,
            "nombre": "" if nombre is None else str(nombre),
            "tramo_id": "" if tramo_id is None else tramo_id,
            "circuito_id": "" if circuito_id is None else circuito_id,
            "linea_id": "" if linea_id is None else linea_id,
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

    base = "revision_abreviatura_vano"
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

    print(f"Total vanos evaluados: {total_evaluados}")
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
