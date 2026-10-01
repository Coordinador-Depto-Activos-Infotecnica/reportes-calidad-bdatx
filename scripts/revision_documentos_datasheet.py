# -*- coding: utf-8 -*-
"""
Revisión de los documentos referenciados en el campo `datasheet` de las tablas
api_subestacion y api_linea.

Los campos de documento del datasheet (planos, diagramas, perfiles, listados y
permisos) se completan con el `id` de un registro de api_documento. Para cada
instalación EN_OPERACION se evalúa cada campo obligatorio:

- `DOCUMENTO_FALTANTE`: el campo no está informado (ausente, `null` o `0`).
- `DOCUMENTO_NO_EXISTE`: el `id` referenciado no existe en api_documento
  (incluye el placeholder `1` que usa la aplicación).
- `DOCUMENTO_INCORRECTO`: el documento existe, pero es genérico (LEEME,
  NO APLICA, PENDIENTE, extensión .txt, etc.) y no corresponde al documento
  específico exigido por la especificación del dato.

Los campos de documento se obtienen de los esqueletos JSON de cada instalación
(utils/esqueletos_instalaciones_json/<instalacion>.json): claves del `datasheet`
cuyo valor de ejemplo es numérico.

Salidas:
    docs/reporte/revision_documentos_datasheet.xlsx
    docs/reporte/revision_documentos_datasheet_detalle.csv
    docs/historico/revision_documentos_datasheet-<YYYY-MM-DD_HH-MM-SS>.xlsx
    docs/historico/revision_documentos_datasheet_detalle-<YYYY-MM-DD_HH-MM-SS>.csv
"""

import os
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

ESQUELETOS_DIR = os.path.join(BASE_DIR, "utils", "esqueletos_instalaciones_json")

# (tabla, archivo de esqueleto)
TABLAS = [
    ("api_subestacion", "subestaciones.json"),
    ("api_linea", "lineas.json"),
]

CODIGO_FALTANTE = "DOCUMENTO_FALTANTE"
CODIGO_NO_EXISTE = "DOCUMENTO_NO_EXISTE"
CODIGO_INCORRECTO = "DOCUMENTO_INCORRECTO"

# Patrones (en name/file_name, sin tildes ni mayúsculas) y extensiones que
# identifican documentos genéricos o no específicos.
PATRONES_INCORRECTOS = [
    "leeme", "no aplica", "no_aplica", "pendiente", "sin informacion",
    "sin_informacion", "generico",
]
EXTENSIONES_INCORRECTAS = {"txt"}

COLUMNAS_DETALLE = [
    "tabla", "codigo", "id", "nombre", "propietario_id", "propietario",
    "campo", "documento_id", "documento_nombre", "detalle",
]


def normalizar(texto):
    texto = str(texto or "").strip().lower()
    return "".join(
        c for c in unicodedata.normalize("NFKD", texto)
        if not unicodedata.combining(c)
    )


def cargar_campos_documento():
    """Devuelve {tabla: [campos]} con las claves de documento del datasheet
    (las que en el esqueleto tienen un valor numérico de ejemplo)."""
    campos = {}
    for tabla, archivo in TABLAS:
        ruta = os.path.join(ESQUELETOS_DIR, archivo)
        with open(ruta, encoding="utf-8") as f:
            data = json.load(f)
        datasheet = data.get("datasheet", {})
        campos[tabla] = [
            k for k, v in datasheet.items()
            if isinstance(v, (int, float)) and not isinstance(v, bool)
        ]
    return campos


def cargar_documentos(conn):
    cur = conn.cursor()
    cur.execute("SELECT id, name, file_name, extension FROM api_documento")
    docs = {r[0]: {"name": r[1], "file_name": r[2], "extension": r[3]} for r in cur.fetchall()}
    cur.close()
    return docs


def es_incorrecto(doc):
    texto = normalizar(doc.get("name")) + " " + normalizar(doc.get("file_name"))
    if any(p in texto for p in PATRONES_INCORRECTOS):
        return True
    return normalizar(doc.get("extension")) in EXTENSIONES_INCORRECTAS


def analizar(conn):
    cur = conn.cursor()
    cur.execute("SELECT id, name FROM api_empresa")
    empresa_map = {r[0]: r[1] for r in cur.fetchall()}
    cur.close()

    docs = cargar_documentos(conn)
    campos_por_tabla = cargar_campos_documento()

    findings = []
    totales = defaultdict(int)
    hallazgos = defaultdict(int)

    for tabla, _ in TABLAS:
        cur = conn.cursor()
        cur.execute(
            f'SELECT id, name, propietario_id, datasheet FROM "{tabla}" '
            f"WHERE status='EN_OPERACION' AND datasheet IS NOT NULL"
        )
        filas = cur.fetchall()
        cur.close()

        for rid, nombre, pid, ds in filas:
            if not isinstance(ds, dict):
                continue
            pname = empresa_map.get(pid, "") or "Sin empresa"

            for campo in campos_por_tabla[tabla]:
                totales[(pname, tabla)] += 1

                valor = ds.get(campo)
                tiene_doc = (
                    isinstance(valor, int) and not isinstance(valor, bool) and valor > 0
                )

                if not tiene_doc:
                    codigo = CODIGO_FALTANTE
                    doc_nombre = ""
                    detalle = f'Campo "{campo}": sin documento registrado'
                else:
                    doc = docs.get(valor)
                    if doc is None:
                        codigo = CODIGO_NO_EXISTE
                        doc_nombre = ""
                        detalle = (
                            f'Campo "{campo}": el documento id {valor} no existe en api_documento'
                        )
                    elif es_incorrecto(doc):
                        codigo = CODIGO_INCORRECTO
                        doc_nombre = str(doc.get("name") or doc.get("file_name") or "")
                        detalle = (
                            f'Campo "{campo}": se completó con el documento genérico '
                            f'"{doc_nombre}" (id {valor}).'
                        )
                    else:
                        continue

                hallazgos[(pname, tabla)] += 1
                findings.append({
                    "tabla": tabla,
                    "codigo": codigo,
                    "id": rid,
                    "nombre": "" if nombre is None else str(nombre),
                    "propietario_id": "" if pid is None else str(pid),
                    "propietario": "" if pname == "Sin empresa" else pname,
                    "campo": campo,
                    "documento_id": valor if tiene_doc else "",
                    "documento_nombre": doc_nombre,
                    "detalle": detalle,
                })

    return findings, totales, hallazgos


def exportar(findings, totales, hallazgos):
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    df_findings = pd.DataFrame(findings, columns=COLUMNAS_DETALLE)

    resumen = defaultdict(int)
    por_campo = defaultdict(int)
    por_propietario = defaultdict(int)
    faltantes = defaultdict(int)
    propietario_id = {}
    for f in findings:
        resumen[(f["tabla"], f["codigo"])] += 1
        por_campo[(f["tabla"], f["campo"], f["codigo"])] += 1
        clave = (f["propietario"], f["codigo"])
        por_propietario[clave] += 1
        propietario_id[f["propietario"]] = f["propietario_id"]
        if f["codigo"] == CODIGO_FALTANTE:
            empresa = f["propietario"] or "Sin empresa"
            faltantes[(empresa, f["tabla"])] += 1

    df_resumen = pd.DataFrame([
        {"tabla": t, "codigo": c, "hallazgos": n}
        for (t, c), n in sorted(resumen.items())
    ])
    df_campo = pd.DataFrame([
        {"tabla": t, "campo": campo, "codigo": c, "hallazgos": n}
        for (t, campo, c), n in sorted(por_campo.items())
    ])
    empresas = sorted({e for (e, _t) in totales} | {e for (e, _t) in hallazgos})
    df_empresa = pd.DataFrame([
        {
            "empresa_name": e,
            "total_registros": totales.get((e, "api_subestacion"), 0) + totales.get((e, "api_linea"), 0),
            "hallazgos": hallazgos.get((e, "api_subestacion"), 0) + hallazgos.get((e, "api_linea"), 0),
            "total_subestacion": totales.get((e, "api_subestacion"), 0),
            "total_linea": totales.get((e, "api_linea"), 0),
            "hallazgos_subestacion": hallazgos.get((e, "api_subestacion"), 0),
            "hallazgos_linea": hallazgos.get((e, "api_linea"), 0),
            "faltantes_subestacion": faltantes.get((e, "api_subestacion"), 0),
            "faltantes_linea": faltantes.get((e, "api_linea"), 0),
        }
        for e in empresas
    ])
    df_propietario = pd.DataFrame([
        {
            "propietario_id": propietario_id.get(pname, ""),
            "propietario": pname,
            "codigo": c,
            "hallazgos": n,
        }
        for (pname, c), n in sorted(por_propietario.items())
    ])

    os.makedirs(REPORTE_DIR, exist_ok=True)
    os.makedirs(HISTORICO_DIR, exist_ok=True)

    base = "revision_documentos_datasheet"
    xlsx_reporte = os.path.join(REPORTE_DIR, f"{base}.xlsx")
    xlsx_historico = os.path.join(HISTORICO_DIR, f"{base}-{ts}{config.sufijo_historico_bd()}.xlsx")
    csv_reporte = os.path.join(REPORTE_DIR, f"{base}_detalle.csv")
    csv_historico = os.path.join(HISTORICO_DIR, f"{base}_detalle-{ts}{config.sufijo_historico_bd()}.csv")

    for ruta in (xlsx_reporte, xlsx_historico):
        with pd.ExcelWriter(ruta, engine="xlsxwriter") as w:
            df_resumen.to_excel(w, sheet_name="Resumen", index=False)
            df_campo.to_excel(w, sheet_name="Por_Campo", index=False)
            df_propietario.to_excel(w, sheet_name="Por_Propietario", index=False)
            df_empresa.to_excel(w, sheet_name="analisis_por_empresa", index=False)
            df_findings.to_excel(w, sheet_name="Detalle", index=False)

    for ruta in (csv_reporte, csv_historico):
        with open(ruta, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(COLUMNAS_DETALLE)
            for item in findings:
                writer.writerow([item[c] for c in COLUMNAS_DETALLE])

    print(f"Total de referencias evaluadas: {sum(totales.values())}")
    print(f"Total hallazgos: {len(findings)}")
    if not df_resumen.empty:
        print(df_resumen.to_string(index=False))
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
