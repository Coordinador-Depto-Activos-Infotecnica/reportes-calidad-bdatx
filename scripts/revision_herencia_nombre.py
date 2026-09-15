# -*- coding: utf-8 -*-
"""
Revisión de calidad de herencia de nombre.

La subestación propaga su nombre "aguas abajo": todo lo que está bajo ella
(patio, casa de servicios generales, elemento común, paño, barra) debe
contener el nombre de la subestación.

Ejemplo: subestación "S/E ANCOA" -> patio "PT S/E ANCOA 220kV" (contiene
"S/E ANCOA").

Relaciones verificadas (FK directa, comparando contra el nombre de la
subestación raíz):

- api_subestacion -> api_patiosubestacion, api_casaserviciosgenerales,
                     api_elementocomunssee (directo)
- api_patiosubestacion -> api_pano, api_barra, api_elementocomunpatiossee
                          (núcleo = subestación vía patio.subestacion_id)

Salidas:
    docs/reporte/revision_herencia_nombre.xlsx
    docs/reporte/revision_herencia_nombre_detalle.csv
    docs/historico/revision_herencia_nombre-<YYYY-MM-DD_HH-MM-SS>.xlsx
    docs/historico/revision_herencia_nombre_detalle-<YYYY-MM-DD_HH-MM-SS>.csv
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

import config

REPORTE_DIR = config.REPORTE_DIR
HISTORICO_DIR = config.HISTORICO_DIR
DB_CONN = config.DB_CONN

# Hijos directos de subestación: (tabla_hijo, columna_fk)
RELACIONES_DIRECTAS = [
    ("api_patiosubestacion", "subestacion_id"),
    ("api_casaserviciosgenerales", "subestacion_id"),
    ("api_elementocomunssee", "subestacion_id"),
]

# Hijos de patio (núcleo = subestación vía patio): (tabla_hijo, columna_fk)
RELACIONES_VIA_PATIO = [
    ("api_pano", "patio_subestacion_id"),
    ("api_barra", "patio_subestacion_id"),
    ("api_elementocomunpatiossee", "patio_subestacion_id"),
]

# Equipamiento (hoja) que se conecta por nodo_id: hijos de paño/barra/celda.
TABLAS_EQUIPO = [
    "api_aislador", "api_bancocondensador", "api_compensadoractivo",
    "api_compensadorestaticoreactivo", "api_condensadoracoplamiento",
    "api_condensadorserie", "api_condensadorsincrono", "api_conexiontierra",
    "api_desconectador", "api_dispositivoreconexion", "api_elementocomunssggcaseta",
    "api_interruptor", "api_medidor", "api_mufa", "api_pararrayo", "api_reactor",
    "api_releproteccion", "api_tablero", "api_trampaonda",
    "api_transformadorcorriente", "api_transformadorpotencial",
    "api_transformadorssaa", "api_transformadorzigzag",
]

COLUMNAS_DETALLE = [
    "relacion", "codigo", "id_hijo", "nombre_hijo", "nombre_subestacion",
    "propietario_id", "propietario",
]


def normalizar(texto):
    return " ".join(str(texto or "").upper().split())


def analizar(conn):
    empresa_map = {}
    cur = conn.cursor()
    cur.execute("SELECT id, name FROM api_empresa")
    empresa_map = {r[0]: r[1] for r in cur.fetchall()}
    cur.close()

    subestaciones = pd.read_sql(
        "SELECT id, name FROM api_subestacion WHERE status='EN_OPERACION'", conn
    )
    sub_name = dict(zip(subestaciones["id"], subestaciones["name"]))

    patios = pd.read_sql(
        "SELECT id, subestacion_id FROM api_patiosubestacion WHERE status='EN_OPERACION'",
        conn,
    )
    patio_sub_id = dict(zip(patios["id"], patios["subestacion_id"]))

    findings = []
    resumen = []

    # 1. Relaciones directas desde subestación
    for hijo_tabla, fk in RELACIONES_DIRECTAS:
        hijos = pd.read_sql(
            f"SELECT id, name, {fk}, propietario_id FROM {hijo_tabla} "
            f"WHERE status='EN_OPERACION' AND {fk} IS NOT NULL",
            conn,
        )
        n_sin = 0
        for _, hijo in hijos.iterrows():
            nombre_sub = sub_name.get(hijo[fk], "")
            if not normalizar(nombre_sub):
                continue
            if normalizar(nombre_sub) not in normalizar(hijo["name"]):
                n_sin += 1
                pid = hijo["propietario_id"]
                findings.append({
                    "relacion": f"api_subestacion -> {hijo_tabla}",
                    "codigo": "HERENCIA_SUBESTACION",
                    "id_hijo": hijo["id"],
                    "nombre_hijo": str(hijo["name"] or ""),
                    "nombre_subestacion": nombre_sub,
                    "propietario_id": "" if (pid is None or pd.isna(pid)) else str(pid),
                    "propietario": "" if (pid is None or pd.isna(pid)) else str(empresa_map.get(pid, "")),
                })
        resumen.append({
            "relacion": f"api_subestacion -> {hijo_tabla}",
            "total_hijos": len(hijos),
            "sin_herencia": n_sin,
        })

    # 2. Relaciones vía patio (núcleo = subestación)
    for hijo_tabla, fk in RELACIONES_VIA_PATIO:
        hijos = pd.read_sql(
            f"SELECT id, name, {fk}, propietario_id FROM {hijo_tabla} "
            f"WHERE status='EN_OPERACION' AND {fk} IS NOT NULL",
            conn,
        )
        n_sin = 0
        for _, hijo in hijos.iterrows():
            sub_id = patio_sub_id.get(hijo[fk])
            nombre_sub = sub_name.get(sub_id, "") if sub_id is not None else ""
            if not normalizar(nombre_sub):
                continue
            if normalizar(nombre_sub) not in normalizar(hijo["name"]):
                n_sin += 1
                pid = hijo["propietario_id"]
                findings.append({
                    "relacion": f"api_patiosubestacion -> {hijo_tabla}",
                    "codigo": "HERENCIA_SUBESTACION",
                    "id_hijo": hijo["id"],
                    "nombre_hijo": str(hijo["name"] or ""),
                    "nombre_subestacion": nombre_sub,
                    "propietario_id": "" if (pid is None or pd.isna(pid)) else str(pid),
                    "propietario": "" if (pid is None or pd.isna(pid)) else str(empresa_map.get(pid, "")),
                })
        resumen.append({
            "relacion": f"api_patiosubestacion -> {hijo_tabla}",
            "total_hijos": len(hijos),
            "sin_herencia": n_sin,
        })

    # 3. Relaciones por nodo compartido (hijos de paño/barra/casa SSGG)
    #    Mapeo nodo_id -> nombre de subestación (vía paño, barra y casa SSGG).
    node_to_sub = {}
    panos = pd.read_sql(
        "SELECT nodo_id, patio_subestacion_id FROM api_pano "
        "WHERE status='EN_OPERACION' AND nodo_id IS NOT NULL",
        conn,
    )
    for _, p in panos.iterrows():
        sub_id = patio_sub_id.get(p["patio_subestacion_id"])
        if sub_id is not None and sub_id in sub_name:
            node_to_sub[p["nodo_id"]] = sub_name[sub_id]

    barras = pd.read_sql(
        "SELECT nodo_id, patio_subestacion_id FROM api_barra "
        "WHERE status='EN_OPERACION' AND nodo_id IS NOT NULL",
        conn,
    )
    for _, b in barras.iterrows():
        sub_id = patio_sub_id.get(b["patio_subestacion_id"])
        if sub_id is not None and sub_id in sub_name:
            node_to_sub[b["nodo_id"]] = sub_name[sub_id]

    casas = pd.read_sql(
        "SELECT nodo_id, subestacion_id FROM api_casaserviciosgenerales "
        "WHERE status='EN_OPERACION' AND nodo_id IS NOT NULL",
        conn,
    )
    for _, c in casas.iterrows():
        if c["subestacion_id"] is not None and c["subestacion_id"] in sub_name:
            node_to_sub[c["nodo_id"]] = sub_name[c["subestacion_id"]]

    for equipo_tabla in TABLAS_EQUIPO:
        equipos = pd.read_sql(
            f"SELECT id, name, nodo_id, propietario_id FROM {equipo_tabla} "
            f"WHERE status='EN_OPERACION' AND nodo_id IS NOT NULL",
            conn,
        )
        n_sin = 0
        for _, eq in equipos.iterrows():
            nombre_sub = node_to_sub.get(eq["nodo_id"], "")
            if not normalizar(nombre_sub):
                continue
            if normalizar(nombre_sub) not in normalizar(eq["name"]):
                n_sin += 1
                pid = eq["propietario_id"]
                findings.append({
                    "relacion": f"nodo -> {equipo_tabla}",
                    "codigo": "HERENCIA_SUBESTACION",
                    "id_hijo": eq["id"],
                    "nombre_hijo": str(eq["name"] or ""),
                    "nombre_subestacion": nombre_sub,
                    "propietario_id": "" if (pid is None or pd.isna(pid)) else str(pid),
                    "propietario": "" if (pid is None or pd.isna(pid)) else str(empresa_map.get(pid, "")),
                })
        resumen.append({
            "relacion": f"nodo -> {equipo_tabla}",
            "total_hijos": len(equipos),
            "sin_herencia": n_sin,
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
            "sin_herencia": n,
        })
    df_propietario = pd.DataFrame(por_propietario)

    os.makedirs(REPORTE_DIR, exist_ok=True)
    os.makedirs(HISTORICO_DIR, exist_ok=True)

    base = "revision_herencia_nombre"
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

    print("Resumen por relación:")
    print(df_resumen.to_string(index=False))
    print()
    print(f"Total sin herencia de nombre: {len(findings)}")
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
