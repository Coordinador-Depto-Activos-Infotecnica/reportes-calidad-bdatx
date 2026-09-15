# -*- coding: utf-8 -*-
"""
Auditoría de completitud de contenedores de la red (Subestación, Patio,
Paño, Casa de Servicios Generales, Armario, Torre y Marcolinea) usando la
BD PostgreSQL local.

Detecta contenedores "vacíos" (sin ningún hijo EN_OPERACION):

- Subestación y Patio  -> por FK directa (subestacion_id / patio_subestacion_id).
- Paño, Casa SSGG y Armario -> por nodo compartido: el contenedor tiene un
  nodo_id, y sus hijos (celda, interruptor, etc.) comparten ese mismo nodo_id.
  Se busca en todas las tablas que poseen nodo_id, excluyendo la propia tabla.
- Torre y Marcolinea -> sin OOCC (vía api_torre_oocc / api_marcolinea_oocc) y
  sin accesorios estructurales (vía api_accesorioestructuralineaaerea usando
  nodo_estructura_id.id / nodo_estructura_id.instalacion).

Salidas:
    docs/reporte/auditoria_contenedores_metricas.xlsx
    docs/reporte/auditoria_contenedores_metricas_detalle.csv
    docs/historico/auditoria_contenedores_metricas-<YYYY-MM-DD_HH-MM-SS>.xlsx
    docs/historico/auditoria_contenedores_metricas_detalle-<YYYY-MM-DD_HH-MM-SS>.csv
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

# Contenedores con relación por FK directa: tabla -> hijos (tabla, columna_fk)
CONTENEDORES_FK = [
    {
        "tabla": "api_subestacion",
        "label": "Subestación",
        "hijos": [
            ("api_patiosubestacion", "subestacion_id"),
            ("api_casaserviciosgenerales", "subestacion_id"),
            ("api_elementocomunssee", "subestacion_id"),
            ("api_panelantirruido", "subestacion_id"),
            ("api_tunel", "subestacion_id"),
            ("api_bienequipovehiculo", "subestacion_id"),
        ],
    },
    {
        "tabla": "api_patiosubestacion",
        "label": "Patio subestación",
        "hijos": [
            ("api_pano", "patio_subestacion_id"),
            ("api_barra", "patio_subestacion_id"),
            ("api_elementocomunpatiossee", "patio_subestacion_id"),
        ],
    },
]

# Contenedores con relación por nodo compartido: (tabla, label)
CONTENEDORES_NODO = [
    ("api_pano", "Paño"),
    ("api_casaserviciosgenerales", "Casa de servicios generales"),
    ("api_armario", "Armario"),
]

# Torres / Marcolineas sin OOCC: (tabla, label, tabla_relacion, columna_fk)
CONTENEDORES_SIN_OOCC = [
    {"tabla": "api_torre", "label": "Torre", "relacion": "api_torre_oocc", "fk": "torre_id"},
    {"tabla": "api_marcolinea", "label": "Marcolinea", "relacion": "api_marcolinea_oocc", "fk": "marcolinea_id"},
]

# Torres / Marcolineas sin accesorios estructurales: (tabla, label, instalacion_referencia)
CONTENEDORES_SIN_ACCESORIOS = [
    {"tabla": "api_torre", "label": "Torre", "instalacion": "api_torre"},
    {"tabla": "api_marcolinea", "label": "Marcolinea", "instalacion": "api_marcolinea"},
]

# Tablas que poseen columna nodo_id (espacio de búsqueda para nodo compartido).
# Se excluye tabla_unificada_nodos (tabla derivada).
TABLAS_NODO = [
    "api_aislador", "api_armario", "api_bancocondensador", "api_barra", "api_camara",
    "api_casaserviciosgenerales", "api_celda", "api_compensadoractivo",
    "api_compensadorestaticoreactivo", "api_condensadoracoplamiento", "api_condensadorserie",
    "api_condensadorsincrono", "api_conexiontierra", "api_desconectador",
    "api_dispositivoreconexion", "api_elementocomunssggcaseta", "api_estacionrepetidora",
    "api_estructura", "api_estructuraestacionrepetidora", "api_interruptor", "api_marcolinea",
    "api_medidor", "api_mufa", "api_pano", "api_pararrayo", "api_patiomufas", "api_reactor",
    "api_releproteccion", "api_tablero", "api_torre", "api_trampaonda",
    "api_transformadorcorriente", "api_transformadorpotencial", "api_transformadorssaa",
    "api_transformadorzigzag",
]

COLUMNAS_DETALLE = [
    "contenedor", "label", "observacion", "codigo", "mecanismo", "id", "nombre",
    "propietario_id", "propietario", "nodo_id",
]


def cargar_empresa_map(conn):
    cur = conn.cursor()
    cur.execute("SELECT id, name FROM api_empresa")
    rows = cur.fetchall()
    cur.close()
    return {r[0]: r[1] for r in rows}


def detectar_vacios_fk(conn, empresa_map):
    vacios = []
    totales = {}
    for cfg in CONTENEDORES_FK:
        tabla = cfg["tabla"]
        label = cfg["label"]
        df = pd.read_sql(
            f"SELECT id, name, propietario_id FROM {tabla} WHERE status='EN_OPERACION'",
            conn,
        )
        totales[tabla] = len(df)

        child_ids = set()
        for hijo, fk in cfg["hijos"]:
            ch = pd.read_sql(
                f"SELECT DISTINCT {fk} AS fk FROM {hijo} "
                f"WHERE status='EN_OPERACION' AND {fk} IS NOT NULL",
                conn,
            )
            child_ids.update(ch["fk"].tolist())

        for _, row in df.iterrows():
            if row["id"] not in child_ids:
                pid = row["propietario_id"]
                vacios.append({
                    "contenedor": tabla,
                    "label": label,
                    "observacion": "Contenedor vacío",
                    "codigo": "CONTENEDOR_VACIO",
                    "mecanismo": "FK directa",
                    "id": row["id"],
                    "nombre": row["name"] or "",
                    "propietario_id": "" if pid is None else str(pid),
                    "propietario": "" if pid is None else str(empresa_map.get(pid, "")),
                    "nodo_id": "",
                })
    return vacios, totales


def detectar_vacios_nodo(conn, empresa_map):
    sharing = defaultdict(set)
    for t in TABLAS_NODO:
        df = pd.read_sql(
            f"SELECT DISTINCT nodo_id FROM {t} "
            f"WHERE status='EN_OPERACION' AND nodo_id IS NOT NULL",
            conn,
        )
        for nid in df["nodo_id"].tolist():
            sharing[nid].add(t)

    vacios = []
    totales = {}
    for tabla, label in CONTENEDORES_NODO:
        df = pd.read_sql(
            f"SELECT id, name, nodo_id, propietario_id FROM {tabla} "
            f"WHERE status='EN_OPERACION'",
            conn,
        )
        totales[tabla] = len(df)
        for _, row in df.iterrows():
            nid = row["nodo_id"]
            pid = row["propietario_id"]
            others = (sharing.get(nid, set()) - {tabla}) if not pd.isna(nid) else set()
            if not others:
                vacios.append({
                    "contenedor": tabla,
                    "label": label,
                    "observacion": "Contenedor vacío",
                    "codigo": "CONTENEDOR_VACIO",
                    "mecanismo": "Nodo compartido",
                    "id": row["id"],
                    "nombre": row["name"] or "",
                    "propietario_id": "" if pid is None else str(pid),
                    "propietario": "" if pid is None else str(empresa_map.get(pid, "")),
                    "nodo_id": "" if pd.isna(nid) else str(int(nid)),
                })
    return vacios, totales


def detectar_sin_oocc(conn, empresa_map):
    vacios = []
    totales = {}
    for cfg in CONTENEDORES_SIN_OOCC:
        tabla = cfg["tabla"]
        label = cfg["label"]
        df = pd.read_sql(
            f"SELECT id, name, propietario_id FROM {tabla} WHERE status='EN_OPERACION'",
            conn,
        )
        totales[tabla] = len(df)

        rel = pd.read_sql(
            f"SELECT DISTINCT {cfg['fk']} AS fk FROM {cfg['relacion']}",
            conn,
        )
        child_ids = set(rel["fk"].tolist())

        for _, row in df.iterrows():
            if row["id"] not in child_ids:
                pid = row["propietario_id"]
                vacios.append({
                    "contenedor": tabla,
                    "label": label,
                    "observacion": "Sin OOCC",
                    "codigo": "TORRE_SIN_OOCC" if tabla == "api_torre" else "MARCO_SIN_OOCC",
                    "mecanismo": "FK directa",
                    "id": row["id"],
                    "nombre": row["name"] or "",
                    "propietario_id": "" if pid is None else str(pid),
                    "propietario": "" if pid is None else str(empresa_map.get(pid, "")),
                    "nodo_id": "",
                })
    return vacios, totales


def detectar_sin_accesorios(conn, empresa_map):
    vacios = []
    totales = {}
    for cfg in CONTENEDORES_SIN_ACCESORIOS:
        tabla = cfg["tabla"]
        label = cfg["label"]
        df = pd.read_sql(
            f"SELECT id, name, propietario_id FROM {tabla} WHERE status='EN_OPERACION'",
            conn,
        )
        totales[tabla] = len(df)

        acc = pd.read_sql(
            f'SELECT DISTINCT "nodo_estructura_id.id"::int AS nid '
            f'FROM api_accesorioestructuralineaaerea '
            f'WHERE status=\'EN_OPERACION\' '
            f'AND "nodo_estructura_id.instalacion" = \'{cfg["instalacion"]}\' '
            f'AND "nodo_estructura_id.id" IS NOT NULL',
            conn,
        )
        child_ids = set(acc["nid"].tolist())

        for _, row in df.iterrows():
            if row["id"] not in child_ids:
                pid = row["propietario_id"]
                vacios.append({
                    "contenedor": tabla,
                    "label": label,
                    "observacion": "Sin accesorios",
                    "codigo": "TORRE_SIN_ACCESORIOS" if tabla == "api_torre" else "MARCO_SIN_ACCESORIOS",
                    "mecanismo": "FK directa",
                    "id": row["id"],
                    "nombre": row["name"] or "",
                    "propietario_id": "" if pid is None else str(pid),
                    "propietario": "" if pid is None else str(empresa_map.get(pid, "")),
                    "nodo_id": "",
                })
    return vacios, totales


def exportar(vacios, totales):
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    df_vacios = pd.DataFrame(vacios, columns=COLUMNAS_DETALLE)

    # Resumen por contenedor y observación
    orden = (
        [(c["tabla"], c["label"], "Contenedor vacío") for c in CONTENEDORES_FK]
        + [(c[0], c[1], "Contenedor vacío") for c in CONTENEDORES_NODO]
        + [(c["tabla"], c["label"], "Sin OOCC") for c in CONTENEDORES_SIN_OOCC]
        + [(c["tabla"], c["label"], "Sin accesorios") for c in CONTENEDORES_SIN_ACCESORIOS]
    )
    conteo = (
        df_vacios.groupby(["contenedor", "observacion"]).size().to_dict()
        if not df_vacios.empty else {}
    )
    filas_resumen = []
    for tabla, label, obs in orden:
        filas_resumen.append({
            "contenedor": tabla,
            "descripcion": label,
            "observacion": obs,
            "total_en_operacion": totales.get(tabla, 0),
            "vacios": int(conteo.get((tabla, obs), 0)),
        })
    df_resumen = pd.DataFrame(filas_resumen)

    # Por propietario
    prop = defaultdict(int)
    for v in vacios:
        key = (v["propietario_id"], v["propietario"])
        prop[key] += 1
    por_propietario = []
    for (pid, pname), n in sorted(prop.items(), key=lambda kv: -kv[1]):
        por_propietario.append({
            "propietario_id": pid,
            "propietario": pname,
            "contenedores_vacios": n,
        })
    df_propietario = pd.DataFrame(por_propietario)

    os.makedirs(REPORTE_DIR, exist_ok=True)
    os.makedirs(HISTORICO_DIR, exist_ok=True)

    base = "auditoria_contenedores_metricas"
    xlsx_reporte = os.path.join(REPORTE_DIR, f"{base}.xlsx")
    xlsx_historico = os.path.join(HISTORICO_DIR, f"{base}-{ts}{config.sufijo_historico_bd()}.xlsx")
    csv_reporte = os.path.join(REPORTE_DIR, f"{base}_detalle.csv")
    csv_historico = os.path.join(HISTORICO_DIR, f"{base}_detalle-{ts}{config.sufijo_historico_bd()}.csv")

    for ruta in (xlsx_reporte, xlsx_historico):
        with pd.ExcelWriter(ruta, engine="xlsxwriter") as w:
            df_resumen.to_excel(w, sheet_name="Resumen", index=False)
            df_propietario.to_excel(w, sheet_name="Por_Propietario", index=False)
            df_vacios.to_excel(w, sheet_name="Detalle", index=False)

    for ruta in (csv_reporte, csv_historico):
        with open(ruta, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(COLUMNAS_DETALLE)
            for v in vacios:
                writer.writerow([v[c] for c in COLUMNAS_DETALLE])

    print("Resumen por contenedor:")
    print(df_resumen.to_string(index=False))
    print()
    print(f"Total hallazgos: {len(vacios)}")
    print(f"Excel reporte: {xlsx_reporte}")
    print(f"CSV reporte:   {csv_reporte}")


def main():
    conn = psycopg2.connect(**DB_CONN)
    try:
        empresa_map = cargar_empresa_map(conn)
        vacios_fk, totales_fk = detectar_vacios_fk(conn, empresa_map)
        vacios_nodo, totales_nodo = detectar_vacios_nodo(conn, empresa_map)
        vacios_oocc, totales_oocc = detectar_sin_oocc(conn, empresa_map)
        vacios_acc, totales_acc = detectar_sin_accesorios(conn, empresa_map)
    finally:
        conn.close()

    vacios = vacios_fk + vacios_nodo + vacios_oocc + vacios_acc
    totales = {**totales_fk, **totales_nodo, **totales_oocc, **totales_acc}

    exportar(vacios, totales)


if __name__ == "__main__":
    main()
