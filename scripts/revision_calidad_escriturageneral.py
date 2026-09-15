# -*- coding: utf-8 -*-
"""
Revisión de calidad de escritura general.

Analiza la columna "name" de las tablas de instalación (con columnas
"name", "status" y "propietario_id") y detecta:

1. Dobles espacios.
2. Espacios al inicio o al final.
3. Uso incorrecto de "kV": si el nombre contiene la unidad, debe escribirse
   siempre "kV" (k minúscula, V mayúscula). Se detectan "KV", "kv" y "Kv".
4. Minúsculas fuera de la unidad tolerada: la única escritura en minúscula
   permitida es "kV"; el resto del nombre debe ir en mayúsculas (respetando
   tildes).
5. Guion con espacios (" - "): se detecta como error, salvo en líneas,
   circuitos y tramos, donde " - " es el separador válido entre extremos.
6. Guion largo (— o –).
7. Separador " - " obligatorio en líneas, circuitos y tramos (NOMBRE_SIN_SEPARADOR).

Salidas:
    docs/reporte/revision_calidad_escriturageneral.xlsx
    docs/reporte/revision_calidad_escriturageneral_detalle.csv
    docs/historico/revision_calidad_escriturageneral-<YYYY-MM-DD_HH-MM-SS>.xlsx
    docs/historico/revision_calidad_escriturageneral_detalle-<YYYY-MM-DD_HH-MM-SS>.csv
"""

import os
import re
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

TABLAS_EXCLUIDAS = {"tabla_unificada_nodos"}

TABLAS_SEPARADOR_GUION_VALIDO = {"api_linea", "api_circuito", "api_tramo"}

TABLAS_ABREV_TB = {"api_transformadorcorriente"}
RE_ABREV_TB = re.compile(r"(?i)[TB][\s.]*[TB]\s*\.?")

COLUMNAS_DETALLE = [
    "tabla", "codigo", "id", "nombre", "tipo", "detalle", "propietario_id", "propietario",
]

TIPOS = {
    "doble_espacio": "Doble espacio",
    "espacios_extremos": "Espacio al inicio o final",
    "kv_incorrecto": "kV mal escrito",
    "minusculas": "Nombre con minúsculas",
    "guion_con_espacios": "Guion con espacios (' - ')",
    "guion_largo": "Guion largo (— o –)",
    "sin_separador": "Sin separador (' - ')",
}

CODIGOS = {
    "doble_espacio": "NOMBRE_DOBLE_ESPACIO",
    "espacios_extremos": "NOMBRE_ESPACIOS_EXTREMOS",
    "kv_incorrecto": "KV_MAL_ESCRITO",
    "minusculas": "NOMBRE_MINUSCULAS",
    "guion_con_espacios": "NOMBRE_GUION_CON_ESPACIOS",
    "guion_largo": "NOMBRE_GUION_LARGO",
    "sin_separador": "NOMBRE_SIN_SEPARADOR",
}


def _kv_wrong(name):
    """Devuelve la primera variante incorrecta de kV, o None si es correcto."""
    if not name:
        return None
    for m in re.finditer(r"[Kk][Vv]", name):
        if m.group() != "kV":
            return m.group()
    return None


def _tiene_minusculas(name):
    if not name:
        return False
    test = str(name).replace("kV", "KV")
    letters = "".join(ch for ch in test if ch.isalpha())
    return letters != letters.upper()


def _tiene_abrev_tb(name):
    if not name:
        return False
    return any("." in m.group() for m in RE_ABREV_TB.finditer(str(name)))


def obtener_tablas_evaluables(conn):
    cur = conn.cursor()
    cur.execute(
        "SELECT t.table_name FROM information_schema.tables t "
        "WHERE t.table_schema='public' AND t.table_type='BASE TABLE' "
        "AND EXISTS (SELECT 1 FROM information_schema.columns c "
        "            WHERE c.table_schema='public' AND c.table_name=t.table_name AND c.column_name='name') "
        "AND EXISTS (SELECT 1 FROM information_schema.columns c "
        "            WHERE c.table_schema='public' AND c.table_name=t.table_name AND c.column_name='status') "
        "AND EXISTS (SELECT 1 FROM information_schema.columns c "
        "            WHERE c.table_schema='public' AND c.table_name=t.table_name AND c.column_name='propietario_id') "
        "ORDER BY t.table_name"
    )
    tables = [r[0] for r in cur.fetchall()]
    cur.close()
    return [t for t in tables if t not in TABLAS_EXCLUIDAS]


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

    findings = []
    resumen_por_tabla = []

    tablas = obtener_tablas_evaluables(conn)

    for tabla in tablas:
        cols = columnas_de(conn, tabla)
        has_id = "id" in cols
        has_prop = "propietario_id" in cols
        has_status = "status" in cols

        if not has_id:
            continue

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

        # 1. Dobles espacios
        doble = df["name"].str.contains("  ", regex=False)
        # 2. Espacios extremos
        extremos = df["name"] != df["name"].str.strip()
        # 3. kV mal escrito
        kv = df["name"].map(_kv_wrong).notna()
        # 4. Guion con espacios (" - "), salvo separador válido en líneas/circuitos/tramos
        permitir_guion_separador = tabla in TABLAS_SEPARADOR_GUION_VALIDO
        guion_esp = (
            pd.Series(False, index=df.index)
            if permitir_guion_separador
            else df["name"].str.contains(" - ", regex=False)
        )
        # 5. Guion largo (em dash / en dash)
        guion_largo = df["name"].str.contains("\u2014|\u2013", regex=True)
        # 6. Minúsculas (solo se tolera la unidad "kV"; se excluye la
        #    abreviatura "t.b." en transformadores de corriente)
        minusculas = df["name"].map(_tiene_minusculas)
        if tabla in TABLAS_ABREV_TB:
            minusculas = minusculas & ~df["name"].map(_tiene_abrev_tb)
        # 7. Separador " - " obligatorio en líneas/circuitos/tramos
        sin_separador = (
            ~df["name"].str.contains(" - ", regex=False)
            if permitir_guion_separador
            else pd.Series(False, index=df.index)
        )

        n_doble = int(doble.sum())
        n_extremos = int(extremos.sum())
        n_kv = int(kv.sum())
        n_guion_esp = int(guion_esp.sum())
        n_guion_largo = int(guion_largo.sum())
        n_minusculas = int(minusculas.sum())
        n_sin_separador = int(sin_separador.sum())

        if n_doble + n_extremos + n_kv + n_guion_esp + n_guion_largo + n_minusculas + n_sin_separador == 0:
            continue

        for _, row in df.iterrows():
            nombre = row["name"]
            pid = row["propietario_id"] if has_prop else None
            pname = "" if (pid is None or pd.isna(pid)) else str(empresa_map.get(pid, ""))

            if "  " in nombre:
                findings.append({
                    "tabla": tabla,
                    "codigo": CODIGOS["doble_espacio"],
                    "id": row["id"],
                    "nombre": nombre,
                    "tipo": TIPOS["doble_espacio"],
                    "detalle": "El nombre contiene dobles espacios.",
                    "propietario_id": "" if (pid is None or pd.isna(pid)) else str(pid),
                    "propietario": pname,
                })
            if nombre != nombre.strip():
                findings.append({
                    "tabla": tabla,
                    "codigo": CODIGOS["espacios_extremos"],
                    "id": row["id"],
                    "nombre": nombre,
                    "tipo": TIPOS["espacios_extremos"],
                    "detalle": "El nombre tiene espacios al inicio o al final.",
                    "propietario_id": "" if (pid is None or pd.isna(pid)) else str(pid),
                    "propietario": pname,
                })
            wrong = _kv_wrong(nombre)
            if wrong:
                findings.append({
                    "tabla": tabla,
                    "codigo": CODIGOS["kv_incorrecto"],
                    "id": row["id"],
                    "nombre": nombre,
                    "tipo": TIPOS["kv_incorrecto"],
                    "detalle": f'Se encontró "{wrong}" en lugar de "kV".',
                    "propietario_id": "" if (pid is None or pd.isna(pid)) else str(pid),
                    "propietario": pname,
                })
            if _tiene_minusculas(nombre) and not (
                tabla in TABLAS_ABREV_TB and _tiene_abrev_tb(nombre)
            ):
                findings.append({
                    "tabla": tabla,
                    "codigo": CODIGOS["minusculas"],
                    "id": row["id"],
                    "nombre": nombre,
                    "tipo": TIPOS["minusculas"],
                    "detalle": "El nombre contiene minúsculas fuera de la unidad 'kV'.",
                    "propietario_id": "" if (pid is None or pd.isna(pid)) else str(pid),
                    "propietario": pname,
                })
            if not permitir_guion_separador and " - " in nombre:
                findings.append({
                    "tabla": tabla,
                    "codigo": CODIGOS["guion_con_espacios"],
                    "id": row["id"],
                    "nombre": nombre,
                    "tipo": TIPOS["guion_con_espacios"],
                    "detalle": 'El guion tiene espacios alrededor (" - "); debe ser "-".',
                    "propietario_id": "" if (pid is None or pd.isna(pid)) else str(pid),
                    "propietario": pname,
                })
            if "\u2014" in nombre or "\u2013" in nombre:
                findings.append({
                    "tabla": tabla,
                    "codigo": CODIGOS["guion_largo"],
                    "id": row["id"],
                    "nombre": nombre,
                    "tipo": TIPOS["guion_largo"],
                    "detalle": "Se utiliza un guion largo (— o –), que es incorrecto.",
                    "propietario_id": "" if (pid is None or pd.isna(pid)) else str(pid),
                    "propietario": pname,
                })
            if permitir_guion_separador and " - " not in nombre:
                findings.append({
                    "tabla": tabla,
                    "codigo": CODIGOS["sin_separador"],
                    "id": row["id"],
                    "nombre": nombre,
                    "tipo": TIPOS["sin_separador"],
                    "detalle": 'El nombre no contiene el separador " - " entre extremos.',
                    "propietario_id": "" if (pid is None or pd.isna(pid)) else str(pid),
                    "propietario": pname,
                })

        resumen_por_tabla.append({
            "tabla": tabla,
            "total_revisados": len(df),
            "dobles_espacios": n_doble,
            "espacios_extremos": n_extremos,
            "kv_incorrecto": n_kv,
            "minusculas": n_minusculas,
            "guion_con_espacios": n_guion_esp,
            "guion_largo": n_guion_largo,
            "sin_separador": n_sin_separador,
            "total": n_doble + n_extremos + n_kv + n_minusculas + n_guion_esp + n_guion_largo + n_sin_separador,
        })

    return findings, resumen_por_tabla


def exportar(findings, resumen_por_tabla):
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    df_findings = pd.DataFrame(findings, columns=COLUMNAS_DETALLE)
    df_resumen = pd.DataFrame(resumen_por_tabla)

    # Por tipo
    if df_findings.empty:
        df_tipo = pd.DataFrame(columns=["tipo", "cantidad"])
    else:
        df_tipo = (
            df_findings.groupby("tipo").size().reset_index(name="cantidad")
            .sort_values("cantidad", ascending=False)
        )

    # Por propietario
    prop = defaultdict(int)
    for f in findings:
        key = (f["propietario_id"], f["propietario"])
        prop[key] += 1
    por_propietario = []
    for (pid, pname), n in sorted(prop.items(), key=lambda kv: -kv[1]):
        por_propietario.append({
            "propietario_id": pid,
            "propietario": pname,
            "total_hallazgos": n,
        })
    df_propietario = pd.DataFrame(por_propietario)

    os.makedirs(REPORTE_DIR, exist_ok=True)
    os.makedirs(HISTORICO_DIR, exist_ok=True)

    base = "revision_calidad_escriturageneral"
    xlsx_reporte = os.path.join(REPORTE_DIR, f"{base}.xlsx")
    xlsx_historico = os.path.join(HISTORICO_DIR, f"{base}-{ts}{config.sufijo_historico_bd()}.xlsx")
    csv_reporte = os.path.join(REPORTE_DIR, f"{base}_detalle.csv")
    csv_historico = os.path.join(HISTORICO_DIR, f"{base}_detalle-{ts}{config.sufijo_historico_bd()}.csv")

    for ruta in (xlsx_reporte, xlsx_historico):
        with pd.ExcelWriter(ruta, engine="xlsxwriter") as w:
            df_resumen.to_excel(w, sheet_name="Resumen", index=False)
            df_tipo.to_excel(w, sheet_name="Por_Tipo", index=False)
            df_propietario.to_excel(w, sheet_name="Por_Propietario", index=False)
            df_findings.to_excel(w, sheet_name="Detalle", index=False)

    for ruta in (csv_reporte, csv_historico):
        with open(ruta, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(COLUMNAS_DETALLE)
            for item in findings:
                writer.writerow([item[c] for c in COLUMNAS_DETALLE])

    print(f"Total hallazgos de escritura: {len(findings)}")
    print(f"Tablas con hallazgos: {len(df_resumen)}")
    if not df_resumen.empty:
        print(df_resumen.sort_values("total", ascending=False).head(15).to_string(index=False))
    print(f"Excel reporte: {xlsx_reporte}")
    print(f"CSV reporte:   {csv_reporte}")


def main():
    conn = psycopg2.connect(**DB_CONN)
    try:
        findings, resumen_por_tabla = analizar(conn)
    finally:
        conn.close()

    exportar(findings, resumen_por_tabla)


if __name__ == "__main__":
    main()
