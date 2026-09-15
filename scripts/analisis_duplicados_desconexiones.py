# -*- coding: utf-8 -*-
import os
import sys
import json
import psycopg2
import pandas as pd
from psycopg2 import sql
from psycopg2 import OperationalError
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.dirname(SCRIPT_DIR)
sys.path.insert(0, BASE_DIR)

import config

JSON_PATH = os.path.join(BASE_DIR, "utils", "relacionamiento_columnas_instalaciones.json")

REPORTE_DIR = config.REPORTE_DIR
HISTORICO_DIR = config.HISTORICO_DIR
DB_CONN = config.DB_CONN

STATUS_EN_OPERACION = "EN_OPERACION"
STATUS_FUERA_OPERACION = "FUERA_OPERACION"


def obtener_columnas(cur, tabla):
    cur.execute(
        "SELECT column_name FROM information_schema.columns "
        "WHERE table_schema = 'public' AND table_name = %s",
        (tabla,),
    )
    return [r[0] for r in cur.fetchall()]


def creacion_columna_instalaciones(DB_CONN):
    conn = psycopg2.connect(**DB_CONN)
    cur = conn.cursor()
    cur.execute(
        "SELECT table_name FROM information_schema.tables "
        "WHERE table_schema = 'public' AND table_type = 'BASE TABLE'"
    )
    tables = [r[0] for r in cur.fetchall()]

    for table_name in tables:
        cur.execute(sql.SQL("ALTER TABLE {}.{} ADD COLUMN IF NOT EXISTS instalacion TEXT").format(
            sql.Identifier("public"), sql.Identifier(table_name)
        ))
        cur.execute(sql.SQL("UPDATE {}.{} SET instalacion = %s").format(
            sql.Identifier("public"), sql.Identifier(table_name)
        ), (table_name,))

    conn.commit()
    cur.close()
    conn.close()


def creacion_columna_empresa_name(DB_CONN):
    conn = psycopg2.connect(**DB_CONN)
    cur = conn.cursor()
    cur.execute(
        "SELECT table_name FROM information_schema.tables "
        "WHERE table_schema = 'public' AND table_type = 'BASE TABLE'"
    )
    tables = [r[0] for r in cur.fetchall()]

    for table_name in tables:
        cols = obtener_columnas(cur, table_name)
        if "propietario_id" not in cols:
            continue
        cur.execute(sql.SQL("ALTER TABLE {}.{} ADD COLUMN IF NOT EXISTS empresa_name TEXT").format(
            sql.Identifier("public"), sql.Identifier(table_name)
        ))
        cur.execute(sql.SQL(
            "UPDATE {}.{} t SET empresa_name = e.name "
            "FROM api_empresa e WHERE t.propietario_id = e.id"
        ).format(sql.Identifier("public"), sql.Identifier(table_name)))

    conn.commit()
    cur.close()
    conn.close()


def _cargar_relacionamiento_json():
    if not os.path.exists(JSON_PATH):
        raise FileNotFoundError(f"No se encontro el archivo en:\n{JSON_PATH}")
    with open(JSON_PATH, "r", encoding="utf-8") as f:
        data = json.load(f)
    nodos_instalaciones = list(data["nodos_instalaciones"])
    tablas_analizar = pd.DataFrame(data["tablas_analizar"])
    diccionario_tabla_relacion = pd.DataFrame(data["diccionario_tabla_relacion"])
    return nodos_instalaciones, tablas_analizar, diccionario_tabla_relacion


def creacion_tabla_nodos(DB_CONN, tablas_a_unir):
    conn = psycopg2.connect(**DB_CONN)
    cur = conn.cursor()

    nueva_tabla = "tabla_unificada_nodos"
    cur.execute(sql.SQL("DROP TABLE IF EXISTS {};").format(sql.Identifier(nueva_tabla)))
    cur.execute(sql.SQL("""
        CREATE TABLE {} (
            id INTEGER,
            name VARCHAR,
            status VARCHAR,
            propietario_id TEXT,
            empresa_name TEXT,
            instalacion VARCHAR,
            nodo_id INTEGER,
            description TEXT
        );
    """).format(sql.Identifier(nueva_tabla)))

    columnas_objetivo = ["id", "name", "status", "propietario_id", "empresa_name", "instalacion", "nodo_id", "description"]

    for tabla in tablas_a_unir:
        cols = obtener_columnas(cur, tabla)
        exprs = [
            sql.Identifier(c) if c in cols else sql.SQL("NULL")
            for c in columnas_objetivo
        ]
        cur.execute(sql.SQL("INSERT INTO {} SELECT {} FROM {};").format(
            sql.Identifier(nueva_tabla),
            sql.SQL(", ").join(exprs),
            sql.Identifier(tabla),
        ))

    conn.commit()
    cur.close()
    conn.close()


def generacion_relacionamiento_bdatx(tablas_analizar, diccionario_tabla_relacion):
    conn = psycopg2.connect(**DB_CONN)
    cur = conn.cursor()

    for _, row in tablas_analizar.iterrows():
        tabla_postgres = row["tablas_postgres"]
        conexiones = [col for col in row.index if col.startswith("conexion") and pd.notna(row[col])]

        for conexion in conexiones:
            columna_conexion = row[conexion]
            relacion = diccionario_tabla_relacion[
                diccionario_tabla_relacion["conexion"] == columna_conexion
            ]

            if relacion.empty:
                print(f"No se encontro relacion para '{columna_conexion}' en '{tabla_postgres}'.")
                continue

            tabla_relacion = relacion["tabla_relacion"].iloc[0]
            col_relacion = relacion["col_relacion"].iloc[0]

            rel_cols = obtener_columnas(cur, tabla_relacion)
            sufijos = [s for s in ("name", "status", "id", "instalacion", "description", "empresa_name") if s in rel_cols]

            for s in sufijos:
                col = f"{columna_conexion}.{s}"
                cur.execute(sql.SQL("ALTER TABLE {}.{} ADD COLUMN IF NOT EXISTS {} TEXT").format(
                    sql.Identifier("public"), sql.Identifier(tabla_postgres), sql.Identifier(col)
                ))

            set_clauses = ",\n    ".join([f'"{columna_conexion}.{s}" = rel.{s}' for s in sufijos])
            sql_update = f"""
                UPDATE "{tabla_postgres}" AS tgt
                SET
                    {set_clauses}
                FROM "{tabla_relacion}" AS rel
                WHERE tgt."{columna_conexion}" = rel."{col_relacion}";
            """
            cur.execute(sql_update)
            conn.commit()
            print(f"[OK] '{tabla_postgres}' actualizada desde '{tabla_relacion}' ({len(sufijos)} columnas).")

    conn.commit()
    cur.close()
    conn.close()


def _cargar_en_operacion(conn, tabla_postgres):
    cols = obtener_columnas(conn.cursor(), tabla_postgres)
    columnas_status = [c for c in cols if "status" in c.lower()]
    col_exacta = next((c for c in columnas_status if c.lower() == "status"), None)
    if not columnas_status or col_exacta is None:
        return None, None, None
    df = pd.read_sql(
        f'SELECT * FROM "{tabla_postgres}" WHERE "{col_exacta}" = %s',
        conn,
        params=(STATUS_EN_OPERACION,),
    )
    otras = [c for c in columnas_status if c.lower() != "status"]
    return df, col_exacta, otras


def _resumen_por_empresa(filtrados, resultados, col_registros):
    totales = []
    for df in filtrados.values():
        if df is None or df.empty or "empresa_name" not in df.columns:
            continue
        totales.append(df[["empresa_name"]])
    if totales:
        df_tot = pd.concat(totales, ignore_index=True)
        df_tot["empresa_name"] = df_tot["empresa_name"].fillna("Sin empresa")
        res_tot = df_tot.groupby("empresa_name").size().reset_index(name="total_registros")
    else:
        res_tot = pd.DataFrame(columns=["empresa_name", "total_registros"])

    problemas = []
    for tabla, df in resultados.items():
        if df is None or df.empty or "empresa_name" not in df.columns:
            continue
        tmp = df[["empresa_name"]].copy()
        tmp["tabla"] = tabla.replace("api_", "")
        tmp["empresa_name"] = tmp["empresa_name"].fillna("Sin empresa")
        problemas.append(tmp)

    if problemas:
        df_prob = pd.concat(problemas, ignore_index=True)
        conteo = df_prob.groupby(["empresa_name", "tabla"]).size().reset_index(name=col_registros)
        matriz = conteo.pivot(index="empresa_name", columns="tabla", values=col_registros).fillna(0).astype(int)
        matriz["Total"] = matriz.sum(axis=1)
        matriz = matriz.reset_index()
    else:
        matriz = pd.DataFrame(columns=["empresa_name", "Total"])

    resumen = res_tot.merge(matriz, on="empresa_name", how="outer")
    resumen = resumen.fillna(0)
    for c in ("total_registros", "Total"):
        if c in resumen.columns:
            resumen[c] = resumen[c].astype(int)
    resumen["porcentaje error"] = resumen.apply(
        lambda r: round(r["Total"] / r["total_registros"] * 100, 2)
        if r["total_registros"] else 0,
        axis=1,
    )
    resumen = resumen.sort_values("Total", ascending=False).reset_index(drop=True)
    return resumen


def _exportar(resultados, cantidades, resumen_empresa, nombre_base):
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    os.makedirs(REPORTE_DIR, exist_ok=True)
    os.makedirs(HISTORICO_DIR, exist_ok=True)
    output_files = [
        os.path.join(REPORTE_DIR, f"{nombre_base}.xlsx"),
        os.path.join(HISTORICO_DIR, f"{nombre_base}-{timestamp}{config.sufijo_historico_bd()}.xlsx"),
    ]
    for output_file in output_files:
        with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
            cantidades.to_excel(writer, sheet_name="analisis", index=False)
            resumen_empresa.to_excel(writer, sheet_name="analisis_por_empresa", index=False)
            for tabla, df in resultados.items():
                df.to_excel(writer, sheet_name=tabla[:31], index=False)


def reporte_elementos_desconectados(DB_CONN, tablas_analizar):
    conn = psycopg2.connect(**DB_CONN)
    resultados = {}
    filtrados = {}
    filas = []

    for _, row in tablas_analizar.iterrows():
        tabla_postgres = row["tablas_postgres"]
        try:
            df_filtrado, col_exacta, otras = _cargar_en_operacion(conn, tabla_postgres)
        except Exception as e:
            print(f"Error al traer la tabla '{tabla_postgres}': {e}")
            continue

        if df_filtrado is None or df_filtrado.empty:
            continue

        if otras:
            mascara = df_filtrado[otras].apply(lambda x: x == STATUS_FUERA_OPERACION).any(axis=1)
            df_final = df_filtrado[mascara]
        else:
            df_final = df_filtrado

        resultados[tabla_postgres] = df_final
        filtrados[tabla_postgres] = df_filtrado
        filas.append({
            "tabla": tabla_postgres,
            "total_registros": len(df_filtrado),
            "registros_desconectados": len(df_final),
        })

    conn.close()
    cantidades = pd.DataFrame(filas)
    cantidades["tabla"] = cantidades["tabla"].str.replace("api_", "", regex=False)
    resumen_empresa = _resumen_por_empresa(filtrados, resultados, "registros_desconectados")
    _exportar(resultados, cantidades, resumen_empresa, "analisis_desconectados")


def reporte_elementos_duplicados(DB_CONN, tablas_analizar):
    conn = psycopg2.connect(**DB_CONN)
    resultados = {}
    filtrados = {}
    filas = []

    for _, row in tablas_analizar.iterrows():
        tabla_postgres = row["tablas_postgres"]
        try:
            df_filtrado, _, _ = _cargar_en_operacion(conn, tabla_postgres)
        except Exception as e:
            print(f"Error al traer la tabla '{tabla_postgres}': {e}")
            continue

        if df_filtrado is None or df_filtrado.empty:
            continue

        if "name" not in df_filtrado.columns:
            print(f"[WARN] '{tabla_postgres}' no tiene columna 'name'. Se omite duplicados.")
            continue

        if tabla_postgres == "api_tramo" and "circuito_id" in df_filtrado.columns:
            df_final = df_filtrado[df_filtrado.duplicated(subset=["name", "circuito_id"], keep=False)]
        else:
            df_final = df_filtrado[df_filtrado.duplicated(subset=["name"], keep=False)]
        resultados[tabla_postgres] = df_final
        filtrados[tabla_postgres] = df_filtrado
        filas.append({
            "tabla": tabla_postgres,
            "total_registros": len(df_filtrado),
            "registros_duplicados": len(df_final),
        })

    conn.close()
    cantidades = pd.DataFrame(filas)
    cantidades["tabla"] = cantidades["tabla"].str.replace("api_", "", regex=False)
    resumen_empresa = _resumen_por_empresa(filtrados, resultados, "registros_duplicados")
    _exportar(resultados, cantidades, resumen_empresa, "analisis_duplicados")


def reporte_elementos_conexion_nula(DB_CONN, tablas_analizar):
    conn = psycopg2.connect(**DB_CONN)
    resultados = {}
    filtrados = {}
    filas = []

    for _, row in tablas_analizar.iterrows():
        tabla_postgres = row["tablas_postgres"]
        columnas_conexion = [row[c] for c in row.index if c.startswith("conexion") and pd.notna(row[c])]

        try:
            df_filtrado, _, _ = _cargar_en_operacion(conn, tabla_postgres)
        except Exception as e:
            print(f"Error al traer la tabla '{tabla_postgres}': {e}")
            continue

        if df_filtrado is None or df_filtrado.empty:
            continue

        columnas_conexion = [c for c in columnas_conexion if c in df_filtrado.columns]
        if columnas_conexion:
            mascara = df_filtrado[columnas_conexion].isnull().any(axis=1)
            df_final = df_filtrado[mascara]
        else:
            df_final = df_filtrado

        resultados[tabla_postgres] = df_final
        filtrados[tabla_postgres] = df_filtrado
        filas.append({
            "tabla": tabla_postgres,
            "total_registros": len(df_filtrado),
            "registros_desconectados": len(df_final),
        })

    conn.close()
    cantidades = pd.DataFrame(filas)
    cantidades["tabla"] = cantidades["tabla"].str.replace("api_", "", regex=False)
    resumen_empresa = _resumen_por_empresa(filtrados, resultados, "registros_conexion_nula")
    _exportar(resultados, cantidades, resumen_empresa, "analisis_conexion_nula")


def main():
    nodos_instalaciones, tablas_analizar, diccionario_tabla_relacion = _cargar_relacionamiento_json()

    print("Iniciando proceso:")
    creacion_columna_instalaciones(DB_CONN)
    print("1- Listo")
    creacion_columna_empresa_name(DB_CONN)
    print("2- Listo")
    creacion_tabla_nodos(DB_CONN, nodos_instalaciones)
    print("3- Listo")
    generacion_relacionamiento_bdatx(tablas_analizar, diccionario_tabla_relacion)
    print("4- Listo")
    reporte_elementos_desconectados(DB_CONN, tablas_analizar)
    print("5- Listo")
    reporte_elementos_duplicados(DB_CONN, tablas_analizar)
    print("6- Listo")
    reporte_elementos_conexion_nula(DB_CONN, tablas_analizar)
    print("7- Listo")


if __name__ == "__main__":
    main()
