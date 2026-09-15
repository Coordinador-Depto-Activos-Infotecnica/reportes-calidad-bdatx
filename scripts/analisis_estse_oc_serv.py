# -*- coding: utf-8 -*-
import os
import sys
import psycopg2
import pandas as pd
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE_DIR)

import config

DB_NAME = config.DB_NAME
PG_CONFIG = config.PG_CONFIG
REPORTE_DIR = config.REPORTE_DIR
HISTORICO_DIR = config.HISTORICO_DIR

STATUS_EN_OPERACION = "EN_OPERACION"

ENTIDADES = {
    "oocc": {
        "sufijo_tabla": "oocc",
        "fk_col": "oocc_id",
        "tabla_relacion": "api_oocc_relacion",
        "tabla_maestra": "api_oocc",
        "sufijo": "oocc",
        "operativo_col": "principal_operativo",
        "estado_col": "estado_principal",
        "texto_sin": "Sin principal operativo",
        "texto_con": "Con principal operativo",
        "sheets": ["api_oocc_relacion", "oocc_no_relacionados", "oocc_relacionados", "api_oocc"],
        "archivo": "api_oocc_relacion_enriquecida",
        "archivo_no_rel": "oocc_no_relacionados",
    },
    "estse": {
        "sufijo_tabla": "estructura_ssee",
        "fk_col": "estructurasubestacion_id",
        "tabla_relacion": "api_estructura_ssee_relacion",
        "tabla_maestra": "api_estructurasubestacion",
        "sufijo": "estse",
        "operativo_col": "dependiente_operativo",
        "estado_col": "estado_dependiente",
        "texto_sin": "Sin dependiente operativo",
        "texto_con": "Con dependiente operativo",
        "sheets": ["api_est_se_relacion", "est_se_no_relacionados", "est_se_relacionados", "api_estse"],
        "archivo": "api_est_se_relacion_enriquecida",
        "archivo_no_rel": "est_se_no_relacionados",
    },
    "serv": {
        "sufijo_tabla": "servidumbre",
        "fk_col": "servidumbre_id",
        "tabla_relacion": "api_servidumbre_relacion",
        "tabla_maestra": "api_servidumbre",
        "sufijo": "servidumbre",
        "operativo_col": "principal_operativo",
        "estado_col": "estado_principal",
        "texto_sin": "Sin principal operativo",
        "texto_con": "Con principal operativo",
        "sheets": ["api_serv_relacion", "serv_no_relacionados", "serv_relacionados", "api_serv"],
        "archivo": "api_serv_relacion_enriquecida",
        "archivo_no_rel": "serv_no_relacionados",
    },
}


def obtener_tablas_publicas(cursor):
    cursor.execute(
        "SELECT table_name FROM information_schema.tables "
        "WHERE table_schema = 'public' AND table_type = 'BASE TABLE'"
    )
    return [r[0] for r in cursor.fetchall()]


def obtener_columnas(cursor, tabla):
    cursor.execute(
        "SELECT column_name FROM information_schema.columns "
        "WHERE table_schema = 'public' AND table_name = %s",
        (tabla,),
    )
    return [r[0] for r in cursor.fetchall()]


def extraer_instalacion(tabla, sufijo_tabla):
    nombre = tabla[len("api_"):]
    if nombre.endswith("_" + sufijo_tabla):
        nombre = nombre[: -(len(sufijo_tabla) + 1)]
    return nombre


def construir_relaciones(conn, cfg):
    cursor = conn.cursor()
    sufijo = cfg["sufijo_tabla"]
    tabla_relacion = cfg["tabla_relacion"]
    tabla_maestra = cfg["tabla_maestra"]

    tablas = obtener_tablas_publicas(cursor)

    for t in tablas:
        if t.endswith("_" + sufijo + "_relacion") and t != tabla_relacion:
            cursor.execute(f'DROP TABLE IF EXISTS "{t}" CASCADE')

    per_inst = [
        t for t in tablas
        if t.startswith("api_")
        and t.endswith("_" + sufijo)
        and t != tabla_maestra
        and t != tabla_relacion
    ]

    for t in per_inst:
        instalacion = extraer_instalacion(t, sufijo)
        instalacion_table = "api_" + instalacion
        relacion_table = f"api_{instalacion}_{sufijo}_relacion"

        if instalacion_table not in tablas:
            print(f"La tabla {instalacion_table} no existe. Saltando {t}...")
            continue

        cols = obtener_columnas(cursor, t)
        source_cols = [c for c in cols if c != "instalacion"]
        select_list = ", ".join(f'"{c}"' for c in source_cols)
        cursor.execute(
            f"CREATE TABLE \"{relacion_table}\" AS "
            f"SELECT {select_list}, '{instalacion_table}' AS instalacion FROM \"{t}\""
        )
        print(f"Tabla {relacion_table} creada correctamente.")

    cursor.execute(f'DROP TABLE IF EXISTS "{tabla_relacion}"')
    cursor.execute(f"""
        CREATE TABLE "{tabla_relacion}" (
            id SERIAL PRIMARY KEY,
            instalacion_id INTEGER,
            {cfg["fk_col"]} INTEGER,
            instalacion TEXT
        )
    """)

    rel_tablas = [
        t for t in obtener_tablas_publicas(cursor)
        if t.endswith("_" + sufijo + "_relacion") and t != tabla_relacion
    ]

    for t in rel_tablas:
        cols = obtener_columnas(cursor, t)
        inst_id_col = next(
            (c for c in cols if c.endswith("_id") and c not in ("id", cfg["fk_col"])),
            None,
        )
        if inst_id_col is None:
            print(f"[WARN] No se encontro columna {{instalacion}}_id en {t}. Se omite.")
            continue
        cursor.execute(
            f'INSERT INTO "{tabla_relacion}" (instalacion_id, {cfg["fk_col"]}, instalacion) '
            f'SELECT {inst_id_col}, {cfg["fk_col"]}, instalacion FROM "{t}"'
        )
        print(f"Datos de {t} insertados en {tabla_relacion}.")

    conn.commit()
    cursor.close()


def enriquecer(conn, cfg):
    cursor = conn.cursor()
    sufijo = cfg["sufijo"]
    fk_col = cfg["fk_col"]
    tabla_relacion = cfg["tabla_relacion"]
    tabla_maestra = cfg["tabla_maestra"]

    columnas_cache = {}

    def columnas_de(tabla):
        if tabla not in columnas_cache:
            columnas_cache[tabla] = obtener_columnas(cursor, tabla)
        return columnas_cache[tabla]

    df_relacion = pd.read_sql(f'SELECT * FROM "{tabla_relacion}"', conn)
    if df_relacion.empty:
        print("Sin relaciones. Se omite enriquecimiento.")
        cursor.close()
        return

    df_empresa = pd.read_sql("SELECT id, name FROM api_empresa", conn)
    propietario_name_map = dict(zip(df_empresa["id"], df_empresa["name"]))

    def mapear_propietario(df, col_id, col_salida):
        if col_id in df.columns:
            df[col_salida] = df[col_id].map(propietario_name_map)
        return df

    frames = []
    for tabla, grp in df_relacion.groupby("instalacion"):
        cols = columnas_de(tabla)
        select = ["id"] + [c for c in ("name", "status", "propietario_id") if c in cols]
        df_inst = pd.read_sql(f'SELECT {", ".join(select)} FROM "{tabla}"', conn)
        merged = grp.merge(
            df_inst, left_on="instalacion_id", right_on="id",
            how="left", suffixes=("", "_inst"),
        )
        frames.append(merged)

    df_final = pd.concat(frames, ignore_index=True)
    df_final.drop(columns=["id_inst"], inplace=True, errors="ignore")
    for c in ("name", "status"):
        if c not in df_final.columns:
            df_final[c] = None

    mapear_propietario(df_final, "propietario_id", "propietario_name")
    df_final.drop(columns=["propietario_id"], inplace=True, errors="ignore")
    if "propietario_name" not in df_final.columns:
        df_final["propietario_name"] = None

    master_cols = columnas_de(tabla_maestra)
    master_select = ["id"] + [c for c in ("name", "status", "propietario_id") if c in master_cols]
    df_master = pd.read_sql(f'SELECT {", ".join(master_select)} FROM "{tabla_maestra}"', conn)
    mapear_propietario(df_master, "propietario_id", "propietario_name")

    merge_select = ["id"] + [c for c in ("name", "status", "propietario_name") if c in df_master.columns]
    df_final = df_final.merge(
        df_master[merge_select], left_on=fk_col, right_on="id",
        how="left", suffixes=("", "_" + sufijo),
    )
    for c in ("name_" + sufijo, "status_" + sufijo, "propietario_name_" + sufijo):
        if c not in df_final.columns:
            df_final[c] = None

    df_no_rel = df_master[~df_master["id"].isin(df_final[fk_col].dropna())].copy()
    if "name" in df_no_rel.columns:
        df_no_rel.rename(columns={"name": "name_" + sufijo}, inplace=True)

    columnas_finales = ["id", "status", "name_" + sufijo]
    if "propietario_id" in df_no_rel.columns:
        df_no_rel = df_no_rel.merge(
            df_empresa, left_on="propietario_id", right_on="id",
            how="left", suffixes=("", "_empresa"),
        )
        df_no_rel.rename(columns={"name": "name_empresa"}, inplace=True)
        columnas_finales += ["propietario_id", "name_empresa"]

    columnas_finales = [c for c in columnas_finales if c in df_no_rel.columns]
    df_no_rel_final = df_no_rel[columnas_finales].copy()
    df_no_rel_final.rename(
        columns={"id": fk_col, "status": "status_" + sufijo}, inplace=True
    )

    df_relacionados = df_final[[fk_col, "status_" + sufijo]].drop_duplicates()
    df_op_count = (
        df_final[df_final["status"] == STATUS_EN_OPERACION]
        .groupby(fk_col)
        .size()
        .reset_index(name=cfg["operativo_col"])
    )
    df_relacionados = df_relacionados.merge(df_op_count, on=fk_col, how="left")
    df_relacionados[cfg["operativo_col"]] = (
        df_relacionados[cfg["operativo_col"]].fillna(0).astype(int)
    )
    df_relacionados[cfg["estado_col"]] = df_relacionados[cfg["operativo_col"]].apply(
        lambda x: cfg["texto_sin"] if x == 0 else cfg["texto_con"]
    )

    if "propietario_name" in df_master.columns:
        df_relacionados = df_relacionados.merge(
            df_master[["id", "propietario_name"]].rename(columns={"id": fk_col}),
            on=fk_col, how="left",
        )

    master_out_cols = ["id"] + [c for c in ("name", "status", "propietario_id", "propietario_name") if c in df_master.columns]
    df_master_out = df_master[master_out_cols].copy()
    df_master_out.rename(columns={"id": fk_col}, inplace=True)
    df_master_out["Relacionado"] = df_master_out[fk_col].isin(
        set(df_relacionados[fk_col])
    ).map({True: "Con relacion", False: "Sin relacion"})

    exportar(df_final, df_no_rel_final, df_relacionados, df_master_out, cfg)
    cursor.close()


def _resumen_no_relacionados(df_no_rel_final, cfg):
    status_col = "status_" + cfg["sufijo"]
    if status_col not in df_no_rel_final.columns or "name_empresa" not in df_no_rel_final.columns:
        return pd.DataFrame(columns=["empresa", "no_relacionados"])

    df = df_no_rel_final[df_no_rel_final[status_col] == STATUS_EN_OPERACION]
    res = (
        df.groupby("name_empresa").size()
        .reset_index(name="no_relacionados")
        .rename(columns={"name_empresa": "empresa"})
    )
    res["empresa"] = res["empresa"].fillna("Sin empresa")
    return res


def _resumen_relacionados(df_relacionados, cfg):
    status_col = "status_" + cfg["sufijo"]
    estado_col = cfg["estado_col"]
    if (
        status_col not in df_relacionados.columns
        or estado_col not in df_relacionados.columns
        or "propietario_name" not in df_relacionados.columns
    ):
        return pd.DataFrame(columns=["empresa", "relacion_desconectada"])

    df = df_relacionados[
        (df_relacionados[status_col] == STATUS_EN_OPERACION)
        & (df_relacionados[estado_col] == cfg["texto_sin"])
    ]
    res = (
        df.groupby("propietario_name").size()
        .reset_index(name="relacion_desconectada")
        .rename(columns={"propietario_name": "empresa"})
    )
    res["empresa"] = res["empresa"].fillna("Sin empresa")
    return res


def _resumen(df_no_rel_final, df_relacionados, cfg):
    no_rel = _resumen_no_relacionados(df_no_rel_final, cfg)
    rel = _resumen_relacionados(df_relacionados, cfg)
    resumen = no_rel.merge(rel, on="empresa", how="outer")
    resumen = resumen.fillna(0)
    for c in ("no_relacionados", "relacion_desconectada"):
        if c in resumen.columns:
            resumen[c] = resumen[c].astype(int)
    return resumen.sort_values("no_relacionados", ascending=False).reset_index(drop=True)


def exportar(df_final, df_no_rel_final, df_relacionados, df_master_out, cfg):
    sheets = cfg["sheets"]
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    resumen = _resumen(df_no_rel_final, df_relacionados, cfg)

    os.makedirs(REPORTE_DIR, exist_ok=True)
    os.makedirs(HISTORICO_DIR, exist_ok=True)

    def escribir(base, carpeta):
        with pd.ExcelWriter(os.path.join(carpeta, base + ".xlsx"), engine="xlsxwriter") as w:
            df_final.to_excel(w, sheet_name=sheets[0], index=False)
            df_no_rel_final.to_excel(w, sheet_name=sheets[1], index=False)
            df_relacionados.to_excel(w, sheet_name=sheets[2], index=False)
            df_master_out.to_excel(w, sheet_name=sheets[3], index=False)
            resumen.to_excel(w, sheet_name="resumen", index=False)
        with pd.ExcelWriter(os.path.join(carpeta, base + "_no_relacionados.xlsx"), engine="xlsxwriter") as w:
            df_no_rel_final.to_excel(w, sheet_name=sheets[1], index=False)

    escribir(cfg["archivo"], REPORTE_DIR)
    escribir(cfg["archivo"] + "-" + ts + config.sufijo_historico_bd(), HISTORICO_DIR)

    print(f"Resultados exportados para '{cfg['sufijo']}' en {REPORTE_DIR}")


def procesar_entidad(cfg):
    conn = psycopg2.connect(database=DB_NAME, **PG_CONFIG)
    try:
        print(f"Procesando entidad: {cfg['sufijo']}")
        construir_relaciones(conn, cfg)
        enriquecer(conn, cfg)
        conn.commit()
    finally:
        conn.close()


def ejecutar_entidades(entidades):
    for nombre in entidades:
        if nombre not in ENTIDADES:
            print(f"[WARN] Entidad desconocida: {nombre}. Disponibles: {list(ENTIDADES.keys())}")
            continue
        procesar_entidad(ENTIDADES[nombre])


def main():
    entidades = sys.argv[1:] if len(sys.argv) > 1 else list(ENTIDADES.keys())
    ejecutar_entidades(entidades)


if __name__ == "__main__":
    main()
