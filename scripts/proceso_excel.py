# -*- coding: utf-8 -*-
import os
import sys
import pandas as pd

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE_DIR)

import config

REPORTE_DIR = config.REPORTE_DIR
AVANCE_DIR = config.AVANCE_DIR

ARCHIVOS = [
    ("analisis_desconectados.xlsx", "registros_desconectados", "desconectados"),
    ("analisis_duplicados.xlsx", "registros_duplicados", "duplicados"),
]


def procesar_archivo(nombre, columna):
    ruta = os.path.join(REPORTE_DIR, nombre)
    if not os.path.exists(ruta):
        print(f"[WARN] No existe {ruta}. Se omite.")
        return None, None

    hojas = pd.read_excel(ruta, sheet_name=None)

    if "analisis" not in hojas:
        print(f"[WARN] No se encontro la hoja 'analisis' en {nombre}.")
        return None, None

    df = hojas["analisis"]
    if columna not in df.columns:
        print(f"[WARN] No se encontro la columna '{columna}' en {nombre}.")
        return None, None

    df["porcentaje error"] = df.apply(
        lambda r: round(r[columna] / r["total_registros"] * 100, 2)
        if r["total_registros"] else 0,
        axis=1,
    )
    hojas["analisis"] = df

    with pd.ExcelWriter(ruta, engine="xlsxwriter") as w:
        for hoja, data in hojas.items():
            data.to_excel(w, sheet_name=hoja, index=False)

    print(f"Procesado {nombre}: columna 'porcentaje error' agregada.")

    df_empresa = hojas.get("analisis_por_empresa")
    return df, df_empresa


def _columna_diaria(df, key_col, fecha):
    if df is None or "porcentaje error" not in df.columns:
        return None
    return df[[key_col, "porcentaje error"]].rename(
        columns={"porcentaje error": fecha}
    )


def _fusionar(existente, hoy, key_col, fecha):
    if existente is None:
        return hoy
    if fecha in existente.columns:
        existente = existente.drop(columns=[fecha])
    return existente.merge(hoy, on=key_col, how="outer")


def construir_avance(registros, fecha):
    os.makedirs(AVANCE_DIR, exist_ok=True)
    ruta = os.path.join(AVANCE_DIR, "avance_error_diario.xlsx")

    hojas = {}
    if os.path.exists(ruta):
        hojas = pd.read_excel(ruta, sheet_name=None)

    for nombre_hoja, key_col, df in registros:
        hoy = _columna_diaria(df, key_col, fecha)
        if hoy is None:
            continue
        hojas[nombre_hoja] = _fusionar(hojas.get(nombre_hoja), hoy, key_col, fecha)

    with pd.ExcelWriter(ruta, engine="xlsxwriter") as w:
        for nombre_hoja, data in hojas.items():
            data.to_excel(w, sheet_name=nombre_hoja, index=False)

    print(f"Avance diario actualizado en {ruta}")


def main():
    registros = []
    for nombre, columna, hoja_avance in ARCHIVOS:
        df_tabla, df_empresa = procesar_archivo(nombre, columna)
        if df_tabla is not None:
            registros.append((hoja_avance, "tabla", df_tabla))
        if df_empresa is not None:
            registros.append((hoja_avance + "_empresa", "empresa_name", df_empresa))

    if registros:
        construir_avance(registros, config.fecha_bd_str("%d-%m-%Y"))


if __name__ == "__main__":
    main()
