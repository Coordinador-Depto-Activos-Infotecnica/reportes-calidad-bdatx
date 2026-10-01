# -*- coding: utf-8 -*-
"""Extrae, a partir del nombre de cada Linea de Transmision (api_linea.name),
las subestaciones de origen y destino, la tension y el sufijo opcional, y
vuelca el resultado a un Excel para revision manual.

El nombre de una linea sigue el patron:
    LT {NOMBRE SE 1} - {NOMBRE SE 2} {tension}kV {opcional}
El separador entre subestaciones puede ser '-', '–' o '—', y el
sufijo de tension puede ir pegado o separado (p.ej. "66kV", "220 kV").
"""
import json
import os
import re
import sys
import unicodedata

import pandas as pd
import psycopg2

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, BASE_DIR)

import config

DB_CONN = config.DB_CONN
UTILS_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
UTILS_DOCS_DIR = os.path.join(UTILS_DIR, "docs")
ABREVIATURAS_SSEE_PATH = os.path.join(UTILS_DIR, "abreviaturas_ssee.json")
ABREVIATURA_LINEA_JSON_PATH = os.path.join(UTILS_DIR, "abreviatura_linea.json")

RE_NOMBRE_LINEA = re.compile(
    r'^LT\s+(?P<se1>.+?)\s*[-–—]\s*(?P<se2>.+?)\s*'
    r'(?P<tension>\d+(?:[.,]\d+)?)\s*[kK][vV]\.?\s*(?P<opcional>.*)$'
)

RE_PREFIJO_SE = re.compile(r'^(S/E|TAP OFF|TAP)\s+', re.IGNORECASE)


def normaliza_nombre(nombre):
    """Normaliza un nombre de subestacion para comparacion: mayusculas, sin
    tildes, sin prefijos S/E, TAP OFF o TAP, espacios colapsados."""
    if not nombre:
        return ""
    texto = nombre.upper().strip()
    texto = unicodedata.normalize('NFKD', texto).encode('ascii', 'ignore').decode('ascii')
    texto = RE_PREFIJO_SE.sub('', texto)
    texto = re.sub(r'\s+', ' ', texto).strip()
    return texto


def normaliza_canonica(nombre):
    """Normaliza un nombre de subestacion preservando el caracter de TAP:
    mayusculas, sin tildes, quita solo el prefijo S/E y unifica "TAP OFF" y
    "TAP" a "TAP". Asi "S/E TAP OFF X" y "TAP OFF X" coinciden entre si, pero
    no colisionan con "S/E X"."""
    if not nombre:
        return ""
    texto = nombre.upper().strip()
    texto = unicodedata.normalize('NFKD', texto).encode('ascii', 'ignore').decode('ascii')
    texto = re.sub(r'^S/E\s+', '', texto)
    texto = re.sub(r'^TAP OFF\s+|^TAP\s+', 'TAP ', texto)
    texto = re.sub(r'\s+', ' ', texto).strip()
    return texto


def parsea_nombre_linea(nombre):
    """Separa el nombre de una linea en SE1, SE2, tension y sufijo opcional.
    Devuelve None si el nombre no calza con el patron esperado."""
    if not nombre:
        return None
    m = RE_NOMBRE_LINEA.match(nombre.strip())
    if not m:
        return None
    return {
        "se1_extraida": m.group("se1").strip(),
        "se2_extraida": m.group("se2").strip(),
        "tension_kv": m.group("tension").replace(",", "."),
        "opcional": m.group("opcional").strip(),
    }


def cargar_datos(conn):
    cur = conn.cursor()
    cur.execute("SELECT id, name FROM api_linea WHERE status='EN_OPERACION' ORDER BY id")
    lineas = cur.fetchall()
    cur.execute("SELECT id, name FROM api_subestacion")
    subestaciones = cur.fetchall()
    cur.close()
    return lineas, subestaciones


def construye_mapa_subestaciones(subestaciones):
    """Devuelve (mapa_canonico, mapa_respaldo). El canonico preserva el TAP y
    se usa primero; el respaldo mantiene el comportamiento historico
    (eliminando tambien TAP OFF/TAP) para las subestaciones sin su par TAP."""
    mapa_canonico = {}
    mapa_respaldo = {}
    for sub_id, nombre in subestaciones:
        mapa_canonico.setdefault(normaliza_canonica(nombre), []).append((sub_id, nombre))
        mapa_respaldo.setdefault(normaliza_nombre(nombre), []).append((sub_id, nombre))
    return mapa_canonico, mapa_respaldo


def busca_subestacion(nombre_extraido, mapa_canonico, mapa_respaldo):
    candidatos = mapa_canonico.get(normaliza_canonica(nombre_extraido))
    if not candidatos:
        candidatos = mapa_respaldo.get(normaliza_nombre(nombre_extraido))
    if not candidatos:
        return None, None
    sub_id, nombre_real = candidatos[0]
    return sub_id, nombre_real


def cargar_abreviaturas_ssee(ruta=ABREVIATURAS_SSEE_PATH):
    """Devuelve (mapa_por_nombre, mapa_por_id). El mapa por id BDATx es el
    preferente porque evita colisiones de nombres homonimos."""
    with open(ruta, "r", encoding="utf-8") as f:
        registros = json.load(f)
    mapa_por_nombre = {}
    mapa_por_id = {}
    for registro in registros:
        nombre = registro.get("Nombre SSEE")
        abreviatura = registro.get("Abreviatura")
        if not abreviatura:
            continue
        if nombre:
            mapa_por_nombre.setdefault(normaliza_nombre(nombre), abreviatura)
        id_bdatx = registro.get("id BDATx")
        if nombre and id_bdatx not in (None, "", 0, "0"):
            mapa_por_id.setdefault(id_bdatx, (nombre, abreviatura))
    return mapa_por_nombre, mapa_por_id


def busca_abreviatura(se_id, nombre_bd, nombre_extraido, mapa_abreviaturas, mapa_abreviaturas_id):
    if se_id is not None:
        registro = mapa_abreviaturas_id.get(se_id)
        if registro:
            nombre_reg, abreviatura = registro
            if normaliza_canonica(nombre_reg) == normaliza_canonica(nombre_bd):
                return abreviatura
    for nombre in (nombre_bd, nombre_extraido):
        clave = normaliza_nombre(nombre)
        if clave in mapa_abreviaturas:
            return mapa_abreviaturas[clave]
    return None


def construye_dataframe(lineas, mapa_canonico, mapa_respaldo, mapa_abreviaturas, mapa_abreviaturas_id):
    filas = []
    for linea_id, nombre in lineas:
        partes = parsea_nombre_linea(nombre)
        fila = {
            "linea_id": linea_id,
            "nombre_linea": nombre,
            "se1_extraida": None,
            "se2_extraida": None,
            "tension_kv": None,
            "opcional": None,
            "se1_id": None,
            "se1_nombre_bd": None,
            "se2_id": None,
            "se2_nombre_bd": None,
            "match_completo": False,
            "patron_reconocido": partes is not None,
            "se1_abreviatura": None,
            "se2_abreviatura": None,
            "nombre_abreviado": None,
        }
        if partes is not None:
            fila.update(partes)
            se1_id, se1_nombre_bd = busca_subestacion(partes["se1_extraida"], mapa_canonico, mapa_respaldo)
            se2_id, se2_nombre_bd = busca_subestacion(partes["se2_extraida"], mapa_canonico, mapa_respaldo)
            fila["se1_id"] = se1_id
            fila["se1_nombre_bd"] = se1_nombre_bd
            fila["se2_id"] = se2_id
            fila["se2_nombre_bd"] = se2_nombre_bd
            fila["match_completo"] = se1_id is not None and se2_id is not None

            se1_abrev = busca_abreviatura(se1_id, se1_nombre_bd, partes["se1_extraida"], mapa_abreviaturas, mapa_abreviaturas_id)
            se2_abrev = busca_abreviatura(se2_id, se2_nombre_bd, partes["se2_extraida"], mapa_abreviaturas, mapa_abreviaturas_id)
            fila["se1_abreviatura"] = se1_abrev
            fila["se2_abreviatura"] = se2_abrev
            if se1_abrev and se2_abrev:
                fila["nombre_abreviado"] = f"{se1_abrev}-{se2_abrev}"
        filas.append(fila)
    return pd.DataFrame(filas)


def genera_excel(df, ruta_salida):
    os.makedirs(os.path.dirname(ruta_salida), exist_ok=True)
    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Lineas_Abreviadas", index=False)
        df[~df["match_completo"]].to_excel(writer, sheet_name="Sin_Match_Completo", index=False)
        df[df["nombre_abreviado"].isna()].to_excel(writer, sheet_name="Sin_Abreviatura", index=False)


def genera_json(df, ruta_salida):
    """Guarda, por linea_id, el nombre_abreviado y los datos que lo componen
    para que otros scripts puedan consumirlo sin pasar por la BD ni el Excel."""
    columnas = [
        "linea_id", "nombre_linea", "se1_nombre_bd", "se2_nombre_bd",
        "se1_abreviatura", "se2_abreviatura", "nombre_abreviado",
    ]
    registros = df[columnas].where(pd.notnull(df[columnas]), None).to_dict(orient="records")
    os.makedirs(os.path.dirname(ruta_salida), exist_ok=True)
    with open(ruta_salida, "w", encoding="utf-8") as f:
        json.dump(registros, f, ensure_ascii=False, indent=2)


def main():
    conn = psycopg2.connect(**DB_CONN)
    try:
        lineas, subestaciones = cargar_datos(conn)
    finally:
        conn.close()

    mapa_canonico, mapa_respaldo = construye_mapa_subestaciones(subestaciones)
    mapa_abreviaturas, mapa_abreviaturas_id = cargar_abreviaturas_ssee()
    df = construye_dataframe(lineas, mapa_canonico, mapa_respaldo, mapa_abreviaturas, mapa_abreviaturas_id)

    ruta_salida = os.path.join(UTILS_DOCS_DIR, "lineas_abreviadas.xlsx")
    genera_excel(df, ruta_salida)
    genera_json(df, ABREVIATURA_LINEA_JSON_PATH)

    total = len(df)
    con_match = int(df["match_completo"].sum())
    sin_patron = int((~df["patron_reconocido"]).sum())
    con_abreviatura = int(df["nombre_abreviado"].notna().sum())
    print(f"Total lineas: {total}")
    print(f"Con match completo de ambas SE: {con_match}")
    print(f"Sin reconocer el patron de nombre: {sin_patron}")
    print(f"Con nombre_abreviado completo: {con_abreviatura}")
    print(f"Excel generado en: {ruta_salida}")
    print(f"JSON generado en: {ABREVIATURA_LINEA_JSON_PATH}")


if __name__ == "__main__":
    main()
