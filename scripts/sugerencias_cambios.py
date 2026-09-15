# -*- coding: utf-8 -*-
"""
Reporte de sugerencias de cambio (simples y complejas).

Genera, en una carpeta con timestamp, un Excel consolidado y un Excel por
empresa, con dos hojas:

- "Simples": correcciones de escritura accionables sobre el nombre, con la
  columna "corregido". Códigos: NOMBRE_MINUSCULAS, NOMBRE_DOBLE_ESPACIO,
  NOMBRE_ESPACIOS_EXTREMOS, NOMBRE_GUION_CON_ESPACIOS, NOMBRE_GUION_LARGO,
  KV_MAL_ESCRITO.
- "Complejas": posible relacionamiento para registros "no relacionados" de
  OOCC/EstSE/Servidumbres (quitar prefijo propio -> detectar instalación por
  prefijo vía utils/prefijos.json -> buscar el nombre en la tabla destino).

Salida:
    docs/SugerenciasCambios/<AAAAMMDD_HHMMSS>/sugerenciascambios.xlsx
    docs/SugerenciasCambios/<AAAAMMDD_HHMMSS>/<empresa>.xlsx
"""

import os
import re
import sys
import json
import warnings
from datetime import datetime

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
import reporte_observaciones_empresa as roe

REPORTE_DIR = config.REPORTE_DIR
DB_CONN = config.DB_CONN
PREFIJOS_PATH = os.path.join(BASE_DIR, "utils", "prefijos.json")

# ---------------------------------------------------------------------------
# Sugerencias simples (correcciones de escritura)
# ---------------------------------------------------------------------------

CODIGOS_SUGERENCIAS = [
    "NOMBRE_MINUSCULAS",
    "NOMBRE_DOBLE_ESPACIO",
    "NOMBRE_ESPACIOS_EXTREMOS",
    "NOMBRE_GUION_CON_ESPACIOS",
    "NOMBRE_GUION_LARGO",
    "KV_MAL_ESCRITO",
]

COLUMNAS_SIMPLES = roe.COLUMNAS + ["corregido"]

_MARCA = "\u0001"


def corregir(nombre, codigo):
    """Devuelve el nombre corregido según la regla indicada."""
    n = str(nombre or "")
    if codigo == "NOMBRE_MINUSCULAS":
        return n.replace("kV", _MARCA).upper().replace(_MARCA, "kV")
    if codigo == "NOMBRE_DOBLE_ESPACIO":
        return re.sub(r" {2,}", " ", n)
    if codigo == "NOMBRE_ESPACIOS_EXTREMOS":
        return n.strip()
    if codigo == "NOMBRE_GUION_CON_ESPACIOS":
        return n.replace(" - ", "-")
    if codigo == "NOMBRE_GUION_LARGO":
        return n.replace("\u2014", "-").replace("\u2013", "-")
    if codigo == "KV_MAL_ESCRITO":
        return re.sub(r"[Kk][Vv]", "kV", n)
    return n


_ORDEN_CORRECCION = [
    "NOMBRE_ESPACIOS_EXTREMOS",
    "NOMBRE_DOBLE_ESPACIO",
    "NOMBRE_GUION_LARGO",
    "NOMBRE_GUION_CON_ESPACIOS",
    "KV_MAL_ESCRITO",
    "NOMBRE_MINUSCULAS",
]


def corregir_todo(nombre, codigos):
    """Aplica en orden las correcciones de todas las reglas presentes."""
    n = str(nombre or "")
    for codigo in _ORDEN_CORRECCION:
        if codigo in codigos:
            n = corregir(n, codigo)
    return n


def consolidar_por_registro(df):
    """Colapsa las observaciones de un mismo registro (tipo_instalacion, id) en
    una sola fila, uniendo los códigos y los detalles, y calculando la
    corrección acumulada."""
    filas = []
    for (tipo, pid), grp in df.groupby(["tipo_instalacion", "id"], sort=False):
        nombre = grp["nombre"].iloc[0]
        empresa = grp["propietario_name"].iloc[0]
        codigos = sorted(grp["observacion"].unique())
        detalles = [str(d) for d in grp["detalle"].tolist() if str(d)]
        filas.append({
            "tipo_instalacion": tipo,
            "id": pid,
            "nombre": nombre,
            "propietario_name": empresa,
            "observacion": ", ".join(codigos),
            "detalle": " | ".join(detalles),
            "corregido": corregir_todo(nombre, codigos),
        })
    return pd.DataFrame(filas, columns=COLUMNAS_SIMPLES)


def generar_sugerencias_simples():
    df = roe.consolidar_observaciones()
    if df.empty:
        return pd.DataFrame(columns=COLUMNAS_SIMPLES)

    filtro = df[df["observacion"].isin(CODIGOS_SUGERENCIAS)].copy()
    filtro = consolidar_por_registro(filtro)
    return filtro.sort_values(["propietario_name", "tipo_instalacion", "id"])


# ---------------------------------------------------------------------------
# Sugerencias complejas (posible relacionamiento)
# ---------------------------------------------------------------------------

FUENTES = [
    {
        "instalacion": "oocc",
        "archivo": "api_oocc_relacion_enriquecida.xlsx",
        "sheet": "oocc_no_relacionados",
        "id_col": "oocc_id",
        "name_col": "name_oocc",
        "status_col": "status_oocc",
        "empresa_col": "name_empresa",
        "prefijos_propios": ["OC"],
    },
    {
        "instalacion": "estse",
        "archivo": "api_est_se_relacion_enriquecida.xlsx",
        "sheet": "est_se_no_relacionados",
        "id_col": "estructurasubestacion_id",
        "name_col": "name_estse",
        "status_col": "status_estse",
        "empresa_col": "name_empresa",
        "prefijos_propios": ["EST", "SUE", "MB", "ML", "MP"],
    },
    {
        "instalacion": "serv",
        "archivo": "api_serv_relacion_enriquecida.xlsx",
        "sheet": "serv_no_relacionados",
        "id_col": "servidumbre_id",
        "name_col": "name_servidumbre",
        "status_col": "status_servidumbre",
        "empresa_col": "name_empresa",
        "prefijos_propios": ["SERV", "TERR"],
    },
]

COLUMNAS_COMPLEJAS = [
    "instalacion", "id", "nombre", "propietario_name",
    "prefijo_detectado", "tabla_sugerida", "id_sugerido", "nombre_sugerido",
    "propietario_sugerido",
]


def cargar_mapa_prefijos():
    with open(PREFIJOS_PATH, encoding="utf-8") as f:
        data = json.load(f)
    mapa = {}
    for e in data:
        tabla = e.get("tabla_psql", "")
        for p in e.get("prefijo", []):
            if p and tabla:
                mapa.setdefault(str(p).upper(), []).append(tabla)
    return mapa


def quitar_prefijo(nombre, prefijos):
    n = str(nombre or "").strip()
    for p in sorted(prefijos, key=len, reverse=True):
        up = str(p).upper()
        if n.upper().startswith(up) and (len(n) == len(up) or n[len(up)].isspace()):
            return n[len(up):].strip()
    return n


def detectar_prefijo(nombre, mapa):
    tokens = str(nombre or "").strip().split()
    if not tokens:
        return None, None
    pref = tokens[0].upper()
    return pref, mapa.get(pref)


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


def pre_cargar_indice(conn, mapa):
    tablas = set()
    for lista in mapa.values():
        tablas.update(lista)

    cur = conn.cursor()
    cur.execute("SELECT id, name FROM api_empresa")
    empresa_map = {r[0]: r[1] for r in cur.fetchall()}
    cur.close()

    indice = {}
    for tabla in tablas:
        cols = columnas_de(conn, tabla)
        if "name" not in cols or "id" not in cols or "status" not in cols:
            continue
        if "empresa_name" in cols:
            select_prop = "empresa_name"
        elif "propietario_id" in cols:
            select_prop = "propietario_id"
        else:
            select_prop = None

        try:
            if select_prop:
                df = pd.read_sql(
                    f'SELECT id, name, "{select_prop}" AS prop FROM "{tabla}" '
                    f"WHERE status='EN_OPERACION'",
                    conn,
                )
            else:
                df = pd.read_sql(
                    f'SELECT id, name FROM "{tabla}" WHERE status=\'EN_OPERACION\'',
                    conn,
                )
                df["prop"] = None
        except Exception:
            continue

        sub = {}
        for _, r in df.iterrows():
            name = str(r["name"] or "").strip()
            if not name:
                continue
            prop = r["prop"]
            if select_prop == "propietario_id":
                prop = "" if prop is None or pd.isna(prop) else empresa_map.get(prop, "")
            else:
                prop = "" if prop is None or pd.isna(prop) else str(prop)
            sub.setdefault(name.upper(), (r["id"], name, prop))
        indice[tabla] = sub
    return indice


def generar_sugerencias_complejas():
    mapa = cargar_mapa_prefijos()

    conn = psycopg2.connect(**DB_CONN)
    try:
        indice = pre_cargar_indice(conn, mapa)
    finally:
        conn.close()

    filas = []
    for cfg in FUENTES:
        ruta = os.path.join(REPORTE_DIR, cfg["archivo"])
        if not os.path.exists(ruta):
            print(f"[WARN] No existe {cfg['archivo']}. Se omite.")
            continue

        df = pd.read_excel(ruta, sheet_name=cfg["sheet"])
        if cfg["status_col"] in df.columns:
            df = df[df[cfg["status_col"]] == "EN_OPERACION"]

        for _, r in df.iterrows():
            nombre = str(r.get(cfg["name_col"]) or "").strip()
            sin_prefijo = quitar_prefijo(nombre, cfg["prefijos_propios"])
            pref, tablas = detectar_prefijo(sin_prefijo, mapa)
            if not pref or not tablas:
                continue

            for tabla in tablas:
                coincidencia = indice.get(tabla, {}).get(sin_prefijo.upper())
                if coincidencia:
                    id_sugerido, nombre_sugerido, propietario_sugerido = coincidencia
                    filas.append({
                        "instalacion": cfg["instalacion"],
                        "id": r.get(cfg["id_col"]),
                        "nombre": nombre,
                        "propietario_name": r.get(cfg["empresa_col"]),
                        "prefijo_detectado": pref,
                        "tabla_sugerida": tabla,
                        "id_sugerido": id_sugerido,
                        "nombre_sugerido": nombre_sugerido,
                        "propietario_sugerido": propietario_sugerido,
                    })
                    break

    df_salida = pd.DataFrame(filas, columns=COLUMNAS_COMPLEJAS)
    return df_salida.sort_values(["instalacion", "tabla_sugerida", "id"])


# ---------------------------------------------------------------------------
# Exportación
# ---------------------------------------------------------------------------

def main():
    df_simples = generar_sugerencias_simples()
    df_complejas = generar_sugerencias_complejas()

    if df_simples.empty and df_complejas.empty:
        print("No hay sugerencias para consolidar.")
        return

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    carpeta = os.path.join(config.SUGERENCIAS_DIR, ts + config.sufijo_carpeta_bd())
    os.makedirs(carpeta, exist_ok=True)

    ruta_completa = os.path.join(carpeta, "sugerenciascambios.xlsx")
    with pd.ExcelWriter(ruta_completa, engine="xlsxwriter") as w:
        df_simples.to_excel(w, sheet_name="Simples", index=False)
        df_complejas.to_excel(w, sheet_name="Complejas", index=False)

    empresas = sorted(
        set(df_simples["propietario_name"]) | set(df_complejas["propietario_name"])
    )
    simples_por_empresa = (
        dict(iter(df_simples.groupby("propietario_name"))) if not df_simples.empty else {}
    )
    complejas_por_empresa = (
        dict(iter(df_complejas.groupby("propietario_name"))) if not df_complejas.empty else {}
    )
    vacio_s = pd.DataFrame(columns=COLUMNAS_SIMPLES)
    vacio_c = pd.DataFrame(columns=COLUMNAS_COMPLEJAS)
    for empresa in empresas:
        ruta = os.path.join(carpeta, roe.sanitizar(empresa) + ".xlsx")
        with pd.ExcelWriter(ruta, engine="xlsxwriter") as w:
            simples_por_empresa.get(empresa, vacio_s).to_excel(
                w, sheet_name="Simples", index=False
            )
            complejas_por_empresa.get(empresa, vacio_c).to_excel(
                w, sheet_name="Complejas", index=False
            )

    conteo_simples = (
        df_simples["observacion"].str.split(", ").explode().value_counts().to_dict()
        if not df_simples.empty else {}
    )
    conteo_complejas = (
        df_complejas.groupby("tabla_sugerida").size().to_dict()
        if not df_complejas.empty else {}
    )

    print(f"Sugerencias simples: {len(df_simples)} registros")
    for codigo, n in sorted(conteo_simples.items(), key=lambda kv: -kv[1]):
        print(f"  {codigo}: {n}")
    print(f"Sugerencias complejas: {len(df_complejas)} registros")
    for tabla, n in sorted(conteo_complejas.items(), key=lambda kv: -kv[1]):
        print(f"  {tabla}: {n}")
    print(f"Empresas con sugerencias: {len(empresas)}")
    print(f"Carpeta: {carpeta}")


if __name__ == "__main__":
    main()
