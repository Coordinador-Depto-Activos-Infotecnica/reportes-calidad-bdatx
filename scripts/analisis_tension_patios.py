# -*- coding: utf-8 -*-
"""
Verificación de coherencia del nivel de tensión en patios de subestación.

Para cada patio EN_OPERACION de api_patiosubestacion se extrae del nombre el
nivel de tensión (formato "PT [SE] XXXkV", tolerando escrituras como "KV",
"kv", "13.2 kV", coma decimal "2,4kV", sufijos "(AUX)"/"-2", etc.) y se mapea
a su letra según utils/niveles_de_tension.json. Luego se valida:

1. Barra (api_barra): el número de tensión del nombre debe ser igual al del
   patio (relación patio_subestacion_id).
2. Paño (api_pano): la letra de nivel de tensión del nombre debe coincidir con
   la letra esperada según el patio (formato "PA [SE] [LETRA][TIPO][N]"; se
   evalúa la primera letra del último token).

Códigos:
    PATIO_TENSION_NO_IDENTIFICADA - patio sin tensión mapeable.
    BARRA_TENSION_NO_COINCIDE     - barra con tensión distinta.
    PANO_TENSION_NO_COINCIDE      - paño con letra incorrecta.

Salidas:
    docs/reporte/analisis_tension_patios.xlsx
    docs/reporte/analisis_tension_patios_detalle.csv
    docs/historico/analisis_tension_patios-<YYYY-MM-DD_HH-MM-SS>.xlsx
    docs/historico/analisis_tension_patios_detalle-<YYYY-MM-DD_HH-MM-SS>.csv
"""

import os
import re
import sys
import json
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

NIVELES_PATH = os.path.join(BASE_DIR, "utils", "niveles_de_tension.json")

CODIGO_PATIO = "PATIO_TENSION_NO_IDENTIFICADA"
CODIGO_BARRA = "BARRA_TENSION_NO_COINCIDE"
CODIGO_PANO = "PANO_TENSION_NO_COINCIDE"

CATALOGO = [
    {
        "codigo": CODIGO_PATIO,
        "aplica_a": "Patio",
        "descripcion": "El nombre del patio no permite identificar un nivel de tensión mapeable al catálogo.",
    },
    {
        "codigo": CODIGO_BARRA,
        "aplica_a": "Barra",
        "descripcion": "El nivel de tensión del nombre de la barra no coincide con el del patio.",
    },
    {
        "codigo": CODIGO_PANO,
        "aplica_a": "Paño",
        "descripcion": "La letra de nivel de tensión del nombre del paño no coincide con la del patio.",
    },
]

COLUMNAS_DETALLE = [
    "tipo_instalacion", "codigo", "id", "nombre", "patio_id", "patio_nombre",
    "tension_patio", "tension_encontrada", "propietario_id", "propietario", "detalle",
]

TOLERANCIA = 0.001


def cargar_niveles():
    with open(NIVELES_PATH, encoding="utf-8") as f:
        data = json.load(f)

    niveles = []
    for e in data["relacion_tension_letra"]:
        tension = str(e.get("tension", "")).replace(",", ".")
        letra = e.get("letra", "")
        numeros = re.findall(r"\d+(?:\.\d+)?", tension)
        if not numeros:
            continue
        valores = [float(x) for x in numeros]
        vmin = vmax = valores[0]
        if len(valores) > 1:
            vmin, vmax = valores[0], valores[1]
        niveles.append((vmin, vmax, letra))
    return niveles


def extraer_tension(name):
    """Devuelve el valor numérico de tensión del nombre, o None si no lo hay."""
    if not name:
        return None
    m = re.search(r"(\d+(?:[.,]\d+)?)\s*[Kk][Vv]", str(name))
    if not m:
        return None
    return round(float(m.group(1).replace(",", ".")), 3)


def letra_para_tension(t, niveles):
    if t is None:
        return None
    for vmin, vmax, letra in niveles:
        if vmin - TOLERANCIA <= t <= vmax + TOLERANCIA:
            return letra
    return None


def extraer_letra_pano(name):
    """Devuelve la primera letra del último token del nombre, o None."""
    if not name:
        return None
    tokens = str(name).strip().upper().split()
    if not tokens:
        return None
    last = tokens[-1]
    if last and last[0].isalpha():
        return last[0]
    return None


def _fmt_tension(t):
    if t is None:
        return ""
    if float(t) == int(t):
        return str(int(t))
    return str(t)


def _key(v):
    return "" if pd.isna(v) else str(int(v))


def analizar(conn):
    cur = conn.cursor()
    cur.execute("SELECT id, name FROM api_empresa")
    empresa_map = {r[0]: r[1] for r in cur.fetchall()}
    cur.close()

    niveles = cargar_niveles()

    patios = pd.read_sql(
        "SELECT id, name, propietario_id FROM api_patiosubestacion "
        "WHERE status='EN_OPERACION'",
        conn,
    )
    barras = pd.read_sql(
        "SELECT id, name, patio_subestacion_id, propietario_id FROM api_barra "
        "WHERE status='EN_OPERACION'",
        conn,
    )
    panos = pd.read_sql(
        "SELECT id, name, patio_subestacion_id, propietario_id FROM api_pano "
        "WHERE status='EN_OPERACION'",
        conn,
    )

    barras_by_patio = defaultdict(list)
    for _, r in barras.iterrows():
        barras_by_patio[_key(r["patio_subestacion_id"])].append(r)

    panos_by_patio = defaultdict(list)
    for _, r in panos.iterrows():
        panos_by_patio[_key(r["patio_subestacion_id"])].append(r)

    findings = []

    for _, p in patios.iterrows():
        patio_id = p["id"]
        patio_name = str(p["name"] or "")
        pid = p["propietario_id"]
        prop = "" if (pid is None or pd.isna(pid)) else str(empresa_map.get(pid, ""))

        t = extraer_tension(patio_name)
        letra = letra_para_tension(t, niveles)

        if t is None or letra is None:
            findings.append({
                "tipo_instalacion": "api_patiosubestacion",
                "codigo": CODIGO_PATIO,
                "id": _key(patio_id),
                "nombre": patio_name,
                "patio_id": "",
                "patio_nombre": "",
                "tension_patio": "",
                "tension_encontrada": _fmt_tension(t),
                "propietario_id": "" if (pid is None or pd.isna(pid)) else str(pid),
                "propietario": prop,
                "detalle": "No se pudo identificar un nivel de tensión mapeable al catálogo en el nombre del patio.",
            })
            continue

        for b in barras_by_patio.get(_key(patio_id), []):
            tb = extraer_tension(b["name"])
            if tb is None or tb != t:
                bpid = b["propietario_id"]
                bprop = "" if (bpid is None or pd.isna(bpid)) else str(empresa_map.get(bpid, ""))
                findings.append({
                    "tipo_instalacion": "api_barra",
                    "codigo": CODIGO_BARRA,
                    "id": _key(b["id"]),
                    "nombre": str(b["name"] or ""),
                    "patio_id": _key(patio_id),
                    "patio_nombre": patio_name,
                    "tension_patio": _fmt_tension(t),
                    "tension_encontrada": _fmt_tension(tb),
                    "propietario_id": "" if (bpid is None or pd.isna(bpid)) else str(bpid),
                    "propietario": bprop,
                    "detalle": f"Tensión de la barra ({_fmt_tension(tb) or 'no detectable'}) no coincide con la del patio ({_fmt_tension(t)}).",
                })

        for pa in panos_by_patio.get(_key(patio_id), []):
            lp = extraer_letra_pano(pa["name"])
            if lp is None or lp != letra:
                apid = pa["propietario_id"]
                aprop = "" if (apid is None or pd.isna(apid)) else str(empresa_map.get(apid, ""))
                findings.append({
                    "tipo_instalacion": "api_pano",
                    "codigo": CODIGO_PANO,
                    "id": _key(pa["id"]),
                    "nombre": str(pa["name"] or ""),
                    "patio_id": _key(patio_id),
                    "patio_nombre": patio_name,
                    "tension_patio": letra,
                    "tension_encontrada": lp or "",
                    "propietario_id": "" if (apid is None or pd.isna(apid)) else str(apid),
                    "propietario": aprop,
                    "detalle": f"Letra de tensión del paño ('{lp or 'no detectable'}') no coincide con la esperada '{letra}'.",
                })

    return findings


def exportar(findings):
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    df_findings = pd.DataFrame(findings, columns=COLUMNAS_DETALLE)

    conteo = df_findings.groupby("codigo").size().to_dict() if not df_findings.empty else {}
    filas_resumen = []
    for c in CATALOGO:
        filas_resumen.append({
            "codigo": c["codigo"],
            "aplica_a": c["aplica_a"],
            "descripcion": c["descripcion"],
            "cantidad": int(conteo.get(c["codigo"], 0)),
        })
    df_resumen = pd.DataFrame(filas_resumen)

    prop = defaultdict(int)
    for f in findings:
        prop[(f["propietario_id"], f["propietario"])] += 1
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

    base = "analisis_tension_patios"
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

    print(f"Total hallazgos de tensión: {len(findings)}")
    print(df_resumen.to_string(index=False))
    print(f"Excel reporte: {xlsx_reporte}")
    print(f"CSV reporte:   {csv_reporte}")


def main():
    conn = psycopg2.connect(**DB_CONN)
    try:
        findings = analizar(conn)
    finally:
        conn.close()

    exportar(findings)


if __name__ == "__main__":
    main()
