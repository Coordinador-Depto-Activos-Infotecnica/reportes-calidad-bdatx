# -*- coding: utf-8 -*-
"""
Genera los entregables globales de calidad de la BDATx (todas las empresas):

- Un informe Word y PDF consolidado, con las mismas métricas, secciones y
  catálogo que el informe por empresa (formato_informe.py), pero con las
  instalaciones observadas y universos sumados a nivel global.
- Un Excel resumen con una fila por empresa (solo las que tienen hallazgos) y
  una columna por métrica (instalaciones observadas y % Error), más una fila
  TOTAL BDATx.

Las métricas cuentan instalaciones distintas (tipo_instalacion + id), no
hallazgos: una instalación con varias observaciones cuenta una vez.

Son entregables adicionales: no reemplazan los informes Word/PDF ni el Excel
por empresa.

Salida:
    docs/ArchivoReporteGlobal/<aaaammdd_hhmmss>_bd<AAAAMMDD>/
        word/BDATx_global.docx
        pdf/BDATx_global.pdf
        resumen_empresas.xlsx
"""

import os
import sys
from datetime import datetime

import pandas as pd

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE_DIR)
sys.path.insert(0, os.path.join(BASE_DIR, "scripts"))

import config
import reporte_observaciones_empresa as roe
import formato_informe as fi

DOCS_DIR = config.DOCS_DIR

GLOBAL = "BDATx_global"
ETIQUETA = "Alcance: BDATx"


def agregar_universos(universos):
    """Suma los universos por empresa en un único alcance global."""
    res = {}
    for clave, valor in universos.items():
        if clave == "Anexos":
            res["Anexos"] = {
                etiqueta: {GLOBAL: sum(int(v) for v in por_empresa.values())}
                for etiqueta, por_empresa in valor.items()
            }
        else:
            res[clave] = {GLOBAL: sum(int(v) for v in valor.values())}
    return res


def _sumar_metricas(destino, metricas):
    destino["inconsistencias"] += metricas["inconsistencias"]
    for campo in ("general", "especifica", "especifica_reglas", "anexos"):
        for clave, valor in metricas[campo].items():
            destino[campo][clave] = destino[campo].get(clave, 0) + valor


def calcular_metricas(df):
    """Métricas por empresa y global (suma, ya que las empresas son disjuntas)."""
    por_empresa = {}
    global_ = {"general": {}, "especifica": {}, "especifica_reglas": {}, "anexos": {},
               "inconsistencias": 0}
    for empresa, df_emp in df.groupby("propietario_name"):
        if not empresa:
            continue
        metricas = fi.metricas_empresa(df_emp)
        por_empresa[empresa] = metricas
        _sumar_metricas(global_, metricas)
    return por_empresa, global_


def _pct_num(observadas, total):
    return round(min(100.0, observadas / total * 100), 2) if total else 0.0


def _fila_metricas(etiqueta, metricas, total, universos, clave):
    """Una fila del resumen: las mismas métricas del PDF (conteo y % Error)."""
    observadas = metricas["inconsistencias"]
    fila = {
        "Empresa": etiqueta,
        "Registros Totales": int(total),
        "Instalaciones observadas": int(observadas),
        "Salud Global (%)": round(max(0.0, 1 - observadas / total) * 100, 2) if total else 100.0,
    }

    for entry in roe.CATALOGO:
        if entry["grupo"] != "General":
            continue
        sub = entry["subcategoria"]
        n = metricas["general"].get(sub, 0)
        denom = universos.get(sub, {}).get(clave, 0)
        fila[f"{sub} - Instalaciones observadas"] = int(n)
        fila[f"{sub} - % Error"] = _pct_num(n, denom)

    for grupo in roe.ESPECIFICAS_POR_INSTALACION:
        inst = grupo["instalacion"]
        total_inst = metricas["especifica"].get(inst, 0)
        denom_g = universos.get(inst, {}).get(clave, 0)
        fila[f"{inst} - Instalaciones observadas"] = int(total_inst)
        fila[f"{inst} - % Error"] = _pct_num(total_inst, denom_g)
        for codigo, etiqueta in grupo["reglas"]:
            n = metricas["especifica_reglas"].get((inst, codigo), 0)
            denom = universos.get(codigo, universos.get(inst, {})).get(clave, 0)
            fila[f"{inst} - {etiqueta} - Instalaciones observadas"] = int(n)
            fila[f"{inst} - {etiqueta} - % Error"] = _pct_num(n, denom)

    u_anexos = universos.get("Anexos", {})
    for eti, tabla in fi.ANEXOS_POR_TABLA:
        n = metricas["anexos"].get(tabla, 0)
        denom = u_anexos.get(eti, {}).get(clave, 0)
        fila[f"{eti} - Instalaciones observadas"] = int(n)
        fila[f"{eti} - % Error"] = _pct_num(n, denom)

    return fila


def _escribir_excel(df, ruta):
    with pd.ExcelWriter(ruta, engine="xlsxwriter") as w:
        df.to_excel(w, sheet_name="Resumen", index=False)
        wb = w.book
        ws = w.sheets["Resumen"]

        fmt_header = wb.add_format({
            "bold": True, "bg_color": "#1E293B", "font_color": "white",
            "text_wrap": True, "valign": "top", "border": 1,
        })
        fmt_int = wb.add_format({"num_format": "#,##0"})
        fmt_pct = wb.add_format({"num_format": '0.00"%"'})
        fmt_txt = wb.add_format({})

        for c, col in enumerate(df.columns):
            ws.write(0, c, col, fmt_header)
            if col == "Empresa":
                ws.set_column(c, c, 32, fmt_txt)
            elif col.endswith("(%)") or col.endswith("% Error"):
                ws.set_column(c, c, 16, fmt_pct)
            else:
                ws.set_column(c, c, 16, fmt_int)

        ws.freeze_panes(1, 1)
        ws.autofilter(0, 0, len(df), len(df.columns) - 1)


def construir_resumen_empresas(por_empresa, registros_totales, universos, empresas,
                               metricas_global, total_global, universos_global, carpeta):
    registros = [
        _fila_metricas(empresa, por_empresa[empresa], registros_totales.get(empresa, 0),
                       universos, empresa)
        for empresa in empresas
    ]
    registros.append(_fila_metricas(
        "TOTAL BDATx", metricas_global, total_global, universos_global, GLOBAL,
    ))

    df = pd.DataFrame(registros)
    ruta = os.path.join(carpeta, "resumen_empresas.xlsx")
    _escribir_excel(df, ruta)
    return ruta


def main():
    df, registros_totales, empresas = fi.cargar_datos()
    universos = fi.cargar_universos()

    por_empresa, metricas_global = calcular_metricas(df)
    total_global = sum(int(v) for v in registros_totales.values())
    universos_global = agregar_universos(universos)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    carpeta = os.path.join(DOCS_DIR, "ArchivoReporteGlobal", ts + config.sufijo_carpeta_bd())
    carpeta_word = os.path.join(carpeta, "word")
    carpeta_pdf = os.path.join(carpeta, "pdf")
    os.makedirs(carpeta_word, exist_ok=True)
    os.makedirs(carpeta_pdf, exist_ok=True)

    fi.generar_documento(
        GLOBAL, metricas_global, {GLOBAL: total_global}, universos_global,
        carpeta_word, etiqueta_meta=ETIQUETA,
    )
    res = fi.convertir_a_pdf(carpeta_word, carpeta_pdf)

    ruta_xlsx = construir_resumen_empresas(
        por_empresa, registros_totales, universos, empresas,
        metricas_global, total_global, universos_global, carpeta,
    )

    print(f"Informe global generado en {carpeta}")
    print(f"Resumen por empresa: {ruta_xlsx}")
    if res["word"] != res["pdf"]:
        print(f"[ERROR] Faltan PDFs: word={res['word']} pdf={res['pdf']}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
