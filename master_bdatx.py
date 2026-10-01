# -*- coding: utf-8 -*-
import os
import re
import sys
import time
import argparse
from datetime import datetime

import config

BASE_DIR = config.BASE_DIR

sys.path.insert(0, os.path.join(BASE_DIR, "scripts"))

import carga_bdatx
import analisis_estse_oc_serv
import analisis_duplicados_desconexiones
import auditoria_lineas_metricas
import auditoria_contenedores_metricas
import analisis_tension_patios
import revision_calidad_escriturageneral
import revision_herencia_nombre
import analisis_prefijo
import revision_abreviatura_torre
import revision_abreviatura_vano
import revision_documentos_datasheet
import proceso_excel
import reporte_observaciones_empresa
import sugerencias_cambios
import formato_informe
import informe_global

ENTIDADES = ["oocc", "estse", "serv"]

# Revisiones de hallazgos: (clave, título, función, args)
LISTA_REVISIONES = [
    ("analisis_estse_oc_serv", "Enriquecimiento OOCC / EstSE / Servidumbres",
     analisis_estse_oc_serv.ejecutar_entidades, (ENTIDADES,)),
    ("analisis_duplicados_desconexiones", "Duplicados, desconectados y conexión nula",
     analisis_duplicados_desconexiones.main, ()),
    ("auditoria_lineas_metricas", "Auditoría de jerarquía de líneas",
     auditoria_lineas_metricas.main, ()),
    ("auditoria_contenedores_metricas", "Contenedores vacíos",
     auditoria_contenedores_metricas.main, ()),
    ("analisis_tension_patios", "Coherencia de tensión en patios",
     analisis_tension_patios.main, ()),
    ("revision_calidad_escriturageneral", "Calidad de escritura general",
     revision_calidad_escriturageneral.main, ()),
    ("revision_herencia_nombre", "Herencia de nombre de la subestación",
     revision_herencia_nombre.main, ()),
    ("analisis_prefijo", "Revisión de prefijos de nombre",
     analisis_prefijo.main, ()),
    ("revision_abreviatura_torre", "Abreviatura de línea en torres",
     revision_abreviatura_torre.main, ()),
    ("revision_abreviatura_vano", "Abreviatura de línea en vanos",
     revision_abreviatura_vano.main, ()),
    ("revision_documentos_datasheet", "Documentos del datasheet (anexos)",
     revision_documentos_datasheet.main, ()),
]


def ejecutar(nombre, fn, *args):
    print(f"[{datetime.now():%H:%M:%S}] Ejecutando {nombre} ...")
    inicio = time.time()
    fn(*args)
    duracion = time.time() - inicio
    print(f"[{datetime.now():%H:%M:%S}] {nombre} OK ({duracion:.1f}s).\n")


def listar_revisiones():
    for i, (clave, titulo, _, _) in enumerate(LISTA_REVISIONES, start=1):
        print(f"  {i:>2}. {titulo}  ({clave})")


def parsear_seleccion(texto):
    texto = (texto or "").strip().lower()
    if texto in ("todos", "all", "*"):
        return list(range(len(LISTA_REVISIONES)))
    seleccion = set()
    for parte in re.split(r"[,\s]+", texto):
        if not parte:
            continue
        if "-" in parte:
            a, b = parte.split("-", 1)
            for i in range(int(a), int(b) + 1):
                seleccion.add(i - 1)
        else:
            seleccion.add(int(parte) - 1)
    return sorted(i for i in seleccion if 0 <= i < len(LISTA_REVISIONES))


def ejecutar_revisiones(indices):
    for i in indices:
        clave, _, fn, args = LISTA_REVISIONES[i]
        ejecutar(clave, fn, *args)


def ejecutar_salidas():
    ejecutar("proceso_excel", proceso_excel.main)
    ejecutar("reporte_observaciones_empresa", reporte_observaciones_empresa.main)
    ejecutar("sugerencias_cambios", sugerencias_cambios.main)
    ejecutar("formato_informe", formato_informe.main)
    ejecutar("informe_global", informe_global.main)


def preguntar_carga_bd():
    print("Origen de la base BDATx:")
    print("  n = no cargar (usar la base actual)")
    print("  d = descargar y cargar la última desde la API")
    print("  a = cargar una base guardada anterior")
    respuesta = input("Elige una opción (n/d/a) [n]: ").strip().lower()
    if respuesta in ("d", "descargar"):
        return "descargar"
    if respuesta in ("a", "anterior", "cargar", "guardada"):
        return "cargar"
    return "ninguno"


def preguntar_modo():
    print("¿Qué deseas ejecutar?")
    print("  1 = Flujo completo (todos los criterios)")
    print("  2 = Uno o más criterios + actualizar reportes")
    print("  3 = Solo actualizar reportes (Excel/Word/PDF)")
    respuesta = input("Elige una opción (1/2/3) [1]: ").strip()
    return {"2": "criterios", "3": "salidas"}.get(respuesta, "completo")


def preguntar_revisiones():
    print("Revisiones disponibles:")
    listar_revisiones()
    respuesta = input("Elige número(s) (p. ej. 3,7 o 9-11 o 'todos'): ").strip()
    return parsear_seleccion(respuesta)


def main():
    parser = argparse.ArgumentParser(description="Orquestador del flujo de reportes BDATx.")
    parser.add_argument("--listar", action="store_true",
                        help="Muestra las revisiones numeradas y sale.")
    parser.add_argument("--revision",
                        help="Números de revisión a ejecutar (p. ej. 3,7 o 9-11 o todos).")
    parser.add_argument("--solo-salidas", action="store_true",
                        help="Solo actualiza reportes (Excel/Word/PDF).")
    args = parser.parse_args()

    if args.listar:
        listar_revisiones()
        return

    inicio_total = time.time()

    try:
        if args.revision is not None:
            modo = "criterios"
            seleccion = parsear_seleccion(args.revision)
        elif args.solo_salidas:
            modo = "salidas"
            seleccion = []
        else:
            modo = preguntar_modo()
            seleccion = preguntar_revisiones() if modo == "criterios" else []

        if modo == "completo":
            modo_bd = preguntar_carga_bd()
            if modo_bd == "descargar":
                ejecutar("carga_bdatx", carga_bdatx.main, "descargar")
            elif modo_bd == "cargar":
                ejecutar("carga_bdatx", carga_bdatx.main, "cargar")
            ejecutar_revisiones(range(len(LISTA_REVISIONES)))
        else:
            if modo == "criterios":
                if not seleccion:
                    print("No se seleccionó ninguna revisión.")
                    return
                ejecutar_revisiones(seleccion)

        ejecutar_salidas()
    except Exception as e:
        print(f"Flujo detenido: {e}")
        sys.exit(1)

    duracion_total = time.time() - inicio_total
    print(f"Flujo completado en {duracion_total:.1f}s.")


if __name__ == "__main__":
    main()
