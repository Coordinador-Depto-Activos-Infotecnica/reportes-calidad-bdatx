# -*- coding: utf-8 -*-
import os
import sys
import time
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
import proceso_excel
import reporte_observaciones_empresa
import sugerencias_cambios
import formato_informe

ENTIDADES = ["oocc", "estse", "serv"]


def ejecutar(nombre, fn, *args):
    print(f"[{datetime.now():%H:%M:%S}] Ejecutando {nombre} ...")
    inicio = time.time()
    fn(*args)
    duracion = time.time() - inicio
    print(f"[{datetime.now():%H:%M:%S}] {nombre} OK ({duracion:.1f}s).\n")


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


def main():
    inicio_total = time.time()

    try:
        modo_bd = preguntar_carga_bd()

        if modo_bd == "descargar":
            ejecutar("carga_bdatx", carga_bdatx.main, "descargar")
        elif modo_bd == "cargar":
            ejecutar("carga_bdatx", carga_bdatx.main, "cargar")

        ejecutar("analisis_estse_oc_serv", analisis_estse_oc_serv.ejecutar_entidades, ENTIDADES)
        ejecutar("analisis_duplicados_desconexiones", analisis_duplicados_desconexiones.main)
        ejecutar("auditoria_lineas_metricas", auditoria_lineas_metricas.main)
        ejecutar("auditoria_contenedores_metricas", auditoria_contenedores_metricas.main)
        ejecutar("analisis_tension_patios", analisis_tension_patios.main)
        ejecutar("revision_calidad_escriturageneral", revision_calidad_escriturageneral.main)
        ejecutar("revision_herencia_nombre", revision_herencia_nombre.main)
        ejecutar("analisis_prefijo", analisis_prefijo.main)
        ejecutar("revision_abreviatura_torre", revision_abreviatura_torre.main)
        ejecutar("revision_abreviatura_vano", revision_abreviatura_vano.main)
        ejecutar("proceso_excel", proceso_excel.main)
        ejecutar("reporte_observaciones_empresa", reporte_observaciones_empresa.main)
        ejecutar("sugerencias_cambios", sugerencias_cambios.main)
        ejecutar("formato_informe", formato_informe.main)
    except Exception as e:
        print(f"Flujo detenido: {e}")
        sys.exit(1)

    duracion_total = time.time() - inicio_total
    print(f"Flujo completado en {duracion_total:.1f}s.")


if __name__ == "__main__":
    main()
