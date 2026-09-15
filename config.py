# -*- coding: utf-8 -*-
import os
import re
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOCS_DIR = os.path.join(BASE_DIR, "docs")

REPORTE_DIR = os.path.join(DOCS_DIR, "reporte")
HISTORICO_DIR = os.path.join(DOCS_DIR, "historico")
AVANCE_DIR = os.path.join(DOCS_DIR, "avance")
SUGERENCIAS_DIR = os.path.join(DOCS_DIR, "SugerenciasCambios")


def _cargar_env(ruta):
    if not os.path.exists(ruta):
        return
    with open(ruta, "r", encoding="utf-8") as f:
        for linea in f:
            linea = linea.strip()
            if not linea or linea.startswith("#") or "=" not in linea:
                continue
            clave, _, valor = linea.partition("=")
            clave = clave.strip()
            valor = valor.strip().strip('"').strip("'")
            os.environ.setdefault(clave, valor)


_cargar_env(os.path.join(BASE_DIR, ".env"))

DB_NAME = os.environ.get("DB_NAME", "bdatx")
DB_USER = os.environ.get("DB_USER", "postgres")
DB_PASSWORD = os.environ.get("DB_PASSWORD", "")
DB_HOST = os.environ.get("DB_HOST", "localhost")
DB_PORT = int(os.environ.get("DB_PORT", "5432"))

PG_CONFIG = {
    "user": DB_USER,
    "password": DB_PASSWORD,
    "host": DB_HOST,
    "port": DB_PORT,
}

DB_CONN = {
    "dbname": DB_NAME,
    "user": DB_USER,
    "password": DB_PASSWORD,
    "host": DB_HOST,
    "port": DB_PORT,
}

PSQL_PATH = os.environ.get("PSQL_PATH", "")
DOWNLOAD_URL = os.environ.get("DOWNLOAD_URL", "")
BASE_FOLDER = os.environ.get("BASE_FOLDER", "")

_FECHA_BD = None
_RE_FECHA = re.compile(r"(\d{4}-\d{2}-\d{2})[ T](\d{2}:\d{2}(?::\d{2})?)")


def set_fecha_bd(valor):
    global _FECHA_BD
    _FECHA_BD = valor


def _leer_comentario_bd():
    try:
        import psycopg2
        conn = psycopg2.connect(**DB_CONN)
        try:
            cur = conn.cursor()
            cur.execute(
                "SELECT shobj_description(oid, 'pg_database') FROM pg_database WHERE datname = %s",
                (DB_NAME,),
            )
            fila = cur.fetchone()
            cur.close()
            return fila[0] if fila else None
        finally:
            conn.close()
    except Exception:
        return None


def fecha_bd():
    """Fecha (snapshot) de la base BDATx cargada. Se obtiene del comentario de
    la base; si no es posible, usa la fecha actual."""
    global _FECHA_BD
    if _FECHA_BD is not None:
        return _FECHA_BD
    comentario = _leer_comentario_bd()
    if comentario:
        m = _RE_FECHA.search(comentario)
        if m:
            for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
                try:
                    _FECHA_BD = datetime.strptime(f"{m.group(1)} {m.group(2)}", fmt)
                    break
                except ValueError:
                    continue
    if _FECHA_BD is None:
        _FECHA_BD = datetime.now()
    return _FECHA_BD


def fecha_bd_str(fmt="%Y-%m-%d"):
    return fecha_bd().strftime(fmt)


def sufijo_historico_bd():
    return "-bd" + fecha_bd_str("%Y-%m-%d")


def sufijo_carpeta_bd():
    return "_bd" + fecha_bd_str("%Y%m%d")
