# -*- coding: utf-8 -*-
import os
import sys
import re
import requests
import gzip
import subprocess
import time
import unicodedata
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE_DIR)

import config

PSQL_PATH = config.PSQL_PATH
DOWNLOAD_URL = config.DOWNLOAD_URL
BASE_FOLDER = config.BASE_FOLDER
DB_NAME = config.DB_NAME
PG_CONFIG = config.PG_CONFIG


def clean_text(text):
    """Normaliza y limpia cadenas para asegurar compatibilidad con UTF-8."""
    return unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")


def download_database(download_url, save_folder):
    """Descarga el archivo desde el enlace proporcionado."""
    os.makedirs(save_folder, exist_ok=True)
    gz_file_path = os.path.join(save_folder, "database.psql.gz")

    response = requests.get(download_url, stream=True, verify=False)
    if response.status_code == 200:
        with open(gz_file_path, "wb") as file:
            for chunk in response.iter_content(chunk_size=8192):
                file.write(chunk)
        print(f"Archivo descargado en {gz_file_path}")
    else:
        raise Exception(f"Error al descargar el archivo: {response.status_code}")
    return gz_file_path


def extract_gz(gz_file_path):
    """Extrae un archivo .gz manteniendo la extensión .psql."""
    extracted_path = gz_file_path.replace(".gz", "")
    with gzip.open(gz_file_path, 'rb') as gz_file:
        with open(extracted_path, 'wb') as extracted_file:
            extracted_file.write(gz_file.read())
    print(f"Archivo extraído en {extracted_path}")
    return extracted_path


def terminate_connections(db_name, pg_config):
    """Termina las conexiones activas a la base de datos."""
    try:
        user = pg_config["user"]
        password = pg_config["password"]
        host = pg_config["host"]
        port = pg_config["port"]

        env = {**os.environ, "PGPASSWORD": password}

        terminate_command = (
            f'"{PSQL_PATH}" -U {user} -h {host} -p {port} '
            f'-c "SELECT pg_terminate_backend(pg_stat_activity.pid) '
            f'FROM pg_stat_activity '
            f'WHERE pg_stat_activity.datname = \'{db_name}\' AND pid <> pg_backend_pid();"'
        )
        result = subprocess.run(terminate_command, shell=True, env=env, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        if result.returncode != 0:
            print(f"Error al terminar conexiones: {result.stderr.decode().strip()}")
        else:
            print(f"Conexiones a la base de datos '{db_name}' terminadas exitosamente.")

    except Exception as e:
        print(f"Error al terminar conexiones: {e}")


def load_psql_to_postgresql(psql_file_path, db_name, pg_config, fecha_snapshot):
    """Carga el archivo .psql a PostgreSQL y fija en el comentario de la base la
    fecha del snapshot (no la de carga)."""
    try:
        user = pg_config["user"]
        password = pg_config["password"]
        host = pg_config["host"]
        port = pg_config["port"]

        env = {**os.environ, "PGPASSWORD": password}

        terminate_connections(db_name, pg_config)

        drop_db_command = (
            f'"{PSQL_PATH}" -U {user} -h {host} -p {port} '
            f'-c "DROP DATABASE IF EXISTS {db_name};"'
        )
        result = subprocess.run(drop_db_command, shell=True, env=env, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        if result.returncode != 0:
            print(f"Error al ejecutar DROP DATABASE: {result.stderr.decode().strip()}")
        else:
            print(f"Base de datos '{db_name}' eliminada exitosamente.")

        create_db_command = (
            f'"{PSQL_PATH}" -U {user} -h {host} -p {port} '
            f'-c "CREATE DATABASE {db_name};"'
        )
        result = subprocess.run(create_db_command, shell=True, env=env, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        if result.returncode != 0:
            print(f"Error al crear la base de datos: {result.stderr.decode().strip()}")
        else:
            print(f"Base de datos '{db_name}' creada exitosamente.")

        load_file_command = (
            f'"{PSQL_PATH}" -U {user} -h {host} -p {port} '
            f'-d {db_name} -f "{psql_file_path}"'
        )
        result = subprocess.run(load_file_command, shell=True, env=env, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        if result.returncode != 0:
            print(f"Error al cargar el archivo .psql: {result.stderr.decode().strip()}")
        else:
            print(f"Archivo .psql cargado exitosamente en la base de datos '{db_name}'.")

        texto = f"BDATx {fecha_snapshot:%Y-%m-%d %H:%M:%S}"
        comment_text = clean_text(texto)
        comment_command = (
            f'"{PSQL_PATH}" -U {user} -h {host} -p {port} -d {db_name} '
            f'-c "COMMENT ON DATABASE {db_name} IS \'{comment_text}\';"'
        )
        result = subprocess.run(comment_command, shell=True, env=env, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        if result.returncode != 0:
            print(f"Error al agregar comentario: {result.stderr.decode().strip()}")
        else:
            print(f"Comentario agregado: '{texto}'")

        config.set_fecha_bd(fecha_snapshot)

    except Exception as e:
        print(f"Error al ejecutar comandos: {e}")


def _fecha_de(nombre):
    s = str(nombre)
    if re.fullmatch(r"\d{12}", s):
        return datetime.strptime(s, "%Y%m%d%H%M")
    m = re.search(r"(\d{4}-\d{2}-\d{2})-(\d{6})", s)
    if m:
        return datetime.strptime(m.group(1) + m.group(2), "%Y-%m-%d%H%M%S")
    m = re.search(r"(\d{4}-\d{2}-\d{2})", s)
    if m:
        return datetime.strptime(m.group(1), "%Y-%m-%d")
    return None


def listar_bases_guardadas():
    bases = []
    if not BASE_FOLDER or not os.path.isdir(BASE_FOLDER):
        return bases
    for nombre in sorted(os.listdir(BASE_FOLDER)):
        ruta = os.path.join(BASE_FOLDER, nombre)
        if os.path.isdir(ruta):
            psql = None
            for cand in ("database.psql", "database.psql.gz"):
                if os.path.exists(os.path.join(ruta, cand)):
                    psql = os.path.join(ruta, cand)
                    break
            if psql is None:
                internos = [f for f in os.listdir(ruta) if f.lower().endswith(".psql")]
                if internos:
                    psql = os.path.join(ruta, internos[0])
            if psql:
                fecha = (_fecha_de(nombre) or _fecha_de(os.path.basename(psql))
                         or datetime.fromtimestamp(os.path.getmtime(psql)))
                bases.append({"etiqueta": nombre, "ruta": psql, "fecha": fecha})
        elif os.path.isfile(ruta) and ruta.lower().endswith(".psql"):
            fecha = _fecha_de(nombre) or datetime.fromtimestamp(os.path.getmtime(ruta))
            bases.append({"etiqueta": nombre, "ruta": ruta, "fecha": fecha})
    bases.sort(key=lambda b: b["fecha"], reverse=True)
    return bases


def elegir_base_guardada():
    bases = listar_bases_guardadas()
    if not bases:
        print("No se encontraron bases guardadas en la carpeta configurada.")
        ruta = input("Ruta del archivo .psql: ").strip().strip('"')
        if ruta and os.path.isfile(ruta):
            return ruta, (_fecha_de(os.path.basename(ruta))
                          or datetime.fromtimestamp(os.path.getmtime(ruta)))
        return None, None

    print("Bases guardadas disponibles:")
    for i, b in enumerate(bases, start=1):
        print(f"  {i}. {b['fecha']:%Y-%m-%d %H:%M}  {b['ruta']}")

    resp = input("Elige un número o escribe la ruta del .psql: ").strip().strip('"')
    if resp.isdigit() and 1 <= int(resp) <= len(bases):
        b = bases[int(resp) - 1]
        return b["ruta"], b["fecha"]
    if resp and os.path.isfile(resp):
        return resp, (_fecha_de(os.path.basename(resp))
                      or datetime.fromtimestamp(os.path.getmtime(resp)))
    return None, None


def cargar_base_guardada():
    ruta, fecha = elegir_base_guardada()
    if not ruta or not os.path.isfile(ruta):
        print("No se seleccionó una base válida.")
        return
    if fecha is None:
        fecha = datetime.now()
    print(f"Cargando base: {ruta}")
    print(f"Fecha del snapshot: {fecha:%Y-%m-%d %H:%M}")
    if ruta.lower().endswith(".gz"):
        ruta = extract_gz(ruta)
    load_psql_to_postgresql(ruta, DB_NAME, PG_CONFIG, fecha)


def main(modo="descargar"):
    if modo == "ninguno":
        return
    if modo == "cargar":
        cargar_base_guardada()
        return

    fecha = datetime.now()
    save_folder = os.path.join(BASE_FOLDER, fecha.strftime("%Y%m%d%H%M"))
    try:
        gz_file_path = download_database(DOWNLOAD_URL, save_folder)
        psql_file_path = extract_gz(gz_file_path)
        load_psql_to_postgresql(psql_file_path, DB_NAME, PG_CONFIG, fecha)
    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    start_time = time.time()
    main()
    end_time = time.time()

    execution_time = end_time - start_time
    hours = execution_time // 3600
    minutes = (execution_time % 3600) // 60
    seconds = execution_time % 60

    print(f"Tiempo de ejecución: {int(hours)} horas, {int(minutes)} minutos, {int(seconds)} segundos")
