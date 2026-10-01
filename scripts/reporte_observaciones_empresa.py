# -*- coding: utf-8 -*-
"""
Consolida todas las observaciones de calidad de los reportes generados en
docs/reporte/ y las entrega en un Excel por empresa.

Columnas: tipo_instalacion, id, nombre, propietario_name, observacion (código), detalle.

Salida: docs/ReporteExcelEmpresa/<aaaammdd_hhmmss>/<empresa>.xlsx
"""

import os
import sys
import re
import hashlib
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

import config

REPORTE_DIR = config.REPORTE_DIR
DOCS_DIR = config.DOCS_DIR

COLUMNAS = ["tipo_instalacion", "id", "nombre", "propietario_name", "observacion", "detalle"]

TIPO_INSTALACION_A_TABLA = {
    "oocc": "api_oocc",
    "estse": "api_estructurasubestacion",
    "serv": "api_servidumbre",
    "lineas": "api_linea",
    "circuitos": "api_circuito",
    "tramos": "api_tramo",
    "vanos": "api_vano",
    "tuneles": "api_tunel",
    "estaciones-repetidoras": "api_estacionrepetidora",
    "accesorios-vanos": "api_accesoriovano",
}

# Catálogo canónico de reglas: refleja exactamente la estructura de
# criterios-revision/criterio_revision_validacion.md (tres grupos con subcategorías).
CATALOGO = [
    {"grupo": "General", "subcategoria": "Duplicidad por nombre",
     "codigos": ["NOMBRE_DUPLICADO"]},
    {"grupo": "General", "subcategoria": "Contenedores sin relacionamiento descendente",
     "codigos": ["CONTENEDOR_VACIO"]},
    {"grupo": "General", "subcategoria": "Registros desconectados",
     "codigos": ["REGISTRO_DESCONECTADO", "CONEXION_NULA"]},
    {"grupo": "General", "subcategoria": "OOCC, Estructuras SE y Servidumbres sin conexión",
     "codigos": ["SIN_RELACION", "RELACION_DESCONECTADA"]},
    {"grupo": "General", "subcategoria": "Revisión de prefijos",
     "codigos": ["PREFIJO_NOMBRE"]},
    {"grupo": "General", "subcategoria": "Revisión de abreviatura de línea en torres",
     "codigos": ["ABREVIATURA_LINEA_TORRE"]},
    {"grupo": "General", "subcategoria": "Revisión de herencia de nombre",
     "codigos": ["HERENCIA_LINEA", "HERENCIA_SUBESTACION"]},
    {"grupo": "General", "subcategoria": "Revisión de sufijos",
     "codigos": ["SUFIJO_NOMBRE"]},
    {"grupo": "General", "subcategoria": "Revisión de escritura general",
     "codigos": ["NOMBRE_MINUSCULAS", "NOMBRE_DOBLE_ESPACIO", "NOMBRE_ESPACIOS_EXTREMOS",
                 "NOMBRE_GUION_CON_ESPACIOS", "NOMBRE_GUION_LARGO", "NOMBRE_SIN_SEPARADOR",
                 "KV_MAL_ESCRITO"]},
    {"grupo": "General", "subcategoria": "Conectividad de circuitos",
     "codigos": ["CIRCUITO_TRAMOS_DESCONECTADOS"]},
    {"grupo": "Anexos", "subcategoria": "Documentos del datasheet",
     "codigos": ["DOCUMENTO_FALTANTE", "DOCUMENTO_NO_EXISTE", "DOCUMENTO_INCORRECTO"]},
    {"grupo": "Específica", "subcategoria": "Completitud específica por instalación",
     "codigos": ["VANO_SIN_ACCESORIOS", "TORRE_SIN_OOCC", "MARCO_SIN_OOCC",
                 "TORRE_SIN_ACCESORIOS", "MARCO_SIN_ACCESORIOS", "VANO_MULTIPLES_ACCESORIOS"]},
    {"grupo": "Específica", "subcategoria": "Nodos específicos",
     "codigos": ["VANO_SIN_NODO1", "VANO_SIN_NODO2", "TRAMO_NODOS_IGUALES"]},
    {"grupo": "Específica", "subcategoria": "Coherencia de tensión en patios",
     "codigos": ["PATIO_TENSION_NO_IDENTIFICADA", "BARRA_TENSION_NO_COINCIDE",
                 "PANO_TENSION_NO_COINCIDE"]},
    {"grupo": "Específica", "subcategoria": "Abreviatura de línea de vanos",
     "codigos": ["ABREVIATURA_LINEA_VANO"]},
]

# Agrupación de la Revisión Específica por instalación (Patio/Vano/Torre/Marcolinea/Tramo).
# Cada instalación lista sus reglas como (código, etiqueta legible).
ESPECIFICAS_POR_INSTALACION = [
    {
        "instalacion": "Patio",
        "reglas": [
            ("BARRA_TENSION_NO_COINCIDE", "Barra con tensión distinta"),
            ("PANO_TENSION_NO_COINCIDE", "Paño con letra de tensión incorrecta"),
            ("PATIO_TENSION_NO_IDENTIFICADA", "Tensión no identificable"),
        ],
    },
    {
        "instalacion": "Vano",
        "reglas": [
            ("VANO_SIN_ACCESORIOS", "Sin accesorios"),
            ("VANO_MULTIPLES_ACCESORIOS", "Múltiples accesorios"),
            ("VANO_SIN_NODO1", "Sin Nodo 1"),
            ("VANO_SIN_NODO2", "Sin Nodo 2"),
            ("ABREVIATURA_LINEA_VANO", "Abreviatura de línea incorrecta"),
        ],
    },
    {
        "instalacion": "Torre",
        "reglas": [
            ("TORRE_SIN_OOCC", "Sin OOCC"),
            ("TORRE_SIN_ACCESORIOS", "Sin accesorios"),
        ],
    },
    {
        "instalacion": "Marcolinea",
        "reglas": [
            ("MARCO_SIN_OOCC", "Sin OOCC"),
            ("MARCO_SIN_ACCESORIOS", "Sin accesorios"),
        ],
    },
    {
        "instalacion": "Tramo",
        "reglas": [
            ("TRAMO_NODOS_IGUALES", "Nodos iguales"),
        ],
    },
]


def clean(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    return str(v)


def sanitizar(nombre):
    n = re.sub(r'[\\/:*?"<>|]', "_", str(nombre or ""))
    n = n.strip().rstrip(".")
    return n[:120] or "Sin_nombre"


def _frame(out_dict):
    return pd.DataFrame(out_dict, columns=COLUMNAS)


# ----------------------------------------------------------------------
# Fuentes
# ----------------------------------------------------------------------

def leer_analisis(archivo, observacion, detalle):
    """analisis_desconectados / analisis_duplicados / analisis_conexion_nula."""
    ruta = os.path.join(REPORTE_DIR, archivo)
    if not os.path.exists(ruta):
        return pd.DataFrame(columns=COLUMNAS)

    xl = pd.ExcelFile(ruta)
    frames = []
    for s in xl.sheet_names:
        if s in ("analisis", "analisis_por_empresa"):
            continue
        df = pd.read_excel(
            ruta, sheet_name=s,
            usecols=["id", "name", "instalacion", "empresa_name"],
        )
        if not df.empty:
            frames.append(df)

    if not frames:
        return pd.DataFrame(columns=COLUMNAS)

    df = pd.concat(frames, ignore_index=True)
    return _frame({
        "tipo_instalacion": df["instalacion"].map(clean),
        "id": df["id"].map(clean),
        "nombre": df["name"].map(clean),
        "propietario_name": df["empresa_name"].map(clean),
        "observacion": observacion,
        "detalle": detalle,
        "fuente": archivo,
    })


STATUS_EN_OPERACION = "EN_OPERACION"


ENRIQUECIDAS = [
    {
        "entidad": "estse",
        "archivo": "api_est_se_relacion_enriquecida.xlsx",
        "fk": "estructurasubestacion_id",
        "sheet_no_rel": "est_se_no_relacionados",
        "name_no_rel": "name_estse",
        "empresa_no_rel": "name_empresa",
        "sheet_rel": "est_se_relacionados",
        "estado_col": "estado_dependiente",
        "texto_sin": "Sin dependiente operativo",
        "empresa_rel": "propietario_name",
        "sheet_maestra": "api_estse",
        "status_col": "status_estse",
    },
    {
        "entidad": "oocc",
        "archivo": "api_oocc_relacion_enriquecida.xlsx",
        "fk": "oocc_id",
        "sheet_no_rel": "oocc_no_relacionados",
        "name_no_rel": "name_oocc",
        "empresa_no_rel": "name_empresa",
        "sheet_rel": "oocc_relacionados",
        "estado_col": "estado_principal",
        "texto_sin": "Sin principal operativo",
        "empresa_rel": "propietario_name",
        "sheet_maestra": "api_oocc",
        "status_col": "status_oocc",
    },
    {
        "entidad": "serv",
        "archivo": "api_serv_relacion_enriquecida.xlsx",
        "fk": "servidumbre_id",
        "sheet_no_rel": "serv_no_relacionados",
        "name_no_rel": "name_servidumbre",
        "empresa_no_rel": "name_empresa",
        "sheet_rel": "serv_relacionados",
        "estado_col": "estado_principal",
        "texto_sin": "Sin principal operativo",
        "empresa_rel": "propietario_name",
        "sheet_maestra": "api_serv",
        "status_col": "status_servidumbre",
    },
]


def leer_enriquecidas():
    frames = []
    for cfg in ENRIQUECIDAS:
        ruta = os.path.join(REPORTE_DIR, cfg["archivo"])
        if not os.path.exists(ruta):
            continue

        # Sin relación (solo EN_OPERACION)
        df_no = pd.read_excel(
            ruta, sheet_name=cfg["sheet_no_rel"],
            usecols=[cfg["fk"], cfg["name_no_rel"], cfg["empresa_no_rel"], cfg["status_col"]],
        )
        if not df_no.empty:
            df_no = df_no[df_no[cfg["status_col"]] == STATUS_EN_OPERACION]
        if not df_no.empty:
            frames.append(_frame({
                "tipo_instalacion": cfg["entidad"],
                "id": df_no[cfg["fk"]].map(clean),
                "nombre": df_no[cfg["name_no_rel"]].map(clean),
                "propietario_name": df_no[cfg["empresa_no_rel"]].map(clean),
                "observacion": "SIN_RELACION",
                "detalle": f"Instalación {cfg['entidad']} sin relación",
                "fuente": cfg["archivo"],
            }))

        # Relación desconectada (solo EN_OPERACION)
        df_rel = pd.read_excel(
            ruta, sheet_name=cfg["sheet_rel"],
            usecols=[cfg["fk"], cfg["estado_col"], cfg["empresa_rel"], cfg["status_col"]],
        )
        if not df_rel.empty:
            df_rel = df_rel[
                (df_rel[cfg["status_col"]] == STATUS_EN_OPERACION)
                & (df_rel[cfg["estado_col"]] == cfg["texto_sin"])
            ]
        if not df_rel.empty:
            nombre_map = {}
            try:
                df_master = pd.read_excel(
                    ruta, sheet_name=cfg["sheet_maestra"],
                    usecols=[cfg["fk"], "name"],
                )
                nombre_map = dict(zip(df_master[cfg["fk"]].map(clean), df_master["name"].map(clean)))
            except Exception:
                pass
            frames.append(_frame({
                "tipo_instalacion": cfg["entidad"],
                "id": df_rel[cfg["fk"]].map(clean),
                "nombre": df_rel[cfg["fk"]].map(lambda x: nombre_map.get(clean(x), "")),
                "propietario_name": df_rel[cfg["empresa_rel"]].map(clean),
                "observacion": "RELACION_DESCONECTADA",
                "detalle": cfg["texto_sin"],
                "fuente": cfg["archivo"],
            }))

    if not frames:
        return pd.DataFrame(columns=COLUMNAS)
    return pd.concat(frames, ignore_index=True)


def leer_auditoria_lineas():
    ruta = os.path.join(REPORTE_DIR, "auditoria_lineas_metricas.xlsx")
    if not os.path.exists(ruta):
        return pd.DataFrame(columns=COLUMNAS)
    df = pd.read_excel(
        ruta, sheet_name="Hallazgos",
        usecols=["codigo", "mensaje", "tipo", "id", "nombre", "propietario"],
    )
    return _frame({
        "tipo_instalacion": df["tipo"].map(clean),
        "id": df["id"].map(clean),
        "nombre": df["nombre"].map(clean),
        "propietario_name": df["propietario"].map(clean),
        "observacion": df["codigo"].map(clean),
        "detalle": df["mensaje"].map(clean),
        "fuente": "auditoria_lineas_metricas.xlsx",
    })


def leer_auditoria_contenedores():
    ruta = os.path.join(REPORTE_DIR, "auditoria_contenedores_metricas.xlsx")
    if not os.path.exists(ruta):
        return pd.DataFrame(columns=COLUMNAS)
    df = pd.read_excel(
        ruta, sheet_name="Detalle",
        usecols=["contenedor", "label", "observacion", "codigo", "id", "nombre", "propietario"],
    )

    def _detalle(r):
        obs = clean(r["observacion"])
        label = clean(r["label"])
        if obs == "Sin OOCC":
            return f"{label} sin obras civiles (OOCC) relacionadas"
        if obs == "Sin accesorios":
            return f"{label} sin accesorios estructurales"
        return f"{label} sin instalaciones aguas abajo"

    return _frame({
        "tipo_instalacion": df["contenedor"].map(clean),
        "id": df["id"].map(clean),
        "nombre": df["nombre"].map(clean),
        "propietario_name": df["propietario"].map(clean),
        "observacion": df["codigo"].map(clean),
        "detalle": df.apply(_detalle, axis=1),
        "fuente": "auditoria_contenedores_metricas.xlsx",
    })


def leer_escritura():
    ruta = os.path.join(REPORTE_DIR, "revision_calidad_escriturageneral.xlsx")
    if not os.path.exists(ruta):
        return pd.DataFrame(columns=COLUMNAS)
    df = pd.read_excel(
        ruta, sheet_name="Detalle",
        usecols=["tabla", "codigo", "id", "nombre", "tipo", "detalle", "propietario"],
    )
    return _frame({
        "tipo_instalacion": df["tabla"].map(clean),
        "id": df["id"].map(clean),
        "nombre": df["nombre"].map(clean),
        "propietario_name": df["propietario"].map(clean),
        "observacion": df["codigo"].map(clean),
        "detalle": df["detalle"].map(clean),
        "fuente": "revision_calidad_escriturageneral.xlsx",
    })


def leer_herencia():
    ruta = os.path.join(REPORTE_DIR, "revision_herencia_nombre.xlsx")
    if not os.path.exists(ruta):
        return pd.DataFrame(columns=COLUMNAS)
    df = pd.read_excel(
        ruta, sheet_name="Detalle",
        usecols=["relacion", "codigo", "id_hijo", "nombre_hijo", "nombre_subestacion", "propietario"],
    )
    return _frame({
        "tipo_instalacion": df["relacion"].map(
            lambda r: str(r).split(" -> ")[-1] if " -> " in str(r) else str(r)
        ),
        "id": df["id_hijo"].map(clean),
        "nombre": df["nombre_hijo"].map(clean),
        "propietario_name": df["propietario"].map(clean),
        "observacion": df["codigo"].map(clean),
        "detalle": df["nombre_subestacion"].map(
            lambda x: f"No contiene el nombre de la subestación {x}"
        ),
        "fuente": "revision_herencia_nombre.xlsx",
    })


def leer_abreviatura_torre():
    ruta = os.path.join(REPORTE_DIR, "revision_abreviatura_torre.xlsx")
    if not os.path.exists(ruta):
        return pd.DataFrame(columns=COLUMNAS)
    df = pd.read_excel(
        ruta, sheet_name="Detalle",
        usecols=["tabla", "codigo", "id", "nombre", "detalle", "propietario"],
    )
    return _frame({
        "tipo_instalacion": df["tabla"].map(clean),
        "id": df["id"].map(clean),
        "nombre": df["nombre"].map(clean),
        "propietario_name": df["propietario"].map(clean),
        "observacion": df["codigo"].map(clean),
        "detalle": df["detalle"].map(clean),
        "fuente": "revision_abreviatura_torre.xlsx",
    })


def leer_abreviatura_vano():
    ruta = os.path.join(REPORTE_DIR, "revision_abreviatura_vano.xlsx")
    if not os.path.exists(ruta):
        return pd.DataFrame(columns=COLUMNAS)
    df = pd.read_excel(
        ruta, sheet_name="Detalle",
        usecols=["tabla", "codigo", "id", "nombre", "detalle", "propietario"],
    )
    return _frame({
        "tipo_instalacion": df["tabla"].map(clean),
        "id": df["id"].map(clean),
        "nombre": df["nombre"].map(clean),
        "propietario_name": df["propietario"].map(clean),
        "observacion": df["codigo"].map(clean),
        "detalle": df["detalle"].map(clean),
        "fuente": "revision_abreviatura_vano.xlsx",
    })


def leer_prefijo():
    ruta = os.path.join(REPORTE_DIR, "analisis_prefijo.xlsx")
    if not os.path.exists(ruta):
        return pd.DataFrame(columns=COLUMNAS)
    df = pd.read_excel(
        ruta, sheet_name="Detalle",
        usecols=["tabla", "codigo", "id", "nombre", "prefijos_esperados", "propietario"],
    )
    return _frame({
        "tipo_instalacion": df["tabla"].map(clean),
        "id": df["id"].map(clean),
        "nombre": df["nombre"].map(clean),
        "propietario_name": df["propietario"].map(clean),
        "observacion": df["codigo"].map(clean),
        "detalle": df["prefijos_esperados"].map(lambda x: f"Prefijos esperados: {x}"),
        "fuente": "analisis_prefijo.xlsx",
    })


def leer_tension_patios():
    ruta = os.path.join(REPORTE_DIR, "analisis_tension_patios.xlsx")
    if not os.path.exists(ruta):
        return pd.DataFrame(columns=COLUMNAS)
    df = pd.read_excel(
        ruta, sheet_name="Detalle",
        usecols=["tipo_instalacion", "codigo", "id", "nombre", "propietario", "detalle"],
    )
    return _frame({
        "tipo_instalacion": df["tipo_instalacion"].map(clean),
        "id": df["id"].map(clean),
        "nombre": df["nombre"].map(clean),
        "propietario_name": df["propietario"].map(clean),
        "observacion": df["codigo"].map(clean),
        "detalle": df["detalle"].map(clean),
        "fuente": "analisis_tension_patios.xlsx",
    })


def leer_documentos_datasheet():
    ruta = os.path.join(REPORTE_DIR, "revision_documentos_datasheet.xlsx")
    if not os.path.exists(ruta):
        return pd.DataFrame(columns=COLUMNAS)
    df = pd.read_excel(
        ruta, sheet_name="Detalle",
        usecols=["tabla", "codigo", "id", "nombre", "propietario", "detalle"],
    )
    return _frame({
        "tipo_instalacion": df["tabla"].map(clean),
        "id": df["id"].map(clean),
        "nombre": df["nombre"].map(clean),
        "propietario_name": df["propietario"].map(clean),
        "observacion": df["codigo"].map(clean),
        "detalle": df["detalle"].map(clean),
        "fuente": "revision_documentos_datasheet.xlsx",
    })


def _normalizar_id(v):
    s = clean(v).strip()
    if s == "":
        return ""
    if re.fullmatch(r"-?\d+\.0+", s):
        s = s.split(".")[0]
    return s


normalizar_id = _normalizar_id


def hash_error(tipo, id_activo, observacion):
    cadena = f"{clean(tipo).strip()}|{_normalizar_id(id_activo)}|{clean(observacion).strip()}"
    return hashlib.sha256(cadena.encode("utf-8")).hexdigest()


def cargar_hashes_justificados():
    conn = psycopg2.connect(**config.DB_ERRORES_CONN)
    try:
        cur = conn.cursor()
        cur.execute("SELECT error_hash FROM auditoria_errores WHERE justificado")
        return {r[0] for r in cur.fetchall()}
    finally:
        conn.close()


def _cargar_decreto_map(conn):
    cur = conn.cursor()
    cur.execute("SELECT id, name FROM api_decreto")
    rows = cur.fetchall()
    cur.close()
    return {r[0]: r[1] for r in rows}


def _tablas_con_decreto(conn, tablas):
    if not tablas:
        return set()
    cur = conn.cursor()
    cur.execute(
        "SELECT table_name FROM information_schema.columns "
        "WHERE table_schema='public' AND column_name='decreto_id' "
        "AND table_name = ANY(%s)",
        (list(tablas),),
    )
    res = {r[0] for r in cur.fetchall()}
    cur.close()
    return res


def agregar_decreto(df):
    """Añade la columna 'decreto' (api_decreto.name) resolviendo el decreto_id
    de cada instalación a partir de su tabla e id."""
    df = df.copy()
    df["decreto"] = ""
    if df.empty:
        return df

    df["_id_norm"] = df["id"].map(_normalizar_id)

    conn = psycopg2.connect(**config.DB_CONN)
    try:
        decreto_map = _cargar_decreto_map(conn)

        ids_por_tabla = {}
        for tipo, g in df.groupby("tipo_instalacion"):
            tabla = TIPO_INSTALACION_A_TABLA.get(tipo, tipo)
            ids = [i for i in g["_id_norm"].tolist() if i]
            if not ids:
                continue
            ids_por_tabla.setdefault(tabla, set()).update(ids)

        for tabla in _tablas_con_decreto(conn, ids_por_tabla.keys()):
            ids = list(ids_por_tabla[tabla])
            cur = conn.cursor()
            cur.execute(
                f'SELECT id::text, decreto_id FROM "{tabla}" '
                f"WHERE decreto_id IS NOT NULL AND id::text = ANY(%s)",
                (ids,),
            )
            id_a_decreto = {str(r[0]): r[1] for r in cur.fetchall()}
            cur.close()

            mask = (df["tipo_instalacion"].map(
                lambda t: TIPO_INSTALACION_A_TABLA.get(t, t)
            ) == tabla) & df["_id_norm"].isin(id_a_decreto)
            df.loc[mask, "decreto"] = df.loc[mask, "_id_norm"].map(
                lambda i: clean(decreto_map.get(id_a_decreto.get(i, ""), ""))
            )
    finally:
        conn.close()

    df.drop(columns=["_id_norm"], inplace=True)
    return df


def _filtrar_justificados(df):
    if df.empty:
        return df
    justificados = cargar_hashes_justificados()
    if not justificados:
        return df
    hashes = pd.Series(
        [
            hash_error(t, i, o)
            for t, i, o in zip(df["tipo_instalacion"], df["id"], df["observacion"])
        ],
        index=df.index,
    )
    return df[~hashes.isin(justificados)]


def consolidar_observaciones(excluir_justificados=True):
    cargadores = [
        ("analisis_desconectados", leer_analisis("analisis_desconectados.xlsx", "REGISTRO_DESCONECTADO", "Registro desconectado (relacionamiento fuera de operación)")),
        ("analisis_duplicados", leer_analisis("analisis_duplicados.xlsx", "NOMBRE_DUPLICADO", "Nombre duplicado dentro de la tabla")),
        ("analisis_conexion_nula", leer_analisis("analisis_conexion_nula.xlsx", "CONEXION_NULA", "Relación sin valor (NULL)")),
        ("enriquecidas", leer_enriquecidas()),
        ("auditoria_lineas", leer_auditoria_lineas()),
        ("auditoria_contenedores", leer_auditoria_contenedores()),
        ("escritura", leer_escritura()),
        ("herencia", leer_herencia()),
        ("prefijo", leer_prefijo()),
        ("abreviatura_torre", leer_abreviatura_torre()),
        ("abreviatura_vano", leer_abreviatura_vano()),
        ("documentos_datasheet", leer_documentos_datasheet()),
        ("tension_patios", leer_tension_patios()),
    ]

    frames = []
    for nombre, df in cargadores:
        if not df.empty:
            frames.append(df)
            print(f"{nombre}: {len(df)} observaciones")
        else:
            print(f"{nombre}: 0 observaciones")

    if not frames:
        print("No hay observaciones para consolidar.")
        return pd.DataFrame(columns=COLUMNAS)

    df = pd.concat(frames, ignore_index=True)
    df["tipo_instalacion"] = df["tipo_instalacion"].map(
        lambda t: TIPO_INSTALACION_A_TABLA.get(t, t)
    )
    df["propietario_name"] = df["propietario_name"].replace("", "Sin empresa")
    if excluir_justificados:
        df = _filtrar_justificados(df)
    return df[COLUMNAS]


def main():
    df = consolidar_observaciones()
    if df.empty:
        return

    df = agregar_decreto(df)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    carpeta = os.path.join(DOCS_DIR, "ReporteExcelEmpresa", ts + config.sufijo_carpeta_bd())
    os.makedirs(carpeta, exist_ok=True)

    n_empresas = 0
    for empresa, grp in df.groupby("propietario_name"):
        ruta = os.path.join(carpeta, sanitizar(empresa) + ".xlsx")
        with pd.ExcelWriter(ruta, engine="xlsxwriter") as w:
            grp[COLUMNAS + ["decreto"]].to_excel(w, sheet_name="Observaciones", index=False)
        n_empresas += 1

    print()
    print(f"Total observaciones: {len(df)}")
    print(f"Empresas con observaciones: {n_empresas}")
    print(f"Carpeta: {carpeta}")


if __name__ == "__main__":
    main()
