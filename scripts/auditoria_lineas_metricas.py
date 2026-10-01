# -*- coding: utf-8 -*-
"""
Auditoría de calidad del árbol de relacionamiento de líneas (Línea -> Circuito
-> Tramo -> Vano/Túnel -> Accesorio) usando la BD PostgreSQL local (bdatx).

Reutiliza la lógica de reglas de arbol_lineas_v11_2_propietario_corregido.py,
pero sin interfaz gráfica y sin consultas a la API: los registros y sus
detalles se obtienen desde las tablas ya cargadas en PostgreSQL.

Salidas:
    docs/reporte/auditoria_lineas_metricas.xlsx
    docs/reporte/auditoria_lineas_metricas_hallazgos.csv
    docs/historico/auditoria_lineas_metricas-<YYYY-MM-DD_HH-MM-SS>.xlsx
    docs/historico/auditoria_lineas_metricas_hallazgos-<YYYY-MM-DD_HH-MM-SS>.csv
"""

import os
import re
import sys
import csv
from datetime import datetime
from collections import defaultdict

import pandas as pd
import psycopg2

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE_DIR)

import config

REPORTE_DIR = config.REPORTE_DIR
HISTORICO_DIR = config.HISTORICO_DIR
DB_CONN = config.DB_CONN

STATUS_EN_OPERACION = "EN_OPERACION"

TIPO_LINEA = "lineas"
TIPO_CIRCUITO = "circuitos"
TIPO_TRAMO = "tramos"
TIPO_VANO = "vanos"
TIPO_TUNEL = "tuneles"
TIPO_ESTACION = "estaciones-repetidoras"
TIPO_ACCESORIO = "accesorios-vanos"


class AuditoriaLineas:
    """Porta las reglas de auditoría del árbol de líneas a datos de PostgreSQL."""

    def __init__(self, records, details, loaded_depth=4):
        self.records = records
        self.details = details
        self.loaded_depth = loaded_depth
        self.audit_required = True
        self.audit_nodes = True

    # ----------------------------------------------------------------
    # Helpers
    # ----------------------------------------------------------------

    @staticmethod
    def _normalize_name(text):
        return " ".join(str(text or "").strip().upper().split())

    @staticmethod
    def _norm_field_name(value):
        return re.sub(r"[^a-z0-9]", "", str(value or "").lower())

    def _field_value(self, record, candidates):
        if not isinstance(record, dict):
            return None
        wanted = {self._norm_field_name(x) for x in candidates}
        for key, value in record.items():
            if self._norm_field_name(key) in wanted:
                return value
        return None

    @staticmethod
    def _value_id(value):
        if value is None:
            return ""
        if isinstance(value, dict):
            for key in ("id", "pk", "value"):
                if value.get(key) is not None:
                    return str(value.get(key))
            return ""
        if isinstance(value, list):
            if not value:
                return ""
            return AuditoriaLineas._value_id(value[0])
        return str(value).strip()

    @staticmethod
    def _value_name(value):
        if isinstance(value, dict):
            for key in ("name", "nombre", "label"):
                if value.get(key):
                    return str(value.get(key))
        return ""

    def _audit_rule_catalog(self):
        return [
            {"codigo": "CONTENEDOR_VACIO", "aplica_a": "Subestación/Patio/Paño/SSGG/Armario/Línea/Circuito/Tramo", "descripcion": "Contenedor sin ninguna instalación aguas abajo.", "criterio": "El contenedor principal debe poseer al menos una instalación en el nivel siguiente."},
            {"codigo": "SUFIJO_NOMBRE", "aplica_a": "Circuito/Tramo", "descripcion": "Inconsistencia en los sufijos Cn (C1/C2) del nombre.", "criterio": "El sufijo Cn del Tramo debe coincidir con el del Circuito padre, y no puede existir C2 o superior sin C1."},
            {"codigo": "HERENCIA_LINEA", "aplica_a": "Circuito", "descripcion": "El nombre del Circuito no hereda el nombre de la Línea padre.", "criterio": "El Nombre de Circuito debe contener el nombre de la Línea (sin los prefijos LT/CTO)."},
            {"codigo": "VANO_SIN_ACCESORIOS", "aplica_a": "Vano", "descripcion": "Vano sin registro de Accesorios Vanos.", "criterio": "La guía indica que cada vano cuenta con un registro de accesorios vanos."},
            {"codigo": "VANO_SIN_NODO1", "aplica_a": "Vano", "descripcion": "Vano sin Nodo 1 detectable.", "criterio": "Nodo 1 es obligatorio para el Vano."},
            {"codigo": "VANO_SIN_NODO2", "aplica_a": "Vano", "descripcion": "Vano sin Nodo 2 detectable.", "criterio": "Nodo 2 es obligatorio para el Vano."},
            {"codigo": "TRAMO_NODOS_IGUALES", "aplica_a": "Tramo", "descripcion": "Nodo 1 y Nodo 2 son el mismo registro.", "criterio": "Los extremos del Tramo deben ser distintos."},
            {"codigo": "CIRCUITO_TRAMOS_DESCONECTADOS", "aplica_a": "Circuito", "descripcion": "Los Tramos forman más de un componente desconectado.", "criterio": "Los segmentos de un Circuito deben formar una secuencia continua."},
            {"codigo": "VANO_MULTIPLES_ACCESORIOS", "aplica_a": "Vano", "descripcion": "Más de un registro Accesorios Vanos para un mismo Vano.", "criterio": "La guía indica un único registro de accesorios por cada vano."},
        ]

    # ----------------------------------------------------------------
    # Reglas básicas (jerarquía, integridad, unicidad, nomenclatura)
    # ----------------------------------------------------------------

    def audit_records(self):
        records = self.records
        findings = []
        if not records:
            return findings

        by_iid = {r["iid"]: r for r in records if r.get("iid")}

        children = defaultdict(list)
        for r in records:
            children[r["parent_iid"]].append(r)

        def add(code, record, message, rule):
            findings.append({
                "codigo": code,
                "nivel": record.get("nivel", ""),
                "tipo": record.get("tipo", ""),
                "id": record.get("id", ""),
                "nombre": record.get("nombre", ""),
                "propietario_id": record.get("propietario_id", ""),
                "propietario": record.get("propietario", ""),
                "padre": f'{record.get("parent_tipo","")} {record.get("parent_id","")} {record.get("parent_nombre","")}'.strip(),
                "mensaje": message,
                "regla": rule,
                "iid": record.get("iid", ""),
            })

        # Línea sin circuitos
        for r in records:
            if r["tipo"].strip().lower() == TIPO_LINEA:
                circuit_children = [c for c in children.get(r["iid"], []) if c["tipo"].strip().lower() == TIPO_CIRCUITO]
                if self.loaded_depth >= 1 and not circuit_children:
                    add("CONTENEDOR_VACIO", r, "Línea sin ningún Circuito relacionado; falta el nivel siguiente.", "El contenedor principal Línea debe poseer al menos un Circuito.")

        # Circuito sin tramos
        for r in records:
            if r["tipo"].strip().lower() == TIPO_CIRCUITO:
                tramo_children = [c for c in children.get(r["iid"], []) if c["tipo"].strip().lower() == TIPO_TRAMO]
                if self.loaded_depth >= 2 and not tramo_children:
                    add("CONTENEDOR_VACIO", r, "Circuito sin ningún tramo relacionado.", "La guía indica que todos los circuitos poseen al menos un tramo.")

        # C2 o superior sin C1
        for line in records:
            if line["tipo"].strip().lower() != TIPO_LINEA:
                continue
            circuitos = [c for c in children.get(line["iid"], []) if c["tipo"].strip().lower() == TIPO_CIRCUITO]
            nums = []
            for c in circuitos:
                m = re.search(r"\bC(?:IRCUITO\s*)?(\d+)\b", self._normalize_name(c["nombre"]))
                if m:
                    nums.append((int(m.group(1)), c))
            if nums and any(n >= 2 for n, _ in nums) and not any(n == 1 for n, _ in nums):
                for n, c in nums:
                    if n >= 2:
                        add("SUFIJO_NOMBRE", c, "Se detectó Circuito 2 o superior sin Circuito 1 bajo la misma línea.", "La guía prohíbe C2 o superior si no existe C1.")

        # Coherencia de nomenclatura Circuito -> Tramo
        def final_circuit_designator(name):
            text = str(name or "").strip().upper()
            match = re.search(r"(?:^|[\s\-\(\)])(C\d+)\s*$", text)
            return match.group(1) if match else ""

        def circuit_designator(name):
            text = str(name or "").strip().upper()
            match = re.search(r"(?:^|[\s\-\(\)])(C\d+)\s*$", text)
            if match:
                return match.group(1)
            tokens = re.findall(r"(?<![A-Z0-9])C(\d+)(?![A-Z0-9])", text)
            if tokens:
                return "C" + tokens[-1]
            return ""

        for r in records:
            if r["tipo"].strip().lower() != TIPO_TRAMO:
                continue
            parent = by_iid.get(r.get("parent_iid", ""))
            if not parent:
                continue
            if parent["tipo"].strip().lower() != TIPO_CIRCUITO:
                continue
            parent_circuit = circuit_designator(parent["nombre"])
            tramo_circuit = final_circuit_designator(r["nombre"])
            if parent_circuit and tramo_circuit and parent_circuit != tramo_circuit:
                add("SUFIJO_NOMBRE", r,
                    f'Error de relacionamiento por nomenclatura: el Circuito padre "{parent["nombre"]}" es {parent_circuit}, pero el Tramo "{r["nombre"]}" termina en {tramo_circuit}.',
                    "La designación C1/C2 del Tramo debe coincidir con la designación C1/C2 del Circuito padre.")

        # Tramo sin instalación del nivel siguiente
        for r in records:
            if r["tipo"].strip().lower() == TIPO_TRAMO:
                next_children = children.get(r["iid"], [])
                if self.loaded_depth >= 3 and not next_children:
                    add("CONTENEDOR_VACIO", r, "Tramo sin instalaciones relacionadas en el nivel siguiente; falta Vano para tramo aéreo o Túnel para tramo subterráneo.", "El contenedor Tramo debe continuar con la instalación correspondiente al tipo de tramo.")

        # Accesorios por vano
        accessory_types = {"accesorios-vanos", "accesorios_vanos", "accesoriosvanos"}
        for r in records:
            if r["tipo"].strip().lower() != TIPO_VANO:
                continue
            acc = [c for c in children.get(r["iid"], []) if c["tipo"].strip().lower() in accessory_types]
            if self.loaded_depth >= 4 and len(acc) == 0:
                add("VANO_SIN_ACCESORIOS", r, "Vano sin registro de Accesorios Vanos; falta la instalación del nivel siguiente.", "La guía indica que cada vano cuenta con un registro de accesorios vanos.")
            elif len(acc) > 1:
                add("VANO_MULTIPLES_ACCESORIOS", r, f"Se detectaron {len(acc)} registros de Accesorios Vanos.", "La guía indica un único registro de accesorios por cada vano.")

        # Herencia del nombre de la Línea en el Circuito (sin prefijos LT/CTO)
        def _sin_prefijo(name, prefijo):
            n = self._normalize_name(name)
            if n.startswith(prefijo):
                return n[len(prefijo):].strip()
            return n

        for r in records:
            if r["tipo"].strip().lower() != TIPO_CIRCUITO:
                continue
            parent = by_iid.get(r.get("parent_iid", ""))
            if not parent or parent["tipo"].strip().lower() != TIPO_LINEA:
                continue
            linea_core = _sin_prefijo(parent["nombre"], "LT ")
            circuito_core = _sin_prefijo(r["nombre"], "CTO ")
            if linea_core and linea_core not in circuito_core:
                add("HERENCIA_LINEA", r,
                    f'El nombre del Circuito "{r["nombre"]}" no hereda el nombre de la Línea padre "{parent["nombre"]}".',
                    "El nombre del Circuito debe contener el nombre de la Línea (sin los prefijos LT/CTO).")

        return findings

    # ----------------------------------------------------------------
    # Auditoría profunda (campos obligatorios, IDs, nodos)
    # ----------------------------------------------------------------

    def _append_deep_audit_findings(self, findings, records):
        if not self.details:
            return findings

        children = defaultdict(list)
        for r in records:
            children[r.get("parent_iid", "")].append(r)

        def add(code_value, record, message, rule):
            findings.append({
                "codigo": code_value,
                "nivel": record.get("nivel", ""),
                "tipo": record.get("tipo", ""),
                "id": record.get("id", ""),
                "nombre": record.get("nombre", ""),
                "propietario_id": record.get("propietario_id", ""),
                "propietario": record.get("propietario", ""),
                "padre": f'{record.get("parent_tipo","")} {record.get("parent_id","")} {record.get("parent_nombre","")}'.strip(),
                "mensaje": message,
                "regla": rule,
                "iid": record.get("iid", ""),
            })

        # Campos obligatorios según tipo
        if self.audit_required:
            required = {
                TIPO_VANO: [
                    ("VANO_SIN_NODO1", ("nodo1", "nodo_1", "nodo1_id", "nodoorigen"), "Nodo 1"),
                    ("VANO_SIN_NODO2", ("nodo2", "nodo_2", "nodo2_id", "nododestino"), "Nodo 2"),
                ],
            }

            for r in records:
                tipo = r["tipo"].strip().lower()
                detail = self.details.get((r["tipo"], r["id"]), {})
                if not detail:
                    continue
                for error_code, candidates, label in required.get(tipo, []):
                    value = self._field_value(detail, candidates)
                    if value is None or value == "" or value == [] or value == {}:
                        add(error_code, r, f'No se detectó el campo obligatorio "{label}" en el detalle del registro.', f"Validación de campo obligatorio para {tipo}.")

        # Nodos y continuidad de tramos
        if self.audit_nodes:
            node_info = {}
            for r in records:
                if r["tipo"].strip().lower() != TIPO_TRAMO:
                    continue
                detail = self.details.get((r["tipo"], r["id"]), {})
                if not detail:
                    continue
                n1_value = self._field_value(detail, ("nodo1", "nodo_1", "nodo1_id", "nodoorigen"))
                n2_value = self._field_value(detail, ("nodo2", "nodo_2", "nodo2_id", "nododestino"))
                n1 = self._value_id(n1_value)
                n2 = self._value_id(n2_value)
                node_info[r["iid"]] = (n1, n2)

                if n1 and n2 and n1 == n2:
                    add("TRAMO_NODOS_IGUALES", r, f"El Tramo posee el mismo Nodo 1 y Nodo 2 ({n1}).", "Los extremos de un Tramo deben representar dos puntos distintos.")

            # Conectividad por circuito
            for circuit in records:
                if circuit["tipo"].strip().lower() != TIPO_CIRCUITO:
                    continue
                tramo_records = [child for child in children.get(circuit["iid"], []) if child["tipo"].strip().lower() == TIPO_TRAMO]
                edges = []
                for tramo in tramo_records:
                    n1, n2 = node_info.get(tramo["iid"], ("", ""))
                    if n1 and n2:
                        edges.append((n1, n2, tramo))
                if len(edges) < 2:
                    continue
                graph = defaultdict(set)
                for n1, n2, _ in edges:
                    graph[n1].add(n2)
                    graph[n2].add(n1)
                unvisited = set(graph)
                components = 0
                while unvisited:
                    components += 1
                    stack = [unvisited.pop()]
                    while stack:
                        node = stack.pop()
                        for neighbor in graph[node]:
                            if neighbor in unvisited:
                                unvisited.remove(neighbor)
                                stack.append(neighbor)
                if components > 1:
                    add("CIRCUITO_TRAMOS_DESCONECTADOS", circuit, f"Los Tramos con información de nodos forman {components} grupos desconectados.", "Los segmentos de un Circuito deben formar una secuencia física continua.")

        return findings

    def run(self):
        findings = self.audit_records()
        findings = self._append_deep_audit_findings(findings, self.records)
        findings.sort(key=lambda x: (
            x["nivel"] if isinstance(x["nivel"], int) else 99,
            x["tipo"],
            x["id"],
            x["codigo"],
        ))
        return findings


# ============================================================
# CARGA DE DATOS DESDE POSTGRESQL
# ============================================================

def cargar(conn):
    def q(sql):
        cur = conn.cursor()
        cur.execute(sql)
        rows = cur.fetchall()
        cur.close()
        return rows

    lineas = q("SELECT id, name, propietario_id FROM api_linea WHERE status='EN_OPERACION'")
    circuitos = q("SELECT id, name, linea_id, propietario_id FROM api_circuito WHERE status='EN_OPERACION'")
    tramos = q("SELECT id, name, circuito_id, nodo1_id, nodo2_id, decreto_id, categoria_vu_id, propietario_id FROM api_tramo WHERE status='EN_OPERACION'")
    vanos = q("SELECT id, name, tramo_id, nodo1_id, nodo2_id, decreto_id, categoria_vu_id, propietario_id FROM api_vano WHERE status='EN_OPERACION'")
    tuneles = q("SELECT id, name, tramo_id, propietario_id FROM api_tunel WHERE status='EN_OPERACION'")
    estaciones = q("SELECT id, name, tramo_id, nodo_id, propietario_id FROM api_estacionrepetidora WHERE status='EN_OPERACION'")
    accesorios = q("SELECT id, name, vano_id, decreto_id, categoria_vu_id, propietario_id FROM api_accesoriovano WHERE status='EN_OPERACION'")
    empresas = q("SELECT id, name FROM api_empresa")
    nodos = q("SELECT id, name FROM api_nodo")

    empresa_map = {r[0]: r[1] for r in empresas}
    nodo_map = {r[0]: r[1] for r in nodos}

    return lineas, circuitos, tramos, vanos, tuneles, estaciones, accesorios, empresa_map, nodo_map


def construir_registros(data):
    lineas, circuitos, tramos, vanos, tuneles, estaciones, accesorios, empresa_map, nodo_map = data

    records = []
    details = {}
    counter = [0]

    def nueva():
        counter[0] += 1
        return str(counter[0])

    def emp(pid):
        return "" if pid is None else str(empresa_map.get(pid, ""))

    def nn(nid):
        return "" if nid is None else str(nodo_map.get(nid, ""))

    circuitos_by_linea = defaultdict(list)
    for c in circuitos:
        circuitos_by_linea[c[2]].append(c)

    tramos_by_circuito = defaultdict(list)
    for t in tramos:
        tramos_by_circuito[t[2]].append(t)

    vanos_by_tramo = defaultdict(list)
    for v in vanos:
        vanos_by_tramo[v[2]].append(v)

    tuneles_by_tramo = defaultdict(list)
    for t in tuneles:
        tuneles_by_tramo[t[2]].append(t)

    estaciones_by_tramo = defaultdict(list)
    for e in estaciones:
        estaciones_by_tramo[e[2]].append(e)

    accesorios_by_vano = defaultdict(list)
    for a in accesorios:
        accesorios_by_vano[a[2]].append(a)

    def det_relacion(fk, nombre=""):
        return {"id": fk, "name": nombre}

    for ln in lineas:
        lid = ln[0]
        linea = {
            "iid": nueva(),
            "parent_iid": "",
            "nivel": 0,
            "tipo": TIPO_LINEA,
            "id": str(lid),
            "nombre": ln[1] or "",
            "propietario_id": "" if ln[2] is None else str(ln[2]),
            "propietario": emp(ln[2]),
            "parent_tipo": "",
            "parent_id": "",
            "parent_nombre": "",
        }
        records.append(linea)

        for c in circuitos_by_linea.get(lid, []):
            cid = c[0]
            circ = {
                "iid": nueva(),
                "parent_iid": linea["iid"],
                "nivel": 1,
                "tipo": TIPO_CIRCUITO,
                "id": str(cid),
                "nombre": c[1] or "",
                "propietario_id": "" if c[3] is None else str(c[3]),
                "propietario": emp(c[3]),
                "parent_tipo": TIPO_LINEA,
                "parent_id": str(lid),
                "parent_nombre": linea["nombre"],
            }
            records.append(circ)

            for t in tramos_by_circuito.get(cid, []):
                tid = t[0]
                tramo = {
                    "iid": nueva(),
                    "parent_iid": circ["iid"],
                    "nivel": 2,
                    "tipo": TIPO_TRAMO,
                    "id": str(tid),
                    "nombre": t[1] or "",
                    "propietario_id": "" if t[7] is None else str(t[7]),
                    "propietario": emp(t[7]),
                    "parent_tipo": TIPO_CIRCUITO,
                    "parent_id": str(cid),
                    "parent_nombre": circ["nombre"],
                }
                records.append(tramo)

                det = {}
                if t[2] is not None:
                    det["circuito"] = det_relacion(t[2])
                if t[3] is not None:
                    det["nodo1"] = det_relacion(t[3], nn(t[3]))
                if t[4] is not None:
                    det["nodo2"] = det_relacion(t[4], nn(t[4]))
                if t[5] is not None:
                    det["decreto"] = det_relacion(t[5])
                if t[6] is not None:
                    det["categoria_vu"] = det_relacion(t[6])
                details[(TIPO_TRAMO, str(tid))] = det

                for v in vanos_by_tramo.get(tid, []):
                    vid = v[0]
                    vano = {
                        "iid": nueva(),
                        "parent_iid": tramo["iid"],
                        "nivel": 3,
                        "tipo": TIPO_VANO,
                        "id": str(vid),
                        "nombre": v[1] or "",
                        "propietario_id": "" if v[7] is None else str(v[7]),
                        "propietario": emp(v[7]),
                        "parent_tipo": TIPO_TRAMO,
                        "parent_id": str(tid),
                        "parent_nombre": tramo["nombre"],
                    }
                    records.append(vano)

                    vdet = {}
                    if v[2] is not None:
                        vdet["tramo"] = det_relacion(v[2])
                    if v[3] is not None:
                        vdet["nodo1"] = det_relacion(v[3], nn(v[3]))
                    if v[4] is not None:
                        vdet["nodo2"] = det_relacion(v[4], nn(v[4]))
                    if v[5] is not None:
                        vdet["decreto"] = det_relacion(v[5])
                    if v[6] is not None:
                        vdet["categoria_vu"] = det_relacion(v[6])
                    details[(TIPO_VANO, str(vid))] = vdet

                    for a in accesorios_by_vano.get(vid, []):
                        aid = a[0]
                        acc = {
                            "iid": nueva(),
                            "parent_iid": vano["iid"],
                            "nivel": 4,
                            "tipo": TIPO_ACCESORIO,
                            "id": str(aid),
                            "nombre": a[1] or "",
                            "propietario_id": "" if a[5] is None else str(a[5]),
                            "propietario": emp(a[5]),
                            "parent_tipo": TIPO_VANO,
                            "parent_id": str(vid),
                            "parent_nombre": vano["nombre"],
                        }
                        records.append(acc)

                        adet = {}
                        if a[2] is not None:
                            adet["vano"] = det_relacion(a[2])
                        if a[3] is not None:
                            adet["decreto"] = det_relacion(a[3])
                        if a[4] is not None:
                            adet["categoria_vu"] = det_relacion(a[4])
                        details[(TIPO_ACCESORIO, str(aid))] = adet

                for t in tuneles_by_tramo.get(tid, []):
                    tuid = t[0]
                    records.append({
                        "iid": nueva(),
                        "parent_iid": tramo["iid"],
                        "nivel": 3,
                        "tipo": TIPO_TUNEL,
                        "id": str(tuid),
                        "nombre": t[1] or "",
                        "propietario_id": "" if t[3] is None else str(t[3]),
                        "propietario": emp(t[3]),
                        "parent_tipo": TIPO_TRAMO,
                        "parent_id": str(tid),
                        "parent_nombre": tramo["nombre"],
                    })

                for e in estaciones_by_tramo.get(tid, []):
                    eid = e[0]
                    records.append({
                        "iid": nueva(),
                        "parent_iid": tramo["iid"],
                        "nivel": 3,
                        "tipo": TIPO_ESTACION,
                        "id": str(eid),
                        "nombre": e[1] or "",
                        "propietario_id": "" if e[4] is None else str(e[4]),
                        "propietario": emp(e[4]),
                        "parent_tipo": TIPO_TRAMO,
                        "parent_id": str(tid),
                        "parent_nombre": tramo["nombre"],
                    })

    return records, details


# ============================================================
# EXPORTACIÓN
# ============================================================

COLUMNAS_HALLAZGO = ["codigo", "nivel", "tipo", "id", "nombre", "propietario_id", "propietario", "padre", "mensaje", "regla"]


def exportar(findings, records, catalog):
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    df_findings = pd.DataFrame(findings, columns=COLUMNAS_HALLAZGO)

    total_registros = len(records)

    # Resumen
    resumen = pd.DataFrame([
        {"Indicador": "Registros auditados", "Valor": total_registros},
        {"Indicador": "Total hallazgos", "Valor": len(findings)},
    ])

    # Por regla (incluye reglas con 0)
    counters = defaultdict(int)
    for f in findings:
        counters[f["codigo"]] += 1

    por_regla = []
    for rule in catalog:
        por_regla.append({
            "codigo": rule["codigo"],
            "aplica_a": rule["aplica_a"],
            "cantidad": counters.get(rule["codigo"], 0),
            "descripcion": rule["descripcion"],
        })
    df_regla = pd.DataFrame(por_regla)

    # Por propietario
    prop_stats = defaultdict(int)
    for f in findings:
        prop_stats[(f["propietario_id"], f["propietario"])] += 1
    por_propietario = []
    for (pid, pname), n in sorted(prop_stats.items(), key=lambda kv: -kv[1]):
        por_propietario.append({
            "propietario_id": pid,
            "propietario": pname,
            "total_hallazgos": n,
        })
    df_propietario = pd.DataFrame(por_propietario)

    # Por tipo
    tipo_stats = defaultdict(int)
    for f in findings:
        tipo_stats[f["tipo"]] += 1
    por_tipo = []
    for tipo, n in sorted(tipo_stats.items(), key=lambda kv: -kv[1]):
        por_tipo.append({
            "tipo": tipo,
            "total_hallazgos": n,
        })
    df_tipo = pd.DataFrame(por_tipo)

    os.makedirs(REPORTE_DIR, exist_ok=True)
    os.makedirs(HISTORICO_DIR, exist_ok=True)

    base = "auditoria_lineas_metricas"
    xlsx_reporte = os.path.join(REPORTE_DIR, f"{base}.xlsx")
    xlsx_historico = os.path.join(HISTORICO_DIR, f"{base}-{ts}{config.sufijo_historico_bd()}.xlsx")
    csv_reporte = os.path.join(REPORTE_DIR, f"{base}_hallazgos.csv")
    csv_historico = os.path.join(HISTORICO_DIR, f"{base}_hallazgos-{ts}{config.sufijo_historico_bd()}.csv")

    for ruta in (xlsx_reporte, xlsx_historico):
        with pd.ExcelWriter(ruta, engine="xlsxwriter") as w:
            resumen.to_excel(w, sheet_name="Resumen", index=False)
            df_regla.to_excel(w, sheet_name="Por_Regla", index=False)
            df_propietario.to_excel(w, sheet_name="Por_Propietario", index=False)
            df_tipo.to_excel(w, sheet_name="Por_Tipo", index=False)
            df_findings.to_excel(w, sheet_name="Hallazgos", index=False)

    for ruta in (csv_reporte, csv_historico):
        with open(ruta, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(COLUMNAS_HALLAZGO)
            for f in findings:
                writer.writerow([f[c] for c in COLUMNAS_HALLAZGO])

    print(f"Registros auditados: {total_registros}")
    print(f"Hallazgos: {len(findings)}")
    print(f"Excel reporte: {xlsx_reporte}")
    print(f"CSV reporte:   {csv_reporte}")


def main():
    conn = psycopg2.connect(**DB_CONN)
    try:
        data = cargar(conn)
    finally:
        conn.close()

    records, details = construir_registros(data)
    auditor = AuditoriaLineas(records, details, loaded_depth=4)
    findings = auditor.run()
    catalog = auditor._audit_rule_catalog()

    exportar(findings, records, catalog)


if __name__ == "__main__":
    main()
