# -*- coding: utf-8 -*-
"""
Convierte un archivo de texto plano (Markdown o .txt) a un documento Word
(.docx). Permite elegir el archivo de forma interactiva (listando los .md/.txt
de esta carpeta) o indicarlo por argumento.

Soporta: encabezados, párrafos, listas con y sin numeración, tablas, citas,
reglas horizontales y formato en línea (negrita, cursiva, código y enlaces).

Uso:
    python utils/scripts/md2word.py
    python utils/scripts/md2word.py "documento.md" [salida.docx] [--pdf]

Los .md/.txt se buscan en la raíz del proyecto y, si no se indica otra ruta de
salida, el .docx se genera en `generados/`. Con `--pdf` se convierte además a
PDF (requiere MS Word instalado).
"""

import os
import re
import sys

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
BUSQUEDA_DIR = BASE_DIR
GENERADOS_DIR = os.path.join(BASE_DIR, "generados")

INLINE_RE = re.compile(
    r"(\[([^\]]+)\]\(([^)]+)\))"
    r"|(\*\*(.+?)\*\*)"
    r"|(?<!\w)(__(.+?)__)(?!\w)"
    r"|(?<!\*)(\*([^*]+)\*)(?!\*)"
    r"|(?<!\w)(_([^_]+)_)(?!\w)"
    r"|(`([^`]+)`)"
)

SEPARADOR_RE = re.compile(r"^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)*\|?\s*$")
HR_RE = re.compile(r"^\s*([-*_])(\s*\1){2,}\s*$")
HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")
UL_RE = re.compile(r"^\s*[-*+]\s+(.*)$")
OL_RE = re.compile(r"^\s*\d+[.)]\s+(.*)$")
CERCA_RE = re.compile(r"^\s*(```|~~~)")


def leer_texto(ruta):
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            with open(ruta, "r", encoding=enc) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    with open(ruta, "r", encoding="utf-8", errors="replace") as f:
        return f.read()


def add_hyperlink(paragraph, url, texto):
    r_id = paragraph.part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    rpr.append(color)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    rpr.append(underline)
    run.append(rpr)
    texto_el = OxmlElement("w:t")
    texto_el.set(qn("xml:space"), "preserve")
    texto_el.text = texto
    run.append(texto_el)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def add_run(paragraph, texto, bold=False, italic=False, mono=False):
    run = paragraph.add_run(texto)
    if bold:
        run.bold = True
    if italic:
        run.italic = True
    if mono:
        run.font.name = "Consolas"
    return run


def add_inline(paragraph, texto, bold=False, italic=False):
    pos = 0
    for m in INLINE_RE.finditer(texto):
        if m.start() > pos:
            add_run(paragraph, texto[pos:m.start()], bold, italic)
        if m.group(1):
            add_hyperlink(paragraph, m.group(3), m.group(2))
        elif m.group(4):
            add_inline(paragraph, m.group(5), True, italic)
        elif m.group(6):
            add_inline(paragraph, m.group(7), True, italic)
        elif m.group(8):
            add_inline(paragraph, m.group(9), bold, True)
        elif m.group(10):
            add_inline(paragraph, m.group(11), bold, True)
        elif m.group(12):
            add_run(paragraph, m.group(13), bold, italic, mono=True)
        pos = m.end()
    if pos < len(texto):
        add_run(paragraph, texto[pos:], bold, italic)


def es_fila_tabla(linea):
    return linea.strip().startswith("|") and linea.strip().endswith("|")


def celdas(linea):
    return [c.strip() for c in linea.strip().strip("|").split("|")]


def add_tabla(doc, filas_md):
    encabezado = celdas(filas_md[0])
    cuerpo = [celdas(f) for f in filas_md[2:]]
    tabla = doc.add_table(rows=1 + len(cuerpo), cols=len(encabezado))
    try:
        tabla.style = "Table Grid"
    except KeyError:
        pass
    for j, texto in enumerate(encabezado):
        parrafo = tabla.cell(0, j).paragraphs[0]
        add_inline(parrafo, texto)
        for run in parrafo.runs:
            run.bold = True
    for i, fila in enumerate(cuerpo, start=1):
        for j in range(len(encabezado)):
            texto = fila[j] if j < len(fila) else ""
            add_inline(tabla.cell(i, j).paragraphs[0], texto)
    doc.add_paragraph()


def add_horizontal(doc):
    parrafo = doc.add_paragraph()
    ppr = parrafo._p.get_or_add_pPr()
    borde = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "999999")
    borde.append(bottom)
    ppr.append(borde)


def convertir(ruta_entrada, ruta_salida):
    lineas = leer_texto(ruta_entrada).splitlines()
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)

    i = 0
    n = len(lineas)
    while i < n:
        linea = lineas[i]

        if CERCA_RE.match(linea):
            i += 1
            while i < n and not CERCA_RE.match(lineas[i]):
                run = doc.add_paragraph().add_run(lineas[i])
                run.font.name = "Consolas"
                i += 1
            i += 1
            continue

        if es_fila_tabla(linea) and i + 1 < n and SEPARADOR_RE.match(lineas[i + 1]):
            filas = []
            while i < n and es_fila_tabla(lineas[i]):
                filas.append(lineas[i])
                i += 1
            add_tabla(doc, filas)
            continue

        m = HEADING_RE.match(linea)
        if m:
            parrafo = doc.add_heading("", level=len(m.group(1)))
            add_inline(parrafo, m.group(2).strip())
            i += 1
            continue

        if HR_RE.match(linea):
            add_horizontal(doc)
            i += 1
            continue

        if linea.strip().startswith(">"):
            bloque = []
            while i < n and lineas[i].strip().startswith(">"):
                bloque.append(lineas[i].strip()[1:].strip())
                i += 1
            add_inline(doc.add_paragraph(style="Quote"), " ".join(x for x in bloque if x))
            continue

        if UL_RE.match(linea):
            while i < n and UL_RE.match(lineas[i]):
                add_inline(doc.add_paragraph(style="List Bullet"), UL_RE.match(lineas[i]).group(1))
                i += 1
            continue

        if OL_RE.match(linea):
            while i < n and OL_RE.match(lineas[i]):
                add_inline(doc.add_paragraph(style="List Number"), OL_RE.match(lineas[i]).group(1))
                i += 1
            continue

        if not linea.strip():
            i += 1
            continue

        bloque = []
        while i < n and lineas[i].strip() and not es_fila_tabla(lineas[i]) \
                and not HEADING_RE.match(lineas[i]) and not HR_RE.match(lineas[i]) \
                and not lineas[i].strip().startswith(">") \
                and not UL_RE.match(lineas[i]) and not OL_RE.match(lineas[i]) \
                and not CERCA_RE.match(lineas[i]):
            bloque.append(lineas[i].strip())
            i += 1
        parrafo = doc.add_paragraph()
        for k, texto_linea in enumerate(bloque):
            if k:
                parrafo.add_run().add_break()
            add_inline(parrafo, texto_linea)

    doc.save(ruta_salida)
    return ruta_salida


def elegir_archivo():
    archivos = sorted(f for f in os.listdir(BUSQUEDA_DIR) if f.lower().endswith((".md", ".txt")))
    for idx, nombre in enumerate(archivos, start=1):
        print(f"  {idx}. {nombre}")
    respuesta = input("Elige un número o escribe la ruta del archivo: ").strip().strip('"')
    if respuesta.isdigit() and 1 <= int(respuesta) <= len(archivos):
        return os.path.join(BUSQUEDA_DIR, archivos[int(respuesta) - 1])
    return respuesta


def convertir_a_pdf(ruta_docx):
    import win32com.client

    ruta_pdf = os.path.splitext(ruta_docx)[0] + ".pdf"
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        documento = word.Documents.Open(ruta_docx)
        try:
            documento.SaveAs(ruta_pdf, FileFormat=17)
        finally:
            documento.Close(False)
    finally:
        word.Quit()
    return ruta_pdf


def main():
    args = sys.argv[1:]
    generar_pdf = "--pdf" in args
    args = [a for a in args if a != "--pdf"]
    if args:
        entrada = args[0]
    else:
        print("Archivos disponibles:")
        entrada = elegir_archivo()

    if not entrada or not os.path.isfile(entrada):
        print(f"No se encontró el archivo: {entrada}")
        sys.exit(1)

    if len(args) > 1:
        salida = args[1]
    else:
        os.makedirs(GENERADOS_DIR, exist_ok=True)
        salida = os.path.join(GENERADOS_DIR, os.path.basename(os.path.splitext(entrada)[0]) + ".docx")
    destino = convertir(entrada, salida)
    print(f"Documento generado: {destino}")

    if generar_pdf:
        print(f"PDF generado: {convertir_a_pdf(destino)}")


if __name__ == "__main__":
    main()
