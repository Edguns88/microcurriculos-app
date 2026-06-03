
from pathlib import Path
import re
import json
import fitz
import pdfplumber
from unidecode import unidecode
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph


FUENTE_CONTENIDO = "Arial Narrow"
TAMANO_CONTENIDO = 8

MODALIDAD_DEFAULT = "Presencial"
TIPOLOGIA_DEFAULT = "Teórico"

LLENAR_RA_PROGRAMA_SOLO_SI_APARECE_EXPLICITO = True
INCLUIR_CONTINUIDAD_EN_PRESENTACION = True
COPIAR_COMPETENCIAS_EN_ITEM_2 = True


def limpiar_basura(texto):
    if not texto:
        return ""
    for basura in ["￾", "\ufeff", "\ufffe", "\x00"]:
        texto = texto.replace(basura, "")
    return texto


def corregir_acentos(texto):
    if not texto:
        return ""
    reemplazos = {
        "´A": "Á", "´E": "É", "´I": "Í", "´O": "Ó", "´U": "Ú",
        "´a": "á", "´e": "é", "´i": "í", "´o": "ó", "´u": "ú",
        "A´": "Á", "E´": "É", "I´": "Í", "O´": "Ó", "U´": "Ú",
        "a´": "á", "e´": "é", "i´": "í", "o´": "ó", "u´": "ú",
        "tama˜nos": "tamaños",
        "peque˜nos": "pequeños",
        "dise˜no": "diseño",
        "desempe˜no": "desempeño",
        "Matem´aticas": "Matemáticas",
        "Programaci´on": "Programación",
        "Informaci´on": "Información",
        "Presentaci´on": "Presentación",
        "Evaluaci´on": "Evaluación",
        "Metodolog´ia": "Metodología",
        "Bibliograf´ia": "Bibliografía",
        "Bibliografa": "Bibliografía",
        "Aut´onomo": "Autónomo",
        "Cr´editos": "Créditos",
        "Sesi´on": "Sesión",
        "Tem´aticas": "Temáticas",
        "´Algebra": "Álgebra",
        "Nu´cleo": "Núcleo",
        "anal´ıtica": "analítica",
        "librer´ıas": "librerías",
        "ser´ıan": "serían",
        "cient´ıfico": "científico",
        "geometr´ıa": "geometría",
        "v´ınculo": "vínculo",
    }
    for a, b in reemplazos.items():
        texto = texto.replace(a, b)
    return texto


def limpiar_texto(texto):
    return corregir_acentos(limpiar_basura(texto or ""))


def norm(texto):
    texto = unidecode(texto or "").upper()
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip()


def leer_pdf(path_pdf):
    doc = fitz.open(path_pdf)
    paginas = []
    for i, page in enumerate(doc, start=1):
        txt = limpiar_texto(page.get_text("text"))
        paginas.append(f"\n--- PAGINA {i} ---\n{txt}")
    doc.close()
    return "\n".join(paginas)


def lineas_utiles(texto):
    return [l.strip() for l in texto.splitlines() if l.strip()]


ENCABEZADOS_CANONICOS = {
    "informacion_general": ["INFORMACION GENERAL"],
    "presentacion": ["PRESENTACION"],
    "continuidad": ["CONTINUIDAD CURRICULAR"],
    "competencias": ["COMPETENCIAS", "COMPETENCIAS Y AFIRMACIONES ASOCIADAS"],
    "resultados": ["RESULTADOS DE APRENDIZAJE"],
    "metodologia": ["METODOLOGIA"],
    "tematicas": ["TEMATICAS O CONTENIDOS", "CONTENIDOS", "CRONOGRAMA"],
    "evaluacion": ["EVALUACION"],
    "bibliografia": ["REFERENCIAS", "BIBLIOGRAFIA"],
}

ORDEN_SECCIONES = list(ENCABEZADOS_CANONICOS.keys())


def detectar_encabezados_partidos(texto):
    lineas = texto.splitlines()
    encontrados = []

    for i, linea in enumerate(lineas):
        actual = norm(linea)
        combinado = actual

        if re.fullmatch(r"\d+\.?", actual) and i + 1 < len(lineas):
            combinado = actual + " " + norm(lineas[i + 1])

        for clave, opciones in ENCABEZADOS_CANONICOS.items():
            for opcion in opciones:
                op = norm(opcion)
                if combinado == op or re.fullmatch(r"\d+\.?\s+" + re.escape(op), combinado):
                    idx = i + 1 if re.fullmatch(r"\d+\.?", actual) else i
                    encontrados.append((idx, clave, lineas[idx].strip()))
                    break

    depurados = []
    for item in sorted(encontrados, key=lambda x: x[0]):
        if not depurados:
            depurados.append(item)
        else:
            ant = depurados[-1]
            if not (ant[1] == item[1] and abs(ant[0] - item[0]) <= 2):
                depurados.append(item)

    return depurados, lineas


def extraer_bloques(texto):
    encabezados, lineas = detectar_encabezados_partidos(texto)
    bloques = {k: "" for k in ORDEN_SECCIONES}

    for pos, (idx, clave, encabezado) in enumerate(encabezados):
        inicio = idx + 1
        fin = encabezados[pos + 1][0] - 1 if pos + 1 < len(encabezados) else len(lineas)
        bloques[clave] = limpiar_texto("\n".join(lineas[inicio:fin]).strip())

    return bloques, encabezados


def extraer_valor_campo(bloque_info, etiqueta, etiquetas_posibles):
    lineas = lineas_utiles(bloque_info)
    etiqueta_n = norm(etiqueta)
    etiquetas_n = [norm(e) for e in etiquetas_posibles]

    for i, linea in enumerate(lineas):
        nlinea = norm(linea)
        if nlinea == etiqueta_n or nlinea.startswith(etiqueta_n + " "):
            valor = linea.replace(etiqueta, "", 1).replace(":", "").replace("´", "").strip()
            if valor:
                return limpiar_texto(valor)

            acumulado = []
            for j in range(i + 1, min(i + 5, len(lineas))):
                nj = norm(lineas[j])
                if any(nj == e or nj.startswith(e + " ") for e in etiquetas_n):
                    break
                acumulado.append(lineas[j])
            return limpiar_texto(" ".join(acumulado).replace("´", "").strip())

    return ""


def extraer_info_general(bloque_info):
    etiquetas = [
        "PROGRAMA", "ÁREA", "AREA", "ASIGNATURA", "CRÉDITOS", "CREDITOS",
        "SEMESTRE", "HORAS PRESENCIALES", "HORAS PRESENCIA- LES",
        "HORAS DE TRABAJO", "AUTÓNOMO", "AUTONOMO", "PROFESOR", "PROFESORES", "CORREO"
    ]

    lineas = lineas_utiles(bloque_info)

    info = {
        "programa": extraer_valor_campo(bloque_info, "PROGRAMA", etiquetas),
        "area": extraer_valor_campo(bloque_info, "ÁREA", etiquetas) or extraer_valor_campo(bloque_info, "AREA", etiquetas),
        "titulo": extraer_valor_campo(bloque_info, "ASIGNATURA", etiquetas),
        "creditos": extraer_valor_campo(bloque_info, "CRÉDITOS", etiquetas) or extraer_valor_campo(bloque_info, "CREDITOS", etiquetas),
        "semestre": extraer_valor_campo(bloque_info, "SEMESTRE", etiquetas),
        "horas_directas": "",
        "horas_independientes": "",
    }

    for i, l in enumerate(lineas):
        if "HORAS PRESENCIA" in norm(l):
            for j in range(i + 1, min(i + 6, len(lineas))):
                if re.fullmatch(r"\d+", lineas[j].strip()):
                    info["horas_directas"] = lineas[j].strip()
                    break
            break

    for i, l in enumerate(lineas):
        if norm(l).startswith("HORAS DE TRABAJO"):
            for j in range(i + 1, min(i + 7, len(lineas))):
                if re.fullmatch(r"\d+", lineas[j].strip()):
                    info["horas_independientes"] = lineas[j].strip()
                    break
            break

    return {k: limpiar_texto(v.replace("´", "").strip()) for k, v in info.items()}


def es_numero_pagina(linea):
    return re.fullmatch(r"\d+", linea.strip()) is not None


def es_linea_seccion_interna(linea):
    n = norm(linea)
    claves = [
        "RESULTADOS DE LA COMPETENCIA", "BIBLIOGRAFIA BASICA",
        "BIBLIOGRAFIA COMPLEMENTARIA", "RECURSOS DIGITALES",
        "INSTRUMENTOS DE EVALUACION", "PORCENTAJES", "FALLAS Y LINEAMIENTOS"
    ]
    return any(n.startswith(c) for c in claves)


def es_linea_numerada(linea):
    return re.match(r"^\d+\.\s+", linea.strip()) is not None


def debe_ser_vineta(linea, contexto_vinetas=False):
    l = linea.strip()
    n = norm(l)

    if l.startswith(("•", "-", "")):
        return True
    if es_linea_numerada(l) or es_linea_seccion_interna(l):
        return False

    verbos = (
        "RESOLVER", "MANIPULAR", "ANALIZAR", "DETERMINAR", "CALCULAR",
        "REALIZAR", "APLICAR", "EXPLICAR", "JUSTIFICAR", "COMUNICAR",
        "UTILIZAR", "EXPLORAR", "FORMULAR", "SUSTENTAR", "PLANTEAR",
        "INTERPRETAR", "COMPRENDER", "DOCUMENTAR", "EXPERIMENTAR", "PROBAR"
    )
    return contexto_vinetas and n.startswith(verbos)


def reflow_bloque(texto, activar_vinetas=True):
    lineas = []
    for l in texto.splitlines():
        l = limpiar_texto(l.strip())
        if not l or l.startswith("--- PAGINA") or es_numero_pagina(l):
            continue
        lineas.append(l)

    parrafos = []
    actual = ""
    contexto_vinetas = False

    for linea in lineas:
        n = norm(linea)

        if n.endswith("IMPLICA:") or n.endswith("INCLUYE:"):
            if actual:
                parrafos.append(("normal", actual.strip()))
            parrafos.append(("normal", linea))
            actual = ""
            contexto_vinetas = True
            continue

        if es_linea_seccion_interna(linea):
            if actual:
                parrafos.append(("normal", actual.strip()))
            parrafos.append(("heading", linea))
            actual = ""
            contexto_vinetas = True
            continue

        if es_linea_numerada(linea):
            if actual:
                parrafos.append(("normal", actual.strip()))
            parrafos.append(("normal", linea))
            actual = ""
            contexto_vinetas = True
            continue

        if activar_vinetas and debe_ser_vineta(linea, contexto_vinetas):
            if actual:
                parrafos.append(("normal", actual.strip()))
            parrafos.append(("bullet", linea.lstrip("•- ").strip()))
            actual = ""
            continue

        if actual.endswith("-"):
            actual = actual[:-1] + linea
        elif actual:
            actual += " " + linea
        else:
            actual = linea

    if actual:
        parrafos.append(("normal", actual.strip()))

    return parrafos


def reflow_bibliografia(texto):
    parrafos = []
    for tipo, p in reflow_bloque(texto, activar_vinetas=False):
        n = norm(p)
        if n.startswith("BIBLIOGRAFIA") or n.startswith("RECURSOS DIGITALES"):
            parrafos.append(("heading", p))
        elif p.strip():
            parrafos.append(("bullet", p.strip().lstrip("•- ")))
    return parrafos


def extraer_tabla_tematicas_pdf(path_pdf):
    candidatas = []

    with pdfplumber.open(path_pdf) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            try:
                tablas = page.extract_tables()
            except Exception:
                tablas = []

            for tabla in tablas or []:
                limpia = []
                for row in tabla:
                    if row is None:
                        continue
                    fila = [("" if c is None else limpiar_texto(str(c).strip())) for c in row]
                    if any(fila):
                        limpia.append(fila)

                if limpia:
                    texto_tabla = " ".join(" ".join(r) for r in limpia)
                    score = sum(
                        1 for x in ["SEMANA", "SESION", "SESIÓN", "FECHA", "CONTENIDO"]
                        if norm(x) in norm(texto_tabla)
                    )
                    if score >= 2:
                        candidatas.append((score, len(limpia), limpia))

    if candidatas:
        candidatas.sort(key=lambda x: (x[0], x[1]), reverse=True)
        return candidatas[0][2]

    return None


def tabla_tematicas_fallback(bloque_tematicas):
    filas = [["Semana", "Sesión", "Fecha", "Contenido"]]
    for linea in bloque_tematicas.splitlines():
        linea = limpiar_texto(linea.strip())
        if not linea or linea.startswith("--- PAGINA") or es_numero_pagina(linea):
            continue
        filas.append(["", "", "", linea])
    return filas


def set_run_font(run, bold=False):
    run.font.name = FUENTE_CONTENIDO
    run.font.size = Pt(TAMANO_CONTENIDO)
    run.bold = bold


def set_paragraph_format(p, tipo="normal"):
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.space_before = Pt(0)
    if tipo == "bullet":
        p.paragraph_format.left_indent = Pt(18)
        p.paragraph_format.first_line_indent = Pt(-9)


def limpiar_parrafo(p):
    p.clear()
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.space_before = Pt(0)


def escribir_parrafo_simple(doc, indice, texto, bold=False):
    p = doc.paragraphs[indice]
    limpiar_parrafo(p)
    r = p.add_run(texto or "")
    set_run_font(r, bold=bold)
    set_paragraph_format(p)


def insertar_parrafos_despues(doc, indice, parrafos):
    base = doc.paragraphs[indice]
    limpiar_parrafo(base)

    ref = base
    for i, (tipo, texto) in enumerate(parrafos):
        if i == 0:
            p = base
        else:
            nuevo_xml = OxmlElement("w:p")
            ref._p.addnext(nuevo_xml)
            p = Paragraph(nuevo_xml, ref._parent)
            ref = p

        limpiar_parrafo(p)

        if tipo == "bullet":
            r = p.add_run("• " + texto)
            set_run_font(r)
            set_paragraph_format(p, "bullet")
        elif tipo == "heading":
            r = p.add_run(texto)
            set_run_font(r, bold=True)
            set_paragraph_format(p)
        else:
            r = p.add_run(texto)
            set_run_font(r)
            set_paragraph_format(p)


def escribir_celda(cell, texto):
    cell.text = texto or ""
    for p in cell.paragraphs:
        for r in p.runs:
            set_run_font(r)
        set_paragraph_format(p)


def diligenciar_tabla_inicial(doc, info):
    tabla = doc.tables[0]

    try:
        escribir_celda(tabla.rows[1].cells[1], info.get("programa", ""))
    except Exception:
        pass

    try:
        escribir_celda(tabla.rows[2].cells[1], MODALIDAD_DEFAULT)
    except Exception:
        pass

    try:
        escribir_celda(tabla.rows[4].cells[1], info.get("creditos", ""))
        escribir_celda(tabla.rows[4].cells[5], info.get("horas_directas", ""))
        escribir_celda(tabla.rows[4].cells[8], info.get("horas_independientes", ""))
    except Exception:
        pass


def insertar_tabla_despues_parrafo(doc, indice_parrafo, data):
    if not data:
        return

    max_cols = max(len(row) for row in data)
    tabla = doc.add_table(rows=1, cols=max_cols)
    tabla.style = "Table Grid"

    for j in range(max_cols):
        tabla.cell(0, j).text = data[0][j] if j < len(data[0]) else ""

    for row in data[1:]:
        cells = tabla.add_row().cells
        for j in range(max_cols):
            cells[j].text = row[j] if j < len(row) else ""

    for row in tabla.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for r in p.runs:
                    set_run_font(r)
                set_paragraph_format(p)

    tbl_xml = tabla._tbl
    doc._body._body.remove(tbl_xml)
    doc.paragraphs[indice_parrafo]._p.addnext(tbl_xml)


def nombre_salida(pdf_path):
    base = Path(pdf_path).stem
    base = re.sub(r"[^A-Za-zÁÉÍÓÚáéíóúÑñ0-9_ -]+", "", base).strip().replace(" ", "_")
    return f"Anexo_{base}_DILIGENCIADO.docx"


def procesar_pdf(template_path, pdf_path, output_dir):
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    texto = leer_pdf(pdf_path)
    bloques, encabezados = extraer_bloques(texto)
    info = extraer_info_general(bloques.get("informacion_general", ""))

    tabla_tematicas = extraer_tabla_tematicas_pdf(pdf_path)
    if tabla_tematicas is None:
        tabla_tematicas = tabla_tematicas_fallback(bloques.get("tematicas", ""))

    doc = Document(template_path)

    escribir_parrafo_simple(doc, 1, info.get("titulo", "").upper(), bold=True)
    diligenciar_tabla_inicial(doc, info)

    insertar_parrafos_despues(doc, 53, reflow_bibliografia(bloques.get("bibliografia", "")))
    insertar_parrafos_despues(doc, 47, reflow_bloque(bloques.get("evaluacion", ""), activar_vinetas=True))
    insertar_parrafos_despues(doc, 41, reflow_bloque(bloques.get("metodologia", ""), activar_vinetas=True))
    insertar_tabla_despues_parrafo(doc, 32, tabla_tematicas)

    presentacion = bloques.get("presentacion", "")
    if INCLUIR_CONTINUIDAD_EN_PRESENTACION and bloques.get("continuidad", "").strip():
        presentacion = presentacion.strip() + "\n\n" + bloques.get("continuidad", "").strip()
    insertar_parrafos_despues(doc, 27, reflow_bloque(presentacion, activar_vinetas=False))

    insertar_parrafos_despues(doc, 21, reflow_bloque(bloques.get("resultados", ""), activar_vinetas=True))

    if LLENAR_RA_PROGRAMA_SOLO_SI_APARECE_EXPLICITO and "RESULTADOS DE APRENDIZAJE DEL PROGRAMA" in norm(texto):
        insertar_parrafos_despues(doc, 15, reflow_bloque(bloques.get("resultados", ""), activar_vinetas=True))

    if COPIAR_COMPETENCIAS_EN_ITEM_2:
        insertar_parrafos_despues(doc, 9, reflow_bloque(bloques.get("competencias", ""), activar_vinetas=True))

    salida = output_dir / nombre_salida(pdf_path)
    doc.save(salida)

    diagnostico = {
        "pdf": str(pdf_path),
        "salida": str(salida),
        "titulo": info.get("titulo", ""),
        "programa": info.get("programa", ""),
        "creditos": info.get("creditos", ""),
        "horas_directas": info.get("horas_directas", ""),
        "horas_independientes": info.get("horas_independientes", ""),
        "secciones": {k: bool(v.strip()) for k, v in bloques.items()},
        "filas_tematicas": len(tabla_tematicas) if tabla_tematicas else 0,
    }

    return salida, diagnostico
