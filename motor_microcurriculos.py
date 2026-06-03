from pathlib import Path
import re
from datetime import datetime

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

# Regla principal del proyecto:
# - El PDF es la fuente principal.
# - La matriz curricular será fuente secundaria en una versión posterior.
# - No se inventa información.
# - No se diligencian profesor ni correo; se dejan para diligenciamiento manual.
INCLUIR_CONTINUIDAD_EN_PRESENTACION = True
COPIAR_COMPETENCIAS_EN_ITEM_2 = True


# ---------------------------------------------------------------------
# Limpieza y normalización
# ---------------------------------------------------------------------

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


def lineas_utiles(texto):
    return [l.strip() for l in texto.splitlines() if l.strip()]


# ---------------------------------------------------------------------
# Lectura PDF y extracción por secciones
# ---------------------------------------------------------------------

def leer_pdf(path_pdf):
    doc = fitz.open(path_pdf)
    paginas = []
    for i, page in enumerate(doc, start=1):
        txt = limpiar_texto(page.get_text("text"))
        paginas.append(f"\n--- PAGINA {i} ---\n{txt}")
    doc.close()
    return "\n".join(paginas)


ENCABEZADOS_CANONICOS = {
    "informacion_general": ["INFORMACION GENERAL", "INFORMACIÓN GENERAL"],
    "presentacion": ["PRESENTACION", "PRESENTACIÓN"],
    "continuidad": ["CONTINUIDAD CURRICULAR"],
    "competencias": ["COMPETENCIAS", "COMPETENCIAS Y AFIRMACIONES ASOCIADAS"],
    "resultados": ["RESULTADOS DE APRENDIZAJE"],
    "metodologia": ["METODOLOGIA", "METODOLOGÍA"],
    "tematicas": ["TEMATICAS O CONTENIDOS", "TEMÁTICAS O CONTENIDOS", "CONTENIDOS", "CRONOGRAMA"],
    "evaluacion": ["EVALUACION", "EVALUACIÓN"],
    "bibliografia": ["REFERENCIAS", "BIBLIOGRAFIA", "BIBLIOGRAFÍA", "REFERENCIAS BIBLIOGRAFICAS", "REFERENCIAS BIBLIOGRÁFICAS"],
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


# ---------------------------------------------------------------------
# Información general
# ---------------------------------------------------------------------

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


def extraer_horas(lineas, patron_inicio):
    """Extrae horas aunque vengan como número solo o como '6 horas / semana'."""
    for i, l in enumerate(lineas):
        if patron_inicio in norm(l):
            # Caso: la misma línea trae el valor
            m = re.search(r"(\d+(?:[\.,]\d+)?)", l)
            if m and not norm(l).startswith(patron_inicio):
                return m.group(1).replace(",", ".")

            # Caso: el valor aparece en líneas posteriores
            for j in range(i + 1, min(i + 7, len(lineas))):
                nj = norm(lineas[j])
                if any(nj.startswith(x) for x in ["PROFESOR", "CORREO", "AREA", "ASIGNATURA", "CREDITOS", "SEMESTRE"]):
                    break
                m = re.search(r"(\d+(?:[\.,]\d+)?)", lineas[j])
                if m:
                    return m.group(1).replace(",", ".")
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
        # Ajuste solicitado: no diligenciar profesor ni correo.
        "profesor": "",
        "correo": "",
    }

    info["horas_directas"] = extraer_horas(lineas, "HORAS PRESENCIA")
    info["horas_independientes"] = extraer_horas(lineas, "HORAS DE TRABAJO")

    return {k: limpiar_texto(str(v).replace("´", "").strip()) for k, v in info.items()}


# ---------------------------------------------------------------------
# Reformateo de bloques
# ---------------------------------------------------------------------

def es_numero_pagina(linea):
    return re.fullmatch(r"\d+", linea.strip()) is not None


def es_linea_seccion_interna(linea):
    n = norm(linea)
    claves = [
        "RESULTADOS DE LA COMPETENCIA", "RESULTADOS ASOCIADOS",
        "BIBLIOGRAFIA BASICA", "BIBLIOGRAFIA COMPLEMENTARIA", "RECURSOS DIGITALES",
        "INSTRUMENTOS DE EVALUACION", "PORCENTAJES", "CRITERIOS Y PORCENTAJES",
        "FALLAS Y LINEAMIENTOS", "RESUMEN DE EVALUACION"
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
        "INTERPRETAR", "COMPRENDER", "DOCUMENTAR", "EXPERIMENTAR", "PROBAR",
        "IDENTIFICAR", "CONSTRUIR", "DISEÑAR", "EVALUAR", "IMPLEMENTAR",
        "PRESENTAR", "ELABORAR", "INTEGRAR"
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
            parrafos.append(("heading", linea))
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


# ---------------------------------------------------------------------
# Temáticas sin fechas
# ---------------------------------------------------------------------

def limpiar_celda_tabla(valor):
    return limpiar_texto(str(valor or "").replace("\n", " ").strip())


def detectar_indices_tematicas(header):
    """Detecta columnas de Semana, Sesión y Contenido. Fecha se ignora."""
    ns = [norm(c) for c in header]
    idx_semana = idx_sesion = idx_contenido = None

    for i, c in enumerate(ns):
        if idx_semana is None and "SEMANA" in c:
            idx_semana = i
        if idx_sesion is None and ("SESION" in c or "SESIÓN" in c):
            idx_sesion = i
        if idx_contenido is None and any(x in c for x in ["CONTENIDO", "TEMA", "TEMAS"]):
            idx_contenido = i

    # Si no detecta contenido, suele ser la última columna no vacía.
    if idx_contenido is None and header:
        idx_contenido = len(header) - 1

    return idx_semana, idx_sesion, idx_contenido


def compactar_tabla_tematicas(tabla_original):
    """Devuelve tabla final: Semana | Sesión | Contenido. Nunca incluye fechas."""
    if not tabla_original:
        return None

    filas_limpias = []
    for row in tabla_original:
        fila = [limpiar_celda_tabla(c) for c in row]
        if any(fila):
            filas_limpias.append(fila)

    if not filas_limpias:
        return None

    header_idx = 0
    mejor_score = -1
    for i, row in enumerate(filas_limpias[:5]):
        nr = norm(" ".join(row))
        score = sum(x in nr for x in ["SEMANA", "SESION", "SESIÓN", "CONTENIDO", "TEMA"])
        if score > mejor_score:
            mejor_score = score
            header_idx = i

    header = filas_limpias[header_idx]
    idx_semana, idx_sesion, idx_contenido = detectar_indices_tematicas(header)

    if idx_contenido is None:
        return None

    resultado = [["Semana", "Sesión", "Contenido"]]
    ultima_semana = ""

    for row in filas_limpias[header_idx + 1:]:
        semana = row[idx_semana] if idx_semana is not None and idx_semana < len(row) else ""
        sesion = row[idx_sesion] if idx_sesion is not None and idx_sesion < len(row) else ""
        contenido = row[idx_contenido] if idx_contenido is not None and idx_contenido < len(row) else ""

        # Evita filas de encabezado repetido.
        if norm(semana) == "SEMANA" or norm(contenido) in ["CONTENIDO", "TEMA"]:
            continue

        # En tablas con celdas combinadas, la semana puede venir vacía.
        if semana:
            ultima_semana = semana
        else:
            semana = ultima_semana

        # Ignorar filas que solo sean festivos/receso si no tienen contenido útil.
        if not any([semana, sesion, contenido]):
            continue

        resultado.append([semana, sesion, contenido])

    return resultado if len(resultado) > 1 else None


def extraer_tabla_tematicas_pdf(path_pdf):
    candidatas = []

    with pdfplumber.open(path_pdf) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            try:
                tablas = page.extract_tables()
            except Exception:
                tablas = []

            for tabla in tablas or []:
                if not tabla:
                    continue
                texto_tabla = " ".join(" ".join(limpiar_celda_tabla(c) for c in row) for row in tabla if row)
                nt = norm(texto_tabla)
                score = sum(x in nt for x in ["SEMANA", "SESION", "SESIÓN", "CONTENIDO", "TEMA"])
                if score >= 2:
                    compacta = compactar_tabla_tematicas(tabla)
                    if compacta:
                        candidatas.append((score, len(compacta), compacta))

    if candidatas:
        candidatas.sort(key=lambda x: (x[0], x[1]), reverse=True)
        return candidatas[0][2]

    return None


# ---------------------------------------------------------------------
# Utilidades Word
# ---------------------------------------------------------------------

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


def escribir_parrafo(p, texto, bold=False, tipo="normal"):
    limpiar_parrafo(p)
    r = p.add_run(texto or "")
    set_run_font(r, bold=bold)
    set_paragraph_format(p, tipo)


def escribir_parrafo_simple(doc, indice, texto, bold=False):
    escribir_parrafo(doc.paragraphs[indice], texto, bold=bold)


def buscar_parrafo(doc, texto_objetivo):
    objetivo = norm(texto_objetivo)
    for i, p in enumerate(doc.paragraphs):
        if norm(p.text) == objetivo:
            return i
    for i, p in enumerate(doc.paragraphs):
        if objetivo in norm(p.text):
            return i
    return None


def insertar_parrafos_despues(doc, indice, parrafos):
    if indice is None:
        return
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


def escribir_celda(cell, texto, bold=False):
    cell.text = texto or ""
    for p in cell.paragraphs:
        for r in p.runs:
            set_run_font(r, bold=bold)
        set_paragraph_format(p)


def texto_tabla(tabla):
    return norm(" ".join(cell.text for row in tabla.rows for cell in row.cells))


def buscar_tabla_por_texto(doc, texto):
    objetivo = norm(texto)
    for tabla in doc.tables:
        if objetivo in texto_tabla(tabla):
            return tabla
    return None


def diligenciar_tabla_inicial(doc, info):
    """Llena la primera tabla de información general sin tocar profesor/correo."""
    if not doc.tables:
        return

    tabla = doc.tables[0]
    fecha_hoy = datetime.today().strftime("%d/%m/%Y")

    # La plantilla puede cambiar de índices; por eso se intenta por coordenadas conocidas
    # y, si falla, no detiene el proceso.
    try:
        # Fecha de actualización: en la plantilla está en la fila 1, celda derecha.
        escribir_celda(tabla.rows[1].cells[-1], fecha_hoy)
    except Exception:
        pass

    try:
        escribir_celda(tabla.rows[2].cells[1], info.get("programa", ""))
    except Exception:
        pass

    try:
        escribir_celda(tabla.rows[3].cells[1], MODALIDAD_DEFAULT)
    except Exception:
        pass

    try:
        # Tipología: se deja textual como valor por defecto en la zona correspondiente.
        # No se marcan checkboxes para no dañar formato institucional.
        pass
    except Exception:
        pass

    try:
        escribir_celda(tabla.rows[4].cells[1], info.get("creditos", ""))
        escribir_celda(tabla.rows[4].cells[5], info.get("horas_directas", ""))
        escribir_celda(tabla.rows[4].cells[8], info.get("horas_independientes", ""))
    except Exception:
        pass


def limpiar_textos_instructivos(doc):
    frases = [
        "Relacione las competencias del programa",
        "Relacione el o los resultados de aprendizaje del programa",
        "Especifique uno o máximo dos resultados de aprendizaje",
        "Describa las características que definen la naturaleza del curso",
        "Precise los contenidos o temáticas",
        "Describa la tipología del curso",
        "Defina los criterios con los cuáles serán evaluados",
        "Incluya la bibliografía básica",
    ]
    for p in doc.paragraphs:
        np = norm(p.text)
        if any(norm(f) in np for f in frases):
            limpiar_parrafo(p)


def obtener_indice_insercion(doc, encabezado):
    idx = buscar_parrafo(doc, encabezado)
    if idx is None:
        return None
    # Busca el primer párrafo no encabezado después del título de sección.
    # En la plantilla ese párrafo es el instructivo, que ya se limpiará y se reutiliza.
    for j in range(idx + 1, min(idx + 6, len(doc.paragraphs))):
        if norm(doc.paragraphs[j].text) and not norm(doc.paragraphs[j].text).startswith("ANEXO MICRO"):
            return j
    return idx + 1 if idx + 1 < len(doc.paragraphs) else idx


def reemplazar_titulo(doc, titulo):
    idx = buscar_parrafo(doc, "Título del curso")
    if idx is not None:
        escribir_parrafo_simple(doc, idx, (titulo or "").upper(), bold=True)


def llenar_tabla_existente_tematicas(doc, data):
    """Llena tabla existente de temáticas. Si no hay data, deja espacio en blanco."""
    if not data:
        return False

    # Buscar una tabla que tenga encabezados de contenido o una tabla posterior al encabezado.
    tabla = None
    for t in doc.tables:
        tt = texto_tabla(t)
        if "SEMANA" in tt and ("SESION" in tt or "SESIÓN" in tt) and ("CONTENIDO" in tt or "TEMA" in tt):
            tabla = t
            break

    # Si la plantilla no tiene tabla detectable, crea una justo después del encabezado.
    if tabla is None:
        idx = buscar_parrafo(doc, "6. Temáticas o contenidos del curso")
        if idx is None:
            return False
        insertar_tabla_despues_parrafo(doc, idx, data)
        return True

    # Ajustar número de columnas a 3 si la tabla existente tiene más columnas.
    # python-docx no elimina columnas de forma simple; se llena solo las tres primeras
    # y se dejan las columnas adicionales vacías.
    while len(tabla.rows) < len(data):
        tabla.add_row()

    # Limpiar filas existentes.
    for row in tabla.rows:
        for cell in row.cells:
            escribir_celda(cell, "")

    for i, fila in enumerate(data):
        if i >= len(tabla.rows):
            tabla.add_row()
        cells = tabla.rows[i].cells
        valores = [fila[0] if len(fila) > 0 else "", fila[1] if len(fila) > 1 else "", fila[2] if len(fila) > 2 else ""]
        for j, valor in enumerate(valores):
            if j < len(cells):
                escribir_celda(cells[j], valor, bold=(i == 0))
        # Vaciar cualquier columna adicional, incluyendo Fecha si existía.
        for j in range(3, len(cells)):
            escribir_celda(cells[j], "")

    return True


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


# ---------------------------------------------------------------------
# Salida y proceso principal
# ---------------------------------------------------------------------

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

    doc = Document(template_path)

    # 1. Título correcto, reemplazando "Título del curso".
    reemplazar_titulo(doc, info.get("titulo", ""))

    # 2. Información general + fecha de actualización automática.
    diligenciar_tabla_inicial(doc, info)

    # 3. Eliminar textos instructivos de la plantilla.
    limpiar_textos_instructivos(doc)

    # 4. Ubicar secciones por texto. Si por alguna razón falla, usa los índices anteriores como respaldo.
    idx_comp_prog = obtener_indice_insercion(doc, "2. Competencias del programa al que se articula el curso") or 9
    idx_ra_prog = obtener_indice_insercion(doc, "3. Resultados de aprendizaje del programa a los que se articula el curso") or 15
    idx_ra_curso = obtener_indice_insercion(doc, "4. Resultados de aprendizaje del curso") or 21
    idx_presentacion = obtener_indice_insercion(doc, "5. Presentación del curso") or 27
    idx_metodologia = obtener_indice_insercion(doc, "7. Metodología") or 41
    idx_evaluacion = obtener_indice_insercion(doc, "8. Evaluación") or 47
    idx_bibliografia = obtener_indice_insercion(doc, "9. Bibliografía del curso") or 53

    # 5. Competencias del programa: por ahora se copian competencias del PDF.
    # Cuando integremos la matriz, este bloque se reemplaza/complementa con la fuente secundaria.
    if COPIAR_COMPETENCIAS_EN_ITEM_2 and bloques.get("competencias", "").strip():
        insertar_parrafos_despues(doc, idx_comp_prog, reflow_bloque(bloques.get("competencias", ""), activar_vinetas=True))

    # 6. Resultados del programa: se deja vacío por ahora si no hay matriz oficial.
    # No se copia automáticamente el mismo resultado del curso en esta sección.
    # Se conserva en blanco para evitar llenar información institucional no explícita.
    if "RESULTADOS DE APRENDIZAJE DEL PROGRAMA" in norm(texto):
        insertar_parrafos_despues(doc, idx_ra_prog, reflow_bloque(bloques.get("resultados", ""), activar_vinetas=True))

    # 7. Resultados de aprendizaje del curso.
    insertar_parrafos_despues(doc, idx_ra_curso, reflow_bloque(bloques.get("resultados", ""), activar_vinetas=True))

    # 8. Presentación + continuidad curricular, si existe.
    presentacion = bloques.get("presentacion", "")
    if INCLUIR_CONTINUIDAD_EN_PRESENTACION and bloques.get("continuidad", "").strip():
        presentacion = presentacion.strip() + "\n\n" + bloques.get("continuidad", "").strip()
    insertar_parrafos_despues(doc, idx_presentacion, reflow_bloque(presentacion, activar_vinetas=False))

    # 9. Temáticas: usar tabla existente; sin fechas. Si no hay tabla en PDF, dejar vacío.
    tematicas_insertadas = llenar_tabla_existente_tematicas(doc, tabla_tematicas)

    # 10. Metodología, evaluación y bibliografía.
    insertar_parrafos_despues(doc, idx_metodologia, reflow_bloque(bloques.get("metodologia", ""), activar_vinetas=True))
    insertar_parrafos_despues(doc, idx_evaluacion, reflow_bloque(bloques.get("evaluacion", ""), activar_vinetas=True))
    insertar_parrafos_despues(doc, idx_bibliografia, reflow_bibliografia(bloques.get("bibliografia", "")))

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
        "profesor": "NO DILIGENCIADO POR REGLA DEL PROYECTO",
        "correo": "NO DILIGENCIADO POR REGLA DEL PROYECTO",
        "fecha_actualizacion": datetime.today().strftime("%d/%m/%Y"),
        "secciones": {k: bool(v.strip()) for k, v in bloques.items()},
        "filas_tematicas": len(tabla_tematicas) if tabla_tematicas else 0,
        "tematicas_insertadas": bool(tematicas_insertadas),
        "nota_tematicas": "Tabla final sin columna Fecha. Si el PDF no trae tabla detectable, queda en blanco.",
    }

    return salida, diagnostico
