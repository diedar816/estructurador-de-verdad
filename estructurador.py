# -*- coding: utf-8 -*-
"""
Estructurador de Información (PDF -> Excel) – MULTI-CASO
- Detecta automáticamente si el PDF tiene texto o requiere OCR.
- Divide el PDF por cada "Caso Noticia:" y escribe una fila por caso en el Excel.
- Une etiquetas partidas (e.g., "Lugar de los\nhechos:" -> "Lugar de los hechos:").
- Reconstruye el Relato aunque esté antes del rótulo.
- Genera ocr_text.txt y debug_extraccion.csv.
"""

import os
import re
import sys
import time
import traceback
import pandas as pd
from pathlib import Path
from PyPDF2 import PdfReader

# ======================== CONFIGURACIÓN ========================
# Siempre trabajamos en la carpeta del script (evita problemas de rutas)
BASE_DIR_CFG = os.path.dirname(os.path.abspath(__file__))
PDF_NAME  = "Estructurador de Información.pdf"
XLSX_NAME = "Estructurado en tabla.xlsx"

# Si el PDF es escaneado y NO tienes Tesseract/Poppler en PATH, ajusta:
TESSERACT_EXE = r"C:\Program Files\Tesseract-OCR\tesseract.exe"  # "" si ya está en PATH
POPPLER_PATH  = None  # p.ej.: r"C:\poppler-24.08.0\Library\bin" o None si está en PATH

# ====================== OCR (cuando haga falta) =================
OCR_OK = True
try:
    import pytesseract
    from pdf2image import convert_from_path
    from PIL import Image, ImageOps, ImageFilter
    if TESSERACT_EXE and os.path.exists(TESSERACT_EXE):
        pytesseract.pytesseract.tesseract_cmd = TESSERACT_EXE
except Exception:
    OCR_OK = False

# ========================== ETIQUETAS ==========================
CASE_FIELDS = [
    "Caso Noticia","Ley de Aplicabilidad","Procedimiento Abreviado?","Priorizado",
    "Tipo Noticia","Delito","Grado Delito","Caracterización","Modalidad","Modo",
    "Fecha de los Hechos","Lugar de los hechos","Relato de los hechos",
    "Municipio Fiscal","Seccional","Unidad de Fiscalía","Despacho",
    "Estado de la asignación","Unidad de Enrutamiento","Estado del caso","Etapa del caso",
]

PERSON_FIELDS = [
    "Calidad","Documento","Número documento","Nombre",
    "Departamento de notificación","Municipio de notificación",
    "Teléfono de notificación","Teléfono móvil","Correo Electrónico","Teléfono Oficina",
]

LABEL_ALIASES = {
    "Ley de Aplicabilidad":       [r"Ley de\s+Aplicabilidad"],
    "Procedimiento Abreviado?":   [r"Procedimiento\s+Abreviado\s*\??"],
    "Relato de los hechos":       [r"Relato de los\s+hechos"],
    "Fecha de los Hechos":        [r"Fecha de los\s+Hechos"],
    "Lugar de los hechos":        [r"Lugar de los\s+hechos"],
    "Unidad de Fiscalía":         [r"Unidad de\s+Fiscal[ií]a"],
    "Correo Electrónico":         [r"Correo\s+Electr[oó]nico"],
    "Número documento":           [r"N[uú]mero\s+documento"],
    "Estado de la asignación":    [r"Estado de la\s+asignaci[oó]n"],
    "Unidad de Enrutamiento":     [r"Unidad de\s+Enrutamiento"],
}

# ========================= UTILIDADES ==========================
def cleanspace(s: str) -> str:
    if s is None:
        return ""
    s = s.replace("\xa0", " ").replace("\u200b", " ")
    s = re.sub(r"[ \t\r\f\v]+", " ", s)
    return s.strip()

def normalize_labels_multiline(text: str) -> str:
    """
    Une etiquetas partidas (p.ej., 'Lugar de los\\nhechos:' -> 'Lugar de los hechos:')
    usando los alias declarados arriba.
    """
    t = text
    for label, patterns in LABEL_ALIASES.items():
        for pat in patterns:
            t = re.sub(rf"(?im){pat}\s*:\s*", f"{label}: ", t)
    return t

def read_pdf_text(pdf_path: str) -> str:
    reader = PdfReader(pdf_path)
    chunks = []
    for page in reader.pages:
        try:
            t = page.extract_text() or ""
        except Exception:
            t = ""
        chunks.append(t)
    return "\n".join(chunks).strip()

def ocr_pdf(pdf_path: str, tmp_img_dir: str) -> str:
    if not OCR_OK:
        raise RuntimeError(
            "OCR no disponible. Instala pytesseract, pillow, pdf2image y además Tesseract + Poppler."
        )
    Path(tmp_img_dir).mkdir(parents=True, exist_ok=True)
    pages = convert_from_path(pdf_path, dpi=300, poppler_path=POPPLER_PATH)
    all_text = []
    for idx, img in enumerate(pages, start=1):
        img = img.convert("L")
        img = ImageOps.autocontrast(img)
        img = img.filter(ImageFilter.SHARPEN)
        txt = pytesseract.image_to_string(img, lang="spa", config="--psm 6")
        if len(txt.strip()) < 10:
            txt = pytesseract.image_to_string(img, lang="eng", config="--psm 6")
        all_text.append(txt)
        try:
            img.save(os.path.join(tmp_img_dir, f"page_{idx:03d}.png"))
        except Exception:
            pass
    return "\n\n---- NUEVA PAGINA ----\n\n".join(all_text)

def extract_line_value(text: str, label: str) -> str:
    pat = rf"(?im)^{re.escape(label)}\s*[:：]\s*(.+)$"
    m = re.search(pat, text)
    return cleanspace(m.group(1)) if m else ""

def find_label_span(text: str, label: str):
    pats = [rf"{re.escape(label)}"] + LABEL_ALIASES.get(label, [])
    for p in pats:
        m = re.search(rf"(?im){p}\s*[:：]?", text)
        if m:
            return m.span()
    return None

def extract_window_value(text: str, label: str, next_labels: list) -> str:
    span = find_label_span(text, label)
    if not span:
        return ""
    start_idx = span[1]
    end_idx = len(text)
    for nxt in next_labels:
        for ptn in [rf"{re.escape(nxt)}"] + LABEL_ALIASES.get(nxt, []):
            m = re.search(rf"(?im)^{ptn}\s*[:：]?", text[start_idx:])
            if m:
                end_idx = min(end_idx, start_idx + m.start())
                break
    value = text[start_idx:end_idx].strip()
    value = re.sub(rf"(?im)^{re.escape(label)}\s*[:：]?\s*", "", value).strip()
    for alias in LABEL_ALIASES.get(label, []):
        value = re.sub(rf"(?im)^{alias}\s*[:：]?\s*", "", value).strip()
    return cleanspace(value)

# ======================= PARSEO DE UN CASO ======================
def parse_case_and_people(raw_text: str) -> dict:
    text = raw_text.replace("\xa0", " ").replace("\u200b", " ")
    text = normalize_labels_multiline(text)

    all_labels = CASE_FIELDS + PERSON_FIELDS

    # Campos del caso (línea o ventana)
    caso = {k: "" for k in CASE_FIELDS}
    for k in CASE_FIELDS:
        v = extract_line_value(text, k)
        if not v:
            next_labels = [x for x in all_labels if x != k]
            v = extract_window_value(text, k, next_labels)
        caso[k] = v

    # Fallback Fecha (dd/mm/yyyy HH:MM[:SS])
    if not caso.get("Fecha de los Hechos"):
        m = re.search(r"(?<!\d)(\d{2}[/-]\d{2}[/-]\d{4}\s+\d{2}:\d{2}(?::\d{2})?)", text)
        if m:
            caso["Fecha de los Hechos"] = m.group(1)

    # Limpiar 'Lugar' si absorbió relato
    if caso.get("Lugar de los hechos"):
        caso["Lugar de los hechos"] = re.split(
            r"(?i)¿\s*qué\s*viene\s*a\s*denunciar|Relato de los hechos",
            caso["Lugar de los hechos"]
        )[0].strip()

    # Relato: mirar hacia atrás si quedó corto
    relato = caso.get("Relato de los hechos", "")
    if len(relato) < 120:
        s_rel = find_label_span(text, "Relato de los hechos")
        s_lug = find_label_span(text, "Lugar de los hechos")
        if s_rel and s_lug and s_lug[1] < s_rel[0]:
            candidate = text[s_lug[1]:s_rel[0]].strip()
            lines = candidate.splitlines()
            start_idx = 0
            for i, ln in enumerate(lines):
                if re.search(r"(?i)¿\s*qué\s*viene\s*a\s*denunciar|hechos\s*:", ln):
                    start_idx = i
                    break
            candidate = "\n".join(lines[start_idx:]).strip()
            if len(candidate) > len(relato):
                relato = candidate
    # Último fallback: desde “¿Qué viene a denunciar?” hasta el siguiente rótulo mayor
    if len(relato) < 120:
        m_start = re.search(r"(?is)¿\s*qué\s*viene\s*a\s*denunciar.*?$", text)
        if m_start:
            tail = text[m_start.start():]
            m_end = re.search(
                r"(?im)^(Municipio Fiscal|Seccional|Estado del caso|Unidad de Fiscal[ií]a|Despacho|Etapa del caso)\s*[:：]",
                tail
            )
            relato = tail[:m_end.start()] if m_end else tail
        relato = cleanspace(relato)
    caso["Relato de los hechos"] = relato

    # Personas
    personas = []
    chunks = re.split(r"(?im)^\s*Calidad\s*[:：]", text)
    for ch in chunks[1:]:
        pdata = {lbl: "" for lbl in PERSON_FIELDS}
        first_line = ch.splitlines()[0] if ch.strip() else ""
        pdata["Calidad"] = cleanspace(first_line)
        for lbl in PERSON_FIELDS[1:]:
            v = extract_line_value(ch, lbl)
            if not v:
                next_labels = [x for x in PERSON_FIELDS if x != lbl]
                v = extract_window_value(ch, lbl, next_labels)
            pdata[lbl] = v
        if any(pdata.values()):
            personas.append(pdata)

    # Aplanar personas
    fila = dict(caso)
    for i, p in enumerate(personas, start=1):
        suf = f"_{i}"
        for lbl in PERSON_FIELDS:
            fila[f"{lbl}{suf}"] = p.get(lbl, "")
    return fila

# =================== DIVISIÓN POR "CASO NOTICIA" ==================
def split_cases(full_text: str) -> list:
    """
    Divide el texto completo en trozos, uno por cada 'Caso Noticia:' encontrado.
    Esto permite que la 'segunda parte' (nuevo caso) quede en la siguiente fila del Excel.
    """
    norm = normalize_labels_multiline(full_text)
    starts = [m.start() for m in re.finditer(r"(?im)^\s*Caso\s+Noticia\s*[:：]", norm)]
    if not starts:
        return [full_text]  # si solo hay un caso
    blocks = []
    for i, s in enumerate(starts):
        e = starts[i+1] if i+1 < len(starts) else len(norm)
        blocks.append(norm[s:e].strip())
    return blocks

# ============================ MAIN ==============================
def main():
    pdf_path  = os.path.join(BASE_DIR_CFG, PDF_NAME)
    xlsx_path = os.path.join(BASE_DIR_CFG, XLSX_NAME)
    log_txt   = os.path.join(BASE_DIR_CFG, "ocr_text.txt")
    debug_csv = os.path.join(BASE_DIR_CFG, "debug_extraccion.csv")
    tmp_img   = os.path.join(BASE_DIR_CFG, "_tmp_pdf_imgs")

    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"No se encontró el PDF en: {pdf_path}")

    # 1) Intentar texto directo
    text_all = read_pdf_text(pdf_path)
    used_ocr = False

    # 2) OCR si no hay texto
    if len(text_all.strip()) < 20:
        used_ocr = True
        text_all = ocr_pdf(pdf_path, tmp_img)

    # Guardar log
    try:
        with open(log_txt, "w", encoding="utf-8") as f:
            f.write(text_all)
    except Exception:
        pass

    # 3) Dividir en casos y procesar cada bloque -> una fila por caso
    case_chunks = split_cases(text_all)
    filas = []
    debug_rows = []

    for idx, chunk in enumerate(case_chunks, start=1):
        fila = parse_case_and_people(chunk)
        filas.append(fila)
        # Para debug
        for k in CASE_FIELDS:
            debug_rows.append({"CasoIndex": idx, "Etiqueta": k, "Valor": fila.get(k, "")})
        for i in range(1, 6):
            for lbl in PERSON_FIELDS:
                debug_rows.append({"CasoIndex": idx, "Etiqueta": f"{lbl}_{i}", "Valor": fila.get(f"{lbl}_{i}", "")})

    # 4) Exportar a Excel (varias filas, una por caso)
    df = pd.DataFrame(filas)
    os.makedirs(os.path.dirname(xlsx_path), exist_ok=True)
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Datos", index=False)
        # Formato
        from openpyxl.utils import get_column_letter
        from openpyxl.styles import Alignment
        ws = writer.sheets["Datos"]
        for col_idx, col in enumerate(df.columns, start=1):
            values = [str(col)] + [str(v) for v in df[col].tolist()]
            maxlen = max(len(v) for v in values)
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, int(maxlen*0.9)), 80)
        if "Relato de los hechos" in df.columns:
            cidx = df.columns.get_loc("Relato de los hechos") + 1
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=cidx, max_col=cidx):
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")

    # 5) CSV de depuración
    pd.DataFrame(debug_rows).to_csv(debug_csv, index=False, encoding="utf-8-sig")

    # 6) Mensajes y apertura
    time.sleep(0.3)
    size = os.path.getsize(xlsx_path) if os.path.exists(xlsx_path) else 0
    print(f"Archivo generado: {xlsx_path} (modo: {'OCR' if used_ocr else 'Texto directo'})")
    print(f"OK: archivo existe y pesa {size} bytes. Casos procesados: {len(case_chunks)}")
    print(f"Depuración: {debug_csv}")
    try:
        os.startfile(xlsx_path)
    except Exception:
        pass

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("\n*** ERROR EJECUTANDO ESTRUCTURADOR (multi-caso) ***")
        print(str(e))
        traceback.print_exc()
        if "pdf2image" in str(e) or "poppler" in str(e).lower():
            print("\nSugerencia: instala Poppler y configura POPPLER_PATH en el script.")
        if "tesseract" in str(e).lower():
            print("\nSugerencia: instala Tesseract OCR y/o ajusta TESSERACT_EXE.")
        if "No module named" in str(e):
            print("\nSugerencia: instala dependencias con:")
            print("python -m pip install PyPDF2 pandas openpyxl python-docx pytesseract pillow pdf2image")
        sys.exit(1)