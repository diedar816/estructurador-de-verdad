
# -*- coding: utf-8 -*-
"""
Estructurador de Información (PDF -> Excel) – MULTI-CASO + ORDEN + CONDICIÓN DE ROL (k_max reservado)
- Reindexa personas por rol POR FILA (INDICIADO -> resto: DENUNCIANTE, VICTIMA, TESTIGO, otros)
- **Reserva** k_max posiciones para INDICIADO justo después de "Relato de los hechos".
  Si una fila tiene menos INDICIADOS que k_max, esas posiciones quedan **vacías** y
  el resto de personas se **desplaza** para empezar a partir de (k_max + 1).
- Evita que DENUNCIANTE/VÍCTIMA/TESTIGO aparezcan después de Relato cuando esa fila no tiene INDICIADO.
"""

import os, re, sys, time, traceback, unicodedata
from pathlib import Path
import pandas as pd
from PyPDF2 import PdfReader

# =================== CONFIG ===================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PDF_NAME  = "Estructurador de Información.pdf"
XLSX_NAME = "Estructurado en tabla.xlsx"

TESSERACT_EXE = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"  # "" si ya está en PATH
POPPLER_PATH  = None

# =================== OCR (si hace falta) ===================
OCR_OK = True
try:
    import pytesseract
    from pdf2image import convert_from_path
    from PIL import Image, ImageOps, ImageFilter
    if TESSERACT_EXE and os.path.exists(TESSERACT_EXE):
        pytesseract.pytesseract.tesseract_cmd = TESSERACT_EXE
except Exception:
    OCR_OK = False

# =================== Campos esperados ===================
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

# =================== Utilidades ===================
import math

def strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s or "") if unicodedata.category(c) != "Mn")

def cleanspace(s: str) -> str:
    if s is None: return ""
    s = s.replace("\xa0"," ").replace("\u200b"," ")
    s = re.sub(r"[ \t\r\f\v]+"," ", s)
    return s.strip()

def normalize_labels_multiline(text: str) -> str:
    t = text
    for label, patterns in LABEL_ALIASES.items():
        for pat in patterns:
            t = re.sub(rf"(?im){pat}\s*:\s*", f"{label}: ", t)
    return t

def read_pdf_text(pdf_path: str) -> str:
    reader = PdfReader(pdf_path)
    chunks = []
    for page in reader.pages:
        try: t = page.extract_text() or ""
        except Exception: t = ""
        chunks.append(t)
    return "\n".join(chunks).strip()

def ocr_pdf(pdf_path: str, tmp_img_dir: str) -> str:
    if not OCR_OK:
        raise RuntimeError("OCR no disponible. Instala pytesseract, pillow, pdf2image y además Tesseract + Poppler.")
    Path(tmp_img_dir).mkdir(parents=True, exist_ok=True)
    pages = convert_from_path(pdf_path, dpi=300, poppler_path=POPPLER_PATH)
    all_text = []
    for idx, img in enumerate(pages, start=1):
        img = img.convert("L"); img = ImageOps.autocontrast(img); img = img.filter(ImageFilter.SHARPEN)
        txt = pytesseract.image_to_string(img, lang="spa", config="--psm 6")
        if len(txt.strip()) < 10:
            txt = pytesseract.image_to_string(img, lang="eng", config="--psm 6")
        all_text.append(txt)
        try: img.save(os.path.join(tmp_img_dir, f"page_{idx:03d}.png"))
        except Exception: pass
    return "\n\n---- NUEVA PAGINA ----\n\n".join(all_text)

def extract_line_value(text: str, label: str) -> str:
    pat = rf"(?im)^{re.escape(label)}\s*[:：]\s*(.+)$"
    m = re.search(pat, text)
    return cleanspace(m.group(1)) if m else ""

def find_label_span(text: str, label: str):
    pats = [rf"{re.escape(label)}"] + LABEL_ALIASES.get(label, [])
    for p in pats:
        m = re.search(rf"(?im){p}\s*[:：]?", text)
        if m: return m.span()
    return None

def extract_window_value(text: str, label: str, next_labels: list) -> str:
    span = find_label_span(text, label)
    if not span: return ""
    start_idx = span[1]; end_idx = len(text)
    for nxt in next_labels:
        for ptn in [rf"{re.escape(nxt)}"] + LABEL_ALIASES.get(nxt, []):
            m = re.search(rf"(?im)^{ptn}\s*[:：]?", text[start_idx:])
            if m:
                end_idx = min(end_idx, start_idx + m.start()); break
    value = text[start_idx:end_idx].strip()
    value = re.sub(rf"(?im)^{re.escape(label)}\s*[:：]?\s*", "", value).strip()
    for alias in LABEL_ALIASES.get(label, []):
        value = re.sub(rf"(?im)^{alias}\s*[:：]?\s*", "", value).strip()
    return cleanspace(value)

# =================== Parseo de un caso ===================
ALL_LABELS = CASE_FIELDS + PERSON_FIELDS

def parse_case_and_people(raw_text: str) -> dict:
    text = raw_text.replace("\xa0"," ").replace("\u200b"," ")
    text = normalize_labels_multiline(text)

    # Caso
    caso = {k: "" for k in CASE_FIELDS}
    for k in CASE_FIELDS:
        v = extract_line_value(text, k)
        if not v:
            v = extract_window_value(text, k, [x for x in ALL_LABELS if x != k])
        caso[k] = v

    # Fecha fallback
    if not caso.get("Fecha de los Hechos"):
        m = re.search(r"(?<!\d)(\d{2}[/-]\d{2}[/-]\d{4}\s+\d{2}:\d{2}(?::\d{2})?)", text)
        if m: caso["Fecha de los Hechos"] = m.group(1)

    # Limpiar Lugar
    if caso.get("Lugar de los hechos"):
        caso["Lugar de los hechos"] = re.split(r"(?i)¿\s*qué\s*viene\s*a\s*denunciar|Relato de los hechos", caso["Lugar de los hechos"])[0].strip()

    # Relato hacia atrás si corto
    relato = caso.get("Relato de los hechos", "")
    if len(relato) < 120:
        s_rel = find_label_span(text, "Relato de los hechos"); s_lug = find_label_span(text, "Lugar de los hechos")
        if s_rel and s_lug and s_lug[1] < s_rel[0]:
            candidate = text[s_lug[1]:s_rel[0]].strip()
            lines = candidate.splitlines(); start_idx = 0
            for i, ln in enumerate(lines):
                if re.search(r"(?i)¿\s*qué\s*viene\s*a\s*denunciar|hechos\s*:", ln): start_idx = i; break
            candidate = "\n".join(lines[start_idx:]).strip()
            if len(candidate) > len(relato): relato = candidate
    if len(relato) < 120:
        m_start = re.search(r"(?is)¿\s*qué\s*viene\s*a\s*denunciar.*?$", text)
        if m_start:
            tail = text[m_start.start():]
            m_end = re.search(r"(?im)^(Municipio Fiscal|Seccional|Estado del caso|Unidad de Fiscal[ií]a|Despacho|Etapa del caso)\s*[:：]", tail)
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
                v = extract_window_value(ch, lbl, PERSON_FIELDS)
            pdata[lbl] = v
        # considera vacíos reales
        if any(str(v).strip().lower() not in ("", "nan", "-") for v in pdata.values()):
            personas.append(pdata)

    fila = dict(caso)
    for i, p in enumerate(personas, start=1):
        suf = f"_{i}"
        for lbl in PERSON_FIELDS:
            fila[f"{lbl}{suf}"] = p.get(lbl, "")
    return fila

# =================== Split por 'Caso Noticia:' ===================
def split_cases(full_text: str) -> list:
    norm = normalize_labels_multiline(full_text)
    starts = [m.start() for m in re.finditer(r"(?im)^\s*Caso\s+Noticia\s*[:：]", norm)]
    if not starts: return [full_text]
    blocks = []
    for i, s in enumerate(starts):
        e = starts[i+1] if i+1 < len(starts) else len(norm)
        blocks.append(norm[s:e].strip())
    return blocks

# =================== Reindex por rol y orden final ===================
ROLE_PRIORITY_REST = ["DENUNCIANTE", "VICTIMA", "TESTIGO"]

def normalize_role(val):
    try:
        import pandas as _pd
        if _pd.isna(val): s = ""
        else: s = str(val)
    except Exception:
        s = str(val) if val is not None else ""
    v = strip_accents(s).upper().strip()
    if v.startswith("INDICIADO"): return "INDICIADO"
    if v.startswith("DENUNCIANTE"): return "DENUNCIANTE"
    if v.startswith("VICTIMA"): return "VICTIMA"
    if v.startswith("TESTIGO"): return "TESTIGO"
    return v

def collect_person_blocks_from_row(row: pd.Series) -> list:
    ns = []
    pat = re.compile(r"_(\d+)$")
    for col in row.index:
        m = pat.search(str(col))
        if m: ns.append(int(m.group(1)))
    ns = sorted(set(ns))

    def _nonempty(x):
        try:
            import pandas as _pd
            if _pd.isna(x): return False
        except Exception: pass
        s = str(x).strip()
        return s != "" and s.lower() != "nan" and s != "-"

    persons = []
    for n in ns:
        block = {f: row.get(f"{f}_{n}", "") for f in PERSON_FIELDS}
        if any(_nonempty(v) for v in block.values()):
            block["__role__"] = normalize_role(block.get("Calidad", ""))
            persons.append(block)
    return persons

def write_person_blocks_to_row(base: dict, persons_by_new_index: list):
    out = dict(base)
    for new_idx, p in persons_by_new_index:
        for f in PERSON_FIELDS:
            out[f"{f}_{new_idx}"] = p.get(f, "")
    return out

def reorder_and_expand(df: pd.DataFrame) -> pd.DataFrame:
    # 1) Recolectar personas por fila y separar por rol
    per_row = []
    for _, r in df.iterrows():
        base = {k: r.get(k, "") for k in r.index if not re.search(r"_(\d+)$", str(k))}
        persons = collect_person_blocks_from_row(r)
        indic = [p for p in persons if p.get("__role__") == "INDICIADO"]
        rest  = [p for p in persons if p.get("__role__") != "INDICIADO"]
        def rest_key(p):
            rp = p.get("__role__", "")
            try: return (ROLE_PRIORITY_REST.index(rp), rp)
            except ValueError: return (len(ROLE_PRIORITY_REST)+1, rp)
        rest_sorted = sorted(rest, key=rest_key)
        per_row.append((base, indic, rest_sorted))

    # 2) Detectar max_n (máximo personas con datos) y k_max (máximo INDICIADOS consecutivos desde el inicio teórico)
    max_n = 0
    k_max = 0
    for base, indic, rest_sorted in per_row:
        m = len(indic) + len(rest_sorted)
        max_n = max(max_n, m)
        k_max = max(k_max, len(indic))
    if max_n == 0: max_n = 1

    # 3) Construir cada fila remapeando índices: reservar 1..k_max para INDICIADO(s).
    new_rows = []
    for base, indic, rest_sorted in per_row:
        pairs = []
        # INDICIADOS ocupan 1..len(indic) y el resto de slots hasta k_max quedan vacíos
        for i, p in enumerate(indic, start=1):
            pairs.append((i, p))
        # Resto arranca en k_max+1
        start_idx = k_max + 1
        for j, p in enumerate(rest_sorted, start=start_idx):
            pairs.append((j, p))
        new_rows.append(write_person_blocks_to_row(base, pairs))

    df2 = pd.DataFrame(new_rows)

    # 4) Orden de columnas con ranuras reservadas
    def person_cols(i):
        return [
            f"Calidad_{i}", f"Documento_{i}", f"Número documento_{i}", f"Nombre_{i}",
            f"Departamento de notificación_{i}", f"Municipio de notificación_{i}",
            f"Teléfono de notificación_{i}", f"Teléfono móvil_{i}", f"Correo Electrónico_{i}", f"Teléfono Oficina_{i}"
        ]

    base_cols = ["Fecha de los Hechos","Caso Noticia","Seccional","Unidad de Fiscalía","Despacho","Unidad de Enrutamiento","Delito"]
    pre_relato = ["Caracterización","Modalidad","Modo","Municipio Fiscal","Lugar de los hechos"]
    relato = ["Relato de los hechos"]
    post_relato_pre_grado = ["Estado del caso","Etapa del caso","Estado de la asignación","Ley de Aplicabilidad","Procedimiento Abreviado?","Priorizado","Tipo Noticia"]
    grado = ["Grado Delito"]

    desired = []
    desired += base_cols
    desired += pre_relato
    desired += relato
    # Reservas 1..k_max (INDICIADO)
    for i in range(1, k_max+1):
        desired += person_cols(i)
    desired += post_relato_pre_grado
    desired += grado
    # Resto de personas
    for i in range(k_max+1, max_n+1):
        desired += person_cols(i)

    # 5) Crear columnas faltantes y reordenar
    for col in desired:
        if col not in df2.columns:
            df2[col] = ""
    restantes = [c for c in df2.columns if c not in desired]
    final_cols = desired + restantes
    return df2.reindex(columns=final_cols)

# =================== MAIN ===================

def main():
    pdf_path  = os.path.join(BASE_DIR, PDF_NAME)
    xlsx_path = os.path.join(BASE_DIR, XLSX_NAME)
    log_txt   = os.path.join(BASE_DIR, "ocr_text.txt")
    dbg_csv   = os.path.join(BASE_DIR, "debug_extraccion.csv")
    tmp_img   = os.path.join(BASE_DIR, "_tmp_pdf_imgs")

    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"No se encontró el PDF: {pdf_path}")

    text_all = read_pdf_text(pdf_path)
    used_ocr = False
    if len(text_all.strip()) < 20:
        used_ocr = True
        text_all = ocr_pdf(pdf_path, tmp_img)

    try:
        with open(log_txt, "w", encoding="utf-8") as f: f.write(text_all)
    except Exception: pass

    case_chunks = split_cases(text_all)

    filas = []
    debug_rows = []
    for idx, chunk in enumerate(case_chunks, start=1):
        fila = parse_case_and_people(chunk)
        filas.append(fila)
        for k in CASE_FIELDS:
            debug_rows.append({"CasoIndex": idx, "Etiqueta": k, "Valor": fila.get(k, "")})
        for i in range(1, 8):
            for lbl in PERSON_FIELDS:
                debug_rows.append({"CasoIndex": idx, "Etiqueta": f"{lbl}_{i}", "Valor": fila.get(f"{lbl}_{i}", "")})

    df = pd.DataFrame(filas)
    df = reorder_and_expand(df)

    os.makedirs(os.path.dirname(xlsx_path), exist_ok=True)
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Datos", index=False)
        from openpyxl.utils import get_column_letter
        from openpyxl.styles import Alignment
        ws = writer.sheets["Datos"]
        for ci, col in enumerate(df.columns, start=1):
            values = [str(col)] + [str(v) for v in df[col].tolist()]
            maxlen = max(len(v) for v in values)
            ws.column_dimensions[get_column_letter(ci)].width = min(max(12, int(maxlen*0.9)), 85)
        if "Relato de los hechos" in df.columns:
            cidx = df.columns.get_loc("Relato de los hechos") + 1
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=cidx, max_col=cidx):
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")

    pd.DataFrame(debug_rows).to_csv(dbg_csv, index=False, encoding="utf-8-sig")

    time.sleep(0.3)
    print(f"Archivo generado: {xlsx_path} (modo: {'OCR' if used_ocr else 'Texto directo'})")
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
