# -*- coding: utf-8 -*-
"""
Estructurador de Información (PDF -> Excel)
- Multi-caso con control de encabezado tolerante a OCR y **exigencia de 21 dígitos**.
- **SIN deduplicación** por defecto: cada aparición de "Caso Noticia:" se convierte en una fila
  (útil para pruebas de carga con casos repetidos). Puede activarse la deduplicación con DEDUPE_BY_ID=True.
- Reconstrucción robusta de Relato (ventana estricta) y poda de Lugar (evitar que absorba Relato).
- Reindex por rol + reservas para INDICIADO(s) tras Relato.
- Excel con auto-ancho, wrap en Relato, encabezado en negrita, congelado.
- NaN-safe; fallback si el Excel está abierto; headers_debug.csv con aceptados/rechazados.
"""

import os, re, sys, time, traceback, unicodedata, datetime
from pathlib import Path
import pandas as pd
from PyPDF2 import PdfReader

# =================== CONFIG ===================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PDF_NAME  = "Estructurador de Información.pdf"
XLSX_NAME = "Estructurado en tabla.xlsx"

# Cambia a True si quieres eliminar duplicados por ID (21 dígitos)
DEDUPE_BY_ID = False

TESSERACT_EXE = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
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

# =================== ETIQUETAS ===================
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

# =================== UTILIDADES ===================

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

# =================== EXTRACCIÓN CON VENTANA ===================

ALL_LABELS = CASE_FIELDS + PERSON_FIELDS
STOP_AFTER_RELATO = [
    "Municipio Fiscal","Seccional","Unidad de Fiscalía","Despacho",
    "Estado de la asignación","Unidad de Enrutamiento","Estado del caso","Etapa del caso",
    "Calidad","Caso Noticia"
]

def find_any_label_position(text: str, label: str):
    patterns = [rf"{re.escape(label)}"] + LABEL_ALIASES.get(label, [])
    for p in patterns:
        m = re.search(rf"(?im){p}\s*[:：]?", text)
        if m: return m.span()
    return None

def extract_between(text: str, start_label: str, stop_labels: list) -> str:
    s_span = find_any_label_position(text, start_label)
    if not s_span: return ""
    start = s_span[1]; end = len(text)
    for sl in stop_labels:
        for p in [rf"{re.escape(sl)}"] + LABEL_ALIASES.get(sl, []):
            m = re.search(rf"(?im)^{p}\s*[:：]?", text[start:])
            if m: end = min(end, start + m.start()); break
    return cleanspace(text[start:end])

# =================== PARSEO POR CASO ===================

def parse_case_and_people(raw_text: str) -> dict:
    text = normalize_labels_multiline(raw_text.replace("\xa0"," ").replace("\u200b"," "))

    caso = {k: "" for k in CASE_FIELDS}
    for k in CASE_FIELDS:
        v = extract_between(text, k, ALL_LABELS)
        if not v:
            m = re.search(rf"(?im)^{re.escape(k)}\s*[:：]\s*(.+)$", text)
            v = cleanspace(m.group(1)) if m else ""
        caso[k] = v

    # Poda Lugar
    lugar = caso.get("Lugar de los hechos", "")
    if lugar:
        lines = [ln.strip() for ln in lugar.splitlines() if ln.strip()]
        lugar_compacto = " ".join(lines[:2])
        if len(lugar_compacto) > 220:
            lugar_compacto = lugar_compacto[:220].rsplit(' ', 1)[0]
        caso["Lugar de los hechos"] = lugar_compacto

    # Relato con ventana estricta
    relato = extract_between(text, "Relato de los hechos", STOP_AFTER_RELATO)
    if not relato:
        m_q = re.search(r"(?s)¿.+$", text)
        if m_q:
            tail = m_q.group(0)
            end = len(tail)
            for sl in STOP_AFTER_RELATO:
                for p in [rf"{re.escape(sl)}"] + LABEL_ALIASES.get(sl, []):
                    m2 = re.search(rf"(?im)^{p}\s*[:：]?", tail)
                    if m2: end = min(end, m2.start()); break
            relato = cleanspace(tail[:end])
    caso["Relato de los hechos"] = relato

    # Personas
    personas = []
    chunks = re.split(r"(?im)^\s*Calidad\s*[:：]", text)
    for ch in chunks[1:]:
        pdata = {lbl: "" for lbl in PERSON_FIELDS}
        first_line = ch.splitlines()[0] if ch.strip() else ""
        pdata["Calidad"] = cleanspace(first_line)
        for lbl in PERSON_FIELDS[1:]:
            v = extract_between(ch, lbl, PERSON_FIELDS)
            if not v:
                m = re.search(rf"(?im)^{re.escape(lbl)}\s*[:：]\s*(.+)$", ch)
                v = cleanspace(m.group(1)) if m else ""
            pdata[lbl] = v
        if any(str(v).strip().lower() not in ("", "nan", "-") for v in pdata.values()):
            personas.append(pdata)

    fila = dict(caso)
    for i, p in enumerate(personas, start=1):
        suf = f"_{i}"
        for lbl in PERSON_FIELDS:
            fila[f"{lbl}{suf}"] = p.get(lbl, "")
    return fila

# =================== CORTE TOLERANTE A OCR ===================
CASE_HEADER_LOOSE = re.compile(r"(?i)caso\s+noticia\s*[:：]\s*([^\n\r]{0,80})")
OCR_FIX_MAP = str.maketrans({'O':'0','o':'0','I':'1','l':'1','S':'5','B':'8'})

def sanitize_case_number(fragment: str) -> str:
    frag = fragment.replace('\xa0',' ').replace('\u200b',' ')
    frag = frag.translate(OCR_FIX_MAP)
    digits = re.sub(r"\D","", frag)
    return digits

def find_valid_case_headers(full_text: str):
    text = normalize_labels_multiline(full_text)
    accepted = []  # (start_index, id21)
    rejected = []  # {pos, snippet, digits, len}
    for m in CASE_HEADER_LOOSE.finditer(text):
        tail = m.group(1)
        num = sanitize_case_number(tail)
        if len(num) == 21:
            accepted.append((m.start(), num))
        else:
            snippet = (tail or '').strip()
            rejected.append({"pos": m.start(), "snippet": snippet, "digits": num, "len": len(num)})
    accepted.sort(key=lambda x: x[0])
    return text, accepted, rejected

def split_cases(full_text: str) -> list:
    text, accepted, rejected = find_valid_case_headers(full_text)
    # Depuración
    try:
        import csv
        with open(os.path.join(BASE_DIR, 'headers_debug.csv'), 'w', encoding='utf-8', newline='') as f:
            w = csv.writer(f); w.writerow(['tipo','pos','id_o_digits','len','snippet'])
            for pos, id21 in accepted: w.writerow(['OK', pos, id21, 21, ''])
            for r in rejected: w.writerow(['RECHAZADO', r['pos'], r['digits'], r['len'], r['snippet']])
    except Exception:
        pass
    if not accepted:
        return [text]
    starts = [pos for pos,_ in accepted]
    blocks = []
    for i, s in enumerate(starts):
        e = starts[i+1] if i+1 < len(starts) else len(text)
        blocks.append(text[s:e].strip())
    return blocks

def extract_case_id(block: str):
    m = CASE_HEADER_LOOSE.search(block)
    if not m: return None
    num = sanitize_case_number(m.group(1))
    return num if len(num) == 21 else None

# =================== REINDEX POR ROL + RESERVAS k_max ===================
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
    ns = sorted({int(m.group(1)) for c in row.index for m in [re.search(r"_(\d+)$", str(c))] if m})
    def _nonempty(x):
        try:
            import pandas as _pd
            if _pd.isna(x): return False
        except Exception: pass
        s = str(x).strip(); return s not in ("", "nan", "-")
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

    max_n = 0
    k_max = 0
    for base, indic, rest_sorted in per_row:
        m = len(indic) + len(rest_sorted)
        max_n = max(max_n, m)
        k_max = max(k_max, len(indic))
    if max_n == 0: max_n = 1

    new_rows = []
    for base, indic, rest_sorted in per_row:
        pairs = []
        for i, p in enumerate(indic, start=1): pairs.append((i, p))
        start_idx = k_max + 1
        for j, p in enumerate(rest_sorted, start=start_idx): pairs.append((j, p))
        new_rows.append(write_person_blocks_to_row(base, pairs))

    df2 = pd.DataFrame(new_rows)

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
    for i in range(1, k_max+1): desired += person_cols(i)
    desired += post_relato_pre_grado
    desired += grado
    for i in range(k_max+1, max_n+1): desired += person_cols(i)

    for col in desired:
        if col not in df2.columns: df2[col] = ""
    restantes = [c for c in df2.columns if c not in desired]
    final_cols = desired + restantes
    return df2.reindex(columns=final_cols)

# =================== EXCEL BLOQUEADO -> NOMBRE ALTERNO ===================

def safe_excel_path(base_path: str) -> str:
    try:
        f = open(base_path, 'ab'); f.close()
        return base_path
    except PermissionError:
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        root, ext = os.path.splitext(base_path)
        alt = f"{root} ({ts}){ext}"
        print(f"Archivo de salida bloqueado, guardando como: {alt}")
        return alt

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

    # Dividir (tolerante) y (opcionalmente) deduplicar
    case_chunks = split_cases(text_all)
    ids_for_log = []
    if DEDUPE_BY_ID:
        seen = set(); chunks_use = []
        for ch in case_chunks:
            cid = extract_case_id(ch)
            ids_for_log.append(cid or "<sin_id>")
            if not cid or cid in seen: continue
            seen.add(cid); chunks_use.append(ch)
    else:
        chunks_use = case_chunks
        for ch in case_chunks:
            ids_for_log.append(extract_case_id(ch) or "<sin_id>")

    uniques = len(set(ids_for_log))
    print(f"Casos detectados (apariciones/IDs únicos): {len(case_chunks)}/{uniques}. DEDUPE={'ON' if DEDUPE_BY_ID else 'OFF'}")

    # Parsear
    filas = []
    debug_rows = []
    for idx, chunk in enumerate(chunks_use, start=1):
        fila = parse_case_and_people(chunk)
        filas.append(fila)
        for k in CASE_FIELDS:
            debug_rows.append({"CasoIndex": idx, "Etiqueta": k, "Valor": fila.get(k, "")})
        for i in range(1, 8):
            for lbl in PERSON_FIELDS:
                debug_rows.append({"CasoIndex": idx, "Etiqueta": f"{lbl}_{i}", "Valor": fila.get(f"{lbl}_{i}", "")})

    print(f"Construyendo DataFrame con {len(filas)} filas...")
    df = pd.DataFrame(filas)
    df = reorder_and_expand(df)

    out_path = safe_excel_path(xlsx_path)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Datos", index=False)
        from openpyxl.utils import get_column_letter
        from openpyxl.styles import Alignment, Font
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
        for cell in ws[1]: cell.font = Font(bold=True)
        ws.freeze_panes = "A2"

    pd.DataFrame(debug_rows).to_csv(dbg_csv, index=False, encoding="utf-8-sig")

    time.sleep(0.3)
    print(f"Archivo generado: {out_path} (modo: {'OCR' if used_ocr else 'Texto directo'})")

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
