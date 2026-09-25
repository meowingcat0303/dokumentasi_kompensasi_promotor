import streamlit as st
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
import requests as _requests
from PIL import Image
import io
import json
import datetime
import uuid
import time
import concurrent.futures
import re
try:
    from streamlit_sortables import sort_items as _sort_items
    HAS_SORTABLES = True
except ImportError:
    HAS_SORTABLES = False

WIB = datetime.timezone(datetime.timedelta(hours=7))

# ── Password Admin ────────────────────────────────────────────────────────────
ADMIN_PASSWORD = "220603"

# ── Config ────────────────────────────────────────────────────────────────────
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
CUSTOMER_SHEET_ID   = "113E5fKvZ0wWloSbQ9IQo8QiCOz5C3zlJr28012PbXz4"
SUBMISSION_SHEET_ID = "1RC7v1fGmcz-9q4VowhBnf767P2N_ptonqlKuRSug0Ko"
IMAGE_MAX_PX  = 1920
IMAGE_QUALITY = 75

PROGRAM_OPTIONS  = ["WSP Permanen", "WSP Non Permanen", "WSP Program", "COI", "Awareness Investment"]
JENIS_OPTIONS    = ["SNT", "Thinplate", "PNT 3D", "Shopsign"]
BRAND_OPTIONS    = ["EVO", "ARJA", "Wismilak Kretek"]
ACTIVITY_OPTIONS = ["Termin 1", "Termin 2"]

TERMIN_MONTHS = 6
REMINDER_DAYS = 21

# ── Form Builder Tab di Google Sheets ─────────────────────────────────────────
FORM_CONFIG_TAB        = "Config_FormBuilder"
FIXED_FIELDS_TAB       = "Config_FixedFields"
SPREADSHEET_CONFIG_TAB = "Config_Spreadsheets"
FORMS_LIST_TAB         = "Config_FormsV2"      # Daftar form baru (multi-form)

# ── Definisi field tetap (urutan & metadata default) ──────────────────────────
DEFAULT_FIXED_FIELDS = [
    {"field_id": "tanggal",   "label": "Tanggal",             "type": "date",     "required": True,  "visible": True,  "options": []},
    {"field_id": "promotor",  "label": "Eksekutor (Promotor)","type": "dropdown", "required": True,  "visible": True,  "options": []},  # opsi dari Config_NamaPromotor
    {"field_id": "program",   "label": "Program Investment",  "type": "dropdown", "required": True,  "visible": True,  "options": PROGRAM_OPTIONS},
    {"field_id": "jenis",     "label": "Jenis Investment",    "type": "dropdown", "required": True,  "visible": True,  "options": JENIS_OPTIONS},
    {"field_id": "brand",     "label": "Brand",               "type": "dropdown", "required": True,  "visible": True,  "options": BRAND_OPTIONS},
    {"field_id": "activity",  "label": "Activity",            "type": "dropdown", "required": True,  "visible": True,  "options": ACTIVITY_OPTIONS},
    {"field_id": "customer",  "label": "Customer",            "type": "customer", "required": True,  "visible": True,  "options": []},
    {"field_id": "gps",       "label": "Lokasi GPS",          "type": "gps",      "required": False, "visible": True,  "options": []},
    {"field_id": "foto_ktp",       "label": "Foto KTP Pemilik",      "type": "photo", "required": True,  "visible": True, "options": []},
    {"field_id": "foto_sunblind",  "label": "Foto Sunblind",          "type": "photo", "required": True,  "visible": True, "options": []},
    {"field_id": "foto_display",   "label": "Foto Display",           "type": "photo", "required": False, "visible": True, "options": []},
    {"field_id": "foto_kompens",   "label": "Foto Bukti Kompensasi", "type": "photo", "required": True,  "visible": True, "options": []},
    {"field_id": "catatan",   "label": "Catatan Tambahan",    "type": "textarea", "required": False, "visible": True,  "options": []},
]

# ── Auth ──────────────────────────────────────────────────────────────────────
@st.cache_resource
def get_credentials():
    info = json.loads(st.secrets["gcp_json"])
    return Credentials.from_service_account_info(info, scopes=SCOPES)

@st.cache_resource(ttl=2700)
def get_gspread():
    return gspread.authorize(get_credentials())

# ── Data loaders ──────────────────────────────────────────────────────────────
def _sheet_to_df(ws) -> pd.DataFrame:
    rows = ws.get_all_values()
    if not rows:
        return pd.DataFrame()
    headers = [str(h).strip() for h in rows[0]]
    return pd.DataFrame(rows[1:], columns=headers)

@st.cache_data(ttl=600)
def load_customers():
    gc = get_gspread()
    ws = gc.open_by_key(CUSTOMER_SHEET_ID).sheet1
    df = _sheet_to_df(ws)
    cols = df.columns.tolist()
    code_col = next((c for c in cols if "customerno" in c.lower() or "kode" in c.lower() or "no" in c.lower()), cols[0])
    name_col = next((c for c in cols if c.lower() == "customer" or "nama" in c.lower() or "toko" in c.lower()), cols[0])
    addr_col = next((c for c in cols if "alamat" in c.lower() or "address" in c.lower()), None)
    sales_group_col    = next((c for c in cols if "sales group" in c.lower() or "salesgroup" in c.lower()), None)
    sales_district_col = next((c for c in cols if "sales district" in c.lower() or "salesdistrict" in c.lower()), None)
    status_col = next((c for c in cols if c.lower() == "status"), None)

    if status_col:
        df_aktif = df[df[status_col].str.strip().str.lower() == "aktif"].copy()
    else:
        df_aktif = df.copy()

    def make_label(r):
        label = f"{r[name_col]} ({r[code_col]})"
        if addr_col:
            label += f" — {str(r[addr_col])[:40]}"
        return label

    df_aktif["_label"] = df_aktif.apply(make_label, axis=1)

    meta = {
        "code_col": code_col,
        "name_col": name_col,
        "addr_col": addr_col,
        "sales_group_col": sales_group_col,
        "sales_district_col": sales_district_col,
        "status_col": status_col,
        "n_total": len(df),
        "n_aktif": len(df_aktif),
    }
    return df_aktif, meta

@st.cache_data(ttl=120)
def load_config_list(tab_name: str) -> list:
    try:
        gc = get_gspread()
        ws = gc.open_by_key(SUBMISSION_SHEET_ID).worksheet(tab_name)
        rows = ws.get_all_values()
        if len(rows) < 2:
            return []
        return sorted([str(r[0]).strip() for r in rows[1:] if r and str(r[0]).strip()])
    except Exception:
        get_gspread.clear()
        try:
            gc = get_gspread()
            ws = gc.open_by_key(SUBMISSION_SHEET_ID).worksheet(tab_name)
            rows = ws.get_all_values()
            if len(rows) < 2:
                return []
            return sorted([str(r[0]).strip() for r in rows[1:] if r and str(r[0]).strip()])
        except Exception as e:
            st.warning(f"Gagal load tab '{tab_name}': {e}")
            return []

@st.cache_data(ttl=120)
def load_all_config():
    promotors = load_config_list("Config_NamaPromotor")
    rayon     = load_config_list("Config_Rayon")
    zona      = load_config_list("Config_Zona")
    return promotors, rayon, zona

@st.cache_data(ttl=60)
def load_submissions() -> pd.DataFrame:
    try:
        gc = get_gspread()
        ws = gc.open_by_key(SUBMISSION_SHEET_ID).worksheet("Submission")
        return _sheet_to_df(ws)
    except Exception as e:
        st.warning(f"Gagal memuat submission: {e}")
        return pd.DataFrame()

# ── Form Builder: Load & Save ke Google Sheets ────────────────────────────────
def ensure_form_config_tab():
    """Pastikan tab Config_FormBuilder ada, buat kalau belum ada."""
    gc = get_gspread()
    wb = gc.open_by_key(SUBMISSION_SHEET_ID)
    try:
        ws = wb.worksheet(FORM_CONFIG_TAB)
    except Exception:
        ws = wb.add_worksheet(title=FORM_CONFIG_TAB, rows=200, cols=20)
        # Header
        headers = [
            "question_id", "order", "section", "label", "type",
            "required", "options", "placeholder", "description",
            "min_value", "max_value", "allow_other", "visible", "field_id",
        ]
        ws.update([headers], "A1")
    return ws

def _default_unified_questions() -> list:
    """Konversi DEFAULT_FIXED_FIELDS menjadi format unified question list."""
    qs = []
    for i, f in enumerate(DEFAULT_FIXED_FIELDS):
        qs.append({
            "question_id": f["field_id"],
            "field_id": f["field_id"],
            "order": i + 1,
            "section": "",
            "label": f["label"],
            "type": f["type"],
            "required": f["required"],
            "options": list(f["options"]),
            "placeholder": "",
            "description": "",
            "min_value": "",
            "max_value": "",
            "allow_other": False,
            "visible": f["visible"],
        })
    return qs

@st.cache_data(ttl=30)
def load_form_config() -> list:
    """Load konfigurasi pertanyaan dari Google Sheets (unified: fixed + tambahan)."""
    try:
        gc = get_gspread()
        wb = gc.open_by_key(SUBMISSION_SHEET_ID)
        try:
            ws = wb.worksheet(FORM_CONFIG_TAB)
        except Exception:
            return _default_unified_questions()
        rows = ws.get_all_values()
        if len(rows) < 2:
            return _default_unified_questions()
        headers = rows[0]
        questions = []
        for r in rows[1:]:
            if not any(r):
                continue
            row_dict = dict(zip(headers, r + [""] * (len(headers) - len(r))))
            q = {
                "question_id": row_dict.get("question_id", str(uuid.uuid4())[:8]),
                "field_id": row_dict.get("field_id", ""),
                "order": int(row_dict.get("order", 0)),
                "section": row_dict.get("section", ""),
                "label": row_dict.get("label", ""),
                "type": row_dict.get("type", "text"),
                "required": row_dict.get("required", "false").lower() == "true",
                "options": [o.strip() for o in row_dict.get("options", "").split("|") if o.strip()],
                "placeholder": row_dict.get("placeholder", ""),
                "description": row_dict.get("description", ""),
                "min_value": row_dict.get("min_value", ""),
                "max_value": row_dict.get("max_value", ""),
                "allow_other": row_dict.get("allow_other", "false").lower() == "true",
                "visible": row_dict.get("visible", "true").lower() != "false",
            }
            questions.append(q)
        questions.sort(key=lambda x: x["order"])
        return questions
    except Exception as e:
        st.warning(f"Gagal load form config: {e}")
        return _default_unified_questions()

def save_form_config(questions: list):
    """Simpan semua pertanyaan ke Google Sheets (unified)."""
    gc = get_gspread()
    ws = ensure_form_config_tab()
    headers = [
        "question_id", "order", "section", "label", "type",
        "required", "options", "placeholder", "description",
        "min_value", "max_value", "allow_other", "visible", "field_id",
    ]
    rows = [headers]
    for i, q in enumerate(questions):
        row = [
            q.get("question_id", str(uuid.uuid4())[:8]),
            str(i + 1),
            q.get("section", ""),
            q.get("label", ""),
            q.get("type", "text"),
            str(q.get("required", False)).lower(),
            "|".join(q.get("options", [])),
            q.get("placeholder", ""),
            q.get("description", ""),
            str(q.get("min_value", "")),
            str(q.get("max_value", "")),
            str(q.get("allow_other", False)).lower(),
            str(q.get("visible", True)).lower(),
            q.get("field_id", ""),
        ]
        rows.append(row)
    ws.clear()
    ws.update(rows, "A1")
    load_form_config.clear()

# ── Fixed Fields: Load & Save ─────────────────────────────────────────────────
@st.cache_data(ttl=30)
def load_fixed_fields() -> list:
    """Load konfigurasi field tetap. Fallback ke DEFAULT_FIXED_FIELDS jika belum ada."""
    try:
        gc = get_gspread()
        wb = gc.open_by_key(SUBMISSION_SHEET_ID)
        try:
            ws = wb.worksheet(FIXED_FIELDS_TAB)
        except Exception:
            return [dict(f) for f in DEFAULT_FIXED_FIELDS]
        rows = ws.get_all_values()
        if len(rows) < 2:
            return [dict(f) for f in DEFAULT_FIXED_FIELDS]
        headers = rows[0]
        result = []
        for r in rows[1:]:
            if not any(r):
                continue
            d = dict(zip(headers, r + [""] * (len(headers) - len(r))))
            result.append({
                "field_id": d.get("field_id", ""),
                "label":    d.get("label", ""),
                "type":     d.get("type", "text"),
                "required": d.get("required", "true").lower() == "true",
                "visible":  d.get("visible", "true").lower() != "false",
                "options":  [o.strip() for o in d.get("options", "").split("|") if o.strip()],
            })
        return result
    except Exception as e:
        st.warning(f"Gagal load fixed fields: {e}")
        return [dict(f) for f in DEFAULT_FIXED_FIELDS]

def save_fixed_fields(fields: list):
    gc = get_gspread()
    wb = gc.open_by_key(SUBMISSION_SHEET_ID)
    try:
        ws = wb.worksheet(FIXED_FIELDS_TAB)
    except Exception:
        ws = wb.add_worksheet(title=FIXED_FIELDS_TAB, rows=100, cols=10)
    headers = ["field_id", "label", "type", "required", "visible", "options"]
    rows = [headers]
    for f in fields:
        rows.append([
            f.get("field_id", ""),
            f.get("label", ""),
            f.get("type", "text"),
            str(f.get("required", True)).lower(),
            str(f.get("visible", True)).lower(),
            "|".join(f.get("options", [])),
        ])
    ws.clear()
    ws.update(rows, "A1")
    load_fixed_fields.clear()

def get_fixed_field(fields: list, field_id: str) -> dict:
    """Ambil config field by ID, fallback ke default."""
    for f in fields:
        if f["field_id"] == field_id:
            return f
    for f in DEFAULT_FIXED_FIELDS:
        if f["field_id"] == field_id:
            return dict(f)
    return {}

# ── Spreadsheet Manager: Load & Save ──────────────────────────────────────────
@st.cache_data(ttl=60)
def load_spreadsheet_list() -> list:
    """Load daftar spreadsheet yang terdaftar. Format: [{name, submission_id, customer_id, active}]"""
    try:
        gc = get_gspread()
        wb = gc.open_by_key(SUBMISSION_SHEET_ID)
        try:
            ws = wb.worksheet(SPREADSHEET_CONFIG_TAB)
        except Exception:
            return []
        rows = ws.get_all_values()
        if len(rows) < 2:
            return []
        headers = rows[0]
        result = []
        for r in rows[1:]:
            if not any(r):
                continue
            d = dict(zip(headers, r + [""] * (len(headers) - len(r))))
            result.append({
                "name":          d.get("name", ""),
                "submission_id": d.get("submission_id", ""),
                "customer_id":   d.get("customer_id", ""),
                "active":        d.get("active", "false").lower() == "true",
                "description":   d.get("description", ""),
            })
        return result
    except Exception:
        return []

def save_spreadsheet_list(sheets: list):
    gc = get_gspread()
    wb = gc.open_by_key(SUBMISSION_SHEET_ID)
    try:
        ws = wb.worksheet(SPREADSHEET_CONFIG_TAB)
    except Exception:
        ws = wb.add_worksheet(title=SPREADSHEET_CONFIG_TAB, rows=100, cols=10)
    headers = ["name", "submission_id", "customer_id", "active", "description"]
    rows = [headers]
    for s in sheets:
        rows.append([
            s.get("name", ""),
            s.get("submission_id", ""),
            s.get("customer_id", ""),
            str(s.get("active", False)).lower(),
            s.get("description", ""),
        ])
    ws.clear()
    ws.update(rows, "A1")
    load_spreadsheet_list.clear()

def get_active_sheet_ids() -> tuple[str, str]:
    """Return (submission_id, customer_id) yang aktif. Fallback ke konstanta default."""
    sheets = load_spreadsheet_list()
    for s in sheets:
        if s.get("active"):
            sub = s.get("submission_id") or SUBMISSION_SHEET_ID
            cus = s.get("customer_id") or CUSTOMER_SHEET_ID
            return sub, cus
    return SUBMISSION_SHEET_ID, CUSTOMER_SHEET_ID

# ── Multi-Form: Parse & Storage ───────────────────────────────────────────────
def parse_spreadsheet_id(url_or_id: str) -> str:
    """Ekstrak Spreadsheet ID dari URL Google Sheets atau kembalikan ID langsung."""
    url_or_id = url_or_id.strip()
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url_or_id)
    if m:
        return m.group(1)
    return url_or_id  # asumsi sudah berupa ID

@st.cache_data(ttl=60)
def load_forms_list() -> list:
    """Load daftar form dari tab Config_FormsV2."""
    try:
        gc = get_gspread()
        wb = gc.open_by_key(SUBMISSION_SHEET_ID)
        try:
            ws = wb.worksheet(FORMS_LIST_TAB)
        except Exception:
            return []
        rows = ws.get_all_values()
        if len(rows) < 2:
            return []
        headers = rows[0]
        result = []
        for r in rows[1:]:
            if not any(r):
                continue
            d = dict(zip(headers, r + [""] * max(0, len(headers) - len(r))))
            result.append({
                "form_id":        d.get("form_id", ""),
                "form_name":      d.get("form_name", ""),
                "description":    d.get("description", ""),
                "spreadsheet_url": d.get("spreadsheet_url", ""),
                "created_at":     d.get("created_at", ""),
                "active":         d.get("active", "true").lower() != "false",
                "apps_script_url": d.get("apps_script_url", ""),
            })
        return result
    except Exception:
        return []

def save_forms_list(forms: list):
    gc = get_gspread()
    wb = gc.open_by_key(SUBMISSION_SHEET_ID)
    try:
        ws = wb.worksheet(FORMS_LIST_TAB)
    except Exception:
        ws = wb.add_worksheet(title=FORMS_LIST_TAB, rows=200, cols=10)
    headers = ["form_id", "form_name", "description", "spreadsheet_url", "created_at", "active", "apps_script_url"]
    rows = [headers]
    for f in forms:
        rows.append([
            f.get("form_id", ""),
            f.get("form_name", ""),
            f.get("description", ""),
            f.get("spreadsheet_url", ""),
            f.get("created_at", ""),
            str(f.get("active", True)).lower(),
            f.get("apps_script_url", ""),
        ])
    ws.clear()
    ws.update(rows, "A1")
    load_forms_list.clear()

@st.cache_data(ttl=30)
def load_form_questions(form_id: str) -> list:
    """Load pertanyaan untuk form tertentu dari tab FormQ_<form_id>."""
    tab_name = f"FormQ_{form_id}"
    try:
        gc = get_gspread()
        wb = gc.open_by_key(SUBMISSION_SHEET_ID)
        try:
            ws = wb.worksheet(tab_name)
        except Exception:
            return []
        rows = ws.get_all_values()
        if len(rows) < 2:
            return []
        headers = rows[0]
        questions = []
        for r in rows[1:]:
            if not any(r):
                continue
            d = dict(zip(headers, r + [""] * max(0, len(headers) - len(r))))
            questions.append({
                "question_id": d.get("question_id", str(uuid.uuid4())[:8]),
                "order":       int(d.get("order", 0) or 0),
                "label":       d.get("label", ""),
                "type":        d.get("type", "text"),
                "required":    d.get("required", "false").lower() == "true",
                "options":     [o.strip() for o in d.get("options", "").split("|") if o.strip()],
                "placeholder": d.get("placeholder", ""),
                "description": d.get("description", ""),
                "min_value":   d.get("min_value", ""),
                "max_value":   d.get("max_value", ""),
                "allow_other": d.get("allow_other", "false").lower() == "true",
                "visible":     d.get("visible", "true").lower() != "false",
                "col_locked":  d.get("col_locked", "false").lower() == "true",
                "col_index":   int(d.get("col_index", 0) or 0),
                "section":     d.get("section", ""),
            })
        questions.sort(key=lambda x: x["order"])
        return questions
    except Exception as e:
        st.warning(f"Gagal load pertanyaan form {form_id}: {e}")
        return []

def save_form_questions(form_id: str, questions: list):
    """Simpan pertanyaan ke tab FormQ_<form_id>."""
    tab_name = f"FormQ_{form_id}"
    gc = get_gspread()
    wb = gc.open_by_key(SUBMISSION_SHEET_ID)
    try:
        ws = wb.worksheet(tab_name)
    except Exception:
        ws = wb.add_worksheet(title=tab_name, rows=300, cols=20)
    headers = [
        "question_id", "order", "label", "type", "required", "options",
        "placeholder", "description", "min_value", "max_value",
        "allow_other", "visible", "col_locked", "col_index", "section",
    ]
    rows = [headers]
    for i, q in enumerate(questions):
        col_index = q.get("col_index", 0)
        if not col_index:
            col_index = i + 1
        rows.append([
            q.get("question_id", str(uuid.uuid4())[:8]),
            str(i + 1),
            q.get("label", ""),
            q.get("type", "text"),
            str(q.get("required", False)).lower(),
            "|".join(q.get("options", [])),
            q.get("placeholder", ""),
            q.get("description", ""),
            str(q.get("min_value", "")),
            str(q.get("max_value", "")),
            str(q.get("allow_other", False)).lower(),
            str(q.get("visible", True)).lower(),
            str(q.get("col_locked", False)).lower(),
            str(col_index),
            q.get("section", ""),
        ])
    ws.clear()
    ws.update(rows, "A1")
    load_form_questions.clear()

def submit_to_public_spreadsheet(spreadsheet_url: str, questions: list, answers: dict):
    """
    Tulis jawaban ke spreadsheet yang dikonfigurasi admin.
    Urutan kolom: sorted by col_index (jika locked) atau by order.
    Baris pertama = header (label pertanyaan), baris berikutnya = data.
    """
    sheet_id = parse_spreadsheet_id(spreadsheet_url)
    gc = get_gspread()
    wb = gc.open_by_key(sheet_id)

    # Ambil atau buat sheet pertama
    try:
        ws = wb.sheet1
    except Exception:
        ws = wb.add_worksheet(title="Responses", rows=5000, cols=50)

    # Bangun urutan kolom: locked pakai col_index, unlocked pakai order
    def _col_key(q):
        if q.get("col_locked") and q.get("col_index", 0) > 0:
            return (0, q["col_index"])    # locked: prioritas absolut
        return (1, q["order"])            # unlocked: ikut urutan

    sorted_qs = sorted([q for q in questions if q.get("visible", True)], key=_col_key)

    # Cek apakah header sudah ada
    existing = ws.get_all_values()
    headers_needed = [q["label"] for q in sorted_qs] + ["Waktu Submit"]

    if not existing or existing[0] != headers_needed:
        # Tulis / update header row
        ws.update([headers_needed], "A1")

    # Bangun data row
    row_data = []
    for q in sorted_qs:
        row_data.append(str(answers.get(q["question_id"], "") or ""))
    row_data.append(datetime.datetime.now(WIB).strftime("%Y-%m-%d %H:%M:%S"))

    ws.append_row(row_data, value_input_option="USER_ENTERED")

# ── Image utils ───────────────────────────────────────────────────────────────
def compress_image(uploaded_file) -> bytes:
    """Kompres gambar. Kalau hasil masih >1.5MB, turunkan quality otomatis."""
    raw = uploaded_file.read() if hasattr(uploaded_file, "read") else uploaded_file
    img = Image.open(io.BytesIO(raw))
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
    # Rotasi otomatis sesuai EXIF (penting untuk foto dari kamera HP)
    try:
        import PIL.ExifTags as _et
        exif = img._getexif()
        if exif:
            orientation_key = next((k for k, v in _et.TAGS.items() if v == "Orientation"), None)
            if orientation_key and orientation_key in exif:
                ori = exif[orientation_key]
                rotate_map = {3: 180, 6: 270, 8: 90}
                if ori in rotate_map:
                    img = img.rotate(rotate_map[ori], expand=True)
    except Exception:
        pass
    ratio = IMAGE_MAX_PX / max(img.size)
    if ratio < 1:
        new_size = (int(img.width * ratio), int(img.height * ratio))
        img = img.resize(new_size, Image.LANCZOS)
    # Auto-turunkan quality kalau file masih besar
    quality = IMAGE_QUALITY
    for _ in range(3):
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=quality, optimize=True)
        if buf.tell() <= 1_500_000 or quality <= 50:
            break
        quality -= 15
    return buf.getvalue()

# ── Drive upload ──────────────────────────────────────────────────────────────
APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbz27id708tLEcf0eGWNC6BrA7TdHiFVfgsPL2b_xGDkWTqBD30tlGWpCXvQ8F2IXIjO/exec"

def upload_to_drive(data: bytes, filename: str, retries: int = 3, script_url: str = "") -> str:
    """Upload ke Drive via Apps Script. Thread-safe (tanpa st.warning)."""
    import base64
    target_url = script_url.strip() if script_url.strip() else APPS_SCRIPT_URL
    b64 = base64.b64encode(data).decode("utf-8")
    last_exc = None
    for attempt in range(retries):
        try:
            resp = _requests.post(
                target_url,
                data={"image": b64, "filename": filename},
                timeout=120,
            )
            resp.raise_for_status()
            result = resp.json()
            if result.get("success"):
                return result["url"]
            raise ValueError(f"Apps Script error: {result}")
        except Exception as e:
            last_exc = e
            if attempt < retries - 1:
                time.sleep(2 * (attempt + 1))
                continue
    raise last_exc

# ── Submission writer ─────────────────────────────────────────────────────────
def append_submission(row: list, retries: int = 3):
    """Tulis baris ke sheet Submission dengan retry."""
    last_exc = None
    for attempt in range(retries):
        try:
            gc = get_gspread()
            ws = gc.open_by_key(SUBMISSION_SHEET_ID).worksheet("Submission")
            ws.append_row(row, value_input_option="USER_ENTERED")
            return
        except Exception as e:
            last_exc = e
            if attempt < retries - 1:
                get_gspread.clear()  # reset koneksi
                time.sleep(2 * (attempt + 1))
                continue
    raise last_exc

# ── Termin / deadline logic ───────────────────────────────────────────────────
def add_months(d: datetime.date, months: int) -> datetime.date:
    month = d.month - 1 + months
    year  = d.year + month // 12
    month = month % 12 + 1
    import calendar
    day = min(d.day, calendar.monthrange(year, month)[1])
    return datetime.date(year, month, day)

def compute_deadline(date_str: str, activity: str) -> datetime.date | None:
    try:
        d = datetime.date.fromisoformat(str(date_str).strip())
        return add_months(d, TERMIN_MONTHS)
    except Exception:
        return None

def countdown_label(deadline: datetime.date, today: datetime.date) -> tuple[str, bool]:
    delta = (deadline - today).days
    if delta < 0:
        return f"Lewat {abs(delta)} hari", True
    elif delta == 0:
        return "Hari ini!", True
    elif delta <= REMINDER_DAYS:
        months_left = delta // 30
        days_left   = delta % 30
        if months_left > 0:
            return f"{months_left} bln {days_left} hr lagi", True
        return f"{delta} hari lagi", True
    else:
        months_left = delta // 30
        days_left   = delta % 30
        if months_left > 0:
            return f"{months_left} bln {days_left} hr lagi", False
        return f"{delta} hari lagi", False

# ── GPS JS ────────────────────────────────────────────────────────────────────
GPS_JS = """
<script>
function getLocation() {
    var statusEl = document.getElementById('gps-status');
    if (!navigator.geolocation) {
        statusEl.innerText = 'Browser tidak mendukung GPS.';
        return;
    }
    statusEl.innerText = 'Mengambil lokasi...';
    navigator.geolocation.getCurrentPosition(
        function(pos) {
            var lat = pos.coords.latitude.toFixed(6);
            var lng = pos.coords.longitude.toFixed(6);
            var acc = Math.round(pos.coords.accuracy);
            var coords = lat + ', ' + lng;
            statusEl.innerText = 'Lokasi ditemukan: ' + coords + ' (akurasi +/- ' + acc + ' m)';
            var target = window.parent.document.querySelector('input[aria-label="koordinat_gps"]');
            if (target) {
                var setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
                setter.call(target, coords);
                target.dispatchEvent(new Event('input', { bubbles: true }));
            }
        },
        function(err) {
            statusEl.innerText = 'Gagal mendapatkan lokasi: ' + err.message;
        },
        { enableHighAccuracy: true, timeout: 20000, maximumAge: 0 }
    );
}
</script>
<button onclick="getLocation()" style="
    background:#1f77b4; color:white; border:none; padding:10px 22px;
    border-radius:6px; font-size:15px; cursor:pointer; margin-bottom:8px; width:100%;">
    Ambil Lokasi GPS Otomatis
</button>
<div id="gps-status" style="color:#444; font-size:13px; margin-top:6px;"></div>
"""

# ══════════════════════════════════════════════════════════════════════════════
# PAGE CONFIG & NAVIGATION
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(page_title="Dokumentasi Investment Lapangan", layout="centered")

# ── Routing berbasis query params ─────────────────────────────────────────────
_qp        = st.query_params
_mode      = _qp.get("mode", "")        # "admin"
_form_id   = _qp.get("form", "")        # form ID publik

if _form_id:
    _PAGE = "public_form"
elif _mode == "admin":
    _PAGE = "admin"
else:
    _PAGE = "input"
    # Sidebar hanya untuk mode default
    page = st.sidebar.radio(
        "Menu",
        ["Input Dokumentasi", "Monitoring (Admin)"],
        key="nav_page",
    )
    if page == "Monitoring (Admin)":
        _PAGE = "admin_sidebar"

# ══════════════════════════════════════════════════════════════════════════════
# PAGE 1 — INPUT DOKUMENTASI
# ══════════════════════════════════════════════════════════════════════════════
if _PAGE == "input":

    st.title("Dokumentasi Investment Lapangan")

    if "submit_info" in st.session_state:
        info = st.session_state.submit_info
        foto_errors = info.get("foto_errors", [])
        if foto_errors:
            st.warning(
                "⚠️ Data tersimpan, tetapi **beberapa foto gagal diupload**:\n\n"
                + "\n".join(f"- {e}" for e in foto_errors)
                + "\n\nHubungi admin dengan ID di bawah untuk upload ulang."
            )
        else:
            st.success("✅ Dokumentasi berhasil dikirim dan tersimpan.")
        st.markdown(f"""
**Ringkasan pengiriman:**

| | |
|---|---|
| ID Unik | `{info['unique_id']}` |
| Outlet | {info['outlet']} ({info['kode']}) |
| Eksekutor | {info['promotor']} |
| Waktu Kirim | {info['waktu']} |

Data sudah masuk ke sistem. Anda dapat menutup halaman ini atau melanjutkan input baru.
""")
        if st.button("Input Dokumentasi Baru", type="primary"):
            del st.session_state["submit_info"]
            st.rerun()
        st.stop()

    with st.spinner("Memuat data..."):
        promotors, rayon_list, zona_list = load_all_config()
        df_cust, meta = load_customers()
        all_questions = load_form_config()

    code_col           = meta["code_col"]
    name_col           = meta["name_col"]
    addr_col           = meta["addr_col"]
    sales_group_col    = meta["sales_group_col"]
    sales_district_col = meta["sales_district_col"]

    st.caption(f"Data customer: **{meta['n_aktif']:,} toko aktif** dari {meta['n_total']:,} total")

    if "gps_coords" not in st.session_state:
        st.session_state.gps_coords = ""

    # Variabel untuk field sistem — diisi saat render pertanyaan
    _vals = {
        "tgl": datetime.datetime.now(WIB).date(),
        "promotor": "—",
        "program": "—",
        "jenis": "—",
        "brand": "—",
        "activity": "—",
        "selected_customer": None,
        "rayon_final": "",
        "zona_final": "",
        "foto_ktp": None,
        "foto_sunblind": None,
        "foto_display": None,
        "foto_kompens": None,
        "catatan": "",
    }
    extra_answers = {}
    _vals["_customer_rendered"] = False  # pastikan rayon/zona mode hanya render sekali

    visible_questions = [q for q in all_questions if q.get("visible", True)]

    # ── Helper: render field sistem ───────────────────────────────────────────
    def _render_system_field(q):
        fid   = q.get("field_id", "")
        label = q["label"] + (" *" if q["required"] else " (opsional)")
        opts  = q.get("options", [])

        if fid == "tanggal":
            _vals["tgl"] = st.date_input(label, value=datetime.datetime.now(WIB).date(), key="tgl_sel")

        elif fid == "promotor":
            _vals["promotor"] = st.selectbox(label, ["— Pilih —"] + promotors, key="promotor_sel")

        elif fid == "program":
            _vals["program"] = st.selectbox(label, ["— Pilih —"] + (opts or PROGRAM_OPTIONS), key="program_sel")

        elif fid == "jenis":
            _vals["jenis"] = st.selectbox(label, ["— Pilih —"] + (opts or JENIS_OPTIONS), key="jenis_sel")

        elif fid == "brand":
            _vals["brand"] = st.selectbox(label, ["— Pilih —"] + (opts or BRAND_OPTIONS), key="brand_sel")

        elif fid == "activity":
            _vals["activity"] = st.selectbox(label, ["— Pilih —"] + (opts or ACTIVITY_OPTIONS), key="activity_sel")

        elif fid == "customer":
            if not _vals["_customer_rendered"]:
                _vals["_customer_rendered"] = True
                st.radio(
                    "Mode Rayon & Zona:",
                    ["Otomatis (dari data customer)", "Manual (pilih sendiri)"],
                    key="rayon_zona_mode",
                    horizontal=True,
                )
            mode_otomatis = st.session_state.get("rayon_zona_mode", "Otomatis (dari data customer)") == "Otomatis (dari data customer)"
            if not mode_otomatis:
                col_r, col_z = st.columns(2)
                with col_r:
                    rayon_sel = st.selectbox("Rayon", ["— Pilih —"] + rayon_list, key="rayon_sel")
                with col_z:
                    zona_sel = st.selectbox("Zona", ["— Pilih —"] + zona_list, key="zona_sel")
            if df_cust.empty:
                st.error("Data customer tidak ditemukan.")
                st.stop()
            chosen_label = st.selectbox(
                label,
                options=["— Pilih —"] + df_cust["_label"].tolist(),
                key="cust_sel",
            )
            if chosen_label != "— Pilih —":
                match = df_cust[df_cust["_label"] == chosen_label]
                if not match.empty:
                    cust = match.iloc[0].to_dict()
                    _vals["selected_customer"] = cust
                    if mode_otomatis:
                        _vals["rayon_final"] = str(cust[sales_group_col]).strip() if sales_group_col else ""
                        _vals["zona_final"]  = str(cust[sales_district_col]).strip() if sales_district_col else ""
                        col_r, col_z = st.columns(2)
                        with col_r:
                            st.info(f"**Rayon:** {_vals['rayon_final'] or '—'}")
                        with col_z:
                            st.info(f"**Zona:** {_vals['zona_final'] or '—'}")
                    else:
                        _vals["rayon_final"] = rayon_sel if rayon_sel != "— Pilih —" else ""
                        _vals["zona_final"]  = zona_sel  if zona_sel  != "— Pilih —" else ""
                    st.success(
                        f"**{cust[name_col]}** ({cust[code_col]})"
                        + (f" — {cust[addr_col]}" if addr_col else "")
                    )

        elif fid == "gps":
            st.markdown(f"**{q['label']}**")
            st.components.v1.html(GPS_JS, height=90)
            gps_input = st.text_input(
                "koordinat_gps",
                value=st.session_state.gps_coords,
                key="gps_manual",
                placeholder="Klik tombol di atas untuk mengisi otomatis",
                label_visibility="collapsed",
            )
            if gps_input:
                st.session_state.gps_coords = gps_input

        elif fid == "foto_ktp":
            _vals["foto_ktp"] = st.file_uploader(label, type=["jpg","jpeg","png"], key="foto_ktp")
            st.caption("Foto dikompres otomatis (max 1920px, JPEG 75%).")

        elif fid == "foto_sunblind":
            _vals["foto_sunblind"] = st.file_uploader(label, type=["jpg","jpeg","png"], key="foto_sunblind")

        elif fid == "foto_display":
            _vals["foto_display"] = st.file_uploader(label, type=["jpg","jpeg","png"], key="foto_display")

        elif fid == "foto_kompens":
            _vals["foto_kompens"] = st.file_uploader(label, type=["jpg","jpeg","png"], key="foto_kompens")

        elif fid == "catatan":
            _vals["catatan"] = st.text_area(label, key="catatan_input")

    # ── Helper: render pertanyaan bebas ───────────────────────────────────────
    def _render_free_field(q):
        qid   = q["question_id"]
        qtype = q["type"]
        label = q["label"] + (" *" if q["required"] else " (opsional)")
        desc  = q.get("description", "")
        if desc:
            st.caption(desc)

        if qtype == "text":
            extra_answers[qid] = st.text_input(label, placeholder=q.get("placeholder",""), key=f"eq_{qid}")
        elif qtype == "textarea":
            extra_answers[qid] = st.text_area(label, placeholder=q.get("placeholder",""), key=f"eq_{qid}")
        elif qtype == "number":
            min_v = float(q["min_value"]) if q.get("min_value") else None
            max_v = float(q["max_value"]) if q.get("max_value") else None
            extra_answers[qid] = st.number_input(label, min_value=min_v, max_value=max_v, key=f"eq_{qid}")
        elif qtype == "date":
            extra_answers[qid] = str(st.date_input(label, key=f"eq_{qid}"))
        elif qtype == "time":
            extra_answers[qid] = str(st.time_input(label, key=f"eq_{qid}"))
        elif qtype == "dropdown":
            opts = ["— Pilih —"] + q.get("options", [])
            if q.get("allow_other"):
                opts.append("Lainnya...")
            sel = st.selectbox(label, opts, key=f"eq_{qid}")
            if sel == "Lainnya...":
                sel = st.text_input("Sebutkan:", key=f"eq_{qid}_other")
            extra_answers[qid] = sel if sel != "— Pilih —" else ""
        elif qtype == "radio":
            opts = q.get("options", [])
            if q.get("allow_other"):
                opts = opts + ["Lainnya..."]
            sel = st.radio(label, opts, key=f"eq_{qid}", horizontal=True)
            if sel == "Lainnya...":
                sel = st.text_input("Sebutkan:", key=f"eq_{qid}_other")
            extra_answers[qid] = sel or ""
        elif qtype == "checkbox":
            opts = q.get("options", [])
            selected_opts = []
            st.markdown(f"**{label}**")
            for opt in opts:
                if st.checkbox(opt, key=f"eq_{qid}_{opt}"):
                    selected_opts.append(opt)
            if q.get("allow_other"):
                other_val = st.text_input("Lainnya:", key=f"eq_{qid}_other")
                if other_val:
                    selected_opts.append(other_val)
            extra_answers[qid] = ", ".join(selected_opts)
        elif qtype == "scale":
            min_v = int(q["min_value"]) if q.get("min_value") else 1
            max_v = int(q["max_value"]) if q.get("max_value") else 5
            extra_answers[qid] = str(st.slider(label, min_value=min_v, max_value=max_v, key=f"eq_{qid}"))
        elif qtype == "photo":
            uploaded = st.file_uploader(label, type=["jpg","jpeg","png"], key=f"eq_{qid}")
            extra_answers[qid] = uploaded
        elif qtype == "yes_no":
            extra_answers[qid] = "Ya" if st.toggle(label, key=f"eq_{qid}") else "Tidak"

    # ── Render semua pertanyaan sesuai urutan ─────────────────────────────────
    current_section = None
    for q in visible_questions:
        sec = q.get("section", "")
        if sec and sec != current_section:
            st.divider()
            st.subheader(sec)
            current_section = sec
        fid = q.get("field_id", "")
        if fid:
            _render_system_field(q)
        else:
            _render_free_field(q)

    st.divider()

    # Shorthand untuk submit
    tgl              = _vals["tgl"]
    promotor         = _vals["promotor"]
    program          = _vals["program"]
    jenis            = _vals["jenis"]
    brand            = _vals["brand"]
    activity         = _vals["activity"]
    selected_customer = _vals["selected_customer"]
    rayon_final      = _vals["rayon_final"]
    zona_final       = _vals["zona_final"]
    foto_ktp         = _vals["foto_ktp"]
    foto_sunblind    = _vals["foto_sunblind"]
    foto_display     = _vals["foto_display"]
    foto_kompens     = _vals["foto_kompens"]
    catatan          = _vals["catatan"]

    # Cari field sistem yang visible & required dari config
    def _q_visible(fid): return any(q.get("field_id")==fid and q.get("visible",True) for q in all_questions)
    def _q_required(fid): return any(q.get("field_id")==fid and q.get("required",True) for q in all_questions)
    def _q_label(fid, default): 
        for q in all_questions:
            if q.get("field_id") == fid:
                return q.get("label", default)
        return default

    # Submit
    if st.button("Submit Dokumentasi", key="btn_submit", type="primary"):
        errors = []
        if _q_visible("promotor") and _q_required("promotor") and promotor == "— Pilih —":
            errors.append(f"{_q_label('promotor','Eksekutor')} wajib dipilih")
        if _q_visible("program") and _q_required("program") and program == "— Pilih —":
            errors.append(f"{_q_label('program','Program Investment')} wajib dipilih")
        if _q_visible("jenis") and _q_required("jenis") and jenis == "— Pilih —":
            errors.append(f"{_q_label('jenis','Jenis Investment')} wajib dipilih")
        if _q_visible("brand") and _q_required("brand") and brand == "— Pilih —":
            errors.append(f"{_q_label('brand','Brand')} wajib dipilih")
        if _q_visible("activity") and _q_required("activity") and activity == "— Pilih —":
            errors.append(f"{_q_label('activity','Activity')} wajib dipilih")
        if _q_visible("customer") and _q_required("customer") and not selected_customer:
            errors.append(f"{_q_label('customer','Customer')} belum dipilih")
        if _q_visible("customer") and selected_customer and not rayon_final:
            errors.append("Rayon tidak ditemukan / belum dipilih")
        if _q_visible("customer") and selected_customer and not zona_final:
            errors.append("Zona tidak ditemukan / belum dipilih")
        if _q_visible("foto_ktp") and _q_required("foto_ktp") and not foto_ktp:
            errors.append(f"{_q_label('foto_ktp','Foto KTP')} wajib diupload")
        if _q_visible("foto_sunblind") and _q_required("foto_sunblind") and not foto_sunblind:
            errors.append(f"{_q_label('foto_sunblind','Foto Sunblind')} wajib diupload")
        if _q_visible("foto_kompens") and _q_required("foto_kompens") and not foto_kompens:
            errors.append(f"{_q_label('foto_kompens','Foto Bukti Kompensasi')} wajib diupload")

        # Validasi pertanyaan bebas yang wajib
        free_visible = [q for q in visible_questions if not q.get("field_id")]
        for q in free_visible:
            if q["required"]:
                qid = q["question_id"]
                val = extra_answers.get(qid, "")
                if q["type"] == "photo":
                    if val is None:
                        errors.append(f"'{q['label']}' wajib diisi")
                elif not val or val == "— Pilih —":
                    errors.append(f"'{q['label']}' wajib diisi")

        if errors:
            for e in errors:
                st.error(e)
        else:
            with st.spinner("Mengupload foto dan menyimpan data, harap tunggu..."):
                ts   = datetime.datetime.now(WIB).strftime("%Y%m%d_%H%M%S")
                uid  = str(uuid.uuid4())[:8]
                kode = selected_customer.get(code_col, "unknown") if selected_customer else "unknown"

                # ── Kumpulkan semua foto yang perlu diupload ──────────────────
                upload_tasks = {}  # key → (bytes, filename)

                def _prep(file, label):
                    if file is None:
                        return
                    fname = f"{ts}_{kode}_{promotor}_{label}_{uid}.jpg"
                    upload_tasks[label] = (compress_image(file), fname)

                _prep(foto_ktp,      "ktp")
                _prep(foto_sunblind, "sunblind")
                _prep(foto_display,  "display")
                _prep(foto_kompens,  "kompensasi")

                free_visible = [q for q in visible_questions if not q.get("field_id")]
                extra_answers_str = {}
                for q in free_visible:
                    qid = q["question_id"]
                    val = extra_answers.get(qid, "")
                    if q["type"] == "photo" and val is not None:
                        _prep(val, f"extra_{qid}")
                    else:
                        extra_answers_str[qid] = str(val) if val else ""

                # ── Upload paralel ─────────────────────────────────────────────
                upload_results = {}   # key → url atau ""
                upload_errors  = []   # label foto yang gagal (non-fatal)

                def _upload_one(item):
                    label, (data, fname) = item
                    try:
                        url = upload_to_drive(data, fname)
                        return label, url
                    except Exception as exc:
                        return label, f"ERROR: {exc}"

                with concurrent.futures.ThreadPoolExecutor(max_workers=4) as pool:
                    futures = pool.map(_upload_one, upload_tasks.items())
                    for label, result in futures:
                        if isinstance(result, str) and result.startswith("ERROR:"):
                            upload_errors.append(f"{label}: {result[7:]}")
                            upload_results[label] = ""
                        else:
                            upload_results[label] = result

                # Isi extra_answers_str untuk foto bebas
                for q in free_visible:
                    qid = q["question_id"]
                    if q["type"] == "photo":
                        extra_answers_str[qid] = upload_results.get(f"extra_{qid}", "")

                unique_id   = f"{ts}_{uid}"
                submit_time = datetime.datetime.now(WIB).strftime("%Y-%m-%d %H:%M:%S")

                row = [
                    unique_id,
                    str(tgl),
                    promotor,
                    program,
                    jenis,
                    brand,
                    activity,
                    rayon_final,
                    zona_final,
                    selected_customer.get(name_col, "") if selected_customer else "",
                    kode,
                    upload_results.get("sunblind", ""),
                    upload_results.get("display",  ""),
                    upload_results.get("kompensasi", ""),
                    upload_results.get("ktp",      ""),
                    st.session_state.gps_coords,
                    catatan,
                    submit_time,
                ]
                for q in free_visible:
                    row.append(extra_answers_str.get(q["question_id"], ""))

                # ── Simpan ke Sheets (dengan retry) ───────────────────────────
                try:
                    append_submission(row)
                except Exception as exc:
                    st.error(
                        f"❌ Data gagal tersimpan ke Sheets: {exc}\n\n"
                        f"**ID unik submission Anda: `{unique_id}`** — screenshot ini dan hubungi admin."
                    )
                    st.stop()

                # ── Sukses — tampilkan peringatan foto gagal (non-fatal) ───────
                st.session_state.submit_info = {
                    "unique_id": unique_id,
                    "outlet":    selected_customer.get(name_col, "-") if selected_customer else "-",
                    "kode":      kode,
                    "promotor":  promotor,
                    "waktu":     submit_time,
                    "foto_errors": upload_errors,
                }
                st.session_state.gps_coords = ""
                # Hanya clear cache submission, bukan semua cache
                load_submissions.clear()
                st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: PUBLIC FORM  (?form=<form_id>)
# ══════════════════════════════════════════════════════════════════════════════
elif _PAGE == "public_form":
    @st.cache_data(ttl=30)
    def _get_form_meta(fid):
        forms = load_forms_list()
        for f in forms:
            if f["form_id"] == fid:
                return f
        return None

    form_meta = _get_form_meta(_form_id)
    if not form_meta or not form_meta.get("active"):
        st.error("❌ Form tidak ditemukan atau sudah tidak aktif.")
        st.stop()

    st.title(form_meta["form_name"])
    if form_meta.get("description"):
        st.caption(form_meta["description"])

    form_questions = load_form_questions(_form_id)
    visible_qs = [q for q in form_questions if q.get("visible", True)]

    if not visible_qs:
        st.warning("Form ini belum memiliki pertanyaan.")
        st.stop()

    pf_answers = {}

    def _render_pf_field(q):
        qid   = q["question_id"]
        qtype = q["type"]
        label = q["label"] + (" *" if q["required"] else " (opsional)")
        if q.get("description"):
            st.caption(q["description"])

        if qtype == "text":
            pf_answers[qid] = st.text_input(label, placeholder=q.get("placeholder",""), key=f"pf_{qid}")
        elif qtype == "textarea":
            pf_answers[qid] = st.text_area(label, placeholder=q.get("placeholder",""), key=f"pf_{qid}")
        elif qtype == "number":
            min_v = float(q["min_value"]) if q.get("min_value") else None
            max_v = float(q["max_value"]) if q.get("max_value") else None
            pf_answers[qid] = st.number_input(label, min_value=min_v, max_value=max_v, key=f"pf_{qid}")
        elif qtype == "date":
            pf_answers[qid] = str(st.date_input(label, key=f"pf_{qid}"))
        elif qtype == "time":
            pf_answers[qid] = str(st.time_input(label, key=f"pf_{qid}"))
        elif qtype == "dropdown":
            opts = ["— Pilih —"] + q.get("options", [])
            if q.get("allow_other"):
                opts.append("Lainnya...")
            sel = st.selectbox(label, opts, key=f"pf_{qid}")
            if sel == "Lainnya...":
                sel = st.text_input("Sebutkan:", key=f"pf_{qid}_other")
            pf_answers[qid] = sel if sel != "— Pilih —" else ""
        elif qtype == "radio":
            opts = q.get("options", [])
            if q.get("allow_other"):
                opts = opts + ["Lainnya..."]
            sel = st.radio(label, opts, key=f"pf_{qid}", horizontal=True)
            if sel == "Lainnya...":
                sel = st.text_input("Sebutkan:", key=f"pf_{qid}_other")
            pf_answers[qid] = sel or ""
        elif qtype == "checkbox":
            selected_opts = []
            st.markdown(f"**{label}**")
            for opt in q.get("options", []):
                if st.checkbox(opt, key=f"pf_{qid}_{opt}"):
                    selected_opts.append(opt)
            if q.get("allow_other"):
                other_val = st.text_input("Lainnya:", key=f"pf_{qid}_other")
                if other_val:
                    selected_opts.append(other_val)
            pf_answers[qid] = ", ".join(selected_opts)
        elif qtype == "scale":
            min_v = int(q["min_value"]) if q.get("min_value") else 1
            max_v = int(q["max_value"]) if q.get("max_value") else 5
            pf_answers[qid] = str(st.slider(label, min_value=min_v, max_value=max_v, key=f"pf_{qid}"))
        elif qtype == "yes_no":
            pf_answers[qid] = "Ya" if st.toggle(label, key=f"pf_{qid}") else "Tidak"
        elif qtype == "photo":
            up = st.file_uploader(label, type=["jpg","jpeg","png"], key=f"pf_{qid}")
            pf_answers[qid] = "[foto]" if up else ""
        elif qtype == "paragraph":
            st.markdown(q.get("description", q["label"]))

    current_section = None
    for q in visible_qs:
        sec = q.get("section", "")
        if sec and sec != current_section:
            st.divider()
            st.subheader(sec)
            current_section = sec
        _render_pf_field(q)

    st.divider()

    if "pf_submitted" in st.session_state and st.session_state.get("pf_form_id") == _form_id:
        st.success("✅ Jawaban Anda berhasil dikirim! Terima kasih.")
        if st.button("Isi ulang form"):
            del st.session_state["pf_submitted"]
            st.rerun()
    else:
        if st.button("Kirim Jawaban", type="primary", key="pf_submit"):
            errors = []
            for q in visible_qs:
                if q["required"]:
                    val = pf_answers.get(q["question_id"], "")
                    if not val or val == "— Pilih —":
                        errors.append(f"'{q['label']}' wajib diisi")
            if errors:
                for e in errors:
                    st.error(e)
            else:
                with st.spinner("Mengirim jawaban..."):
                    try:
                        submit_to_public_spreadsheet(
                            form_meta["spreadsheet_url"],
                            visible_qs,
                            pf_answers,
                        )
                        st.session_state["pf_submitted"] = True
                        st.session_state["pf_form_id"] = _form_id
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Gagal mengirim: {e}")

elif _PAGE in ("admin", "admin_sidebar"):
    pass  # handled below

if _PAGE in ("admin", "admin_sidebar"):

    st.title("Panel Admin")

    # ── Login check ───────────────────────────────────────────────────────────
    if "admin_authenticated" not in st.session_state:
        st.session_state.admin_authenticated = False

    if not st.session_state.admin_authenticated:
        st.subheader("Login Admin")
        pwd_input = st.text_input("Password", type="password", key="admin_pwd_input")
        if st.button("Masuk", key="btn_admin_login"):
            if pwd_input == ADMIN_PASSWORD:
                st.session_state.admin_authenticated = True
                st.rerun()
            else:
                st.error("Password salah.")
        st.stop()

    # ── Logout ────────────────────────────────────────────────────────────────
    col_logout, _ = st.columns([1, 4])
    with col_logout:
        if st.button("Keluar", key="btn_logout"):
            st.session_state.admin_authenticated = False
            st.rerun()

    # ── Info link admin ───────────────────────────────────────────────────────
    if _PAGE == "admin":
        _base = st.query_params.get("_base_url", "")
        st.caption(f"🔗 URL Admin: `?mode=admin` — bagikan hanya ke admin")

    # ── Tab Admin ─────────────────────────────────────────────────────────────
    tab_formmanager, tab_monitoring, tab_formbuilder, tab_settings = st.tabs([
        "📋 Form Manager",
        "📊 Monitoring Submission",
        "📝 Form Builder (Input Dokumentasi)",
        "⚙️ Pengaturan",
    ])

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 0: FORM MANAGER
    # ══════════════════════════════════════════════════════════════════════════
    with tab_formmanager:
        st.subheader("Form Manager")
        st.caption("Buat dan kelola form publik seperti Google Form. Setiap form punya link unik.")

        if "fm_forms" not in st.session_state:
            st.session_state.fm_forms = load_forms_list()
        if "fm_active_form" not in st.session_state:
            st.session_state.fm_active_form = None  # form_id yang sedang diedit

        fm_forms = st.session_state.fm_forms

        # ── Tampilkan link form jika ada yang diedit ──────────────────────────
        if st.session_state.fm_active_form:
            fid = st.session_state.fm_active_form
            meta = next((f for f in fm_forms if f["form_id"] == fid), None)
            if not meta:
                st.session_state.fm_active_form = None
                st.rerun()

            # ── Header editor ─────────────────────────────────────────────────
            col_back, col_title = st.columns([1, 4])
            with col_back:
                if st.button("← Kembali", key="fm_back"):
                    st.session_state.fm_active_form = None
                    st.rerun()
            with col_title:
                st.markdown(f"### ✏️ Edit Form: **{meta['form_name']}**")

            # Link form
            try:
                base = st.get_option("browser.serverAddress") or "localhost"
                port = st.get_option("browser.serverPort") or 8501
                form_link = f"http://{base}:{port}/?form={fid}"
            except Exception:
                form_link = f"/?form={fid}"

            st.info(f"🔗 **Link Form Publik:** `?form={fid}`\n\nSalin URL browser ini dan ganti akhirannya dengan `?form={fid}`")

            # Info spreadsheet
            with st.expander("📊 Pengaturan Spreadsheet & Info Form", expanded=True):
                c1, c2 = st.columns(2)
                with c1:
                    new_name = st.text_input("Nama Form", value=meta["form_name"], key="fm_edit_name")
                    if new_name != meta["form_name"]:
                        meta["form_name"] = new_name
                with c2:
                    new_desc = st.text_input("Deskripsi (opsional)", value=meta.get("description",""), key="fm_edit_desc")
                    if new_desc != meta.get("description",""):
                        meta["description"] = new_desc

                new_url = st.text_input(
                    "🔗 URL Spreadsheet Penerima Jawaban *",
                    value=meta.get("spreadsheet_url",""),
                    key="fm_edit_url",
                    placeholder="https://docs.google.com/spreadsheets/d/xxx/edit",
                    help="Paste URL lengkap Google Sheets. Pastikan service account sudah diberi akses editor ke spreadsheet ini.",
                )
                if new_url != meta.get("spreadsheet_url",""):
                    meta["spreadsheet_url"] = new_url

                if meta.get("spreadsheet_url"):
                    sid = parse_spreadsheet_id(meta["spreadsheet_url"])
                    st.caption(f"Sheet ID: `{sid}`")
                else:
                    st.warning("⚠️ Spreadsheet belum dikonfigurasi. Form tidak bisa menerima jawaban.")

                st.divider()
                st.markdown("**📷 Konfigurasi Upload Foto** (opsional — isi jika form ada pertanyaan foto)")
                new_script_url = st.text_input(
                    "URL Google Apps Script untuk upload foto ke Drive",
                    value=meta.get("apps_script_url", ""),
                    key="fm_edit_script_url",
                    placeholder="https://script.google.com/macros/s/xxx/exec",
                    help=(
                        "Diperlukan jika form memiliki pertanyaan bertipe Foto. "
                        "Buat Apps Script baru di Google Drive, deploy sebagai web app, "
                        "lalu paste URL-nya di sini. "
                        "Kosongkan untuk menggunakan URL default (sama dengan Form Input Dokumentasi)."
                    ),
                )
                if new_script_url != meta.get("apps_script_url", ""):
                    meta["apps_script_url"] = new_script_url
                if meta.get("apps_script_url"):
                    st.caption(f"✅ Apps Script URL dikonfigurasi")
                else:
                    st.caption(f"ℹ️ Kosong → pakai URL default: `{APPS_SCRIPT_URL[:60]}...`")

                col_sv, col_toggle = st.columns(2)
                with col_sv:
                    if st.button("💾 Simpan Info Form", key="fm_save_meta", type="primary"):
                        save_forms_list(fm_forms)
                        st.session_state.fm_forms = load_forms_list()
                        st.success("✅ Info form disimpan.")
                with col_toggle:
                    toggle_lbl = "🟢 Aktif" if meta.get("active") else "🔴 Nonaktif"
                    if st.button(toggle_lbl, key="fm_toggle_active"):
                        meta["active"] = not meta.get("active", True)
                        save_forms_list(fm_forms)
                        st.session_state.fm_forms = load_forms_list()
                        st.rerun()

            st.divider()
            st.markdown("### 📝 Pertanyaan")

            # ── Session state pertanyaan form yang diedit ─────────────────────
            fq_key = f"fm_questions_{fid}"
            fq_dirty_key = f"fm_dirty_{fid}"
            if fq_key not in st.session_state:
                with st.spinner("Memuat pertanyaan..."):
                    st.session_state[fq_key] = load_form_questions(fid)
            if fq_dirty_key not in st.session_state:
                st.session_state[fq_dirty_key] = False

            fq_questions = st.session_state[fq_key]

            # ── Toolbar pertanyaan ────────────────────────────────────────────
            tcol1, tcol2, tcol3 = st.columns([3, 1, 1])
            with tcol1:
                fq_new_type = st.selectbox(
                    "Tipe pertanyaan baru:",
                    options=[
                        ("text",      "📝 Jawaban Singkat"),
                        ("textarea",  "📄 Paragraf"),
                        ("number",    "🔢 Angka"),
                        ("date",      "📅 Tanggal"),
                        ("time",      "🕐 Waktu"),
                        ("dropdown",  "▼ Dropdown"),
                        ("radio",     "⭕ Pilihan Ganda"),
                        ("checkbox",  "☑️ Kotak Centang"),
                        ("scale",     "⭐ Skala / Rating"),
                        ("yes_no",    "✅ Ya / Tidak"),
                        ("photo",     "📷 Upload Foto"),
                        ("paragraph", "📌 Teks Judul/Instruksi"),
                    ],
                    format_func=lambda x: x[1],
                    key="fq_new_type",
                )
            with tcol2:
                st.write("")
                if st.button("➕ Tambah", key="fq_add", use_container_width=True):
                    new_q = {
                        "question_id": str(uuid.uuid4())[:8],
                        "order": len(fq_questions) + 1,
                        "label": "Pertanyaan Baru",
                        "type": fq_new_type[0],
                        "required": False,
                        "options": [],
                        "placeholder": "",
                        "description": "",
                        "min_value": "",
                        "max_value": "",
                        "allow_other": False,
                        "visible": True,
                        "col_locked": False,
                        "col_index": len(fq_questions) + 1,
                        "section": "",
                    }
                    fq_questions.append(new_q)
                    st.session_state[fq_dirty_key] = True
                    st.rerun()
            with tcol3:
                st.write("")
                fq_save_btn = st.button(
                    "💾 Simpan",
                    key="fq_save",
                    type="primary",
                    use_container_width=True,
                    disabled=not st.session_state[fq_dirty_key],
                )

            if st.session_state[fq_dirty_key]:
                st.info("⚠️ Ada perubahan belum disimpan.")

            # ── Drag & Drop urutan ────────────────────────────────────────────
            st.markdown("**🔀 Atur Urutan Pertanyaan (Drag & Drop)**")

            if HAS_SORTABLES and fq_questions:
                sort_items_list = [f"{q['question_id']}::{q['label']}" for q in fq_questions]
                sorted_result = _sort_items(sort_items_list, direction="vertical", key=f"sort_{fid}")
                if sorted_result != sort_items_list:
                    id_to_q = {q["question_id"]: q for q in fq_questions}
                    new_order = []
                    for item in sorted_result:
                        qid_part = item.split("::")[0]
                        if qid_part in id_to_q:
                            new_order.append(id_to_q[qid_part])
                    if new_order:
                        # Update col_index untuk yang tidak di-lock
                        for i, q in enumerate(new_order):
                            q["order"] = i + 1
                            if not q.get("col_locked"):
                                q["col_index"] = i + 1
                        st.session_state[fq_key] = new_order
                        st.session_state[fq_dirty_key] = True
                        st.rerun()
            elif not HAS_SORTABLES:
                st.caption("ℹ️ Install `streamlit-sortables` untuk drag & drop. Gunakan tombol ⬆️⬇️ di bawah.")

            st.divider()

            # ── Daftar pertanyaan (editable) ──────────────────────────────────
            FQ_TYPE_LABELS = {
                "text": "Jawaban Singkat", "textarea": "Paragraf", "number": "Angka",
                "date": "Tanggal", "time": "Waktu", "dropdown": "Dropdown",
                "radio": "Pilihan Ganda", "checkbox": "Kotak Centang", "scale": "Skala",
                "yes_no": "Ya / Tidak", "photo": "Upload Foto", "paragraph": "Judul/Instruksi",
            }
            FQ_TYPE_ICONS = {
                "text": "📝", "textarea": "📄", "number": "🔢", "date": "📅", "time": "🕐",
                "dropdown": "▼", "radio": "⭕", "checkbox": "☑️", "scale": "⭐",
                "yes_no": "✅", "photo": "📷", "paragraph": "📌",
            }

            fq_to_delete = None

            for fq_idx, fq in enumerate(fq_questions):
                qid    = fq["question_id"]
                qtype  = fq["type"]
                icon   = FQ_TYPE_ICONS.get(qtype, "📝")
                tlabel = FQ_TYPE_LABELS.get(qtype, qtype)
                req_m  = " *" if fq.get("required") else ""
                vis_m  = "" if fq.get("visible", True) else " 🚫"
                lock_m = " 🔒" if fq.get("col_locked") else ""

                col_letter = chr(64 + fq.get("col_index", fq_idx + 1)) if fq.get("col_index", fq_idx+1) <= 26 else f"col{fq.get('col_index','?')}"

                with st.expander(
                    f"{icon} {fq_idx+1}. {fq['label']}{req_m}{vis_m}{lock_m}   —   *{tlabel}*   `→ Kol. {col_letter}`",
                    expanded=False,
                ):
                    # Kontrol urutan
                    rctrl = st.columns([1, 1, 1, 1, 1, 1])
                    with rctrl[0]:
                        if st.button("⬆️", key=f"fq_up_{qid}", disabled=fq_idx==0):
                            fq_questions[fq_idx], fq_questions[fq_idx-1] = fq_questions[fq_idx-1], fq_questions[fq_idx]
                            # Update order & col_index untuk non-locked
                            for ii, q in enumerate(fq_questions):
                                q["order"] = ii + 1
                                if not q.get("col_locked"):
                                    q["col_index"] = ii + 1
                            st.session_state[fq_dirty_key] = True
                            st.rerun()
                    with rctrl[1]:
                        if st.button("⬇️", key=f"fq_dn_{qid}", disabled=fq_idx==len(fq_questions)-1):
                            fq_questions[fq_idx], fq_questions[fq_idx+1] = fq_questions[fq_idx+1], fq_questions[fq_idx]
                            for ii, q in enumerate(fq_questions):
                                q["order"] = ii + 1
                                if not q.get("col_locked"):
                                    q["col_index"] = ii + 1
                            st.session_state[fq_dirty_key] = True
                            st.rerun()
                    with rctrl[2]:
                        vis_lbl = "👁️ Tampil" if fq.get("visible", True) else "🙈 Sembunyikan"
                        if st.button(vis_lbl, key=f"fq_vis_{qid}"):
                            fq["visible"] = not fq.get("visible", True)
                            st.session_state[fq_dirty_key] = True
                            st.rerun()
                    with rctrl[3]:
                        req_lbl = "✅ Wajib" if fq.get("required") else "❌ Tidak Wajib"
                        if st.button(req_lbl, key=f"fq_req_{qid}"):
                            fq["required"] = not fq.get("required", False)
                            st.session_state[fq_dirty_key] = True
                            st.rerun()
                    with rctrl[4]:
                        lock_lbl = "🔒 Kunci Kolom" if fq.get("col_locked") else "🔓 Bebas"
                        if st.button(lock_lbl, key=f"fq_lock_{qid}"):
                            fq["col_locked"] = not fq.get("col_locked", False)
                            if fq["col_locked"] and not fq.get("col_index"):
                                fq["col_index"] = fq_idx + 1
                            st.session_state[fq_dirty_key] = True
                            st.rerun()
                    with rctrl[5]:
                        if st.button("🗑️ Hapus", key=f"fq_del_{qid}"):
                            fq_to_delete = fq_idx

                    # Kunci kolom — edit manual
                    if fq.get("col_locked"):
                        ci_val = fq.get("col_index", fq_idx+1)
                        new_ci = st.number_input(
                            f"📌 Kolom spreadsheet (1=A, 2=B, ...)",
                            min_value=1, max_value=100,
                            value=int(ci_val),
                            key=f"fq_ci_{qid}",
                            help="Angka ini menentukan kolom di spreadsheet. Dikunci = tidak berubah meski urutan form diubah.",
                        )
                        col_letter_show = chr(64 + new_ci) if new_ci <= 26 else f"col{new_ci}"
                        st.caption(f"→ Akan ditulis ke **Kolom {col_letter_show}** spreadsheet")
                        if new_ci != fq.get("col_index"):
                            fq["col_index"] = new_ci
                            st.session_state[fq_dirty_key] = True

                    st.divider()

                    # Edit label, section, tipe
                    ec1, ec2 = st.columns(2)
                    with ec1:
                        new_label = st.text_input("Label Pertanyaan", value=fq["label"], key=f"fq_label_{qid}")
                        if new_label != fq["label"]:
                            fq["label"] = new_label; st.session_state[fq_dirty_key] = True
                    with ec2:
                        new_sec = st.text_input("Nama Seksi (opsional)", value=fq.get("section",""), key=f"fq_sec_{qid}", placeholder="cth: Data Pribadi")
                        if new_sec != fq.get("section",""):
                            fq["section"] = new_sec; st.session_state[fq_dirty_key] = True

                    new_type_sel = st.selectbox(
                        "Tipe Jawaban",
                        options=list(FQ_TYPE_LABELS.keys()),
                        index=list(FQ_TYPE_LABELS.keys()).index(qtype) if qtype in FQ_TYPE_LABELS else 0,
                        format_func=lambda x: f"{FQ_TYPE_ICONS.get(x,'')} {FQ_TYPE_LABELS.get(x,x)}",
                        key=f"fq_type_{qid}",
                    )
                    if new_type_sel != qtype:
                        fq["type"] = new_type_sel; st.session_state[fq_dirty_key] = True; st.rerun()

                    new_desc = st.text_input("Deskripsi / Instruksi", value=fq.get("description",""), key=f"fq_desc_{qid}")
                    if new_desc != fq.get("description",""):
                        fq["description"] = new_desc; st.session_state[fq_dirty_key] = True

                    if qtype in ("text","textarea","number"):
                        new_ph = st.text_input("Placeholder", value=fq.get("placeholder",""), key=f"fq_ph_{qid}")
                        if new_ph != fq.get("placeholder",""):
                            fq["placeholder"] = new_ph; st.session_state[fq_dirty_key] = True

                    if qtype in ("number","scale"):
                        mc1, mc2 = st.columns(2)
                        with mc1:
                            new_min = st.text_input("Nilai Min", value=str(fq.get("min_value","")), key=f"fq_min_{qid}")
                            if new_min != str(fq.get("min_value","")):
                                fq["min_value"] = new_min; st.session_state[fq_dirty_key] = True
                        with mc2:
                            new_max = st.text_input("Nilai Max", value=str(fq.get("max_value","")), key=f"fq_max_{qid}")
                            if new_max != str(fq.get("max_value","")):
                                fq["max_value"] = new_max; st.session_state[fq_dirty_key] = True

                    if qtype in ("dropdown","radio","checkbox"):
                        st.markdown("**Pilihan Jawaban** (satu per baris)")
                        cur_opts = "\n".join(fq.get("options",[]))
                        new_opts_raw = st.text_area("Opsi", value=cur_opts, key=f"fq_opts_{qid}", height=120, label_visibility="collapsed")
                        new_opts = [o.strip() for o in new_opts_raw.splitlines() if o.strip()]
                        if new_opts != fq.get("options",[]):
                            fq["options"] = new_opts; st.session_state[fq_dirty_key] = True
                        allow_other = st.checkbox("Izinkan jawaban 'Lainnya...'", value=fq.get("allow_other",False), key=f"fq_other_{qid}")
                        if allow_other != fq.get("allow_other",False):
                            fq["allow_other"] = allow_other; st.session_state[fq_dirty_key] = True

            if fq_to_delete is not None:
                fq_questions.pop(fq_to_delete)
                st.session_state[fq_dirty_key] = True
                st.rerun()

            if fq_save_btn:
                with st.spinner("Menyimpan pertanyaan..."):
                    # Assign col_index untuk yang belum locked
                    for ii, q in enumerate(fq_questions):
                        if not q.get("col_locked"):
                            q["col_index"] = ii + 1
                        q["order"] = ii + 1
                    save_form_questions(fid, fq_questions)
                    st.session_state[fq_dirty_key] = False
                st.success("✅ Pertanyaan berhasil disimpan!")
                time.sleep(1)
                st.rerun()

        else:
            # ── List semua form ────────────────────────────────────────────────
            col_new, col_reload_fm = st.columns([3, 1])
            with col_new:
                new_form_name = st.text_input("Nama form baru:", placeholder="cth: Survey Kepuasan Pelanggan", key="fm_new_name")
            with col_reload_fm:
                st.write("")
                if st.button("🔄 Reload", key="fm_reload", use_container_width=True):
                    load_forms_list.clear()
                    st.session_state.fm_forms = load_forms_list()
                    st.rerun()

            if st.button("➕ Buat Form Baru", key="fm_create", type="primary", disabled=not new_form_name.strip()):
                new_fid = str(uuid.uuid4())[:12].replace("-","")
                new_form = {
                    "form_id":        new_fid,
                    "form_name":      new_form_name.strip(),
                    "description":    "",
                    "spreadsheet_url": "",
                    "created_at":     datetime.datetime.now(WIB).strftime("%Y-%m-%d %H:%M:%S"),
                    "active":         True,
                }
                fm_forms.append(new_form)
                save_forms_list(fm_forms)
                st.session_state.fm_forms = load_forms_list()
                st.session_state.fm_active_form = new_fid
                st.rerun()

            st.divider()

            if not fm_forms:
                st.markdown(
                    """<div style="text-align:center;padding:40px;color:#888;border:2px dashed #ddd;border-radius:12px">
                    <div style="font-size:48px">📋</div>
                    <div style="font-size:16px;margin-top:12px">Belum ada form. Buat form pertama Anda!</div>
                    </div>""",
                    unsafe_allow_html=True,
                )
            else:
                fm_to_del = None
                for fi, fm in enumerate(fm_forms):
                    status = "🟢 Aktif" if fm.get("active") else "🔴 Nonaktif"
                    url_ok = "✅" if fm.get("spreadsheet_url") else "⚠️ Belum ada spreadsheet"
                    with st.expander(f"{'🟢' if fm.get('active') else '🔴'} **{fm['form_name']}** — {url_ok}", expanded=False):
                        st.caption(f"ID: `{fm['form_id']}` | Dibuat: {fm.get('created_at','?')} | {status}")
                        if fm.get("description"):
                            st.markdown(fm["description"])
                        st.info(f"🔗 Link form: tambahkan `?form={fm['form_id']}` ke URL aplikasi ini")
                        fc1, fc2, fc3 = st.columns(3)
                        with fc1:
                            if st.button("✏️ Edit Form", key=f"fm_edit_{fi}", type="primary", use_container_width=True):
                                st.session_state.fm_active_form = fm["form_id"]
                                st.rerun()
                        with fc2:
                            tog = "Nonaktifkan" if fm.get("active") else "Aktifkan"
                            if st.button(tog, key=f"fm_tog_{fi}", use_container_width=True):
                                fm["active"] = not fm.get("active", True)
                                save_forms_list(fm_forms)
                                st.session_state.fm_forms = load_forms_list()
                                st.rerun()
                        with fc3:
                            if st.button("🗑️ Hapus", key=f"fm_del_{fi}", use_container_width=True):
                                fm_to_del = fi
                if fm_to_del is not None:
                    fm_forms.pop(fm_to_del)
                    save_forms_list(fm_forms)
                    st.session_state.fm_forms = load_forms_list()
                    st.rerun()

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 1: MONITORING
    # ══════════════════════════════════════════════════════════════════════════
    with tab_monitoring:
        st.subheader("Monitoring Investment")
        today = datetime.datetime.now(WIB).date()

        with st.spinner("Memuat data..."):
            df_sub = load_submissions()

        if df_sub.empty:
            st.warning("Belum ada data submission.")
        else:
            df_sub.columns = [c.strip() for c in df_sub.columns]

            st.subheader("Filter")
            col1, col2, col3 = st.columns(3)

            with col1:
                eksekutor_list = ["Semua"] + sorted(df_sub["Eksekutor"].dropna().unique().tolist()) if "Eksekutor" in df_sub.columns else ["Semua"]
                filter_eks = st.selectbox("Eksekutor", eksekutor_list, key="filter_eks")

            with col2:
                activity_list = ["Semua"] + sorted(df_sub["Activity"].dropna().unique().tolist()) if "Activity" in df_sub.columns else ["Semua"]
                filter_act = st.selectbox("Activity (Termin)", activity_list, key="filter_act")

            with col3:
                program_list = ["Semua"] + sorted(df_sub["Program Investment"].dropna().unique().tolist()) if "Program Investment" in df_sub.columns else ["Semua"]
                filter_prog = st.selectbox("Program", program_list, key="filter_prog")

            df_view = df_sub.copy()
            if filter_eks  != "Semua" and "Eksekutor"         in df_view.columns: df_view = df_view[df_view["Eksekutor"]         == filter_eks]
            if filter_act  != "Semua" and "Activity"          in df_view.columns: df_view = df_view[df_view["Activity"]          == filter_act]
            if filter_prog != "Semua" and "Program Investment" in df_view.columns: df_view = df_view[df_view["Program Investment"] == filter_prog]

            st.divider()

            deadlines   = []
            countdowns  = []
            warnings    = []
            label_types = []

            for _, row in df_view.iterrows():
                activity = str(row.get("Activity", "")).strip()
                date_str = str(row.get("Date", "")).strip()
                deadline = compute_deadline(date_str, activity)

                if deadline is None:
                    deadlines.append("")
                    countdowns.append("")
                    warnings.append(False)
                    label_types.append("")
                else:
                    label_types.append("Jatuh Tempo Termin 2" if activity == "Termin 1" else "Kontrak Berakhir")
                    deadlines.append(str(deadline))
                    cd, warn = countdown_label(deadline, today)
                    countdowns.append(cd)
                    warnings.append(warn)

            df_view = df_view.copy()
            df_view["Tipe Deadline"]    = label_types
            df_view["Tanggal Deadline"] = deadlines
            df_view["Countdown"]        = countdowns
            df_view["_warn"]            = warnings

            total_rows   = len(df_view)
            warn_count   = df_view["_warn"].sum()
            termin1_done = (df_view["Activity"] == "Termin 1").sum() if "Activity" in df_view.columns else 0
            termin2_done = (df_view["Activity"] == "Termin 2").sum() if "Activity" in df_view.columns else 0

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total Submission", total_rows)
            c2.metric("Termin 1", termin1_done)
            c3.metric("Termin 2", termin2_done)
            c4.metric("Perlu Perhatian", int(warn_count))

            st.divider()
            st.subheader("Detail Submission & Deadline")

            display_cols = [c for c in [
                "Date", "Eksekutor", "Program Investment", "Jenis Investment",
                "Brand", "Activity", "Rayon", "Zona", "Nama Outlet", "Kode Customer",
                "Tipe Deadline", "Tanggal Deadline", "Countdown",
            ] if c in df_view.columns]

            st.dataframe(
                df_view[display_cols],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Countdown": st.column_config.TextColumn("Countdown", help="Kuning = <= 21 hari atau sudah lewat"),
                }
            )

            st.caption("Baris dengan latar kuning = deadline <= 21 hari atau sudah lewat.")

            rows_html = ""
            for _, row in df_view[display_cols + ["_warn"]].iterrows():
                bg = ' style="background:#fff3cd"' if row["_warn"] else ""
                cells = "".join(f"<td style='padding:4px 8px;border:1px solid #ddd'>{row[c]}</td>" for c in display_cols)
                rows_html += f"<tr{bg}>{cells}</tr>"

            header_html = "".join(f"<th style='padding:4px 8px;border:1px solid #ddd;background:#f0f2f6;text-align:left'>{c}</th>" for c in display_cols)

            table_html = f"""
            <div style="overflow-x:auto;max-height:600px;overflow-y:auto">
            <table style="border-collapse:collapse;width:100%;font-size:13px">
                <thead><tr>{header_html}</tr></thead>
                <tbody>{rows_html}</tbody>
            </table>
            </div>
            """
            st.components.v1.html(table_html, height=620, scrolling=True)

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 2: FORM BUILDER
    # ══════════════════════════════════════════════════════════════════════════
    with tab_formbuilder:
        st.subheader("Form Builder")
        st.caption("Atur semua pertanyaan — urutan, label, tipe, wajib/tidak — bebas seperti Google Form.")

        # Inisialisasi session state
        if "fb_questions" not in st.session_state:
            with st.spinner("Memuat konfigurasi form..."):
                st.session_state.fb_questions = load_form_config()
        if "fb_dirty" not in st.session_state:
            st.session_state.fb_dirty = False

        questions = st.session_state.fb_questions

        # ── Toolbar ───────────────────────────────────────────────────────────
        col_add, col_reload, col_save = st.columns([2, 1, 1])

        with col_add:
            new_type = st.selectbox(
                "Tipe pertanyaan baru:",
                options=[
                    ("text",     "📝 Jawaban Singkat"),
                    ("textarea", "📄 Paragraf"),
                    ("number",   "🔢 Angka"),
                    ("date",     "📅 Tanggal"),
                    ("time",     "🕐 Waktu"),
                    ("dropdown", "▼ Dropdown (Pilihan)"),
                    ("radio",    "⭕ Pilihan Ganda"),
                    ("checkbox", "☑️ Kotak Centang"),
                    ("scale",    "⭐ Skala / Rating"),
                    ("yes_no",   "✅ Ya / Tidak"),
                    ("photo",    "📷 Upload Foto"),
                ],
                format_func=lambda x: x[1],
                key="fb_new_type",
            )

        with col_reload:
            st.write("")
            if st.button("🔄 Reload", key="fb_reload", use_container_width=True):
                load_form_config.clear()
                st.session_state.fb_questions = load_form_config()
                st.session_state.fb_dirty = False
                st.rerun()

        with col_save:
            st.write("")
            save_btn = st.button(
                "💾 Simpan Semua",
                key="fb_save",
                type="primary",
                use_container_width=True,
                disabled=not st.session_state.fb_dirty,
            )

        # Pilih posisi insert
        n_q = len(questions)
        pos_options = [f"Posisi {i+1} — sebelum '{questions[i]['label']}'" for i in range(n_q)]
        pos_options.append(f"Posisi {n_q+1} — paling bawah")
        insert_pos = st.selectbox(
            "Tambahkan pertanyaan baru di:",
            options=list(range(n_q + 1)),
            index=n_q,
            format_func=lambda i: pos_options[i],
            key="fb_insert_pos",
        )

        if st.button("➕ Tambah Pertanyaan", key="fb_add"):
            new_q = {
                "question_id": str(uuid.uuid4())[:8],
                "field_id": "",
                "order": insert_pos + 1,
                "section": "",
                "label": "Pertanyaan Baru",
                "type": new_type[0],
                "required": False,
                "options": [],
                "placeholder": "",
                "description": "",
                "min_value": "",
                "max_value": "",
                "allow_other": False,
                "visible": True,
            }
            st.session_state.fb_questions.insert(insert_pos, new_q)
            st.session_state.fb_dirty = True
            st.rerun()

        if st.session_state.fb_dirty:
            st.info("⚠️ Ada perubahan yang belum disimpan. Klik **Simpan Semua**.")

        st.divider()

        # ── Daftar Pertanyaan ─────────────────────────────────────────────────
        TYPE_LABELS = {
            "text": "Jawaban Singkat", "textarea": "Paragraf", "number": "Angka",
            "date": "Tanggal", "time": "Waktu", "dropdown": "Dropdown",
            "radio": "Pilihan Ganda", "checkbox": "Kotak Centang", "scale": "Skala / Rating",
            "yes_no": "Ya / Tidak", "photo": "Upload Foto",
            # tipe sistem
            "customer": "Pilih Customer", "gps": "Lokasi GPS",
        }
        TYPE_ICONS = {
            "text": "📝", "textarea": "📄", "number": "🔢", "date": "📅", "time": "🕐",
            "dropdown": "▼", "radio": "⭕", "checkbox": "☑️", "scale": "⭐",
            "yes_no": "✅", "photo": "📷", "customer": "🏪", "gps": "📍",
        }
        # field_id sistem — kunci DIHAPUS, semua bisa diedit bebas
        EDITABLE_OPTIONS_IDS = {"program","jenis","brand","activity"}

        if not questions:
            st.markdown(
                """<div style="text-align:center;padding:40px;color:#888;border:2px dashed #ddd;border-radius:12px;margin:20px 0">
                <div style="font-size:48px;margin-bottom:12px">📋</div>
                <div style="font-size:16px">Belum ada pertanyaan.</div>
                <div style="font-size:13px;margin-top:8px">Klik <b>Tambah Pertanyaan</b> untuk memulai.</div>
                </div>""",
                unsafe_allow_html=True,
            )

        to_delete = None

        for idx, q in enumerate(questions):
            qid    = q["question_id"]
            fid    = q.get("field_id", "")
            qtype  = q["type"]
            icon   = TYPE_ICONS.get(qtype, "📝")
            tlabel = TYPE_LABELS.get(qtype, qtype)
            vis_m  = "" if q.get("visible", True) else " 🚫"
            req_m  = " *" if q.get("required") else ""
            sys_hint = f" [{fid}]" if fid else ""

            with st.expander(
                f"{icon} {idx+1}. {q['label']}{req_m}{vis_m}{sys_hint}   —   *{tlabel}*",
                expanded=False,
            ):
                # ── Kontrol urutan & aksi ─────────────────────────────────────
                row_ctrl = st.columns([1, 1, 1, 1, 1])
                with row_ctrl[0]:
                    if st.button("⬆️", key=f"fb_up_{qid}", help="Naikan", disabled=idx == 0):
                        questions[idx], questions[idx-1] = questions[idx-1], questions[idx]
                        st.session_state.fb_dirty = True
                        st.rerun()
                with row_ctrl[1]:
                    if st.button("⬇️", key=f"fb_dn_{qid}", help="Turunkan", disabled=idx == len(questions)-1):
                        questions[idx], questions[idx+1] = questions[idx+1], questions[idx]
                        st.session_state.fb_dirty = True
                        st.rerun()
                with row_ctrl[2]:
                    vis_lbl = "👁️ Tampil" if q.get("visible", True) else "🙈 Sembunyikan"
                    if st.button(vis_lbl, key=f"fb_vis_{qid}"):
                        q["visible"] = not q.get("visible", True)
                        st.session_state.fb_dirty = True
                        st.rerun()
                with row_ctrl[3]:
                    req_lbl = "✅ Wajib" if q.get("required") else "❌ Tidak Wajib"
                    if st.button(req_lbl, key=f"fb_req_{qid}"):
                        q["required"] = not q.get("required", False)
                        st.session_state.fb_dirty = True
                        st.rerun()
                with row_ctrl[4]:
                    if st.button("🗑️ Hapus", key=f"fb_del_{qid}", type="secondary"):
                        to_delete = idx

                if fid:
                    st.caption(f"ℹ️ Field internal: `{fid}` — logika submit form utama terkait field ini.")

                st.divider()

                # ── Edit label & section ──────────────────────────────────────
                c1, c2 = st.columns(2)
                with c1:
                    new_label = st.text_input("Label Pertanyaan", value=q["label"], key=f"fb_label_{qid}")
                    if new_label != q["label"]:
                        q["label"] = new_label
                        st.session_state.fb_dirty = True
                with c2:
                    new_section = st.text_input("Nama Seksi (opsional)", value=q.get("section",""), key=f"fb_sec_{qid}", placeholder="cth: Informasi Tambahan")
                    if new_section != q.get("section",""):
                        q["section"] = new_section
                        st.session_state.fb_dirty = True

                # Tipe jawaban — semua bisa diubah
                if not fid:
                    new_type_sel = st.selectbox(
                        "Tipe Jawaban",
                        options=list(TYPE_LABELS.keys()),
                        index=list(TYPE_LABELS.keys()).index(qtype) if qtype in TYPE_LABELS else 0,
                        format_func=lambda x: f"{TYPE_ICONS.get(x,'')} {TYPE_LABELS.get(x,x)}",
                        key=f"fb_type_{qid}",
                    )
                    if new_type_sel != qtype:
                        q["type"] = new_type_sel
                        st.session_state.fb_dirty = True
                        st.rerun()

                new_desc = st.text_input("Deskripsi / Instruksi (opsional)", value=q.get("description",""), key=f"fb_desc_{qid}", placeholder="Teks kecil di bawah label")
                if new_desc != q.get("description",""):
                    q["description"] = new_desc
                    st.session_state.fb_dirty = True

                # Placeholder
                if qtype in ("text","textarea","number"):
                    new_ph = st.text_input("Placeholder", value=q.get("placeholder",""), key=f"fb_ph_{qid}")
                    if new_ph != q.get("placeholder",""):
                        q["placeholder"] = new_ph
                        st.session_state.fb_dirty = True

                # Min/Max
                if qtype in ("number","scale"):
                    mc1, mc2 = st.columns(2)
                    with mc1:
                        new_min = st.text_input("Nilai Minimum", value=str(q.get("min_value","")), key=f"fb_min_{qid}")
                        if new_min != str(q.get("min_value","")):
                            q["min_value"] = new_min; st.session_state.fb_dirty = True
                    with mc2:
                        new_max = st.text_input("Nilai Maksimum", value=str(q.get("max_value","")), key=f"fb_max_{qid}")
                        if new_max != str(q.get("max_value","")):
                            q["max_value"] = new_max; st.session_state.fb_dirty = True

                # Opsi dropdown/radio/checkbox
                if qtype in ("dropdown","radio","checkbox"):
                    st.markdown("**Pilihan Jawaban** (satu per baris)")
                    current_opts = "\n".join(q.get("options",[]))
                    new_opts_raw = st.text_area("Opsi", value=current_opts, key=f"fb_opts_{qid}", height=120, label_visibility="collapsed")
                    new_opts = [o.strip() for o in new_opts_raw.splitlines() if o.strip()]
                    if new_opts != q.get("options",[]):
                        q["options"] = new_opts; st.session_state.fb_dirty = True
                    allow_other = st.checkbox("Izinkan jawaban 'Lainnya...'", value=q.get("allow_other",False), key=f"fb_other_{qid}")
                    if allow_other != q.get("allow_other",False):
                        q["allow_other"] = allow_other; st.session_state.fb_dirty = True

        # Hapus setelah iterasi
        if to_delete is not None:
            st.session_state.fb_questions.pop(to_delete)
            st.session_state.fb_dirty = True
            st.rerun()

        # Simpan
        if save_btn:
            with st.spinner("Menyimpan konfigurasi form ke Google Sheets..."):
                save_form_config(st.session_state.fb_questions)
                st.session_state.fb_dirty = False
            st.success("✅ Konfigurasi form berhasil disimpan!")
            time.sleep(1)
            st.rerun()

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 3: PENGATURAN
    # ══════════════════════════════════════════════════════════════════════════
    with tab_settings:
        st.subheader("Pengaturan Sistem")

        st.markdown("### 📋 Kelola Daftar Config")
        st.caption("Tambah atau hapus data di tab Config Google Sheets.")

        setting_tab = st.selectbox(
            "Pilih konfigurasi yang ingin dikelola:",
            options=[
                ("Config_NamaPromotor", "Daftar Nama Promotor / Eksekutor"),
                ("Config_Rayon", "Daftar Rayon"),
                ("Config_Zona", "Daftar Zona"),
            ],
            format_func=lambda x: x[1],
            key="setting_tab_sel",
        )

        tab_key  = setting_tab[0]
        tab_name = setting_tab[1]

        st.markdown(f"#### {tab_name}")

        current_items = load_config_list(tab_key)

        col_list, col_edit = st.columns([1, 1])

        with col_list:
            st.markdown("**Data saat ini:**")
            if current_items:
                for item in current_items:
                    st.markdown(f"- {item}")
            else:
                st.caption("Belum ada data.")

        with col_edit:
            st.markdown("**Edit daftar** (satu per baris):")
            current_text = "\n".join(current_items)
            new_text = st.text_area(
                "Edit",
                value=current_text,
                key=f"setting_edit_{tab_key}",
                height=200,
                label_visibility="collapsed",
            )
            if st.button("💾 Simpan Perubahan", key=f"setting_save_{tab_key}", type="primary"):
                new_items = [i.strip() for i in new_text.splitlines() if i.strip()]
                try:
                    with st.spinner("Menyimpan..."):
                        gc = get_gspread()
                        wb = gc.open_by_key(SUBMISSION_SHEET_ID)
                        try:
                            ws = wb.worksheet(tab_key)
                        except Exception:
                            ws = wb.add_worksheet(title=tab_key, rows=300, cols=5)
                        ws.clear()
                        header = tab_key.replace("Config_", "")
                        rows_to_write = [[header]] + [[item] for item in new_items]
                        ws.update(rows_to_write, "A1")
                        load_config_list.clear()
                        load_all_config.clear()
                    st.success(f"✅ {tab_name} berhasil disimpan ({len(new_items)} item)")
                    time.sleep(1)
                    st.rerun()
                except Exception as e:
                    st.error(f"Gagal menyimpan: {e}")

        st.divider()

        st.markdown("### 🔧 Pengaturan Umum")

        with st.expander("🗓️ Pengaturan Deadline Termin"):
            st.info(
                f"Saat ini: Jarak antar termin = **{TERMIN_MONTHS} bulan**, "
                f"Peringatan kuning = **≤ {REMINDER_DAYS} hari**."
            )
            st.caption(
                "Untuk mengubah nilai ini, edit konstanta `TERMIN_MONTHS` dan `REMINDER_DAYS` "
                "di bagian atas file App.py, lalu restart aplikasi."
            )

        with st.expander("🗑️ Clear Cache"):
            st.caption("Paksa reload semua data dari Google Sheets.")
            if st.button("Clear semua cache", key="btn_clear_cache"):
                st.cache_data.clear()
                st.cache_resource.clear()
                st.success("Cache dibersihkan. Data akan di-reload saat halaman berikutnya dimuat.")

        with st.expander("📊 Info Spreadsheet"):
            st.markdown(f"- **Customer Sheet ID:** `{CUSTOMER_SHEET_ID}`")
            st.markdown(f"- **Submission Sheet ID:** `{SUBMISSION_SHEET_ID}`")
            st.markdown(f"- **Form Builder Tab:** `{FORM_CONFIG_TAB}`")

        st.divider()

        # ── Spreadsheet Manager ───────────────────────────────────────────────
        st.markdown("### 🗂️ Spreadsheet Manager")
        st.caption(
            "Daftarkan beberapa Spreadsheet dan pilih mana yang aktif. "
            "Submission & data customer akan menggunakan spreadsheet yang ditandai **Aktif**."
        )

        if "sm_sheets" not in st.session_state:
            st.session_state.sm_sheets = load_spreadsheet_list()
        if "sm_dirty" not in st.session_state:
            st.session_state.sm_dirty = False

        sm_sheets = st.session_state.sm_sheets

        # Tampilkan default (dari konstanta) jika list kosong
        active_sub, active_cus = get_active_sheet_ids()
        st.info(
            f"**Spreadsheet aktif saat ini:**\n\n"
            f"- Submission: `{active_sub}`\n"
            f"- Customer: `{active_cus}`"
        )

        col_sm_add, col_sm_save = st.columns([2, 1])
        with col_sm_add:
            if st.button("➕ Tambah Spreadsheet", key="sm_add"):
                sm_sheets.append({
                    "name": "Spreadsheet Baru",
                    "submission_id": "",
                    "customer_id": "",
                    "active": False,
                    "description": "",
                })
                st.session_state.sm_dirty = True
                st.rerun()
        with col_sm_save:
            sm_save_btn = st.button(
                "💾 Simpan", key="sm_save", type="primary",
                disabled=not st.session_state.sm_dirty,
            )

        if st.session_state.sm_dirty:
            st.warning("Ada perubahan belum disimpan.")

        to_del_sm = None
        for si, ss in enumerate(sm_sheets):
            active_badge = " 🟢 **AKTIF**" if ss.get("active") else ""
            with st.expander(f"📋 {ss.get('name','Tanpa Nama')}{active_badge}", expanded=False):
                new_name = st.text_input("Nama", value=ss.get("name",""), key=f"sm_name_{si}")
                if new_name != ss.get("name",""):
                    ss["name"] = new_name; st.session_state.sm_dirty = True

                new_desc = st.text_input("Keterangan (opsional)", value=ss.get("description",""), key=f"sm_desc_{si}")
                if new_desc != ss.get("description",""):
                    ss["description"] = new_desc; st.session_state.sm_dirty = True

                new_sub = st.text_input(
                    "Submission Spreadsheet ID",
                    value=ss.get("submission_id",""),
                    key=f"sm_sub_{si}",
                    placeholder="cth: 1RC7v1fGmcz-9q4Vow...",
                    help="Salin dari URL Google Sheets: docs.google.com/spreadsheets/d/**ID**/edit",
                )
                if new_sub != ss.get("submission_id",""):
                    ss["submission_id"] = new_sub; st.session_state.sm_dirty = True

                new_cus = st.text_input(
                    "Customer Spreadsheet ID",
                    value=ss.get("customer_id",""),
                    key=f"sm_cus_{si}",
                    placeholder="cth: 113E5fKvZ0wWlo... (kosongkan = gunakan default)",
                )
                if new_cus != ss.get("customer_id",""):
                    ss["customer_id"] = new_cus; st.session_state.sm_dirty = True

                row_sm = st.columns([2, 1])
                with row_sm[0]:
                    if not ss.get("active"):
                        if st.button("🟢 Jadikan Aktif", key=f"sm_activate_{si}", type="primary"):
                            for other in sm_sheets:
                                other["active"] = False
                            ss["active"] = True
                            st.session_state.sm_dirty = True
                            st.rerun()
                    else:
                        if st.button("⬜ Nonaktifkan", key=f"sm_deactivate_{si}"):
                            ss["active"] = False
                            st.session_state.sm_dirty = True
                            st.rerun()
                with row_sm[1]:
                    if st.button("🗑️ Hapus", key=f"sm_del_{si}"):
                        to_del_sm = si

        if to_del_sm is not None:
            sm_sheets.pop(to_del_sm)
            st.session_state.sm_dirty = True
            st.rerun()

        if sm_save_btn:
            with st.spinner("Menyimpan daftar spreadsheet..."):
                save_spreadsheet_list(sm_sheets)
                st.cache_data.clear()
                st.session_state.sm_dirty = False
            st.success("✅ Daftar spreadsheet disimpan. Cache di-refresh.")
            time.sleep(1)
            st.rerun()
