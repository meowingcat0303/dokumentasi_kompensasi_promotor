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
FORM_CONFIG_TAB = "Config_FormBuilder"

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
            "min_value", "max_value", "allow_other", "visible",
        ]
        ws.update([headers], "A1")
    return ws

@st.cache_data(ttl=30)
def load_form_config() -> list:
    """Load konfigurasi pertanyaan dari Google Sheets."""
    try:
        gc = get_gspread()
        wb = gc.open_by_key(SUBMISSION_SHEET_ID)
        try:
            ws = wb.worksheet(FORM_CONFIG_TAB)
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
            row_dict = dict(zip(headers, r + [""] * (len(headers) - len(r))))
            q = {
                "question_id": row_dict.get("question_id", str(uuid.uuid4())[:8]),
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
        return []

def save_form_config(questions: list):
    """Simpan semua pertanyaan ke Google Sheets."""
    gc = get_gspread()
    ws = ensure_form_config_tab()
    headers = [
        "question_id", "order", "section", "label", "type",
        "required", "options", "placeholder", "description",
        "min_value", "max_value", "allow_other", "visible",
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
        ]
        rows.append(row)
    ws.clear()
    ws.update(rows, "A1")
    load_form_config.clear()

# ── Image utils ───────────────────────────────────────────────────────────────
def compress_image(uploaded_file) -> bytes:
    img = Image.open(uploaded_file)
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
    ratio = IMAGE_MAX_PX / max(img.size)
    if ratio < 1:
        new_size = (int(img.width * ratio), int(img.height * ratio))
        img = img.resize(new_size, Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=IMAGE_QUALITY, optimize=True)
    return buf.getvalue()

# ── Drive upload ──────────────────────────────────────────────────────────────
APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbz27id708tLEcf0eGWNC6BrA7TdHiFVfgsPL2b_xGDkWTqBD30tlGWpCXvQ8F2IXIjO/exec"

def upload_to_drive(data: bytes, filename: str, retries: int = 3) -> str:
    import base64
    b64 = base64.b64encode(data).decode("utf-8")
    last_exc = None
    for attempt in range(retries):
        try:
            resp = _requests.post(
                APPS_SCRIPT_URL,
                data={"image": b64, "filename": filename},
                timeout=90,
            )
            if not resp.ok:
                st.warning(f"Upload attempt {attempt+1} gagal: HTTP {resp.status_code} — {resp.text[:300]}")
                resp.raise_for_status()
            result = resp.json()
            if result.get("success"):
                return result["url"]
            raise ValueError(f"Apps Script error: {result}")
        except (_requests.exceptions.HTTPError, _requests.exceptions.Timeout) as e:
            last_exc = e
            if attempt < retries - 1:
                time.sleep(3 * (attempt + 1))
                continue
            raise last_exc

# ── Submission writer ─────────────────────────────────────────────────────────
def append_submission(row: list):
    gc = get_gspread()
    ws = gc.open_by_key(SUBMISSION_SHEET_ID).worksheet("Submission")
    ws.append_row(row, value_input_option="USER_ENTERED")

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

page = st.sidebar.radio(
    "Menu",
    ["Input Dokumentasi", "Monitoring (Admin)"],
    key="nav_page",
)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE 1 — INPUT DOKUMENTASI
# ══════════════════════════════════════════════════════════════════════════════
if page == "Input Dokumentasi":

    st.title("Dokumentasi Investment Lapangan")

    if "submit_info" in st.session_state:
        info = st.session_state.submit_info
        st.success("Dokumentasi berhasil dikirim dan tersimpan.")
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

    code_col           = meta["code_col"]
    name_col           = meta["name_col"]
    addr_col           = meta["addr_col"]
    sales_group_col    = meta["sales_group_col"]
    sales_district_col = meta["sales_district_col"]

    st.caption(f"Data customer: **{meta['n_aktif']:,} toko aktif** dari {meta['n_total']:,} total")
    st.divider()

    # Bagian 1: Identitas & Program
    st.subheader("Identitas & Detail Program")

    col1, col2 = st.columns(2)
    with col1:
        tgl = st.date_input("Tanggal", value=datetime.datetime.now(WIB).date(), key="tgl_sel")
    with col2:
        promotor = st.selectbox("Eksekutor (Promotor)", ["— Pilih —"] + promotors, key="promotor_sel")

    col3, col4 = st.columns(2)
    with col3:
        program = st.selectbox("Program Investment", ["— Pilih —"] + PROGRAM_OPTIONS, key="program_sel")
    with col4:
        jenis = st.selectbox("Jenis Investment", ["— Pilih —"] + JENIS_OPTIONS, key="jenis_sel")

    col5, col6 = st.columns(2)
    with col5:
        brand = st.selectbox("Brand", ["— Pilih —"] + BRAND_OPTIONS, key="brand_sel")
    with col6:
        activity = st.selectbox("Activity", ["— Pilih —"] + ACTIVITY_OPTIONS, key="activity_sel")

    st.divider()

    # Bagian 2: Customer, Rayon & Zona
    st.subheader("Customer, Rayon & Zona")

    st.radio(
        "Mode Rayon & Zona:",
        ["Otomatis (dari data customer)", "Manual (pilih sendiri)"],
        key="rayon_zona_mode",
        horizontal=True,
    )
    mode_otomatis = st.session_state.rayon_zona_mode == "Otomatis (dari data customer)"

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
        "Cari & Pilih Customer",
        options=["— Pilih —"] + df_cust["_label"].tolist(),
        key="cust_sel",
    )

    selected_customer = None
    rayon_final = ""
    zona_final  = ""

    if chosen_label != "— Pilih —":
        match = df_cust[df_cust["_label"] == chosen_label]
        if not match.empty:
            selected_customer = match.iloc[0].to_dict()

            if mode_otomatis:
                rayon_final = str(selected_customer[sales_group_col]).strip() if sales_group_col else ""
                zona_final  = str(selected_customer[sales_district_col]).strip() if sales_district_col else ""
                col_r, col_z = st.columns(2)
                with col_r:
                    st.info(f"**Rayon:** {rayon_final or '—'}")
                with col_z:
                    st.info(f"**Zona:** {zona_final or '—'}")
            else:
                rayon_final = rayon_sel if rayon_sel != "— Pilih —" else ""
                zona_final  = zona_sel  if zona_sel  != "— Pilih —" else ""

            st.success(
                f"**{selected_customer[name_col]}** ({selected_customer[code_col]})"
                + (f" — {selected_customer[addr_col]}" if addr_col else "")
            )

    st.divider()

    # Bagian 3: GPS
    st.subheader("Lokasi GPS")
    st.components.v1.html(GPS_JS, height=90)

    if "gps_coords" not in st.session_state:
        st.session_state.gps_coords = ""

    gps_input = st.text_input(
        "koordinat_gps",
        value=st.session_state.gps_coords,
        key="gps_manual",
        placeholder="Klik tombol di atas untuk mengisi otomatis",
        label_visibility="collapsed",
    )
    if gps_input:
        st.session_state.gps_coords = gps_input

    st.divider()

    # Bagian 4: Foto
    st.subheader("Upload Foto")
    foto_ktp      = st.file_uploader("Foto KTP Pemilik *",       type=["jpg","jpeg","png"], key="foto_ktp")
    foto_sunblind = st.file_uploader("Foto Sunblind *",          type=["jpg","jpeg","png"], key="foto_sunblind")
    foto_display  = st.file_uploader("Foto Display (opsional)",  type=["jpg","jpeg","png"], key="foto_display")
    foto_kompens  = st.file_uploader("Foto Bukti Kompensasi *",  type=["jpg","jpeg","png"], key="foto_kompens")
    st.caption("Foto dikompres otomatis (max 1920px, JPEG 75%) sebelum diupload.")

    # ── Pertanyaan Tambahan dari Form Builder ──────────────────────────────────
    extra_questions = load_form_config()
    extra_answers   = {}

    visible_questions = [q for q in extra_questions if q.get("visible", True)]

    if visible_questions:
        # Kelompokkan per section
        sections = {}
        for q in visible_questions:
            sec = q.get("section", "") or "Pertanyaan Tambahan"
            sections.setdefault(sec, []).append(q)

        for sec_name, qs in sections.items():
            st.divider()
            st.subheader(sec_name)
            for q in qs:
                qid   = q["question_id"]
                label = q["label"] + (" *" if q["required"] else " (opsional)")
                desc  = q.get("description", "")
                if desc:
                    st.caption(desc)

                qtype = q["type"]

                if qtype == "text":
                    extra_answers[qid] = st.text_input(
                        label, placeholder=q.get("placeholder", ""), key=f"eq_{qid}"
                    )

                elif qtype == "textarea":
                    extra_answers[qid] = st.text_area(
                        label, placeholder=q.get("placeholder", ""), key=f"eq_{qid}"
                    )

                elif qtype == "number":
                    min_v = float(q["min_value"]) if q.get("min_value") else None
                    max_v = float(q["max_value"]) if q.get("max_value") else None
                    extra_answers[qid] = st.number_input(
                        label, min_value=min_v, max_value=max_v, key=f"eq_{qid}"
                    )

                elif qtype == "date":
                    extra_answers[qid] = str(st.date_input(label, key=f"eq_{qid}"))

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
                    selected = []
                    st.markdown(f"**{label}**")
                    for opt in opts:
                        if st.checkbox(opt, key=f"eq_{qid}_{opt}"):
                            selected.append(opt)
                    if q.get("allow_other"):
                        other_val = st.text_input("Lainnya:", key=f"eq_{qid}_other")
                        if other_val:
                            selected.append(other_val)
                    extra_answers[qid] = ", ".join(selected)

                elif qtype == "scale":
                    min_v = int(q["min_value"]) if q.get("min_value") else 1
                    max_v = int(q["max_value"]) if q.get("max_value") else 5
                    extra_answers[qid] = str(st.slider(label, min_value=min_v, max_value=max_v, key=f"eq_{qid}"))

                elif qtype == "photo":
                    uploaded = st.file_uploader(label, type=["jpg","jpeg","png"], key=f"eq_{qid}")
                    extra_answers[qid] = uploaded  # handled at submit

                elif qtype == "time":
                    extra_answers[qid] = str(st.time_input(label, key=f"eq_{qid}"))

                elif qtype == "yes_no":
                    extra_answers[qid] = "Ya" if st.toggle(label, key=f"eq_{qid}") else "Tidak"

    catatan = st.text_area("Catatan Tambahan (opsional)", key="catatan_input")
    st.divider()

    # Submit
    if st.button("Submit Dokumentasi", key="btn_submit", type="primary"):
        errors = []
        if promotor  == "— Pilih —": errors.append("Eksekutor wajib dipilih")
        if program   == "— Pilih —": errors.append("Program Investment wajib dipilih")
        if jenis     == "— Pilih —": errors.append("Jenis Investment wajib dipilih")
        if brand     == "— Pilih —": errors.append("Brand wajib dipilih")
        if activity  == "— Pilih —": errors.append("Activity wajib dipilih")
        if not selected_customer:    errors.append("Customer belum dipilih")
        if not rayon_final:          errors.append("Rayon tidak ditemukan / belum dipilih")
        if not zona_final:           errors.append("Zona tidak ditemukan / belum dipilih")
        if not foto_ktp:             errors.append("Foto KTP wajib diupload")
        if not foto_sunblind:        errors.append("Foto Sunblind wajib diupload")
        if not foto_kompens:         errors.append("Foto Bukti Kompensasi wajib diupload")

        # Validasi pertanyaan tambahan yang wajib
        for q in visible_questions:
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
                ts  = datetime.datetime.now(WIB).strftime("%Y%m%d_%H%M%S")
                uid = str(uuid.uuid4())[:8]
                kode = selected_customer.get(code_col, "unknown")

                def safe_upload(file, label):
                    if file is None:
                        return ""
                    fname = f"{ts}_{kode}_{promotor}_{label}_{uid}.jpg"
                    time.sleep(1)
                    return upload_to_drive(compress_image(file), fname)

                url_ktp      = safe_upload(foto_ktp,      "ktp")
                url_sunblind = safe_upload(foto_sunblind, "sunblind")
                url_display  = safe_upload(foto_display,  "display")
                url_kompens  = safe_upload(foto_kompens,  "kompensasi")

                # Upload foto dari pertanyaan tambahan
                extra_answers_str = {}
                for q in visible_questions:
                    qid = q["question_id"]
                    val = extra_answers.get(qid, "")
                    if q["type"] == "photo" and val is not None:
                        url_extra = safe_upload(val, f"extra_{qid}")
                        extra_answers_str[qid] = url_extra
                    else:
                        extra_answers_str[qid] = str(val) if val else ""

                unique_id   = f"{ts}_{uid}"
                submit_time = datetime.datetime.now(WIB).strftime("%Y-%m-%d %H:%M:%S")

                # Bangun row: kolom tetap + kolom dinamis dari form builder
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
                    selected_customer.get(name_col, ""),
                    kode,
                    url_sunblind,
                    url_display,
                    url_kompens,
                    url_ktp,
                    st.session_state.gps_coords,
                    catatan,
                    submit_time,
                ]
                # Tambahkan jawaban pertanyaan ekstra
                for q in visible_questions:
                    row.append(extra_answers_str.get(q["question_id"], ""))

                append_submission(row)
                st.session_state.submit_info = {
                    "unique_id": unique_id,
                    "outlet": selected_customer.get(name_col, "-"),
                    "kode": kode,
                    "promotor": promotor,
                    "waktu": submit_time,
                }
                st.session_state.gps_coords = ""
                st.cache_data.clear()
                st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 2 — MONITORING & ADMIN PANEL
# ══════════════════════════════════════════════════════════════════════════════
elif page == "Monitoring (Admin)":

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

    # ── Tab Admin ─────────────────────────────────────────────────────────────
    tab_monitoring, tab_formbuilder, tab_settings = st.tabs([
        "📊 Monitoring Submission",
        "📝 Form Builder",
        "⚙️ Pengaturan",
    ])

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
        st.subheader("Form Builder — Pertanyaan Tambahan")
        st.caption(
            "Tambahkan, edit, atau hapus pertanyaan yang akan muncul di halaman Input Dokumentasi. "
            "Mirip seperti Google Form — atur tipe jawaban, wajib/tidak, hingga pilihan."
        )

        # Inisialisasi session state form builder
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
                    ("text",     "Jawaban Singkat"),
                    ("textarea", "Paragraf"),
                    ("number",   "Angka"),
                    ("date",     "Tanggal"),
                    ("time",     "Waktu"),
                    ("dropdown", "Dropdown (Pilihan)"),
                    ("radio",    "Pilihan Ganda"),
                    ("checkbox", "Kotak Centang"),
                    ("scale",    "Skala / Rating"),
                    ("yes_no",   "Ya / Tidak"),
                    ("photo",    "Upload Foto"),
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

        if st.button("➕ Tambah Pertanyaan", key="fb_add", use_container_width=False):
            new_q = {
                "question_id": str(uuid.uuid4())[:8],
                "order": len(questions) + 1,
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
            st.session_state.fb_questions.append(new_q)
            st.session_state.fb_dirty = True
            st.rerun()

        if st.session_state.fb_dirty:
            st.info("⚠️ Ada perubahan yang belum disimpan. Klik **Simpan Semua** untuk menyimpan ke Google Sheets.")

        st.divider()

        # ── Daftar Pertanyaan ─────────────────────────────────────────────────
        TYPE_LABELS = {
            "text":     "Jawaban Singkat",
            "textarea": "Paragraf",
            "number":   "Angka",
            "date":     "Tanggal",
            "time":     "Waktu",
            "dropdown": "Dropdown",
            "radio":    "Pilihan Ganda",
            "checkbox": "Kotak Centang",
            "scale":    "Skala / Rating",
            "yes_no":   "Ya / Tidak",
            "photo":    "Upload Foto",
        }

        TYPE_ICONS = {
            "text":     "📝",
            "textarea": "📄",
            "number":   "🔢",
            "date":     "📅",
            "time":     "🕐",
            "dropdown": "▼",
            "radio":    "⭕",
            "checkbox": "☑️",
            "scale":    "⭐",
            "yes_no":   "✅",
            "photo":    "📷",
        }

        if not questions:
            st.markdown(
                """
                <div style="text-align:center;padding:40px;color:#888;border:2px dashed #ddd;border-radius:12px;margin:20px 0">
                    <div style="font-size:48px;margin-bottom:12px">📋</div>
                    <div style="font-size:16px">Belum ada pertanyaan tambahan.</div>
                    <div style="font-size:13px;margin-top:8px">Klik <b>Tambah Pertanyaan</b> untuk memulai.</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        to_delete = None

        for idx, q in enumerate(questions):
            qid   = q["question_id"]
            qtype = q["type"]
            icon  = TYPE_ICONS.get(qtype, "📝")
            tlabel = TYPE_LABELS.get(qtype, qtype)
            visible_marker = "" if q.get("visible", True) else " 🚫"
            required_marker = " *" if q.get("required") else ""

            with st.expander(
                f"{icon} {idx+1}. {q['label']}{required_marker}{visible_marker}   —   *{tlabel}*",
                expanded=False,
            ):
                # ── Baris kontrol atas ───────────────────────────────────────
                row_ctrl = st.columns([1, 1, 1, 1, 1])
                with row_ctrl[0]:
                    if st.button("⬆️", key=f"fb_up_{qid}", help="Naikan urutan", disabled=idx == 0):
                        questions[idx], questions[idx-1] = questions[idx-1], questions[idx]
                        st.session_state.fb_dirty = True
                        st.rerun()
                with row_ctrl[1]:
                    if st.button("⬇️", key=f"fb_dn_{qid}", help="Turunkan urutan", disabled=idx == len(questions)-1):
                        questions[idx], questions[idx+1] = questions[idx+1], questions[idx]
                        st.session_state.fb_dirty = True
                        st.rerun()
                with row_ctrl[2]:
                    vis_label = "👁️ Tampil" if q.get("visible", True) else "🙈 Disembunyikan"
                    if st.button(vis_label, key=f"fb_vis_{qid}"):
                        q["visible"] = not q.get("visible", True)
                        st.session_state.fb_dirty = True
                        st.rerun()
                with row_ctrl[3]:
                    req_label = "✅ Wajib" if q.get("required") else "❌ Tidak Wajib"
                    if st.button(req_label, key=f"fb_req_{qid}"):
                        q["required"] = not q.get("required", False)
                        st.session_state.fb_dirty = True
                        st.rerun()
                with row_ctrl[4]:
                    if st.button("🗑️ Hapus", key=f"fb_del_{qid}", type="secondary"):
                        to_delete = idx

                st.divider()

                # ── Field konfigurasi ────────────────────────────────────────
                c1, c2 = st.columns(2)
                with c1:
                    new_label = st.text_input("Label Pertanyaan", value=q["label"], key=f"fb_label_{qid}")
                    if new_label != q["label"]:
                        q["label"] = new_label
                        st.session_state.fb_dirty = True

                with c2:
                    new_section = st.text_input("Nama Seksi (grup)", value=q.get("section", ""), key=f"fb_sec_{qid}",
                                                 placeholder="cth: Data Tambahan")
                    if new_section != q.get("section", ""):
                        q["section"] = new_section
                        st.session_state.fb_dirty = True

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

                new_desc = st.text_input(
                    "Deskripsi / Instruksi (opsional)",
                    value=q.get("description", ""),
                    key=f"fb_desc_{qid}",
                    placeholder="Teks kecil di bawah label pertanyaan",
                )
                if new_desc != q.get("description", ""):
                    q["description"] = new_desc
                    st.session_state.fb_dirty = True

                # Placeholder (untuk text / textarea / number)
                if qtype in ("text", "textarea", "number"):
                    new_ph = st.text_input(
                        "Placeholder",
                        value=q.get("placeholder", ""),
                        key=f"fb_ph_{qid}",
                    )
                    if new_ph != q.get("placeholder", ""):
                        q["placeholder"] = new_ph
                        st.session_state.fb_dirty = True

                # Min/Max (number, scale)
                if qtype in ("number", "scale"):
                    mc1, mc2 = st.columns(2)
                    with mc1:
                        new_min = st.text_input("Nilai Minimum", value=str(q.get("min_value", "")), key=f"fb_min_{qid}")
                        if new_min != str(q.get("min_value", "")):
                            q["min_value"] = new_min
                            st.session_state.fb_dirty = True
                    with mc2:
                        new_max = st.text_input("Nilai Maksimum", value=str(q.get("max_value", "")), key=f"fb_max_{qid}")
                        if new_max != str(q.get("max_value", "")):
                            q["max_value"] = new_max
                            st.session_state.fb_dirty = True

                # Options (dropdown, radio, checkbox)
                if qtype in ("dropdown", "radio", "checkbox"):
                    st.markdown("**Pilihan Jawaban** (satu per baris)")
                    current_opts = "\n".join(q.get("options", []))
                    new_opts_raw = st.text_area(
                        "Pilihan (satu per baris)",
                        value=current_opts,
                        key=f"fb_opts_{qid}",
                        height=120,
                        label_visibility="collapsed",
                    )
                    new_opts = [o.strip() for o in new_opts_raw.splitlines() if o.strip()]
                    if new_opts != q.get("options", []):
                        q["options"] = new_opts
                        st.session_state.fb_dirty = True

                    allow_other = st.checkbox(
                        "Izinkan jawaban 'Lainnya...' (input bebas)",
                        value=q.get("allow_other", False),
                        key=f"fb_other_{qid}",
                    )
                    if allow_other != q.get("allow_other", False):
                        q["allow_other"] = allow_other
                        st.session_state.fb_dirty = True

                # Preview
                st.markdown("---")
                st.caption("👀 **Preview pertanyaan:**")
                preview_label = q["label"] + (" **(wajib)**" if q.get("required") else " *(opsional)*")
                if q.get("description"):
                    st.caption(q["description"])

                with st.container():
                    if qtype == "text":
                        st.text_input(preview_label, disabled=True, placeholder=q.get("placeholder", ""), key=f"prev_{qid}")
                    elif qtype == "textarea":
                        st.text_area(preview_label, disabled=True, placeholder=q.get("placeholder", ""), key=f"prev_{qid}")
                    elif qtype == "number":
                        st.number_input(preview_label, disabled=True, key=f"prev_{qid}")
                    elif qtype == "date":
                        st.date_input(preview_label, disabled=True, key=f"prev_{qid}")
                    elif qtype == "time":
                        st.time_input(preview_label, disabled=True, key=f"prev_{qid}")
                    elif qtype == "dropdown":
                        opts_prev = ["— Pilih —"] + q.get("options", [])
                        if q.get("allow_other"):
                            opts_prev.append("Lainnya...")
                        st.selectbox(preview_label, opts_prev, disabled=True, key=f"prev_{qid}")
                    elif qtype == "radio":
                        opts_prev = q.get("options", [])
                        if q.get("allow_other"):
                            opts_prev = opts_prev + ["Lainnya..."]
                        st.radio(preview_label, opts_prev or ["(belum ada pilihan)"], disabled=True, key=f"prev_{qid}", horizontal=True)
                    elif qtype == "checkbox":
                        st.markdown(f"**{preview_label}**")
                        for opt in (q.get("options", []) or ["(belum ada pilihan)"]):
                            st.checkbox(opt, disabled=True, key=f"prev_{qid}_{opt}")
                    elif qtype == "scale":
                        min_v = int(q.get("min_value") or 1)
                        max_v = int(q.get("max_value") or 5)
                        st.slider(preview_label, min_value=min_v, max_value=max_v, disabled=True, key=f"prev_{qid}")
                    elif qtype == "yes_no":
                        st.toggle(preview_label, disabled=True, key=f"prev_{qid}")
                    elif qtype == "photo":
                        st.file_uploader(preview_label, type=["jpg","jpeg","png"], disabled=True, key=f"prev_{qid}")

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
