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

WIB = datetime.timezone(datetime.timedelta(hours=7))

# ── Config ──────────────────────────────────────────────────────────────────
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
CUSTOMER_SHEET_ID   = "113E5fKvZ0wWloSbQ9IQo8QiCOz5C3zlJr28012PbXz4"
SUBMISSION_SHEET_ID = "1RC7v1fGmcz-9q4VowhBnf767P2N_ptonqlKuRSug0Ko"

# GID tiap tab di SUBMISSION_SHEET (cek dari URL sheet saat tab aktif: ...#gid=XXXX)
# sheet1 customer = gid 0 by default
CUSTOMER_GID            = 0
CONFIG_PROMOTOR_GID     = None   # diisi manual jika tahu GID-nya, fallback ke nama tab
CONFIG_RAYON_GID        = None
CONFIG_ZONA_GID         = None

def _csv_url(sheet_id: str, gid: int = 0) -> str:
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
IMAGE_MAX_PX  = 1920
IMAGE_QUALITY = 75

PROGRAM_OPTIONS  = ["WSP Permanen", "WSP Non Permanen", "WSP Program", "COI", "Awareness Investment"]
JENIS_OPTIONS    = ["SNT", "Thinplate", "PNT 3D", "Shopsign"]
BRAND_OPTIONS    = ["EVO", "ARJA", "Wismilak Kretek"]
ACTIVITY_OPTIONS = ["Termin 1", "Termin 2"]

# ── Auth ─────────────────────────────────────────────────────────────────────
@st.cache_resource
def get_credentials():
    info = json.loads(st.secrets["gcp_json"])
    return Credentials.from_service_account_info(info, scopes=SCOPES)

@st.cache_resource
def get_gspread():
    return gspread.authorize(get_credentials())

# ── Data loaders ─────────────────────────────────────────────────────────────
def _read_csv_url(url: str) -> pd.DataFrame:
    """Baca Google Sheet via CSV export — jauh lebih cepat dari API, tanpa OAuth."""
    try:
        df = pd.read_csv(url, dtype=str).fillna("")
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception as e:
        st.error(f"Gagal membaca data dari Google Sheets: {e}\n\nPastikan sheet sudah di-share 'Anyone with link → Viewer'.")
        return pd.DataFrame()

@st.cache_data(ttl=600)
def load_customers():
    url = _csv_url(CUSTOMER_SHEET_ID, CUSTOMER_GID)
    df = _read_csv_url(url)
    if df.empty:
        return df, {}
    cols = df.columns.tolist()
    code_col = next((c for c in cols if "customerno" in c.lower() or "kode" in c.lower() or "no" in c.lower()), cols[0])
    name_col = next((c for c in cols if c.lower() == "customer" or "nama" in c.lower() or "toko" in c.lower()), cols[0])
    addr_col = next((c for c in cols if "alamat" in c.lower() or "address" in c.lower()), None)
    sales_group_col    = next((c for c in cols if "sales group" in c.lower() or "salesgroup" in c.lower()), None)
    sales_district_col = next((c for c in cols if "sales district" in c.lower() or "salesdistrict" in c.lower()), None)
    status_col = next((c for c in cols if c.lower() == "status"), None)

    # Pre-filter aktif & build labels here (inside cache = runs once)
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

@st.cache_data(ttl=300)
def _get_sheet_gids(sheet_id: str) -> dict:
    """Ambil mapping nama tab → GID via gspread (sekali, di-cache lama)."""
    try:
        gc = get_gspread()
        sh = gc.open_by_key(sheet_id)
        return {ws.title: ws.id for ws in sh.worksheets()}
    except Exception:
        return {}

@st.cache_data(ttl=120)
def load_config_list(tab_name: str) -> list:
    """Baca config tab via CSV export. GID di-resolve otomatis."""
    gids = _get_sheet_gids(SUBMISSION_SHEET_ID)
    gid  = gids.get(tab_name)
    if gid is None:
        st.warning(f"Tab '{tab_name}' tidak ditemukan di sheet.")
        return []
    url = _csv_url(SUBMISSION_SHEET_ID, gid)
    df  = _read_csv_url(url)
    if df.empty:
        return []
    first_col = df.iloc[:, 0]
    return sorted([str(v).strip() for v in first_col if str(v).strip()])

@st.cache_data(ttl=120)
def load_all_config():
    """Load promotors, rayon, zona — GID lookup hanya 1x, data via CSV."""
    promotors = load_config_list("Config_NamaPromotor")
    rayon     = load_config_list("Config_Rayon")
    zona      = load_config_list("Config_Zona")
    return promotors, rayon, zona

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

# ── Drive upload via Apps Script ──────────────────────────────────────────────
APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbz27id708tLEcf0eGWNC6BrA7TdHiFVfgsPL2b_xGDkWTqBD30tlGWpCXvQ8F2IXIjO/exec"

def upload_to_drive(data: bytes, filename: str) -> str:
    import base64
    b64 = base64.b64encode(data).decode("utf-8")
    resp = _requests.post(APPS_SCRIPT_URL, data={"image": b64, "filename": filename}, timeout=60)
    resp.raise_for_status()
    result = resp.json()
    if result.get("success"):
        return result["url"]
    raise ValueError(f"Apps Script error: {result}")

# ── Submission writer ─────────────────────────────────────────────────────────
def append_submission(row: list):
    gc = get_gspread()
    ws = gc.open_by_key(SUBMISSION_SHEET_ID).worksheet("Submission")
    ws.append_row(row, value_input_option="USER_ENTERED")

# ── GPS helper ────────────────────────────────────────────────────────────────
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
# UI
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(page_title="Dokumentasi Investment Lapangan", layout="centered")
st.title("Dokumentasi Investment Lapangan")

# ── Halaman konfirmasi setelah submit ─────────────────────────────────────────
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

# ── Load semua data di awal (paralel via cache) ───────────────────────────────
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

# ══════════════════════════════════════════════════════════════════════════════
# FORM — satu deret, tanpa step
# ══════════════════════════════════════════════════════════════════════════════

# ── Bagian 1: Identitas & Program ─────────────────────────────────────────────
st.subheader("Identitas & Detail Program")

col1, col2 = st.columns(2)
with col1:
    promotor = st.selectbox("Eksekutor (Promotor)", ["— Pilih —"] + promotors, key="promotor_sel")
with col2:
    tgl = st.date_input("Tanggal", value=datetime.datetime.now(WIB).date(), key="tgl_sel")

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

# ── Bagian 2: Customer, Rayon & Zona ─────────────────────────────────────────
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

# ── Bagian 3: GPS ─────────────────────────────────────────────────────────────
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

# ── Bagian 4: Foto ────────────────────────────────────────────────────────────
st.subheader("Upload Foto")
foto_ktp      = st.file_uploader("Foto KTP Pemilik *",       type=["jpg","jpeg","png"], key="foto_ktp")
foto_sunblind = st.file_uploader("Foto Sunblind *",          type=["jpg","jpeg","png"], key="foto_sunblind")
foto_display  = st.file_uploader("Foto Display (opsional)",  type=["jpg","jpeg","png"], key="foto_display")
foto_kompens  = st.file_uploader("Foto Bukti Kompensasi *",  type=["jpg","jpeg","png"], key="foto_kompens")
st.caption("Foto dikompres otomatis (max 1920px, JPEG 75%) sebelum diupload.")

catatan = st.text_area("Catatan Tambahan (opsional)", key="catatan_input")

st.divider()

# ── Submit ────────────────────────────────────────────────────────────────────
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
                return upload_to_drive(compress_image(file), fname)

            url_ktp      = safe_upload(foto_ktp,      "ktp")
            url_sunblind = safe_upload(foto_sunblind, "sunblind")
            url_display  = safe_upload(foto_display,  "display")
            url_kompens  = safe_upload(foto_kompens,  "kompensasi")

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
