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

# ── Config ──────────────────────────────────────────────────────────────────
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
]
CUSTOMER_SHEET_ID   = "113E5fKvZ0wWloSbQ9IQo8QiCOz5C3zlJr28012PbXz4"
SUBMISSION_SHEET_ID = "1RC7v1fGmcz-9q4VowhBnf767P2N_ptonqlKuRSug0Ko"
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
def _sheet_to_df(ws) -> pd.DataFrame:
    rows = ws.get_all_values()
    if not rows:
        return pd.DataFrame()
    headers = [str(h).strip() for h in rows[0]]
    return pd.DataFrame(rows[1:], columns=headers)

@st.cache_data(ttl=300)
def load_customers():
    gc = get_gspread()
    ws = gc.open_by_key(CUSTOMER_SHEET_ID).sheet1
    return _sheet_to_df(ws)

@st.cache_data(ttl=60)
def load_config_list(tab_name: str) -> list:
    gc = get_gspread()
    ws = gc.open_by_key(SUBMISSION_SHEET_ID).worksheet(tab_name)
    rows = ws.get_all_values()
    if len(rows) < 2:
        return []
    return sorted([str(r[0]).strip() for r in rows[1:] if r and str(r[0]).strip()])

def load_promotors():  return load_config_list("Config_NamaPromotor")
def load_rayon():      return load_config_list("Config_Rayon")
def load_zona():       return load_config_list("Config_Zona")

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
    resp = _requests.post(
        APPS_SCRIPT_URL,
        data={"image": b64, "filename": filename},
        timeout=60,
    )
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

# ── GPS helper (inject JS, baca dari query param) ─────────────────────────────
GPS_JS = """
<script>
function getLocation() {
    if (!navigator.geolocation) {
        document.getElementById('gps-status').innerText = '⚠️ Browser tidak mendukung GPS.';
        return;
    }
    document.getElementById('gps-status').innerText = '📡 Mengambil lokasi...';
    navigator.geolocation.getCurrentPosition(
        function(pos) {
            var lat = pos.coords.latitude.toFixed(6);
            var lng = pos.coords.longitude.toFixed(6);
            var acc = Math.round(pos.coords.accuracy);
            document.getElementById('gps-status').innerText =
                '✅ Lokasi ditemukan: ' + lat + ', ' + lng + ' (±' + acc + 'm)';
            // Kirim ke Streamlit via query param trick
            window.parent.postMessage({
                type: 'streamlit:setComponentValue',
                value: lat + ',' + lng
            }, '*');
            // Fallback: isi hidden input lalu submit form
            var inp = window.parent.document.querySelector('input[data-testid="stTextInput"][aria-label="gps_hidden"]');
            if (inp) {
                var nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
                nativeInputValueSetter.call(inp, lat + ',' + lng);
                inp.dispatchEvent(new Event('input', { bubbles: true }));
            }
        },
        function(err) {
            document.getElementById('gps-status').innerText = '❌ GPS error: ' + err.message;
        },
        { enableHighAccuracy: true, timeout: 15000 }
    );
}
</script>
<button onclick="getLocation()" style="
    background:#1f77b4; color:white; border:none; padding:10px 20px;
    border-radius:6px; font-size:15px; cursor:pointer; margin-bottom:8px;">
    📍 Ambil Lokasi GPS
</button>
<div id="gps-status" style="color:#444; font-size:13px; margin-top:4px;"></div>
"""

# ── UI ────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Dokumentasi Investment Lapangan", page_icon="📋", layout="centered")
st.title("📋 Dokumentasi Investment Lapangan")

# Session state init
for key, default in [
    ("step", 1),
    ("selected_customer", None),
    ("gps_coords", ""),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# ══════════════════════════════════════════════════════════════════════════════
# STEP 1 — Identitas & Detail Program
# ══════════════════════════════════════════════════════════════════════════════
with st.expander("① Identitas & Detail Program", expanded=(st.session_state.step == 1)):
    promotors = load_promotors()

    col1, col2 = st.columns(2)
    with col1:
        promotor = st.selectbox("Eksekutor (Promotor)", ["— Pilih —"] + promotors, key="promotor_sel")
    with col2:
        tgl = st.date_input("Tanggal", value=datetime.date.today(), key="tgl_sel")

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

    if st.button("Lanjut →", key="btn_step1"):
        missing = [
            f for f, v in [
                ("Eksekutor", promotor), ("Program Investment", program),
                ("Jenis Investment", jenis), ("Brand", brand), ("Activity", activity),
            ] if v == "— Pilih —"
        ]
        if missing:
            st.warning(f"Lengkapi dahulu: {', '.join(missing)}")
        else:
            st.session_state.step = 2
            st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# STEP 2 — Pilih Customer, Rayon & Zona
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.step >= 2:
    with st.expander("② Pilih Customer, Rayon & Zona", expanded=(st.session_state.step == 2)):
        # Rayon & Zona dari config
        rayon_list = load_rayon()
        zona_list  = load_zona()

        col_r, col_z = st.columns(2)
        with col_r:
            rayon = st.selectbox("Rayon", ["— Pilih —"] + rayon_list, key="rayon_sel")
        with col_z:
            zona = st.selectbox("Zona", ["— Pilih —"] + zona_list, key="zona_sel")

        st.divider()

        # Customer search
        df_cust = load_customers()

        if df_cust.empty:
            st.error("❌ Data customer tidak ditemukan.")
            st.stop()

        cols = df_cust.columns.tolist()
        code_col = next((c for c in cols if "customerno" in c.lower() or "kode" in c.lower() or "no" in c.lower()), cols[0])
        name_col = next((c for c in cols if c.lower() == "customer" or "nama" in c.lower() or "toko" in c.lower()), cols[0])
        addr_col = next((c for c in cols if "alamat" in c.lower() or "address" in c.lower()), None)

        search_mode = st.radio("Cari berdasarkan:", ["Nama Toko", "Kode Customer"], horizontal=True, key="search_mode")
        search_col  = name_col if search_mode == "Nama Toko" else code_col
        search      = st.text_input("🔍 Ketik untuk mencari", key="search_input")

        mask    = df_cust[search_col].astype(str).str.contains(search, case=False, na=False) if search else pd.Series([True] * len(df_cust))
        results = df_cust[mask]
        st.caption(f"Menampilkan {min(len(results), 50)} dari {len(df_cust)} data")

        display_cols = [c for c in [code_col, name_col, addr_col] if c]
        st.dataframe(results[display_cols].head(50), use_container_width=True, hide_index=True)

        if not results.empty:
            def make_label(r):
                return f"{r[name_col]} ({r[code_col]})"
            option_labels = ["— Pilih —"] + [make_label(r) for _, r in results.head(50).iterrows()]
            chosen_label  = st.selectbox("Pilih Customer", option_labels, key="cust_sel")

            if chosen_label != "— Pilih —":
                idx = option_labels.index(chosen_label) - 1
                row = results.head(50).iloc[idx]
                st.success(
                    f"**{row[name_col]}** ({row[code_col]})"
                    + (f" — {row[addr_col]}" if addr_col else "")
                )

                if st.button("Konfirmasi →", key="btn_step2"):
                    if rayon == "— Pilih —" or zona == "— Pilih —":
                        st.warning("Pilih Rayon dan Zona terlebih dahulu.")
                    else:
                        st.session_state.selected_customer = row.to_dict()
                        st.session_state.step = 3
                        st.rerun()
        else:
            st.warning("Tidak ada hasil. Coba kata kunci lain.")

# ══════════════════════════════════════════════════════════════════════════════
# STEP 3 — Upload Foto, Lokasi GPS & Submit
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.step >= 3:
    with st.expander("③ Upload Foto, Lokasi & Submit", expanded=(st.session_state.step == 3)):
        cust = st.session_state.selected_customer

        if cust:
            cust_cols = list(cust.keys())
            code_col = next((c for c in cust_cols if "customerno" in c.lower() or "kode" in c.lower() or "no" in c.lower()), cust_cols[0])
            name_col = next((c for c in cust_cols if c.lower() == "customer" or "nama" in c.lower() or "toko" in c.lower()), cust_cols[0])
            addr_col = next((c for c in cust_cols if "alamat" in c.lower() or "address" in c.lower()), None)

            st.markdown(f"""
| Field | Value |
|---|---|
| Eksekutor | {st.session_state.get('promotor_sel', '-')} |
| Tanggal | {st.session_state.get('tgl_sel', '-')} |
| Program | {st.session_state.get('program_sel', '-')} |
| Jenis | {st.session_state.get('jenis_sel', '-')} |
| Brand | {st.session_state.get('brand_sel', '-')} |
| Activity | {st.session_state.get('activity_sel', '-')} |
| Rayon | {st.session_state.get('rayon_sel', '-')} |
| Zona | {st.session_state.get('zona_sel', '-')} |
| Kode Customer | `{cust.get(code_col, '-')}` |
| Nama Outlet | {cust.get(name_col, '-')} |
| Alamat | {cust.get(addr_col, '-') if addr_col else '-'} |
""")

        st.divider()

        # ── GPS ───────────────────────────────────────────────────────────────
        st.subheader("📍 Lokasi GPS")
        st.components.v1.html(GPS_JS, height=90)

        # Input tersembunyi — user bisa paste manual jika GPS gagal
        gps_input = st.text_input(
            "Koordinat (lat,lng) — terisi otomatis setelah klik tombol GPS di atas",
            value=st.session_state.gps_coords,
            key="gps_manual",
            placeholder="Contoh: -6.595038, 106.816635",
        )
        if gps_input:
            st.session_state.gps_coords = gps_input

        st.divider()

        # ── Foto ──────────────────────────────────────────────────────────────
        st.subheader("📸 Upload Foto")
        foto_ktp      = st.file_uploader("🪪 Foto KTP Pemilik *",          type=["jpg","jpeg","png"], key="foto_ktp")
        foto_sunblind = st.file_uploader("🪟 Foto Sunblind *",             type=["jpg","jpeg","png"], key="foto_sunblind")
        foto_display  = st.file_uploader("🖼️ Foto Display (opsional)",      type=["jpg","jpeg","png"], key="foto_display")
        foto_kompens  = st.file_uploader("💰 Foto Bukti Kompensasi *",      type=["jpg","jpeg","png"], key="foto_kompens")

        st.caption("Foto dikompres otomatis (max 1920px, JPEG 75%) sebelum diupload.")

        jumlah  = st.number_input("Jumlah Kompensasi (Rp)", min_value=0, step=1000, key="jumlah_input")
        catatan = st.text_area("Catatan Tambahan (opsional)", key="catatan_input")

        # ── Submit ────────────────────────────────────────────────────────────
        if st.button("✅ Submit Dokumentasi", key="btn_submit", type="primary"):
            errors = []
            if not foto_ktp:      errors.append("Foto KTP wajib diupload")
            if not foto_sunblind: errors.append("Foto Sunblind wajib diupload")
            if not foto_kompens:  errors.append("Foto Bukti Kompensasi wajib diupload")
            if not cust:          errors.append("Customer belum dipilih")

            if errors:
                for e in errors:
                    st.error(f"❌ {e}")
            else:
                with st.spinner("Mengupload foto & menyimpan data..."):
                    ts  = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    uid = str(uuid.uuid4())[:8]
                    promotor_name = st.session_state.get("promotor_sel", "unknown")
                    kode = cust.get(code_col, "unknown")

                    def safe_upload(file, label):
                        if file is None:
                            return ""
                        try:
                            fname = f"{ts}_{kode}_{promotor_name}_{label}_{uid}.jpg"
                            return upload_to_drive(compress_image(file), fname)
                        except Exception as e:
                            st.error(f"❌ Gagal upload {label}: {e}")
                            raise

                    url_ktp      = safe_upload(foto_ktp,      "ktp")
                    url_sunblind = safe_upload(foto_sunblind, "sunblind")
                    url_display  = safe_upload(foto_display,  "display")
                    url_kompens  = safe_upload(foto_kompens,  "kompensasi")

                    # Unique ID
                    unique_id = f"{ts}_{uid}"

                    row = [
                        unique_id,                                                         # Unique Id
                        str(st.session_state.get("tgl_sel", datetime.date.today())),      # Date
                        promotor_name,                                                      # Eksekutor
                        st.session_state.get("program_sel", ""),                           # Program Investment
                        st.session_state.get("jenis_sel", ""),                             # Jenis Investment
                        st.session_state.get("brand_sel", ""),                             # Brand
                        st.session_state.get("activity_sel", ""),                          # Activity
                        st.session_state.get("rayon_sel", ""),                             # Rayon
                        st.session_state.get("zona_sel", ""),                              # Zona
                        cust.get(name_col, ""),                                            # Nama Outlet
                        kode,                                                               # Kode Customer
                        url_sunblind,                                                       # Foto Sunblind
                        url_display,                                                        # Foto Display
                        url_kompens,                                                        # Foto Kompensasi
                        url_ktp,                                                            # Foto KTP
                        st.session_state.gps_coords,                                       # Lokasi (LatLong)
                        jumlah,                                                             # Jumlah Kompensasi
                        catatan,                                                            # Catatan
                        datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),            # Timestamp submit
                    ]
                    append_submission(row)

                st.success("✅ Dokumentasi berhasil disimpan!")
                st.balloons()

                # Reset session
                for k in ["step", "selected_customer", "gps_coords"]:
                    del st.session_state[k]
                st.cache_data.clear()
                st.rerun()
