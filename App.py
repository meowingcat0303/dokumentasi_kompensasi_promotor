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

@st.cache_data(ttl=60)
def load_promotors():
    return load_config_list("Config_NamaPromotor")

@st.cache_data(ttl=60)
def load_rayon():
    return load_config_list("Config_Rayon")

@st.cache_data(ttl=60)
def load_zona():
    return load_config_list("Config_Zona")

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

# ── UI ────────────────────────────────────────────────────────────────────────
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

# Session state init
for key, default in [
    ("step", 1),
    ("selected_customer", None),
    ("gps_coords", ""),
    ("rayon_zona_mode", "Otomatis (dari data customer)"),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# ══════════════════════════════════════════════════════════════════════════════
# STEP 1 — Identitas & Detail Program
# ══════════════════════════════════════════════════════════════════════════════
with st.expander("1. Identitas & Detail Program", expanded=(st.session_state.step == 1)):
    promotors = load_promotors()

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
    with st.expander("2. Pilih Customer, Rayon & Zona", expanded=(st.session_state.step == 2)):

        # ── Mode selector ────────────────────────────────────────────────────
        st.radio(
            "Mode Rayon & Zona:",
            ["Otomatis (dari data customer)", "Manual (pilih sendiri)"],
            key="rayon_zona_mode",
            horizontal=True,
            help="Otomatis: Rayon & Zona diisi otomatis sesuai data customer yang dipilih.\nManual: Pilih Rayon & Zona secara mandiri dari dropdown.",
        )
        mode_otomatis = st.session_state.rayon_zona_mode == "Otomatis (dari data customer)"

        # ── Rayon & Zona (hanya tampil di mode manual) ───────────────────────
        if not mode_otomatis:
            rayon_list = load_rayon()
            zona_list  = load_zona()

            if not rayon_list or not zona_list:
                with st.expander("Debug info Config sheet", expanded=True):
                    st.write(f"Config_Rayon: {len(rayon_list)} item ditemukan: {rayon_list}")
                    st.write(f"Config_Zona: {len(zona_list)} item ditemukan: {zona_list}")
                    st.caption("Pastikan nama tab di sheet persis: Config_Rayon dan Config_Zona (case-sensitive), dan data dimulai dari baris 2.")

            col_r, col_z = st.columns(2)
            with col_r:
                rayon = st.selectbox("Rayon", ["— Pilih —"] + rayon_list, key="rayon_sel")
            with col_z:
                zona = st.selectbox("Zona", ["— Pilih —"] + zona_list, key="zona_sel")
        else:
            # Placeholder — akan diisi setelah customer dipilih
            rayon = st.session_state.get("auto_rayon", "")
            zona  = st.session_state.get("auto_zona", "")

            if rayon or zona:
                col_r, col_z = st.columns(2)
                with col_r:
                    st.info(f"**Rayon:** {rayon}" if rayon else "Rayon akan terisi otomatis")
                with col_z:
                    st.info(f"**Zona:** {zona}" if zona else "Zona akan terisi otomatis")

        st.divider()

        # ── Customer search ──────────────────────────────────────────────────
        df_cust = load_customers()

        if df_cust.empty:
            st.error("Data customer tidak ditemukan.")
            st.stop()

        cols = df_cust.columns.tolist()
        code_col = next((c for c in cols if "customerno" in c.lower() or "kode" in c.lower() or "no" in c.lower()), cols[0])
        name_col = next((c for c in cols if c.lower() == "customer" or "nama" in c.lower() or "toko" in c.lower()), cols[0])
        addr_col = next((c for c in cols if "alamat" in c.lower() or "address" in c.lower()), None)

        # Deteksi kolom SALES GROUP (Rayon) dan SALES DISTRICT (Zona) dari sheet customer
        sales_group_col   = next((c for c in cols if "sales group" in c.lower() or "salesgroup" in c.lower()), None)
        sales_district_col = next((c for c in cols if "sales district" in c.lower() or "salesdistrict" in c.lower()), None)

        search_mode = st.radio("Cari berdasarkan:", ["Nama Toko", "Kode Customer"], horizontal=True, key="search_mode")
        search_col  = name_col if search_mode == "Nama Toko" else code_col
        search      = st.text_input("Ketik untuk mencari", key="search_input")

        mask    = df_cust[search_col].astype(str).str.contains(search, case=False, na=False) if search else pd.Series([True] * len(df_cust))
        results = df_cust[mask]
        st.caption(f"Menampilkan {min(len(results), 50)} dari {len(df_cust)} data")

        display_cols = [c for c in [code_col, name_col, addr_col, sales_group_col, sales_district_col] if c]
        st.dataframe(results[display_cols].head(50), use_container_width=True, hide_index=True)

        if not results.empty:
            def make_label(r):
                return f"{r[name_col]} ({r[code_col]})"
            option_labels = ["— Pilih —"] + [make_label(r) for _, r in results.head(50).iterrows()]
            chosen_label  = st.selectbox("Pilih Customer", option_labels, key="cust_sel")

            if chosen_label != "— Pilih —":
                idx = option_labels.index(chosen_label) - 1
                row = results.head(50).iloc[idx]

                # Jika mode otomatis, ambil rayon & zona dari baris customer
                if mode_otomatis:
                    auto_rayon = str(row[sales_group_col]).strip() if sales_group_col and sales_group_col in row else ""
                    auto_zona  = str(row[sales_district_col]).strip() if sales_district_col and sales_district_col in row else ""
                    st.session_state.auto_rayon = auto_rayon
                    st.session_state.auto_zona  = auto_zona
                    rayon = auto_rayon
                    zona  = auto_zona

                    col_r, col_z = st.columns(2)
                    with col_r:
                        st.success(f"**Rayon (otomatis):** {rayon or '—'}")
                    with col_z:
                        st.success(f"**Zona (otomatis):** {zona or '—'}")

                st.success(
                    f"**{row[name_col]}** ({row[code_col]})"
                    + (f" — {row[addr_col]}" if addr_col else "")
                )

                if st.button("Konfirmasi →", key="btn_step2"):
                    if not mode_otomatis and (rayon == "— Pilih —" or zona == "— Pilih —"):
                        st.warning("Pilih Rayon dan Zona terlebih dahulu.")
                    elif mode_otomatis and (not rayon or not zona):
                        st.warning(
                            "Rayon atau Zona tidak ditemukan di data customer ini. "
                            "Coba gunakan Mode Manual atau periksa data sheet customer."
                        )
                    else:
                        st.session_state.selected_customer = row.to_dict()
                        # Simpan rayon & zona yang akan dipakai ke session state
                        st.session_state["confirmed_rayon"] = rayon
                        st.session_state["confirmed_zona"]  = zona
                        st.session_state.step = 3
                        st.rerun()
        else:
            st.warning("Tidak ada hasil. Coba kata kunci lain.")

# ══════════════════════════════════════════════════════════════════════════════
# STEP 3 — Upload Foto, Lokasi GPS & Submit
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.step >= 3:
    with st.expander("3. Upload Foto, Lokasi & Submit", expanded=(st.session_state.step == 3)):
        cust = st.session_state.selected_customer

        # Ambil rayon & zona yang sudah dikonfirmasi
        confirmed_rayon = st.session_state.get("confirmed_rayon", "")
        confirmed_zona  = st.session_state.get("confirmed_zona", "")

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
| Rayon | {confirmed_rayon or '-'} |
| Zona | {confirmed_zona or '-'} |
| Kode Customer | `{cust.get(code_col, '-')}` |
| Nama Outlet | {cust.get(name_col, '-')} |
| Alamat | {cust.get(addr_col, '-') if addr_col else '-'} |
""")

        st.divider()

        # ── GPS ───────────────────────────────────────────────────────────────
        st.subheader("Lokasi GPS")
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

        st.divider()

        # ── Foto ──────────────────────────────────────────────────────────────
        st.subheader("Upload Foto")
        foto_ktp      = st.file_uploader("Foto KTP Pemilik *",        type=["jpg","jpeg","png"], key="foto_ktp")
        foto_sunblind = st.file_uploader("Foto Sunblind *",           type=["jpg","jpeg","png"], key="foto_sunblind")
        foto_display  = st.file_uploader("Foto Display (opsional)",   type=["jpg","jpeg","png"], key="foto_display")
        foto_kompens  = st.file_uploader("Foto Bukti Kompensasi *",   type=["jpg","jpeg","png"], key="foto_kompens")

        st.caption("Foto dikompres otomatis (max 1920px, JPEG 75%) sebelum diupload.")

        catatan = st.text_area("Catatan Tambahan (opsional)", key="catatan_input")

        # ── Submit ────────────────────────────────────────────────────────────
        if st.button("Submit Dokumentasi", key="btn_submit", type="primary"):
            errors = []
            if not foto_ktp:      errors.append("Foto KTP wajib diupload")
            if not foto_sunblind: errors.append("Foto Sunblind wajib diupload")
            if not foto_kompens:  errors.append("Foto Bukti Kompensasi wajib diupload")
            if not cust:          errors.append("Customer belum dipilih")

            if errors:
                for e in errors:
                    st.error(e)
            else:
                with st.spinner("Mengupload foto dan menyimpan data, harap tunggu..."):
                    ts  = datetime.datetime.now(WIB).strftime("%Y%m%d_%H%M%S")
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
                            st.error(f"Gagal upload {label}: {e}")
                            raise

                    url_ktp      = safe_upload(foto_ktp,      "ktp")
                    url_sunblind = safe_upload(foto_sunblind, "sunblind")
                    url_display  = safe_upload(foto_display,  "display")
                    url_kompens  = safe_upload(foto_kompens,  "kompensasi")

                    unique_id = f"{ts}_{uid}"
                    submit_time = datetime.datetime.now(WIB).strftime("%Y-%m-%d %H:%M:%S")

                    row = [
                        unique_id,
                        str(st.session_state.get("tgl_sel", datetime.datetime.now(WIB).date())),
                        promotor_name,
                        st.session_state.get("program_sel", ""),
                        st.session_state.get("jenis_sel", ""),
                        st.session_state.get("brand_sel", ""),
                        st.session_state.get("activity_sel", ""),
                        confirmed_rayon,
                        confirmed_zona,
                        cust.get(name_col, ""),
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
                        "outlet": cust.get(name_col, "-"),
                        "kode": kode,
                        "promotor": promotor_name,
                        "waktu": submit_time,
                    }

                for k in ["step", "selected_customer", "gps_coords", "auto_rayon", "auto_zona", "confirmed_rayon", "confirmed_zona"]:
                    if k in st.session_state:
                        del st.session_state[k]
                st.cache_data.clear()
                st.rerun()
