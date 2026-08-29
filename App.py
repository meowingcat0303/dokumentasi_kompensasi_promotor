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

# ── Config ────────────────────────────────────────────────────────────────────
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
CUSTOMER_SHEET_ID   = "113E5fKvZ0wWloSbQ9IQo8QiCOz5C3zlJr28012PbXz4"
SUBMISSION_SHEET_ID = "1CL-_9EDLEEuj-x54lrCgNLZYdhuNyncnY8wKylwdg54"
IMAGE_MAX_PX  = 1920
IMAGE_QUALITY = 75

TERMIN_OPTIONS = ["Termin 1", "Termin 2"]
WSP_OPTIONS    = ["WSP Bronze", "WSP Silver"]

APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbz68G8gdX7BZEuPRHhrYpHeumsjqIQap5vjz7OwQzStYPAhjcjDsaPyf3QfMKQAbK7R/exec"

SCORING = {
    "Stock Minimal": 10,
    "Kebersihan":    10,
    "Planogram":     20,
    "Stiker":        10,
}

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
    code_col           = next((c for c in cols if "customerno" in c.lower() or "kode" in c.lower() or "no" in c.lower()), cols[0])
    name_col           = next((c for c in cols if c.lower() == "customer" or "nama" in c.lower() or "toko" in c.lower()), cols[0])
    addr_col           = next((c for c in cols if "alamat" in c.lower() or "address" in c.lower()), None)
    sales_group_col    = next((c for c in cols if "sales group" in c.lower() or "salesgroup" in c.lower()), None)
    sales_district_col = next((c for c in cols if "sales district" in c.lower() or "salesdistrict" in c.lower()), None)
    status_col         = next((c for c in cols if c.lower() == "status"), None)
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
    salesman = load_config_list("Config_NamaSalesman")
    rayon    = load_config_list("Config_Rayon")
    zona     = load_config_list("Config_Zona")
    return salesman, rayon, zona

@st.cache_data(ttl=120)
def load_existing_submissions() -> pd.DataFrame:
    try:
        gc = get_gspread()
        ws = gc.open_by_key(SUBMISSION_SHEET_ID).worksheet("Submission")
        return _sheet_to_df(ws)
    except Exception:
        return pd.DataFrame()

def check_duplicate_termin(df_sub: pd.DataFrame, kode_customer: str, termin: str) -> bool:
    if df_sub.empty:
        return False
    cols = df_sub.columns.tolist()
    kode_col   = next((c for c in cols if "kode" in c.lower() or "customer" in c.lower()), None)
    termin_col = next((c for c in cols if "termin" in c.lower()), None)
    if not kode_col or not termin_col:
        return False
    mask = (
        df_sub[kode_col].str.strip() == str(kode_customer).strip()
    ) & (
        df_sub[termin_col].str.strip() == str(termin).strip()
    )
    return mask.any()

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

# ── Drive upload via Apps Script (form-encoded, sama persis dengan Dokumentasi Promotor) ──
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

# ══════════════════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(page_title="Penilaian WSP", layout="centered")
st.title("Penilaian WSP")

# ── Halaman konfirmasi setelah submit ─────────────────────────────────────────
if "submit_info" in st.session_state:
    info = st.session_state.submit_info
    st.success("Penilaian berhasil dikirim dan tersimpan.")
    st.markdown(f"""
**Ringkasan Penilaian:**

| | |
|---|---|
| ID Unik | `{info['unique_id']}` |
| Outlet | {info['outlet']} ({info['kode']}) |
| Salesman | {info['salesman']} |
| Termin | {info['termin']} |
| Tipe WSP | {info['tipe_wsp']} |
| Total Nilai | **{info['total']} / 50** |
| Waktu Kirim | {info['waktu']} |

Data sudah masuk ke sistem.
""")
    if st.button("Input Penilaian Baru", type="primary"):
        del st.session_state["submit_info"]
        st.session_state.pop("confirm_override", None)
        st.rerun()
    st.stop()

# ── Load data ─────────────────────────────────────────────────────────────────
with st.spinner("Memuat data..."):
    salesman_list, rayon_list, zona_list = load_all_config()
    df_cust, meta = load_customers()
    df_sub = load_existing_submissions()

code_col           = meta["code_col"]
name_col           = meta["name_col"]
addr_col           = meta["addr_col"]
sales_group_col    = meta["sales_group_col"]
sales_district_col = meta["sales_district_col"]

st.caption(f"Data customer: **{meta['n_aktif']:,} toko aktif** dari {meta['n_total']:,} total")
st.divider()

# ── Bagian 1: Identitas ───────────────────────────────────────────────────────
st.subheader("Identitas")

col1, col2 = st.columns(2)
with col1:
    tgl = st.date_input("Tanggal", value=datetime.datetime.now(WIB).date(), key="tgl_sel")
with col2:
    salesman = st.selectbox("Salesman", ["— Pilih —"] + salesman_list, key="salesman_sel")

col3, col4 = st.columns(2)
with col3:
    termin = st.selectbox("Termin Penilaian", ["— Pilih —"] + TERMIN_OPTIONS, key="termin_sel")
with col4:
    tipe_wsp = st.selectbox("Tipe WSP", ["— Pilih —"] + WSP_OPTIONS, key="tipe_wsp_sel")

st.divider()

# ── Bagian 2: Customer ────────────────────────────────────────────────────────
st.subheader("Customer")

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
rayon_final  = ""
zona_final   = ""
is_duplicate = False

if chosen_label != "— Pilih —":
    match = df_cust[df_cust["_label"] == chosen_label]
    if not match.empty:
        selected_customer = match.iloc[0].to_dict()
        kode_customer = selected_customer.get(code_col, "")

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

        if termin != "— Pilih —":
            is_duplicate = check_duplicate_termin(df_sub, kode_customer, termin)
            if is_duplicate:
                st.warning(
                    f"⚠️ Outlet ini sudah pernah disubmit untuk **{termin}**. "
                    "Kamu tetap bisa submit, tapi pastikan ini bukan duplikat."
                )

st.divider()

# ── Bagian 3: Penilaian ───────────────────────────────────────────────────────
st.subheader("Penilaian")

scores = {}
for aspek, maks in SCORING.items():
    scores[aspek] = st.number_input(
        f"{aspek} (0 – {maks})",
        min_value=0, max_value=maks, value=0, step=1,
        key=f"score_{aspek}",
    )

total_nilai = sum(scores.values())
st.metric("Total Nilai", f"{total_nilai} / 50")

st.divider()

# ── Bagian 4: GPS ─────────────────────────────────────────────────────────────
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

# ── Bagian 5: Foto ────────────────────────────────────────────────────────────
st.subheader("Foto Dispenser")
foto_dispenser = st.file_uploader("Foto Dispenser *", type=["jpg", "jpeg", "png"], key="foto_dispenser")
st.caption("Foto dikompres otomatis (max 1920px, JPEG 75%) sebelum diupload.")

catatan = st.text_area("Catatan Tambahan (opsional)", key="catatan_input")

st.divider()

# ── Submit ────────────────────────────────────────────────────────────────────
if is_duplicate:
    confirm_override = st.checkbox(
        "Saya memahami bahwa outlet ini sudah pernah disubmit untuk termin ini, dan ingin tetap melanjutkan.",
        key="confirm_override",
    )
else:
    confirm_override = True

if st.button("Submit Penilaian", key="btn_submit", type="primary"):
    errors = []
    if salesman  == "— Pilih —": errors.append("Salesman wajib dipilih")
    if termin    == "— Pilih —": errors.append("Termin Penilaian wajib dipilih")
    if tipe_wsp  == "— Pilih —": errors.append("Tipe WSP wajib dipilih")
    if not selected_customer:    errors.append("Customer belum dipilih")
    if not rayon_final:          errors.append("Rayon tidak ditemukan / belum dipilih")
    if not zona_final:           errors.append("Zona tidak ditemukan / belum dipilih")
    if not foto_dispenser:       errors.append("Foto Dispenser wajib diupload")
    if is_duplicate and not confirm_override:
        errors.append("Centang konfirmasi untuk melanjutkan submit data duplikat")

    if errors:
        for e in errors:
            st.error(e)
    else:
        with st.spinner("Mengupload foto dan menyimpan data, harap tunggu..."):
            ts   = datetime.datetime.now(WIB).strftime("%Y%m%d_%H%M%S")
            uid  = str(uuid.uuid4())[:8]
            kode = selected_customer.get(code_col, "unknown")

            fname = f"{ts}_{kode}_{salesman}_dispenser_{uid}.jpg"
            try:
                url_dispenser = upload_to_drive(compress_image(foto_dispenser), fname)
            except Exception as e:
                st.error(f"❌ Gagal upload foto:\n\n{e}")
                st.stop()

            unique_id   = f"{ts}_{uid}"
            submit_time = datetime.datetime.now(WIB).strftime("%Y-%m-%d %H:%M:%S")

            row = [
                unique_id,
                str(tgl),
                salesman,
                termin,
                tipe_wsp,
                kode,
                selected_customer.get(name_col, ""),
                rayon_final,
                zona_final,
                scores["Stock Minimal"],
                scores["Kebersihan"],
                scores["Planogram"],
                scores["Stiker"],
                total_nilai,
                url_dispenser,
                st.session_state.gps_coords,
                catatan,
                submit_time,
            ]
            try:
                append_submission(row)
            except Exception as e:
                st.error(f"❌ Foto berhasil diupload, tapi gagal simpan ke Google Sheets:\n\n{e}")
                st.stop()

            st.session_state.submit_info = {
                "unique_id": unique_id,
                "outlet":    selected_customer.get(name_col, "-"),
                "kode":      kode,
                "salesman":  salesman,
                "termin":    termin,
                "tipe_wsp":  tipe_wsp,
                "total":     total_nilai,
                "waktu":     submit_time,
            }
            st.session_state.gps_coords = ""
            st.cache_data.clear()
            st.rerun()
