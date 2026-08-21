import streamlit as st
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from PIL import Image
import io
import json
import datetime
import uuid

# ── Config ──────────────────────────────────────────────────────────────────
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
CUSTOMER_SHEET_ID   = "1RC7v1fGmcz-9q4VowhBnf767P2N_ptonqlKuRSug0Ko"  # Sheet data customer
SUBMISSION_SHEET_ID = "1RC7v1fGmcz-9q4VowhBnf767P2N_ptonqlKuRSug0Ko"  # Sheet output (sama, tab berbeda)
DRIVE_FOLDER_ID     = "19ZIk2g8hsr6dmU2KW2y78Q4KS3Y47WGX"

IMAGE_MAX_PX  = 1920
IMAGE_QUALITY = 75

# ── Auth ─────────────────────────────────────────────────────────────────────
@st.cache_resource
def get_credentials():
    info = json.loads(st.secrets["gcp_json"])
    return Credentials.from_service_account_info(info, scopes=SCOPES)

@st.cache_resource
def get_gspread():
    return gspread.authorize(get_credentials())

@st.cache_resource
def get_drive():
    return build("drive", "v3", credentials=get_credentials())

# ── Data loaders ─────────────────────────────────────────────────────────────
def _sheet_to_df(ws) -> pd.DataFrame:
    """Baca sheet → DataFrame tanpa masalah tipe kolom di pandas baru."""
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
def load_promotors():
    gc = get_gspread()
    ws = gc.open_by_key(SUBMISSION_SHEET_ID).worksheet("Config_NamaPromotor")
    rows = ws.get_all_values()
    if len(rows) < 2:
        return []
    # Kolom pertama, skip header baris 1
    return sorted([str(r[0]).strip() for r in rows[1:] if r and str(r[0]).strip()])

@st.cache_data(ttl=60)
def load_programs():
    gc = get_gspread()
    ws = gc.open_by_key(SUBMISSION_SHEET_ID).worksheet("Config_NamaProgram")
    rows = ws.get_all_values()
    if len(rows) < 2:
        return []
    return sorted([str(r[0]).strip() for r in rows[1:] if r and str(r[0]).strip()])

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
def upload_to_drive(data: bytes, filename: str, folder_id: str) -> str:
    service = get_drive()
    meta = {"name": filename, "parents": [folder_id]}
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype="image/jpeg", resumable=False)
    f = service.files().create(body=meta, media_body=media, fields="id, webViewLink").execute()
    # Make publicly readable so admin bisa lihat
    service.permissions().create(
        fileId=f["id"],
        body={"type": "anyone", "role": "reader"},
    ).execute()
    return f.get("webViewLink", "")

# ── Submission writer ─────────────────────────────────────────────────────────
def append_submission(row: list):
    gc = get_gspread()
    ws = gc.open_by_key(SUBMISSION_SHEET_ID).worksheet("Submissions")
    ws.append_row(row, value_input_option="USER_ENTERED")

# ── UI ────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Dokumentasi Kompensasi Lapangan", page_icon="📋", layout="centered")
st.title("📋 Dokumentasi Kompensasi Lapangan")

# Step tracker
if "step" not in st.session_state:
    st.session_state.step = 1
if "selected_customer" not in st.session_state:
    st.session_state.selected_customer = None

# ── STEP 1: Pilih Promotor & Program ─────────────────────────────────────────
with st.expander("① Identitas Promotor", expanded=(st.session_state.step == 1)):
    promotors = load_promotors()
    programs  = load_programs()

    promotor = st.selectbox("Nama Promotor", ["— Pilih —"] + promotors, key="promotor_sel")
    program  = st.selectbox("Jenis Program", ["— Pilih —"] + programs, key="program_sel")

    if st.button("Lanjut →", key="btn_step1"):
        if promotor == "— Pilih —" or program == "— Pilih —":
            st.warning("Pilih promotor dan jenis program terlebih dahulu.")
        else:
            st.session_state.step = 2
            st.rerun()

# ── STEP 2: Cari & Pilih Customer ─────────────────────────────────────────────
if st.session_state.step >= 2:
    with st.expander("② Pilih Customer", expanded=(st.session_state.step == 2)):
        df_cust = load_customers()

        search = st.text_input("🔍 Cari kode / nama toko / alamat", key="search_input")

        if search:
            mask = df_cust.apply(
                lambda col: col.astype(str).str.contains(search, case=False, na=False)
            ).any(axis=1)
            results = df_cust[mask]
        else:
            results = df_cust.head(50)

        st.caption(f"Menampilkan {len(results)} baris")

        # Detect kolom kunci
        cols = df_cust.columns.tolist()
        code_col = next((c for c in cols if "kode" in c.lower()), cols[0])
        name_col = next((c for c in cols if "nama" in c.lower() or "toko" in c.lower()), cols[1] if len(cols) > 1 else cols[0])
        addr_col = next((c for c in cols if "alamat" in c.lower() or "address" in c.lower()), None)

        display_cols = [c for c in [code_col, name_col, addr_col] if c]
        st.dataframe(results[display_cols], use_container_width=True, hide_index=True)

        # Dropdown pilih dari hasil search
        if not results.empty:
            options = results[code_col].astype(str).tolist()
            chosen_code = st.selectbox("Pilih Kode Customer", ["— Pilih —"] + options, key="cust_sel")

            if chosen_code != "— Pilih —":
                row = results[results[code_col].astype(str) == chosen_code].iloc[0]
                st.success(f"**{row[name_col]}**" + (f" — {row[addr_col]}" if addr_col else ""))

                if st.button("Konfirmasi Customer →", key="btn_step2"):
                    st.session_state.selected_customer = row.to_dict()
                    st.session_state.step = 3
                    st.rerun()

# ── STEP 3: Upload Foto & Submit ──────────────────────────────────────────────
if st.session_state.step >= 3:
    with st.expander("③ Upload Dokumentasi & Submit", expanded=(st.session_state.step == 3)):
        cust = st.session_state.selected_customer

        if cust:
            cols = list(cust.keys())
            code_col = next((c for c in cols if "kode" in c.lower()), cols[0])
            name_col = next((c for c in cols if "nama" in c.lower() or "toko" in c.lower()), cols[1] if len(cols) > 1 else cols[0])
            addr_col = next((c for c in cols if "alamat" in c.lower() or "address" in c.lower()), None)

            st.markdown(f"""
| Field | Value |
|---|---|
| Kode Customer | `{cust.get(code_col, '-')}` |
| Nama Toko | {cust.get(name_col, '-')} |
| Alamat | {cust.get(addr_col, '-') if addr_col else '-'} |
| Promotor | {st.session_state.get('promotor_sel', '-')} |
| Program | {st.session_state.get('program_sel', '-')} |
""")

        jumlah = st.number_input("Jumlah Kompensasi (Rp)", min_value=0, step=1000, key="jumlah_input")
        catatan = st.text_area("Catatan Tambahan (opsional)", key="catatan_input")

        foto_bayar = st.file_uploader("📸 Foto Bukti Pembayaran *", type=["jpg", "jpeg", "png"], key="foto_bayar")
        foto_kontrak = st.file_uploader("📄 Foto Kontrak (opsional)", type=["jpg", "jpeg", "png"], key="foto_kontrak")
        foto_ekstra = st.file_uploader("📷 Foto Tambahan (opsional)", type=["jpg", "jpeg", "png"], key="foto_ekstra")

        st.caption("Foto akan dikompres otomatis (max 1920px, JPEG 75%) sebelum diupload.")

        if st.button("✅ Submit Dokumentasi", key="btn_submit", type="primary"):
            if not foto_bayar:
                st.error("Foto bukti pembayaran wajib diupload.")
            elif not cust:
                st.error("Customer belum dipilih.")
            else:
                with st.spinner("Mengupload & menyimpan data..."):
                    ts  = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    uid = str(uuid.uuid4())[:8]
                    promotor_name = st.session_state.get("promotor_sel", "unknown")
                    kode = cust.get(code_col, "unknown")

                    def safe_upload(file, label):
                        if file is None:
                            return ""
                        fname = f"{ts}_{kode}_{promotor_name}_{label}_{uid}.jpg"
                        data  = compress_image(file)
                        return upload_to_drive(data, fname, DRIVE_FOLDER_ID)

                    url_bayar   = safe_upload(foto_bayar,   "pembayaran")
                    url_kontrak = safe_upload(foto_kontrak, "kontrak")
                    url_ekstra  = safe_upload(foto_ekstra,  "ekstra")

                    row = [
                        datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        promotor_name,
                        st.session_state.get("program_sel", ""),
                        kode,
                        cust.get(name_col, ""),
                        cust.get(addr_col, "") if addr_col else "",
                        jumlah,
                        catatan,
                        url_bayar,
                        url_kontrak,
                        url_ekstra,
                    ]
                    append_submission(row)

                st.success("✅ Dokumentasi berhasil disimpan!")
                st.balloons()

                # Reset
                st.session_state.step = 1
                st.session_state.selected_customer = None
                st.cache_data.clear()
                if st.button("📝 Input Baru", key="btn_reset"):
                    st.rerun()
