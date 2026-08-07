import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import datetime
import os

st.set_page_config(page_title="Aplikasi Rekap Pengajuan Transfer WSF", layout="wide", page_icon="📊")

st.title("📊 Aplikasi Rekap Otomatis & Verifikasi Rekening Supplier WSF")
st.markdown("""
Aplikasi ini digunakan untuk **merekap seluruh data pengajuan transfer** dari file Excel WSF secara otomatis per baris lengkap sekaligus **memverifikasi nomor rekening supplier** terhadap **Database Master Rekening Supplier**.
""")

st.sidebar.header("📁 Upload File")
uploaded_files = st.sidebar.file_uploader("Upload File Form Pengajuan Transfer (.xlsx)", type=["xlsx"], accept_multiple_files=True)

EXCLUDE_WORDS = [
    "DIAJUKAN OLEH", "MENGETAHUI", "ALFIAN", "ACHMAD KOHAR", 
    "KANTOR PUSAT", "ADM. TRADING", "MANAJER SBB", "SBB KANTOR PUSAT",
    "TOTAL", "GRAND TOTAL", "FORM PENGAJUAN BANK KELUAR", "SARANA BERKAH BERSAMA", "SUPPLIER"
]

def clean_str(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s.replace(" ", "").replace("-", "")

# ==========================================
# 1. LOAD DATABASE MASTER REKENING
# ==========================================
@st.cache_data
def load_master_database():
    db_filename = None
    for f in os.listdir('.'):
        if 'database' in f.lower() and 'rekening' in f.lower() and f.endswith('.xlsx'):
            db_filename = f
            break
            
    if db_filename and os.path.exists(db_filename):
        try:
            df_db = pd.read_excel(db_filename)
            db_map = {}
            for _, r in df_db.iterrows():
                supp_name = str(r['Supplier']).strip().upper() if pd.notna(r['Supplier']) else ""
                norek_val = clean_str(r['No. Rekening'])
                namarek_val = str(r['Atas Nama']).strip() if pd.notna(r['Atas Nama']) else ""
                
                if supp_name:
                    if supp_name not in db_map:
                        db_map[supp_name] = []
                    db_map[supp_name].append({'norek': norek_val, 'namarek': namarek_val})
            return db_map, db_filename
        except Exception as e:
            return None, str(e)
    return None, "File Database tidak ditemukan di repository/folder utama"

master_db, db_info = load_master_database()

if master_db:
    st.sidebar.success(f"✅ Database Terhubung: {db_info}")
else:
    st.sidebar.error(f"❌ Status DB: {db_info}")

# ==========================================
# 2. PROSES FILE FPT (LOGIKA ASLI HARI/TANGGAL)
# ==========================================
def process_file(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    all_rows = []
    
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
        
        current_hari = ""
        current_tanggal = ""
        
        num_rows = len(df)
        num_cols = len(df.columns)
        
        i = 0
        while i < num_rows:
            row_vals = [str(val).strip() if pd.notna(val) else "" for val in df.iloc[i].values]
            row_str_upper = " ".join(row_vals).upper()
            
            # --- EXTRACT HARI & TANGGAL (MEMERGUKAN LOGIKA VERSI ANDA) ---
            for col_idx in range(num_cols - 1):
                raw_cell = df.iloc[i, col_idx]
                cell_val = str(raw_cell).strip() if pd.notna(raw_cell) else ""
                
                next_cell = df.iloc[i, col_idx + 1] if col_idx + 1 < num_cols else ""
                
                # Format Tanggal jika bertipe datetime/Timestamp
                if isinstance(next_cell, (pd.Timestamp, datetime.datetime)):
                    next_val_str = next_cell.strftime('%d/%m/%Y')
                else:
                    next_val_str = str(next_cell).strip() if pd.notna(next_cell) else ""

                # Pindai HARI
                if "HARI" in cell_val.upper() and ":" in cell_val.upper():
                    current_hari = cell_val.split(":")[-1].strip()
                    if not current_hari and next_val_str:
                        current_hari = next_val_str
                elif cell_val.upper() == "HARI" and next_val_str:
                    current_hari = next_val_str.replace(":", "").strip()
                    
                # Pindai TANGGAL
                if "TANGGAL" in cell_val.upper() and ":" in cell_val.upper():
                    raw_tgl = cell_val.split(":")[-1].strip()
                    if not raw_tgl and next_val_str:
                        raw_tgl = next_val_str
                    current_tanggal = raw_tgl
                elif cell_val.upper() == "TANGGAL" and next_val_str:
                    current_tanggal = next_val_str.replace(":", "").strip()

            # Lewati baris pengesahan / footer / total
            if any(exc in row_str_upper for exc in EXCLUDE_WORDS):
                i += 1
                continue
                
            # Identifikasi Kolom Supplier (Col 1 jika offset, Col 0 jika standar)
            col_supp = 1 if (num_cols > 10 and str(df.iloc[i, 1]).strip() != "") else 0
            supp_val = str(df.iloc[i, col_supp]).strip() if pd.notna(df.iloc[i, col_supp]) else ""
            
            if supp_val and not any(exc in supp_val.upper() for exc in EXCLUDE_WORDS):
                try:
                    bw = str(df.iloc[i, col_supp + 1]).strip() if pd.notna(df.iloc[i, col_supp + 1]) else ""
                    tonase = df.iloc[i, col_supp + 2]
                    harga = df.iloc[i, col_supp + 3]
                    total = df.iloc[i, col_supp + 4]
                    bank = str(df.iloc[i, col_supp + 5]).strip() if pd.notna(df.iloc[i, col_supp + 5]) else ""
                    norek = str(df.iloc[i, col_supp + 6]).strip() if pd.notna(df.iloc[i, col_supp + 6]) else ""
                    nominal = df.iloc[i, col_supp + 7] if col_supp + 7 < num_cols else total
                    ket = str(df.iloc[i, col_supp + 8]).strip() if col_supp + 8 < num_cols and pd.notna(df.iloc[i, col_supp + 8]) else ""
                    
                    # Ambil Nama Rekening dari baris di bawahnya
                    namarek = ""
                    if i + 1 < num_rows:
                        next_row_val = str(df.iloc[i + 1, col_supp + 6]).strip() if pd.notna(df.iloc[i + 1, col_supp + 6]) else ""
                        if next_row_val and not any(exc in next_row_val.upper() for exc in EXCLUDE_WORDS):
                            namarek = next_row_val
                    
                    tonase_num = pd.to_numeric(tonase, errors='coerce')
                    total_num = pd.to_numeric(total, errors='coerce')
                    nominal_num = pd.to_numeric(nominal, errors='coerce')
                    harga_num = pd.to_numeric(harga, errors='coerce')
                    
                    if pd.notna(tonase_num) or pd.notna(total_num):
                        all_rows.append({
                            'SHEET': sheet_name,
                            'HARI': current_hari,
                            'TANGGAL': current_tanggal,
                            'SUPPLIER': supp_val,
                            'BW / UKURAN': bw,
                            'TONASE (KG)': tonase_num,
                            'HARGA (RP)': harga_num,
                            'TOTAL (RP)': total_num,
                            'BANK': bank,
                            'NOMOR REKENING': norek,
                            'NAMA REKENING': namarek,
                            'NOMINAL TRANSFER (RP)': nominal_num if pd.notna(nominal_num) else total_num,
                            'KETERANGAN / STATUS': ket
                        })
                except Exception:
                    pass
            i += 1
            
    return pd.DataFrame(all_rows)

# ==========================================
# 3. TAMPILAN UTAMA & VERIFIKASI
# ==========================================
if uploaded_files:
    dfs = [process_file(f) for f in uploaded_files]
    final_df = pd.concat(dfs, ignore_index=True)
    
    # ------------------------------------------
    # VERIFIKASI DENGAN DATABASE MASTER REKENING
    # ------------------------------------------
    if master_db is not None and not final_df.empty:
        status_verifikasi = []
        referensi_db = []

        for _, row in final_df.iterrows():
            supp_fpt = str(row['SUPPLIER']).strip().upper()
            norek_fpt = clean_str(row['NOMOR REKENING'])
            
            # Cari Nama Supplier di Master DB
            matched_supplier = None
            for db_supp_key in master_db:
                if db_supp_key in supp_fpt or supp_fpt in db_supp_key:
                    matched_supplier = db_supp_key
                    break
            
            if matched_supplier:
                list_rek = master_db[matched_supplier]
                is_match = False
                ref_texts = []
                for rek_data in list_rek:
                    ref_texts.append(f"{rek_data['norek']} ({rek_data['namarek']})")
                    if norek_fpt == rek_data['norek']:
                        is_match = True
                        
                referensi_db.append(" | ".join(ref_texts))
                
                if is_match:
                    status_verifikasi.append("✅ MATCH / SESUAI")
                else:
                    status_verifikasi.append("⚠️ BEDA REKENING!")
            else:
                status_verifikasi.append("❓ SUPPLIER TIDAK TERDAFTAR")
                referensi_db.append("-")

        final_df['STATUS REKENING (VERIFIKASI)'] = status_verifikasi
        final_df['REKENING DATABASE (REFERENSI)'] = referensi_db

    st.success(f"Berhasil merekap total **{len(final_df):,}** baris transaksi!")
    
    # Alert indikator jika ada beda rekening
    if 'STATUS REKENING (VERIFIKASI)' in final_df.columns:
        jml_beda = (final_df['STATUS REKENING (VERIFIKASI)'] == "⚠️ BEDA REKENING!").sum()
        jml_tidak_terdaftar = (final_df['STATUS REKENING (VERIFIKASI)'] == "❓ SUPPLIER TIDAK TERDAFTAR").sum()
        
        c1, c2 = st.columns(2)
        if jml_beda > 0:
            c1.error(f"⚠️ Ditemukan **{jml_beda}** pengajuan dengan **NOMOR REKENING BEDA!**")
        if jml_tidak_terdaftar > 0:
            c2.warning(f"❓ Ditemukan **{jml_tidak_terdaftar}** Supplier belum ada di Database Master.")

    # Filter controls
    col1, col2, col3 = st.columns(3)
    with col1:
        search_supp = st.text_input("Filter Nama Supplier:")
    with col2:
        status_opts = ["Semua"] + list(final_df['KETERANGAN / STATUS'].dropna().unique())
        selected_status = st.selectbox("Filter Status/Keterangan:", status_opts)
    with col3:
        sheet_opts = ["Semua"] + list(final_df['SHEET'].dropna().unique())
        selected_sheet = st.selectbox("Filter Sheet:", sheet_opts)
        
    filtered_df = final_df.copy()
    if search_supp:
        filtered_df = filtered_df[filtered_df['SUPPLIER'].astype(str).str.contains(search_supp, case=False, na=False)]
    if selected_status != "Semua":
        filtered_df = filtered_df[filtered_df['KETERANGAN / STATUS'] == selected_status]
    if selected_sheet != "Semua":
        filtered_df = filtered_df[filtered_df['SHEET'] == selected_sheet]
        
    st.dataframe(filtered_df, use_container_width=True)
    
    # Download Excel Button
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        filtered_df.to_excel(writer, index=False, sheet_name='Rekap & Verifikasi')
    excel_data = output.getvalue()
    
    st.download_button(
        label="📥 Download Hasil Rekapitulasi & Verifikasi (.xlsx)",
        data=excel_data,
        file_name="REKAP_TOTAL_VERIFIKASI_TRANSFER_WSF.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
