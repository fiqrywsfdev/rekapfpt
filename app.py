import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import os

st.set_page_config(page_title="Aplikasi Rekap & Verifikasi Transfer WSF", layout="wide", page_icon="📊")

st.title("📊 Aplikasi Rekap Otomatis & Verifikasi Rekening Supplier WSF")
st.markdown("""
Aplikasi ini secara otomatis **merekap seluruh data pengajuan transfer (FPT)** dan melakukan **verifikasi/validasi nomor rekening** pengajuan terhadap **Database Master Rekening Supplier**.
""")

st.sidebar.header("📁 Upload File")
uploaded_fpt = st.sidebar.file_uploader("Upload File FPT Pengajuan (.xlsx)", type=["xlsx"], accept_multiple_files=True)

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

# 1. Baca Database Master secara Otomatis dari folder aplikasi (Aman dari perbedaan huruf besar/kecil)
@st.cache_data
def load_master_database():
    db_filename = None
    # Pencarian file fleksibel
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
    return None, "File tidak ditemukan"

master_db, db_info = load_master_database()

if master_db:
    st.sidebar.success(f"✅ Database Master Terhubung! ({db_info})")
else:
    st.sidebar.error(f"❌ Database belum terbaca. Status: {db_info}")

def process_file(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    all_rows = []
    
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
        current_hari, current_tanggal = "", ""
        num_rows, num_cols = len(df), len(df.columns)
        
        i = 0
        while i < num_rows:
            row_vals = [str(val).strip() if pd.notna(val) else "" for val in df.iloc[i].values]
            row_str_upper = " ".join(row_vals).upper()
            
            for col_idx in range(num_cols):
                cell_val = str(df.iloc[i, col_idx]).strip()
                if "HARI" in cell_val.upper() and ":" in cell_val:
                    current_hari = cell_val.split(":")[-1].strip()
                if "TANGGAL" in cell_val.upper() and ":" in cell_val:
                    current_tanggal = cell_val.split(":")[-1].strip()

            if any(exc in row_str_upper for exc in ["DIAJUKAN OLEH", "MENGETAHUI", "ALFIAN", "ACHMAD KOHAR", "ADM. TRADING", "MANAJER SBB"]):
                i += 1
                continue
                
            col_supp = 1 if (num_cols > 10 and df.iloc[i, 1] is not None and str(df.iloc[i, 1]).strip() != "") else 0
            supp_val = str(df.iloc[i, col_supp]).strip() if pd.notna(df.iloc[i, col_supp]) else ""
            
            if supp_val and not any(exc in supp_val.upper() for exc in EXCLUDE_WORDS):
                try:
                    bw = str(df.iloc[i, col_supp + 1]).strip() if pd.notna(df.iloc[i, col_supp + 1]) else ""
                    tonase = pd.to_numeric(df.iloc[i, col_supp + 2], errors='coerce')
                    harga = pd.to_numeric(df.iloc[i, col_supp + 3], errors='coerce')
                    total = pd.to_numeric(df.iloc[i, col_supp + 4], errors='coerce')
                    bank = str(df.iloc[i, col_supp + 5]).strip() if pd.notna(df.iloc[i, col_supp + 5]) else ""
                    norek = str(df.iloc[i, col_supp + 6]).strip() if pd.notna(df.iloc[i, col_supp + 6]) else ""
                    nominal = pd.to_numeric(df.iloc[i, col_supp + 7], errors='coerce') if col_supp + 7 < num_cols else total
                    ket = str(df.iloc[i, col_supp + 8]).strip() if col_supp + 8 < num_cols and pd.notna(df.iloc[i, col_supp + 8]) else ""
                    
                    namarek = ""
                    if i + 1 < num_rows:
                        next_row_val = str(df.iloc[i + 1, col_supp + 6]).strip() if pd.notna(df.iloc[i + 1, col_supp + 6]) else ""
                        if next_row_val and not any(exc in next_row_val.upper() for exc in EXCLUDE_WORDS):
                            namarek = next_row_val
                    
                    if pd.notna(tonase) or pd.notna(total):
                        all_rows.append({
                            'SHEET': sheet_name,
                            'HARI': current_hari,
                            'TANGGAL': current_tanggal,
                            'SUPPLIER': supp_val,
                            'BW / UKURAN': bw,
                            'TONASE (KG)': tonase,
                            'HARGA (RP)': harga,
                            'TOTAL (RP)': total,
                            'BANK': bank,
                            'NOMOR REKENING': norek,
                            'NAMA REKENING': namarek,
                            'NOMINAL TRANSFER (RP)': nominal if pd.notna(nominal) else total,
                            'KETERANGAN / STATUS': ket
                        })
                except Exception:
                    pass
            i += 1
            
    return pd.DataFrame(all_rows)

if uploaded_fpt:
    dfs = [process_file(f) for f in uploaded_fpt]
    final_df = pd.concat(dfs, ignore_index=True)
    final_df.insert(0, 'NO', range(1, 1 + len(final_df)))
    
    if master_db is not None:
        status_verifikasi = []
        referensi_db = []

        for _, row in final_df.iterrows():
            supp_fpt = str(row['SUPPLIER']).strip().upper()
            norek_fpt = clean_str(row['NOMOR REKENING'])
            
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

    st.success(f"Berhasil merekap **{len(final_df):,}** baris transaksi!")
    
    if 'STATUS REKENING (VERIFIKASI)' in final_df.columns:
        jml_beda = (final_df['STATUS REKENING (VERIFIKASI)'] == "⚠️ BEDA REKENING!").sum()
        jml_tidak_terdaftar = (final_df['STATUS REKENING (VERIFIKASI)'] == "❓ SUPPLIER TIDAK TERDAFTAR").sum()
        
        col1, col2 = st.columns(2)
        if jml_beda > 0:
            col1.error(f"⚠️ Ditemukan **{jml_beda}** pengajuan dengan REKENING BEDA/SALAH!")
        if jml_tidak_terdaftar > 0:
            col2.warning(f"❓ Terdapat **{jml_tidak_terdaftar}** Supplier yang belum ada di Database!")
    
    st.dataframe(final_df, use_container_width=True)
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        final_df.to_excel(writer, index=False, sheet_name='Rekap Total & Verifikasi')
    
    st.download_button(
        label="📥 Download Hasil Rekapitulasi & Verifikasi (.xlsx)",
        data=output.getvalue(),
        file_name="REKAP_TOTAL_VERIFIKASI_TRANSFER_WSF.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
