import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import datetime

st.set_page_config(page_title="Aplikasi Rekap Pengajuan Transfer WSF", layout="wide", page_icon="📊")

st.title("📊 Aplikasi Rekap Otomatis Form Pengajuan Transfer WSF")
st.markdown("""
Aplikasi ini digunakan untuk **merekap seluruh data pengajuan transfer** dari file Excel WSF secara otomatis per baris lengkap.
Hasil rekap memuat: **Hari, Tanggal, Supplier, BW/Ukuran, Tonase (Kg), Harga (Rp), Total (Rp), Bank, Nomor Rekening, Nama Rekening, Nominal Transfer (Rp), dan Keterangan/Status (DP/Pelunasan/Frozen/PUJ/WBS)**.

*Kata-kata pengesahan seperti "Diajukan oleh", "Mengetahui", "Alfian", "Achmad Kohar", "Adm. Trading", dll. otomatis dibersihkan.*
""")

uploaded_files = st.file_uploader("Upload File Form Pengajuan Transfer (.xlsx)", type=["xlsx"], accept_multiple_files=True)

EXCLUDE_WORDS = [
    "DIAJUKAN OLEH", "MENGETAHUI", "ALFIAN", "ACHMAD KOHAR", 
    "KANTOR PUSAT", "ADM. TRADING", "MANAJER SBB", "SBB KANTOR PUSAT",
    "TOTAL", "GRAND TOTAL", "FORM PENGAJUAN BANK KELUAR", "SARANA BERKAH BERSAMA"
]

def process_file(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    all_rows = []
    
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
        
        current_hari = ""
        current_tanggal = None
        
        num_rows = len(df)
        num_cols = len(df.columns)
        
        # Check alignment shift
        # In 'RPA TF', data starts at col 1. In 'RPA RK', data starts at col 0.
        i = 0
        while i < num_rows:
            row_vals = [str(val).strip() if pd.notna(val) else "" for val in df.iloc[i].values]
            row_str_upper = " ".join(row_vals).upper()
            
            # Extract Hari & Tanggal header
            for col_idx in range(num_cols - 1):
                cell_val = str(df.iloc[i, col_idx]).strip()
                next_val = str(df.iloc[i, col_idx + 1]).strip() if col_idx + 1 < num_cols else ""
                
                if "HARI" in cell_val.upper() and ":" in cell_val.upper():
                    current_hari = cell_val.split(":")[-1].strip()
                    if not current_hari and next_val:
                        current_hari = next_val
                elif cell_val.upper() == "HARI" and next_val:
                    current_hari = next_val.replace(":", "").strip()
                    
                if "TANGGAL" in cell_val.upper() and ":" in cell_val.upper():
                    raw_tgl = cell_val.split(":")[-1].strip()
                    if not raw_tgl and next_val:
                        raw_tgl = next_val
                    current_tanggal = raw_tgl
                elif cell_val.upper() == "TANGGAL" and next_val:
                    current_tanggal = next_val.replace(":", "").strip()

            # Identify data row start
            # Check if any excluded words match
            if any(exc in row_str_upper for exc in EXCLUDE_WORDS):
                i += 1
                continue
                
            # Find candidate data row
            # Case 1: RPA TF offset (col 1 is supplier)
            col_supp = 1 if (num_cols > 10 and str(df.iloc[i, 1]).strip() != "") else 0
            
            supp_val = str(df.iloc[i, col_supp]).strip() if pd.notna(df.iloc[i, col_supp]) else ""
            
            if supp_val and supp_val.upper() not in EXCLUDE_WORDS and not supp_val.upper().startswith("SUPPLIER"):
                # Check if next columns have numeric values (tonase / harga / total)
                try:
                    bw = str(df.iloc[i, col_supp + 1]).strip() if pd.notna(df.iloc[i, col_supp + 1]) else ""
                    tonase = df.iloc[i, col_supp + 2]
                    harga = df.iloc[i, col_supp + 3]
                    total = df.iloc[i, col_supp + 4]
                    bank = str(df.iloc[i, col_supp + 5]).strip() if pd.notna(df.iloc[i, col_supp + 5]) else ""
                    norek = str(df.iloc[i, col_supp + 6]).strip() if pd.notna(df.iloc[i, col_supp + 6]) else ""
                    nominal = df.iloc[i, col_supp + 7] if col_supp + 7 < num_cols else total
                    ket = str(df.iloc[i, col_supp + 8]).strip() if col_supp + 8 < num_cols and pd.notna(df.iloc[i, col_supp + 8]) else ""
                    
                    # Next row usually contains Nama Rekening
                    namarek = ""
                    if i + 1 < num_rows:
                        next_row_val = str(df.iloc[i + 1, col_supp + 6]).strip() if pd.notna(df.iloc[i + 1, col_supp + 6]) else ""
                        if next_row_val and next_row_val.upper() not in EXCLUDE_WORDS:
                            namarek = next_row_val
                    
                    # Ensure tonase and total are numbers
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
                        i += 1 # skip next sub-row if processed
                except Exception as e:
                    pass
            i += 1
            
    return pd.DataFrame(all_rows)

if uploaded_files:
    dfs = []
    for f in uploaded_files:
        df_res = process_file(f)
        dfs.append(df_res)
        
    final_df = pd.concat(dfs, ignore_index=True)
    
    st.success(f"Berhasil merekap total **{len(final_df):,}** baris transaksi!")
    
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
        filtered_df.to_excel(writer, index=False, sheet_name='Rekap Total')
    excel_data = output.getvalue()
    
    st.download_button(
        label="📥 Download Hasil Rekapitulasi (.xlsx)",
        data=excel_data,
        file_name="REKAP_TOTAL_PENGAJUAN_TRANSFER_WSF.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
