import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="Aplikasi Rekap Pengajuan Transfer WSF", layout="wide", page_icon="📊")

st.title("📊 Aplikasi Rekap Otomatis Form Pengajuan Transfer WSF")
st.markdown("Upload file Excel pengajuan transfer (.xlsx) untuk merekap secara otomatis per baris lengkap.")

EXCLUDE_WORDS = [
    "DIAJUKAN OLEH", "MENGETAHUI", "ALFIAN", "ACHMAD KOHAR", 
    "KANTOR PUSAT", "ADM. TRADING", "MANAJER SBB", "SBB KANTOR PUSAT",
    "TOTAL", "GRAND TOTAL", "FORM PENGAJUAN BANK KELUAR", "SARANA BERKAH BERSAMA", "SUPPLIER"
]

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
            
            # Ambil Hari & Tanggal Header
            for col_idx in range(num_cols):
                cell_val = str(df.iloc[i, col_idx]).strip()
                if "HARI" in cell_val.upper() and ":" in cell_val:
                    current_hari = cell_val.split(":")[-1].strip()
                if "TANGGAL" in cell_val.upper() and ":" in cell_val:
                    current_tanggal = cell_val.split(":")[-1].strip()

            # Filter kata pengesahan / tanda tangan
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
                    
                    # Mengambil Nama Rekening dari baris di bawahnya
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

uploaded_files = st.file_uploader("Upload File Form Pengajuan Transfer (.xlsx)", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    dfs = [process_file(f) for f in uploaded_files]
    final_df = pd.concat(dfs, ignore_index=True)
    final_df.insert(0, 'NO', range(1, 1 + len(final_df)))
    
    st.success(f"Berhasil merekap **{len(final_df):,}** baris transaksi secara otomatis!")
    st.dataframe(final_df, use_container_width=True)
    
    # Fitur Download File Hasil Rekap
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        final_df.to_excel(writer, index=False, sheet_name='Rekap Total')
    
    st.download_button(
        label="📥 Download Hasil Rekapitulasi (.xlsx)",
        data=output.getvalue(),
        file_name="HASIL_REKAP_PENGAJUAN_TRANSFER.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
