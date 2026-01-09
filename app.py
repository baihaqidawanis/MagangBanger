import streamlit as st
import pandas as pd
import os

# ==============================================================================
# IMPORT FUNGSI ETL DARI main.py
# ==============================================================================
import main

# ==============================================================================
# CONFIG & PAGE SETUP
# ==============================================================================
st.set_page_config(page_title="Dashboard ETL Pendapatan", layout="centered", page_icon="📊")

st.title("Generate Laporan Pendapatan")
st.caption("Automated ETL Pipeline untuk Monitoring Laporan Excel")

# ==============================================================================
# HELPER FUNCTIONS
# ==============================================================================
def save_uploaded_file(uploaded_file, folder_path, filename):
    """Simpan file upload dari Streamlit ke folder lokal"""
    if uploaded_file is not None:
        file_path = os.path.join(folder_path, filename)
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        return file_path
    return None

def process_all_data(config):
    """Wrapper untuk ETL pipeline dari main.py dengan config Streamlit"""
    main.CONFIG.update(config)
    
    try:
        m1, m2, raw, kode_nama_map = main.extract_data()
        final = main.transform_data(m1, m2, raw, kode_nama_map)
        main.load_data(final)
        return config["OUTPUT_FILE"]
    except Exception as e:
        st.error(f"❌ Error ETL: {e}")
        return None

# ==============================================================================
# MAIN UI
# ==============================================================================

# --- SIDEBAR: FILE UPLOADS ---
with st.sidebar:
    st.header("📂 Upload File")
    template_file = st.file_uploader("Template Excel", type=["xlsx"], help="File template BANGER")
    input_file = st.file_uploader("Lampiran Pendapatan", type=["xlsx", "xlsb"], help="File lampiran pendapatan bulanan")
    rekap_file = st.file_uploader("Master Rekap", type=["xlsx"], help="File validasi kode produk")
    pelanggan_file = st.file_uploader("Data Pelanggan", type=["xlsx"], help="File data pelanggan BANGER")
    opt_file = st.file_uploader("Data OPT", type=["xlsx"], help="File data OPT BANGER")

    st.divider()
    st.header("⚙️ Konfigurasi")
    
    # List 12 bulan lengkap
    BULAN_LIST = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", 
                  "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
    
    bulan_lalu = st.selectbox("Bulan Lalu (Template)", BULAN_LIST, index=9, 
                              help="Bulan dari file template (misal: Oktober untuk template 10 BANGER OKTOBER.xlsx)")
    bulan_ini = st.selectbox("Bulan Ini (Data Raw)", BULAN_LIST, index=10, 
                             help="Bulan dari file data raw (misal: November untuk file 11 Lampiran Pendapatan November.xlsx)")
    output_filename = st.text_input("Nama File Output", "Laporan_Final")

# --- MAIN AREA ---
st.markdown("### Proses ETL")
st.info("Upload semua file yang dibutuhkan di sidebar, lalu klik tombol Generate")

process_btn = st.button("GENERATE LAPORAN", type="primary", use_container_width=True)

if process_btn:
    if not all([template_file, input_file, rekap_file, pelanggan_file, opt_file]):
        st.error("⚠️ Mohon lengkapi semua file upload di sidebar!")
    else:
        # Setup folder temporary
        folder = "temp_upload"
        os.makedirs(folder, exist_ok=True)
        
        # Config untuk ETL
        config = {
            "TEMPLATE_FILE": save_uploaded_file(template_file, folder, "template.xlsx"),
            "INPUT_FILE": save_uploaded_file(input_file, folder, "input.xlsx"),
            "FILE_REKAP": save_uploaded_file(rekap_file, folder, "rekap.xlsx"),
            "FILE_PELANGGAN": save_uploaded_file(pelanggan_file, folder, "pelanggan.xlsx"),
            "FILE_OPT": save_uploaded_file(opt_file, folder, "opt.xlsx"),
            "OUTPUT_FILE": os.path.join(folder, output_filename + ".xlsx"),
            "BULAN_LALU": bulan_lalu,
            "BULAN_INI": bulan_ini,
            "DASHBOARD_HEADER_ROW": 2, 
            "DASHBOARD_DATA_START": 3
        }
        
        # Jalankan ETL
        with st.spinner("⏳ Sedang memproses ETL Pipeline..."):
            output_path = process_all_data(config)
        
        if output_path and os.path.exists(output_path):
            st.success("✅ Laporan berhasil dibuat!")
            
            # Download button
            with open(output_path, "rb") as f:
                st.download_button(
                    " Download Laporan Excel",
                    f,
                    file_name=output_filename + ".xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            st.balloons()
        else:
            st.error("❌ Gagal membuat laporan. Silakan cek file input dan coba lagi.")

# --- FOOTER ---
st.markdown("---")

