import streamlit as st
import pandas as pd
import sqlite3
import io
import time
import os
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- KONFIGURASI FOLDER & DATABASE ---
DB_NAME = "lpki.db"
UPLOAD_DIR = "uploads"

if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)

def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS profil_lembaga (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nama_lembaga TEXT,
            nspp TEXT,
            alamat TEXT,
            pimpinan TEXT,
            akta_notaris TEXT,
            sk_ijin TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS satuan_pendidikan (
            npsn TEXT PRIMARY KEY,
            nama_satuan TEXT NOT NULL,
            jenjang TEXT,
            status TEXT,
            kepala_sekolah TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS pengajar (
            id_pengajar TEXT PRIMARY KEY,
            nama TEXT NOT NULL,
            bidang_pengajaran TEXT,
            no_hp TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS santri (
            nis TEXT PRIMARY KEY,
            nama TEXT NOT NULL,
            kelas TEXT,
            asrama_kamar TEXT,
            wali_murid TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS sarpras (
            kode_barang TEXT PRIMARY KEY,
            nama_barang TEXT NOT NULL,
            kondisi TEXT,
            jumlah INTEGER
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS galeri_foto (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kategori TEXT,
            nama_file TEXT,
            keterangan TEXT
        )
    ''')
    conn.commit()
    conn.close()

init_db()

def run_query(query, params=()):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute(query, params)
    conn.commit()
    conn.close()

def load_data(table_name):
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
    conn.close()
    return df

# --- FUNGSI GENERATE WORD ---
def generate_word_portfolio():
    doc = Document()

    df_profil = load_data("profil_lembaga")
    nama_lembaga = df_profil['nama_lembaga'].values[0] if not df_profil.empty else "LEMBAGA PENDIDIKAN KEAGAMAAN ISLAM"
    
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = title_p.add_run(f"PORTOFOLIO & PROFIL LEMBAGA\n{nama_lembaga.upper()}")
    run_title.bold = True
    run_title.font.size = Pt(16)
    run_title.font.color.rgb = RGBColor(0, 102, 204)

    doc.add_paragraph("-------------------------------------------------------------------------------------------------------------------")

    # 1. Profil
    doc.add_heading("1. Profil & Legalitas Lembaga", level=1)
    if not df_profil.empty:
        p = df_profil.iloc[0]
        doc.add_paragraph(f"Nama Lembaga: {p['nama_lembaga']}")
        doc.add_paragraph(f"NSPP: {p['nspp'] if p['nspp'] else '-'}")
        doc.add_paragraph(f"Pimpinan / Ketua: {p['pimpinan'] if p['pimpinan'] else '-'}")
        doc.add_paragraph(f"Alamat: {p['alamat'] if p['alamat'] else '-'}")
        doc.add_paragraph(f"Akta Notaris: {p['akta_notaris'] if p['akta_notaris'] else '-'}")
        doc.add_paragraph(f"SK Izin Operasional: {p['sk_ijin'] if p['sk_ijin'] else '-'}")
    else:
        doc.add_paragraph("Belum ada data profil lembaga.")

    # 2. Satuan Pendidikan
    doc.add_heading("2. Satuan Pendidikan / Unit Keagamaan", level=1)
    df_satuan = load_data("satuan_pendidikan")
    if not df_satuan.empty:
        table = doc.add_table(rows=1, cols=5)
        table.style = 'Table Grid'
        hdr = table.rows[0].cells
        hdr[0].text, hdr[1].text, hdr[2].text, hdr[3].text, hdr[4].text = 'NPSN', 'Nama Satuan', 'Jenjang', 'Status', 'Kepala Sekolah'
        for _, row in df_satuan.iterrows():
            r = table.add_row().cells
            r[0].text, r[1].text, r[2].text, r[3].text, r[4].text = str(row['npsn']), str(row['nama_satuan']), str(row['jenjang']), str(row['status']), str(row['kepala_sekolah'])
    else:
        doc.add_paragraph("Belum ada data satuan pendidikan.")

    # 3. Pengajar
    doc.add_heading("3. Data Pengajar / Ustadz", level=1)
    df_pengajar = load_data("pengajar")
    if not df_pengajar.empty:
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        hdr = table.rows[0].cells
        hdr[0].text, hdr[1].text, hdr[2].text, hdr[3].text = 'ID Pengajar', 'Nama Lengkap', 'Bidang Pengajaran', 'No. HP'
        for _, row in df_pengajar.iterrows():
            r = table.add_row().cells
            r[0].text, r[1].text, r[2].text, r[3].text = str(row['id_pengajar']), str(row['nama']), str(row['bidang_pengajaran']), str(row['no_hp'])
    else:
        doc.add_paragraph("Belum ada data pengajar.")

    # 4. Santri
    doc.add_heading("4. Data Santri", level=1)
    df_santri = load_data("santri")
    if not df_santri.empty:
        doc.add_paragraph(f"Total Santri Terdaftar: {len(df_santri)} Orang")
        table = doc.add_table(rows=1, cols=5)
        table.style = 'Table Grid'
        hdr = table.rows[0].cells
        hdr[0].text, hdr[1].text, hdr[2].text, hdr[3].text, hdr[4].text = 'NIS', 'Nama Santri', 'Kelas', 'Kamar/Asrama', 'Wali Murid'
        for _, row in df_santri.iterrows():
            r = table.add_row().cells
            r[0].text, r[1].text, r[2].text, r[3].text, r[4].text = str(row['nis']), str(row['nama']), str(row['kelas']), str(row['asrama_kamar']), str(row['wali_murid'])
    else:
        doc.add_paragraph("Belum ada data santri.")

    # 5. Sarpras
    doc.add_heading("5. Inventaris Sarana & Prasarana", level=1)
    df_sarpras = load_data("sarpras")
    if not df_sarpras.empty:
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        hdr = table.rows[0].cells
        hdr[0].text, hdr[1].text, hdr[2].text, hdr[3].text = 'Kode Barang', 'Nama Barang', 'Kondisi', 'Jumlah'
        for _, row in df_sarpras.iterrows():
            r = table.add_row().cells
            r[0].text, r[1].text, r[2].text, r[3].text = str(row['kode_barang']), str(row['nama_barang']), str(row['kondisi']), str(row['jumlah'])
    else:
        doc.add_paragraph("Belum ada data sarpras.")

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- TAMPILAN APLIKASI STREAMLIT & TEMA KUSTOM ---
st.set_page_config(page_title="SIM Lembaga Pendidikan Keagamaan Islam", page_icon="🕌", layout="wide")

st.markdown("""
    <style>
    /* Latar belakang halaman utama */
    .stApp {
        background-color: #f0f7f4;
        color: #000000 !important;
    }
    
    /* Memaksa teks di area utama berwarna HITAM */
    .stApp p, .stApp span, .stApp label, .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp div {
        color: #000000 !important;
    }

    /* Kotak input berwarna abu-abu terang */
    .stTextInput input, .stTextArea textarea, .stSelectbox select, .stNumberInput input {
        color: #000000 !important;
        background-color: #e2e8f0 !important;
        border: 1px solid #cbd5e1 !important;
    }

    /* Mengubah latar belakang sidebar menjadi hijau tua */
    [data-testid="stSidebar"] {
        background-color: #1b4d3e !important;
    }
    
    /* Menjaga teks di dalam sidebar tetap PUTIH */
    [data-testid="stSidebar"] p, 
    [data-testid="stSidebar"] span, 
    [data-testid="stSidebar"] label, 
    [data-testid="stSidebar"] h1, 
    [data-testid="stSidebar"] h2, 
    [data-testid="stSidebar"] h3 {
        color: #ffffff !important;
    }

    /* Mengubah SEMUA TOMBOL (termasuk tombol Simpan) menjadi WARNA BIRU */
    .stButton > button {
        background-color: #2563eb !important;
        color: #ffffff !important;
        border-radius: 8px;
        border: none !important;
    }
    .stButton > button:hover {
        background-color: #1d4ed8 !important;
    }
    
    /* Kartu / Kotak Form Input */
    div.stForm {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 10px;
        border: 1px solid #cbd5e1;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    </style>
""", unsafe_allow_html=True)

st.sidebar.title("🕌 SIM LPKI")
menu = st.sidebar.selectbox("Pilih Menu", [
    "Dashboard", 
    "Profil Lembaga", 
    "Satuan Pendidikan", 
    "Data Pengajar", 
    "Data Santri", 
    "Data Sarpras",
    "📸 Galeri & Dokumentasi",
    "📄 Cetak Portofolio Word"
])

# 1. DASHBOARD
if menu == "Dashboard":
    conn = sqlite3.connect(DB_NAME)
    df_profil = pd.read_sql_query("SELECT nama_lembaga FROM profil_lembaga ORDER BY id DESC LIMIT 1", conn)
    conn.close()
    
    nama_lembaga = df_profil['nama_lembaga'].values[0] if not df_profil.empty else "Lembaga Pendidikan Keagamaan Islam"
    
    st.header("📊 Dashboard Rekapitulasi")
    st.subheader(f"🕌 {nama_lembaga}")
    st.divider()

    df_satuan = load_data("satuan_pendidikan")
    df_pengajar = load_data("pengajar")
    df_santri = load_data("santri")
    df_sarpras = load_data("sarpras")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Satuan Pendidikan", len(df_satuan))
    col2.metric("Total Pengajar", len(df_pengajar))
    col3.metric("Total Santri", len(df_santri))
    col4.metric("Total Sarpras", len(df_sarpras))

    st.success("✅ Aplikasi siap digunakan dan data tersimpan dengan aman!")

# 2. PROFIL LEMBAGA
elif menu == "Profil Lembaga":
    st.header("🕌 Profil & Legalitas Lembaga Pendidikan Keagamaan Islam")
    conn = sqlite3.connect(DB_NAME)
    df_profil = pd.read_sql_query("SELECT * FROM profil_lembaga ORDER BY id DESC LIMIT 1", conn)
    conn.close()
    
    with st.form("form_profil"):
        st.subheader("Perbarui Informasi Lembaga")
        def_nama = df_profil['nama_lembaga'].values[0] if not df_profil.empty else ""
        def_nspp = df_profil['nspp'].values[0] if not df_profil.empty else ""
        def_alamat = df_profil['alamat'].values[0] if not df_profil.empty else ""
        def_pimpinan = df_profil['pimpinan'].values[0] if not df_profil.empty else ""
        def_akta = df_profil['akta_notaris'].values[0] if not df_profil.empty else ""
        def_sk = df_profil['sk_ijin'].values[0] if not df_profil.empty else ""
        
        nama_lembaga = st.text_input("Nama Lembaga Pendidikan", value=def_nama)
        nspp = st.text_input("Nomor Statistik / NSPP / NSQ", value=def_nspp)
        alamat = st.text_area("Alamat Lengkap", value=def_alamat)
        pimpinan = st.text_input("Nama Pimpinan / Ketua", value=def_pimpinan)
        akta_notaris = st.text_input("Nomor & Tanggal Akta Notaris", value=def_akta, placeholder="Contoh: Akta No. 12 / 15 Januari 2020")
        sk_ijin = st.text_input("Nomor & Tanggal SK Ijin Operasional", value=def_sk, placeholder="Contoh: SK.123/Kemenag/2021 / 10 Mei 2021")
        
        if st.form_submit_button("Simpan / Perbarui Profil"):
            if nama_lembaga:
                run_query("DELETE FROM profil_lembaga")
                run_query("INSERT INTO profil_lembaga (nama_lembaga, nspp, alamat, pimpinan, akta_notaris, sk_ijin) VALUES (?, ?, ?, ?, ?, ?)", 
                          (nama_lembaga, nspp, alamat, pimpinan, akta_notaris, sk_ijin))
                st.success("✅ Profil lembaga berhasil diperbarui!")
                st.rerun()
            else:
                st.warning("Nama Lembaga Pendidikan wajib diisi.")

    if not df_profil.empty:
        st.divider()
        st.subheader("📄 Informasi Resmi Lembaga")
        p = df_profil.iloc[0]
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"**Nama Lembaga:**\n{p['nama_lembaga']}")
            st.markdown(f"**Nomor Statistik:**\n{p['nspp'] if p['nspp'] else '-'}")
            st.markdown(f"**Pimpinan / Ketua:**\n{p['pimpinan'] if p['pimpinan'] else '-'}")
        with col2:
            st.markdown(f"**Alamat:**\n{p['alamat'] if p['alamat'] else '-'}")
            st.markdown(f"**Akta Notaris:**\n{p['akta_notaris'] if p['akta_notaris'] else '-'}")
            st.markdown(f"**SK Ijin Operasional:**\n{p['sk_ijin'] if p['sk_ijin'] else '-'}")

# 3. SATUAN PENDIDIKAN
elif menu == "Satuan Pendidikan":
    st.header("🏫 Manajemen Satuan Pendidikan / Unit Keagamaan")
    tab1, tab2 = st.tabs(["➕ Input Manual", "📁 Impor dari Excel"])

    with tab1:
        with st.form("form_satuan", clear_on_submit=True):
            npsn = st.text_input("NPSN / Nomor Induk Satuan (Kosongkan jika non-formal)")
            nama_satuan = st.text_input("Nama Satuan (Contoh: TPQ Al-Ikhlas / Madin Darul Ilmi)")
            jenjang = st.selectbox("Jenjang Pendidikan", ["Madrasah Diniyah Takmiliyah (MDT)", "Taman Pendidikan Al-Qur'an (TPQ)", "Pesantren / LPTQ", "Majelis Taklim", "Lainnya"])
            status = st.selectbox("Status Lembaga", ["Swasta", "Negeri"])
            kepala_sekolah = st.text_input("Nama Kepala / Pimpinan Unit")
            
            if st.form_submit_button("Simpan Satuan Pendidikan"):
                if nama_satuan:
                    if not npsn:
                        npsn = f"NON-{int(time.time())}"
                    try:
                        run_query("INSERT INTO satuan_pendidikan VALUES (?, ?, ?, ?, ?)", (npsn, nama_satuan, jenjang, status, kepala_sekolah))
                        st.success(f"Satuan pendidikan **{nama_satuan}** berhasil disimpan!")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error(f"NPSN '{npsn}' sudah terdaftar.")

    with tab2:
        st.subheader("Unggah Data Satuan Pendidikan via Excel (.xlsx)")
        df_template_satuan = pd.DataFrame(columns=["npsn", "nama_satuan", "jenjang", "status", "kepala_sekolah"])
        output_satuan = io.BytesIO()
        with pd.ExcelWriter(output_satuan, engine='openpyxl') as writer:
            df_template_satuan.to_excel(writer, sheet_name='Satuan Pendidikan', index=False)
        st.download_button("📥 Unduh Template Excel", data=output_satuan.getvalue(), file_name="template_satuan_pendidikan.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        st.divider()
        uploaded_file_satuan = st.file_uploader("Pilih file Excel (.xlsx)", type=["xlsx"], key="up_satuan")
        if uploaded_file_satuan is not None:
            try:
                df_upload_satuan = pd.read_excel(uploaded_file_satuan, engine='openpyxl')
                if st.button("Proses Impor Satuan Pendidikan"):
                    conn = sqlite3.connect(DB_NAME)
                    df_upload_satuan.to_sql('satuan_pendidikan', conn, if_exists='append', index=False)
                    conn.close()
                    st.success("✅ Data satuan pendidikan berhasil diimpor!")
                    st.rerun()
            except Exception as e:
                st.error(f"Terjadi kesalahan: {e}")

    st.divider()
    st.subheader("📋 Daftar Unit Satuan Pendidikan")
    st.dataframe(load_data("satuan_pendidikan"), use_container_width=True)

# 4. DATA PENGAJAR
elif menu == "Data Pengajar":
    st.header("👨‍🏫 Manajemen Data Pengajar / Ustadz")
    tab1, tab2, tab3 = st.tabs(["➕ Input Manual", "📁 Impor dari Excel", "✏️ Edit / Hapus Data"])
    
    with tab1:
        with st.form("form_pengajar", clear_on_submit=True):
            id_pengajar = st.text_input("ID Pengajar")
            nama = st.text_input("Nama Lengkap & Gelar")
            bidang = st.text_input("Bidang Pengajaran / Kitab")
            hp = st.text_input("No. HP / WhatsApp")
            if st.form_submit_button("Simpan Data Pengajar"):
                if id_pengajar and nama:
                    try:
                        run_query("INSERT INTO pengajar VALUES (?, ?, ?, ?)", (id_pengajar, nama, bidang, hp))
                        st.success(f"Data pengajar **{nama}** berhasil disimpan!")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error(f"ID Pengajar '{id_pengajar}' sudah terdaftar.")

    with tab2:
        st.subheader("Unggah Data Pengajar Masal via Excel (.xlsx)")
        df_template_pengajar = pd.DataFrame(columns=["id_pengajar", "nama", "bidang_pengajaran", "no_hp"])
        output_pengajar = io.BytesIO()
        with pd.ExcelWriter(output_pengajar, engine='openpyxl') as writer:
            df_template_pengajar.to_excel(writer, sheet_name='Data Pengajar', index=False)
        st.download_button("📥 Unduh Template Excel", data=output_pengajar.getvalue(), file_name="template_data_pengajar.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        st.divider()
        uploaded_file_pengajar = st.file_uploader("Pilih file Excel (.xlsx) data pengajar", type=["xlsx"], key="up_pengajar")
        if uploaded_file_pengajar is not None:
            try:
                df_upload_pengajar = pd.read_excel(uploaded_file_pengajar, engine='openpyxl')
                if st.button("Proses Impor Pengajar"):
                    conn = sqlite3.connect(DB_NAME)
                    df_upload_pengajar.to_sql('pengajar', conn, if_exists='append', index=False)
                    conn.close()
                    st.success("✅ Data pengajar berhasil diimpor!")
                    st.rerun()
            except Exception as e:
                st.error(f"Terjadi kesalahan: {e}")

    with tab3:
        st.subheader("Pilih Pengajar yang Ingin Diedit atau Dihapus")
        df_pengajar_s = load_data("pengajar")
        if not df_pengajar_s.empty:
            pilih_pengajar = st.selectbox("Pilih ID & Nama Pengajar", options=df_pengajar_s['id_pengajar'].tolist(), format_func=lambda x: f"{x} - {df_pengajar_s.loc[df_pengajar_s['id_pengajar'] == x, 'nama'].values[0]}")
            
            data_terpilih = df_pengajar_s[df_pengajar_s['id_pengajar'] == pilih_pengajar].iloc[0]
            
            with st.form("form_edit_pengajar"):
                new_nama = st.text_input("Nama Lengkap & Gelar", value=data_terpilih['nama'])
                new_bidang = st.text_input("Bidang Pengajaran", value=data_terpilih['bidang_pengajaran'])
                new_hp = st.text_input("No. HP / WhatsApp", value=data_terpilih['no_hp'])
                
                col_e1, col_e2 = st.columns(2)
                update_btn = col_e1.form_submit_button("💾 Perbarui Data")
                delete_btn = col_e2.form_submit_button("🗑️ Hapus Data", type="primary")
                
                if update_btn:
                    run_query("UPDATE pengajar SET nama = ?, bidang_pengajaran = ?, no_hp = ? WHERE id_pengajar = ?", (new_nama, new_bidang, new_hp, pilih_pengajar))
                    st.success("✅ Data pengajar berhasil diperbarui!")
                    st.rerun()
                elif delete_btn:
                    run_query("DELETE FROM pengajar WHERE id_pengajar = ?", (pilih_pengajar,))
                    st.success("🗑️ Data pengajar berhasil dihapus!")
                    st.rerun()
        else:
            st.info("Belum ada data pengajar.")

    st.divider()
    st.subheader("Daftar Pengajar Aktif")
    st.dataframe(load_data("pengajar"), use_container_width=True)

# 5. DATA SANTRI
elif menu == "Data Santri":
    st.header("👦👧 Manajemen Data Santri")
    tab1, tab2, tab3 = st.tabs(["➕ Input Manual", "📁 Impor dari Excel", "✏️ Edit / Hapus Data"])
    
    with tab1:
        with st.form("form_santri", clear_on_submit=True):
            nis = st.text_input("NIS / Nomor Induk")
            nama = st.text_input("Nama Lengkap")
            kelas = st.text_input("Kelas / Tingkat")
            asrama_kamar = st.text_input("Kamar / Asrama (Opsional)")
            wali = st.text_input("Nama Orang Tua / Wali")
            if st.form_submit_button("Simpan Data Santri"):
                if nis and nama:
                    try:
                        run_query("INSERT INTO santri VALUES (?, ?, ?, ?, ?)", (nis, nama, kelas, asrama_kamar, wali))
                        st.success(f"Data santri **{nama}** berhasil disimpan!")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error(f"NIS '{nis}' sudah terdaftar.")

    with tab2:
        st.subheader("Unggah Data Santri Masal via Excel (.xlsx)")
        df_template_santri = pd.DataFrame(columns=["nis", "nama", "kelas", "asrama_kamar", "wali_murid"])
        output_santri = io.BytesIO()
        with pd.ExcelWriter(output_santri, engine='openpyxl') as writer:
            df_template_santri.to_excel(writer, sheet_name='Data Santri', index=False)
        st.download_button("📥 Unduh Template Excel", data=output_santri.getvalue(), file_name="template_data_santri.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        st.divider()
        uploaded_file = st.file_uploader("Pilih file Excel (.xlsx) data santri", type=["xlsx"], key="up_santri")
        if uploaded_file is not None:
            try:
                df_upload = pd.read_excel(uploaded_file, engine='openpyxl')
                if st.button("Proses Impor Santri"):
                    conn = sqlite3.connect(DB_NAME)
                    df_upload.to_sql('santri', conn, if_exists='append', index=False)
                    conn.close()
                    st.success("✅ Data santri berhasil diimpor!")
                    st.rerun()
            except Exception as e:
                st.error(f"Terjadi kesalahan: {e}")

    with tab3:
        st.subheader("Pilih Santri yang Ingin Diedit atau Dihapus")
        df_santri_s = load_data("santri")
        if not df_santri_s.empty:
            pilih_santri = st.selectbox("Pilih NIS & Nama Santri", options=df_santri_s['nis'].tolist(), format_func=lambda x: f"{x} - {df_santri_s.loc[df_santri_s['nis'] == x, 'nama'].values[0]}")
            
            santri_terpilih = df_santri_s[df_santri_s['nis'] == pilih_santri].iloc[0]
            
            with st.form("form_edit_santri"):
                new_nama_s = st.text_input("Nama Lengkap", value=santri_terpilih['nama'])
                new_kelas = st.text_input("Kelas / Tingkat", value=santri_terpilih['kelas'])
                new_kamar = st.text_input("Kamar / Asrama", value=santri_terpilih['asrama_kamar'])
                new_wali = st.text_input("Nama Orang Tua / Wali", value=santri_terpilih['wali_murid'])
                
                col_s1, col_s2 = st.columns(2)
                update_s_btn = col_s1.form_submit_button("💾 Perbarui Data")
                delete_s_btn = col_s2.form_submit_button("🗑️ Hapus Data", type="primary")
                
                if update_s_btn:
                    run_query("UPDATE santri SET nama = ?, kelas = ?, asrama_kamar = ?, wali_murid = ? WHERE nis = ?", (new_nama_s, new_kelas, new_kamar, new_wali, pilih_santri))
                    st.success("✅ Data santri berhasil diperbarui!")
                    st.rerun()
                elif delete_s_btn:
                    run_query("DELETE FROM santri WHERE nis = ?", (pilih_santri,))
                    st.success("🗑️ Data santri berhasil dihapus!")
                    st.rerun()
        else:
            st.info("Belum ada data santri.")

    st.divider()
    st.subheader("Daftar Santri Aktif")
    st.dataframe(load_data("santri"), use_container_width=True)

# 6. DATA SARPRAS
elif menu == "Data Sarpras":
    st.header("🏢 Manajemen Sarana & Prasarana")
    with st.form("form_sarpras", clear_on_submit=True):
        kode = st.text_input("Kode Barang")
        nama_brg = st.text_input("Nama Barang")
        kondisi = st.selectbox("Kondisi", ["Baik", "Rusak Ringan", "Rusak Berat"])
        jumlah = st.number_input("Jumlah Unit", min_value=1, step=1)
        if st.form_submit_button("Simpan Data Sarpras"):
            if kode and nama_brg:
                try:
                    run_query("INSERT INTO sarpras VALUES (?, ?, ?, ?)", (kode, nama_brg, kondisi, int(jumlah)))
                    st.success(f"Data sarpras **{nama_brg}** berhasil disimpan!")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error(f"Kode Barang '{kode}' sudah terdaftar.")

    st.subheader("Daftar Inventaris Sarpras")
    st.dataframe(load_data("sarpras"), use_container_width=True)

# 6.5 GALERI & DOKUMENTASI FOTO
elif menu == "📸 Galeri & Dokumentasi":
    st.header("📸 Galeri & Dokumentasi Visual Lembaga")
    tab1, tab2 = st.tabs(["➕ Unggah Foto Baru", "🖼️ Lihat & Kelola Galeri"])
    
    with tab1:
        with st.form("form_galeri", clear_on_submit=True):
            kategori = st.selectbox("Pilih Kategori Foto", [
                "Papan Nama Lembaga", 
                "Gedung / Asrama", 
                "Ruang Belajar / Kelas", 
                "Kegiatan Pembelajaran KBM"
            ])
            keterangan = st.text_input("Keterangan / Deskripsi Foto")
            file_foto = st.file_uploader("Pilih Berkas Foto", type=["jpg", "jpeg", "png"])
            
            if st.form_submit_button("Unggah Foto"):
                if file_foto is not None:
                    nama_file_unik = f"{int(time.time())}_{file_foto.name}"
                    file_path = os.path.join(UPLOAD_DIR, nama_file_unik)
                    
                    with open(file_path, "wb") as f:
                        f.write(file_foto.getbuffer())
                        
                    run_query("INSERT INTO galeri_foto (kategori, nama_file, keterangan) VALUES (?, ?, ?)", 
                              (kategori, nama_file_unik, keterangan))
                    st.success(f"✅ Foto kategori **{kategori}** berhasil diunggah!")
                    st.rerun()
                else:
                    st.warning("Pilih file foto terlebih dahulu.")

    with tab2:
        st.subheader("Daftar Dokumentasi Visual Tersimpan")
        conn = sqlite3.connect(DB_NAME)
        df_galeri = pd.read_sql_query("SELECT * FROM galeri_foto", conn)
        conn.close()
        
        if not df_galeri.empty:
            filter_kat = st.selectbox("Filter Kategori", ["Semua Kategori", "Papan Nama Lembaga", "Gedung / Asrama", "Ruang Belajar / Kelas", "Kegiatan Pembelajaran KBM"])
            
            if filter_kat != "Semua Kategori":
                df_tampil = df_galeri[df_galeri['kategori'] == filter_kat]
            else:
                df_tampil = df_galeri
                
            if not df_tampil.empty:
                for index, row in df_tampil.iterrows():
                    st.markdown(f"### Kategori: {row['kategori']}")
                    st.write(f"**Keterangan:** {row['keterangan'] if row['keterangan'] else '-'}")
                    
                    img_path = os.path.join(UPLOAD_DIR, row['nama_file'])
                    if os.path.exists(img_path):
                        st.image(img_path, width=400)
                    else:
                        st.error("File foto fisik tidak ditemukan di direktori.")
                        
                    if st.button(f"🗑️ Hapus Foto ID {row['id']}", key=f"del_foto_{row['id']}"):
                        if os.path.exists(img_path):
                            os.remove(img_path)
                        run_query("DELETE FROM galeri_foto WHERE id = ?", (row['id'],))
                        st.success("Foto berhasil dihapus!")
                        st.rerun()
                    st.divider()
            else:
                st.info("Tidak ada foto pada kategori ini.")
        else:
            st.info("Belum ada foto dokumentasi yang diunggah.")

# 7. CETAK PORTOFOLIO WORD
elif menu == "📄 Cetak Portofolio Word":
    st.header("📄 Cetak Portofolio Lembaga (.docx)")
    st.write("Unduh rangkuman seluruh data lembaga keagamaan islam dalam bentuk dokumen Word yang rapi dan siap dicetak.")
    st.divider()
    
    df_profil = load_data("profil_lembaga")
    nama_lembaga = df_profil['nama_lembaga'].values[0] if not df_profil.empty else "Lembaga"
    
    doc_file = generate_word_portfolio()
    
    st.download_button(
        label="📥 Download Portofolio Word (.docx)",
        data=doc_file,
        file_name=f"Portofolio_{nama_lembaga.replace(' ', '_')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary"
    )
