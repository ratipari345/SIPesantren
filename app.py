import streamlit as st
import pandas as pd
import sqlite3

# --- KONFIGURASI DATABASE ---
# Jika dijalankan di Streamlit Cloud, database akan tersimpan secara lokal di server Streamlit
DB_NAME = "pesantren.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS profil_pesantren (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nama_pp TEXT,
            nspp TEXT,
            alamat TEXT,
            pengasuh TEXT,
            akta_notaris TEXT,
            sk_ijin TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS ustad (
            id_ustad TEXT PRIMARY KEY,
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
            kamar TEXT,
            wali_santri TEXT
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
        CREATE TABLE IF NOT EXISTS satuan_pendidikan (
            npsn TEXT PRIMARY KEY,
            nama_satuan TEXT NOT NULL,
            jenjang TEXT,
            status TEXT,
            kepala_sekolah TEXT
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

# --- TAMPILAN APLIKASI STREAMLIT ---
st.set_page_config(page_title="SIM Pondok Pesantren", page_icon="🏫", layout="wide")

st.sidebar.title("🏫 SIM Pesantren")
menu = st.sidebar.selectbox("Pilih Menu", [
    "Dashboard", 
    "Profil Pesantren", 
    "Satuan Pendidikan", 
    "Data Ustad", 
    "Data Santri", 
    "Data Sarpras"
])
# 1. DASHBOARD
# 1. DASHBOARD
if menu == "Dashboard":
    # Ambil nama pesantren terbaru dari database profil
    conn = sqlite3.connect(DB_NAME)
    df_profil = pd.read_sql_query("SELECT nama_pp FROM profil_pesantren ORDER BY id DESC LIMIT 1", conn)
    conn.close()
    
    # Tentukan nama pesantren (jika belum diisi, tampilkan nama default)
    nama_pesantren = df_profil['nama_pp'].values[0] if not df_profil.empty else "Pondok Pesantren"
    
    # Header Dashboard dengan Nama Lembaga
    st.header(f"📊 Dashboard Rekapitulasi")
    st.subheader(f"🏫 {nama_pesantren}")
    st.divider()

    df_ustad = load_data("ustad")
    df_santri = load_data("santri")
    df_sarpras = load_data("sarpras")

    col1, col2, col3 = st.columns(3)
    col1.metric("Total Ustad", len(df_ustad))
    col2.metric("Total Santri", len(df_santri))
    col3.metric("Total Sarpras", len(df_sarpras))

    st.success("✅ Aplikasi siap digunakan dan data tersimpan dengan aman!")
# 2. PROFIL PESANTREN
elif menu == "Profil Pesantren":
    st.header("🏫 Profil & Legalitas Pondok Pesantren")
    
    conn = sqlite3.connect(DB_NAME)
    df_profil = pd.read_sql_query("SELECT * FROM profil_pesantren ORDER BY id DESC LIMIT 1", conn)
    conn.close()
    
    with st.form("form_profil"):
        st.subheader("Perbarui Informasi Lembaga")
        
        def_nama = df_profil['nama_pp'].values[0] if not df_profil.empty else ""
        def_nspp = df_profil['nspp'].values[0] if not df_profil.empty else ""
        def_alamat = df_profil['alamat'].values[0] if not df_profil.empty else ""
        def_pengasuh = df_profil['pengasuh'].values[0] if not df_profil.empty else ""
        def_akta = df_profil['akta_notaris'].values[0] if not df_profil.empty else ""
        def_sk = df_profil['sk_ijin'].values[0] if not df_profil.empty else ""
        
        nama_pp = st.text_input("Nama Pondok Pesantren", value=def_nama)
        nspp = st.text_input("Nomor Statistik Pesantren (NSPP)", value=def_nspp)
        alamat = st.text_area("Alamat Lengkap", value=def_alamat)
        pengasuh = st.text_input("Nama Pengasuh / Pimpinan", value=def_pengasuh)
        akta_notaris = st.text_input("Nomor & Tanggal Akta Notaris", value=def_akta, placeholder="Contoh: Akta No. 12 / 15 Januari 2020")
        sk_ijin = st.text_input("Nomor & Tanggal SK Ijin Operasional", value=def_sk, placeholder="Contoh: SK.123/Kemenag/2021 / 10 Mei 2021")
        
        if st.form_submit_button("Simpan / Perbarui Profil"):
            if nama_pp:
                run_query("DELETE FROM profil_pesantren")
                run_query("INSERT INTO profil_pesantren (nama_pp, nspp, alamat, pengasuh, akta_notaris, sk_ijin) VALUES (?, ?, ?, ?, ?, ?)", 
                          (nama_pp, nspp, alamat, pengasuh, akta_notaris, sk_ijin))
                st.success("✅ Profil pesantren berhasil diperbarui!")
                st.rerun()
            else:
                st.warning("Nama Pondok Pesantren wajib diisi.")

    if not df_profil.empty:
        st.divider()
        st.subheader("📄 Informasi Resmi Lembaga")
        p = df_profil.iloc[0]
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"**Nama Pondok Pesantren:**\n{p['nama_pp']}")
            st.markdown(f"**Nomor Statistik (NSPP):**\n{p['nspp'] if p['nspp'] else '-'}")
            st.markdown(f"**Pengasuh / Pimpinan:**\n{p['pengasuh'] if p['pengasuh'] else '-'}")
        with col2:
            st.markdown(f"**Alamat:**\n{p['alamat'] if p['alamat'] else '-'}")
            st.markdown(f"**Akta Notaris:**\n{p['akta_notaris'] if p['akta_notaris'] else '-'}")
            st.markdown(f"**SK Ijin Operasional:**\n{p['sk_ijin'] if p['sk_ijin'] else '-'}")
    else:
        st.info("ℹ️ Belum ada data profil yang dimasukkan. Silakan isi form di atas.")
        # 3. SATUAN PENDIDIKAN
elif menu == "Satuan Pendidikan":
    st.header("🏫 Manajemen Satuan Pendidikan / Unit Sekolah")
    st.info("Gunakan menu ini untuk mendaftarkan unit pendidikan yang berada di bawah naungan pondok pesantren (Contoh: MTs Darul Huda, MA Darul Huda, TPQ, dll).")

    # Tab pilihan: Input Manual atau Impor Excel
    tab1, tab2 = st.tabs(["➕ Input Manual", "📁 Impor dari Excel"])

    with tab1:
        with st.form("form_satuan", clear_on_submit=True):
            npsn = st.text_input("NPSN / Nomor Induk Satuan (Kosongkan jika non-formal)")
            nama_satuan = st.text_input("Nama Satuan Pendidikan (Contoh: MTs Pesantren)")
            jenjang = st.selectbox("Jenjang Pendidikan", ["RA / TK", "MI / SD", "MTs / SMP", "MA / SMA / SMK", "Madrasah Diniyah", "TPQ / TKA", "Lainnya"])
            status = st.selectbox("Status Lembaga", ["Swasta", "Negeri"])
            kepala_sekolah = st.text_input("Nama Kepala / Pimpinan Unit")
            
            if st.form_submit_button("Simpan Satuan Pendidikan"):
                if nama_satuan:
                    # Jika NPSN kosong, berikan kode unik otomatis
                    if not npsn:
                        import time
                        npsn = f"NON-{int(time.time())}"
                    try:
                        run_query("INSERT INTO satuan_pendidikan VALUES (?, ?, ?, ?, ?)", 
                                  (npsn, nama_satuan, jenjang, status, kepala_sekolah))
                        st.success(f"Satuan pendidikan **{nama_satuan}** berhasil disimpan!")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error(f"NPSN '{npsn}' sudah terdaftar.")
                else:
                    st.warning("Nama Satuan Pendidikan wajib diisi.")

    with tab2:
        st.subheader("Unggah Data Satuan Pendidikan via Excel")
        
        # Template download untuk satuan pendidikan
        df_template_satuan = pd.DataFrame(columns=["npsn", "nama_satuan", "jenjang", "status", "kepala_sekolah"])
        output_satuan = io.BytesIO()
        with pd.ExcelWriter(output_satuan, engine='xlsxwriter') as writer:
            df_template_satuan.to_excel(writer, sheet_name='Satuan Pendidikan', index=False)
        processed_data_satuan = output_satuan.getvalue()
        
        st.download_button(
            label="📥 Unduh Template Excel Satuan Pendidikan",
            data=processed_data_satuan,
            file_name="template_satuan_pendidikan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.divider()
        uploaded_file_satuan = st.file_uploader("Pilih file Excel (.xlsx)", type=["xlsx", "csv"], key="up_satuan")
        
        if uploaded_file_satuan is not None:
            try:
                if uploaded_file_satuan.name.endswith('.csv'):
                    df_upload_satuan = pd.read_csv(uploaded_file_satuan)
                else:
                    df_upload_satuan = pd.read_excel(uploaded_file_satuan)
                
                st.write("Pratinjau Data:", df_upload_satuan.head())
                
                if st.button("Proses Impor Satuan Pendidikan"):
                    conn = sqlite3.connect(DB_NAME)
                    df_upload_satuan.to_sql('satuan_pendidikan', conn, if_exists='append', index=False)
                    conn.close()
                    st.success("✅ Data satuan pendidikan berhasil diimpor!")
                    st.rerun()
            except Exception as e:
                st.error(f"Terjadi kesalahan: {e}")

    st.divider()
    st.subheader("📋 Daftar Unit Satuan Pendidikan di Pesantren")
    df_satuan = load_data("satuan_pendidikan")
    st.dataframe(df_satuan, use_container_width=True)

# 4. DATA USTAD
elif menu == "Data Ustad":
    st.header("👨‍🏫 Manajemen Data Ustad / Pengajar")
    
    # Tab pilihan: Input Manual atau Impor Excel
    tab1, tab2 = st.tabs(["➕ Input Manual", "📁 Impor dari Excel"])
    
    with tab1:
        with st.form("form_ustad", clear_on_submit=True):
            id_ustad = st.text_input("ID Ustad")
            nama = st.text_input("Nama Lengkap & Gelar")
            bidang = st.text_input("Bidang Pengajaran / Kitab")
            hp = st.text_input("No. HP / WhatsApp")
            if st.form_submit_button("Simpan Data Ustad"):
                if id_ustad and nama:
                    try:
                        run_query("INSERT INTO ustad VALUES (?, ?, ?, ?)", (id_ustad, nama, bidang, hp))
                        st.success(f"Data ustad **{nama}** berhasil disimpan!")
                    except sqlite3.IntegrityError:
                        st.error(f"ID Ustad '{id_ustad}' sudah terdaftar.")
                else:
                    st.warning("ID dan Nama wajib diisi.")

    with tab2:
        st.subheader("Unggah Data Ustad Masal via Excel/CSV")
        st.info("Pastikan kolom pada file Excel Anda berurutan: **id_ustad, nama, bidang_pengajaran, no_hp**")
        
        uploaded_file_ustad = st.file_uploader("Pilih file Excel (.xlsx) atau CSV untuk Ustad", type=["xlsx", "csv"], key="upload_ustad")
        
        if uploaded_file_ustad is not None:
            try:
                # Membaca file menggunakan pandas
                if uploaded_file_ustad.name.endswith('.csv'):
                    df_upload_ustad = pd.read_csv(uploaded_file_ustad)
                else:
                    df_upload_ustad = pd.read_excel(uploaded_file_ustad)
                
                st.write("Pratinjau Data Ustad yang akan diimpor:", df_upload_ustad.head())
                
                if st.button("Proses dan Masukkan ke Database Ustad"):
                    conn = sqlite3.connect(DB_NAME)
                    df_upload_ustad.to_sql('ustad', conn, if_exists='append', index=False)
                    conn.close()
                    st.success("✅ Data ustad berhasil diimpor secara massal!")
                    st.rerun()
            except Exception as e:
                st.error(f"Terjadi kesalahan saat membaca file: {e}")

    st.subheader("Daftar Ustad Aktif")
    st.dataframe(load_data("ustad"), use_container_width=True)

# 5. DATA SANTRI
elif menu == "Data Santri":
    st.header("👦👧 Manajemen Data Santri")
    
    # Tab pilihan: Input Manual atau Impor Excel
    tab1, tab2 = st.tabs(["➕ Input Manual", "📁 Impor dari Excel"])
    
    with tab1:
        with st.form("form_santri", clear_on_submit=True):
            nis = st.text_input("NIS (Nomor Induk Santri)")
            nama = st.text_input("Nama Lengkap Santri")
            kelas = st.text_input("Kelas / Jenjang")
            kamar = st.text_input("Kamar Asrama")
            wali = st.text_input("Nama Wali Santri")
            if st.form_submit_button("Simpan Data Santri"):
                if nis and nama:
                    try:
                        run_query("INSERT INTO santri VALUES (?, ?, ?, ?, ?)", (nis, nama, kelas, kamar, wali))
                        st.success(f"Data santri **{nama}** berhasil disimpan!")
                    except sqlite3.IntegrityError:
                        st.error(f"NIS '{nis}' sudah terdaftar.")
                else:
                    st.warning("NIS dan Nama wajib diisi.")

    with tab2:
        st.subheader("Unggah Data Santri Masal via Excel/CSV")
        st.info("Pastikan kolom pada file Excel Anda berurutan: **nis, nama, kelas, kamar, wali_santri**")
        
        uploaded_file = st.file_uploader("Pilih file Excel (.xlsx) atau CSV", type=["xlsx", "csv"])
        
        if uploaded_file is not None:
            try:
                # Membaca file menggunakan pandas
                if uploaded_file.name.endswith('.csv'):
                    df_upload = pd.read_csv(uploaded_file)
                else:
                    df_upload = pd.read_excel(uploaded_file)
                
                st.write("Pratinjau Data yang akan diimpor:", df_upload.head())
                
                if st.button("Proses dan Masukkan ke Database"):
                    conn = sqlite3.connect(DB_NAME)
                    # Masukkan data ke tabel santri secara otomatis (jika NIS sama, akan diabaikan/di-replace)
                    df_upload.to_sql('santri', conn, if_exists='append', index=False)
                    conn.close()
                    st.success("✅ Data santri berhasil diimpor secara massal!")
                    st.rerun()
            except Exception as e:
                st.error(f"Terjadi kesalahan saat membaca file: {e}")

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
                except sqlite3.IntegrityError:
                    st.error(f"Kode Barang '{kode}' sudah terdaftar.")
            else:
                st.warning("Kode dan Nama Barang wajib diisi.")

    st.subheader("Daftar Inventaris Sarpras")
    st.dataframe(load_data("sarpras"), use_container_width=True)
