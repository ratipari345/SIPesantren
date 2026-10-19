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
menu = st.sidebar.selectbox("Pilih Menu", ["Dashboard", "Profil Pesantren", "Data Ustad", "Data Santri", "Data Sarpras"])

# 1. DASHBOARD
if menu == "Dashboard":
    st.header("📊 Dashboard Rekapitulasi Pondok Pesantren")
    df_ustad = load_data("ustad")
    df_santri = load_data("santri")
    df_sarpras = load_data("sarpras")

    col1, col2, col3 = st.columns(3)
    col1.metric("Total Ustad", len(df_ustad))
    col2.metric("Total Santri", len(df_santri))
    col3.metric("Total Sarpras", len(df_sarpras))

    st.success("✅ Aplikasi siap digunakan!")

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

# 3. DATA USTAD
elif menu == "Data Ustad":
    st.header("👨‍🏫 Manajemen Data Ustad / Pengajar")
    with st.form("form_ustad", clear_on_submit=True):
        id_ustad = st.text_input("ID Ustad")
        nama = st.text_input("Nama Lengkap & Gelar")
        bidang = st.text_input("Bidang Pengajaran / Kitab")
        hp = st.text_input("No. HP / WhatsApp")
        if st.form_submit_button("Simpan Data Ustad"):
            if id_ustad and nama:
                try:
                    run_query("INSERT INTO ustad VALUES (?, ?, ?, ?)", (id_ustad, nama, bidang, hp))
                    st.success(f"Data **{nama}** berhasil disimpan!")
                except sqlite3.IntegrityError:
                    st.error(f"ID Ustad '{id_ustad}' sudah terdaftar.")
            else:
                st.warning("ID dan Nama wajib diisi.")

    st.subheader("Daftar Ustad")
    st.dataframe(load_data("ustad"), use_container_width=True)

# 4. DATA SANTRI
elif menu == "Data Santri":
    st.header("👦👧 Manajemen Data Santri")
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

    st.subheader("Daftar Santri Aktif")
    st.dataframe(load_data("santri"), use_container_width=True)

# 5. DATA SARPRAS
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
