import streamlit as st
import pandas as pd
import sqlite3
import io

# --- KONFIGURASI DATABASE ---
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
        CREATE TABLE IF NOT EXISTS satuan_pendidikan (
            npsn TEXT PRIMARY KEY,
            nama_satuan TEXT NOT NULL,
            jenjang TEXT,
            status TEXT,
            kepala_sekolah TEXT
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
