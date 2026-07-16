import os
import urllib.parse
import pandas as pd
from sqlalchemy import create_engine, text
from dotenv import load_dotenv

# Konfigurasi
load_dotenv()
DB_USER = os.getenv("DB_USER", "postgres")
DB_PASS = os.getenv("DB_PASS", "password")
DB_HOST = os.getenv("DB_HOST", "localhost")
DB_PORT = os.getenv("DB_PORT", "5432")
DB_NAME = os.getenv("DB_NAME", "DB_Pipeline")

FILE_INPUT = "Data_Baru_Mapped.xlsx"

encoded_password = urllib.parse.quote_plus(DB_PASS)
engine = create_engine(f'postgresql://{DB_USER}:{encoded_password}@{DB_HOST}:{DB_PORT}/{DB_NAME}')

print("Membaca referensi dari Excel untuk proses penghapusan...")
df_pipe = pd.read_excel(FILE_INPUT, sheet_name='Pipelines')
df_sa = pd.read_excel(FILE_INPUT, sheet_name='SolutionArchitects')
df_pr = pd.read_excel(FILE_INPUT, sheet_name='Principals')
df_cust = pd.read_excel(FILE_INPUT, sheet_name='Customers')

# Langkah 1: Hapus data Transaksional (Pipelines & Relasinya)
with engine.begin() as conn:
    print("Menghapus transaksi pipelines dan relasinya...")
    for project_name in df_pipe['project_name'].dropna().unique():
        # Cari ID dari pipeline yang project_name-nya ada di excel
        res = conn.execute(text("SELECT id FROM pipelines WHERE project_name = :p"), {"p": project_name}).fetchall()
        p_ids = [r[0] for r in res]
        
        for pid in p_ids:
            # Hapus semua relasi child terlebih dahulu
            conn.execute(text("DELETE FROM proposal_temps WHERE pipeline_id = :pid"), {"pid": pid})
            conn.execute(text("DELETE FROM proposals WHERE pipeline_id = :pid"), {"pid": pid})
            conn.execute(text("DELETE FROM principal_pipelines WHERE pipeline_id = :pid"), {"pid": pid})
            # Hapus pipeline-nya
            conn.execute(text("DELETE FROM pipelines WHERE id = :pid"), {"pid": pid})
            
print("Data transaksional berhasil dihapus.")

# Langkah 2: Hapus data Master (Hanya jika tidak ada pipeline lain yang memakainya)
with engine.connect() as conn:
    print("Mencoba menghapus data master (Principals, SA, Customers)...")
    
    for p_name in df_pr['principal_name'].dropna().unique():
        try:
            conn.execute(text("DELETE FROM principals WHERE principal_name = :n"), {"n": p_name})
            conn.commit()
        except Exception:
            conn.rollback() # Abaikan jika principal ini masih dipakai oleh project lama
            
    for sa_name in df_sa['name'].dropna().unique():
        try:
            conn.execute(text("DELETE FROM solution_architects WHERE name = :n"), {"n": sa_name})
            conn.commit()
        except Exception:
            conn.rollback()

    for c_name in df_cust['customer_name'].dropna().unique():
        try:
            conn.execute(text("DELETE FROM customers WHERE customer_name = :n"), {"n": c_name})
            conn.commit()
        except Exception:
            conn.rollback()

print("Penghapusan data spesifik selesai!")