import os
import re
import pandas as pd
from sqlalchemy import create_engine, text
from dotenv import load_dotenv

# =======================================================
# 1. KONFIGURASI DATABASE & EXCEL
# =======================================================
load_dotenv()
DB_USER = os.getenv("DB_USER", "postgres")
DB_PASS = os.getenv("DB_PASS", "password")
DB_HOST = os.getenv("DB_HOST", "localhost")
DB_PORT = os.getenv("DB_PORT", "5432")
DB_NAME = os.getenv("DB_NAME", "DB_Pipeline")

# Sesuaikan dengan nama file Excel Anda
FILE_INPUT = os.getenv("EXCEL_SOURCE_PATH", "Data Baru.xlsx")

if not os.path.exists(FILE_INPUT):
    raise ValueError(f"File '{FILE_INPUT}' tidak ditemukan!")

# =======================================================
# 2. LOAD & MAPPING DATA EXCEL
# =======================================================
print(f"Membaca file: {FILE_INPUT}...")
try:
    df_customer_raw = pd.read_excel(FILE_INPUT, sheet_name='Inisial Customer')
    df_proptek = pd.read_excel(FILE_INPUT, sheet_name='Proptek Modified')
except ValueError:
    df_proptek = pd.read_excel(FILE_INPUT, sheet_name='Proptek')

cust_initial_col = [c for c in df_customer_raw.columns if 'Inisial' in c][0]
cust_name_col = [c for c in df_customer_raw.columns if 'Nama' in c][0]
customer_map = dict(zip(
    df_customer_raw[cust_initial_col].dropna().astype(str).str.strip().str.upper(),
    df_customer_raw[cust_name_col].dropna().astype(str).str.strip()
))

df_proptek = df_proptek.dropna(how='all')

# =======================================================
# 3. KONEKSI & PROSES DATABASE (CEK & INSERT)
# =======================================================
engine = create_engine(f'postgresql://{DB_USER}:{DB_PASS}@{DB_HOST}:{DB_PORT}/{DB_NAME}')
print("Terkoneksi ke PostgreSQL. Memulai proses sinkronisasi (Cek duplikasi)...")

prop_counter = 1

# Menggunakan block transaksi agar aman
with engine.begin() as conn:
    for idx, row in df_proptek.iterrows():
        # ---------------------------------------------------
        # A. MASTER CUSTOMER
        # ---------------------------------------------------
        init = str(row.get('Customer Initial', '')).strip().upper()
        if not init or init.lower() in ['nan', 'none', 'null', '']:
            continue  # Lewati jika tidak ada inisial customer di baris ini

        cust_name = customer_map.get(init, init)
        
        # Cek apakah Customer sudah ada
        res_cust = conn.execute(
            text("SELECT id FROM customers WHERE LOWER(customer_name) = LOWER(:name)"), 
            {"name": cust_name}
        ).scalar()
        
        if res_cust:
            cust_id = res_cust  # Gunakan ID yang sudah ada (Tidak perlu Insert)
        else:
            # Insert tanpa kolom ID agar DB generate UUID otomatis, lalu tangkap ID-nya
            cust_id = conn.execute(
                text("INSERT INTO customers (customer_name, sector) VALUES (:name, '-') RETURNING id"),
                {"name": cust_name}
            ).scalar()

        # ---------------------------------------------------
        # B. MASTER SOLUTION ARCHITECT
        # ---------------------------------------------------
        sa_name = str(row.get('Solution Architect', '')).strip()
        sa_id = None
        if sa_name and sa_name.lower() not in ['nan', 'none', 'null', '']:
            res_sa = conn.execute(
                text("SELECT id FROM solution_architects WHERE LOWER(name) = LOWER(:name)"),
                {"name": sa_name}
            ).scalar()
            
            if res_sa:
                sa_id = res_sa
            else:
                email = f"{re.sub(r'[^a-z0-9]', '', sa_name.lower())}@dummy.com"
                sa_id = conn.execute(
                    text("INSERT INTO solution_architects (name, email) VALUES (:name, :email) RETURNING id"),
                    {"name": sa_name, "email": email}
                ).scalar()

        # ---------------------------------------------------
        # C. MASTER PRINCIPAL
        # ---------------------------------------------------
        raw_p = str(row.get('Principal', '')).strip()
        principal_ids = []
        if raw_p and raw_p.lower() not in ['nan', 'none', 'null', '']:
            # Handle jika ada lebih dari 1 principal dalam 1 sel (dipisah koma/dan)
            for p in re.split(r',|&', raw_p):
                p = p.strip()
                if p and p.lower() not in ['nan', 'none', 'null', '']:
                    res_p = conn.execute(
                        text("SELECT id FROM principals WHERE LOWER(principal_name) = LOWER(:name)"),
                        {"name": p}
                    ).scalar()
                    
                    if res_p:
                        principal_ids.append(res_p)
                    else:
                        new_p_id = conn.execute(
                            text("INSERT INTO principals (principal_name) VALUES (:name) RETURNING id"),
                            {"name": p}
                        ).scalar()
                        principal_ids.append(new_p_id)

        # ---------------------------------------------------
        # D. TRANSAKSI PIPELINE
        # ---------------------------------------------------
        judul = str(row.get('Judul Proposal', row.get('Title', '-'))).strip()
        sol = str(row.get('Solution', '-')).strip()
        status = str(row.get('Statu', row.get('Status', '-'))).strip()
        
        project_name = f"{judul} ({sol})" if sol not in ["-", "nan", ""] else judul
        
        # Cek Pipeline (berdasarkan nama project dan customer_id)
        res_pipe = conn.execute(
            text("SELECT id FROM pipelines WHERE project_name = :p AND customer_id = :c"),
            {"p": project_name, "c": cust_id}
        ).scalar()

        if res_pipe:
            # Jika pipeline sudah ada, ambil ID-nya dan JANGAN INSERT
            pipe_id = res_pipe
        else:
            # Jika belum ada, masukkan data baru. Project revenue diset mutlak 0
            pipe_id = conn.execute(
                text("""
                    INSERT INTO pipelines (
                        project_name, solution_type, project_revenue, next_action_plan, 
                        est_time_line, current_status, last_update, sales_stage, sales_type, 
                        customer_id, solution_architect_id
                    ) VALUES (
                        :p, :st, :rev, '-', '-', :cs, '-', '-', '-', :cid, :said
                    ) RETURNING id
                """),
                {
                    "p": project_name, "st": sol, "rev": 0, "cs": status, 
                    "cid": cust_id, "said": sa_id
                }
            ).scalar()

            # ---------------------------------------------------
            # E. PROPOSAL & PROPOSAL TEMP 
            # ---------------------------------------------------
            no_prop = str(row.get('Nomor Proposal', '')).strip()
            if no_prop.lower() in ['nan', 'none', 'null', '', '-']:
                no_prop = f"PROP-{prop_counter:04d}"
                prop_counter += 1

            # Cek Proposal Temp
            res_temp = conn.execute(
                text("SELECT id FROM proposal_temps WHERE pipeline_id = :pid AND no_proposal = :np"),
                {"pid": pipe_id, "np": no_prop}
            ).scalar()
            if not res_temp:
                conn.execute(
                    text("INSERT INTO proposal_temps (no_proposal, status, pipeline_id) VALUES (:np, 'pending', :pid)"),
                    {"np": no_prop, "pid": pipe_id}
                )

            # Cek Proposal (MENGISI file, rfa, rfi DENGAN STRING KOSONG '')
            res_prop = conn.execute(
                text("SELECT id FROM proposals WHERE pipeline_id = :pid AND no_proposal = :np"),
                {"pid": pipe_id, "np": no_prop}
            ).scalar()
            if not res_prop:
                conn.execute(
                    text("""
                        INSERT INTO proposals (title, no_proposal, is_deleted, pipeline_id, file, rfa, rfi) 
                        VALUES (:t, :np, false, :pid, '', '', '')
                    """),
                    {"t": judul, "np": no_prop, "pid": pipe_id}
                )

        # ---------------------------------------------------
        # F. PRINCIPAL PIPELINES (Junction Table)
        # ---------------------------------------------------
        # Relasikan pipeline dengan principal (bisa >1 principal per baris)
        for pid in principal_ids:
            res_pp = conn.execute(
                text("SELECT id FROM principal_pipelines WHERE pipeline_id = :pipe_id AND principal_id = :prin_id"),
                {"pipe_id": pipe_id, "prin_id": pid}
            ).scalar()
            
            if not res_pp: # Insert hanya jika relasi belum ada
                conn.execute(
                    text("INSERT INTO principal_pipelines (project_name, principal_id, pipeline_id) VALUES (:pn, :prin_id, :pipe_id)"),
                    {"pn": project_name, "prin_id": pid, "pipe_id": pipe_id}
                )

print("Berhasil! Seluruh data dari Excel telah di-import ke PostgreSQL tanpa duplikasi.")