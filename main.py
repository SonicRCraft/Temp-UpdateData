import os
import re
import pandas as pd
from sqlalchemy import create_engine, text
from dotenv import load_dotenv

# =======================================================
# 1. PERSIAPAN & LOAD EXCEL
# =======================================================
load_dotenv()
file_path = os.getenv("EXCEL_SOURCE_PATH", "Test Add Database.xlsx")

if not os.path.exists(file_path):
    raise ValueError(f"Path file Excel '{file_path}' tidak ditemukan.")

print(f"Membaca file: {file_path}...")

try:
    df_customer = pd.read_excel(file_path, sheet_name='Inisial Customer')
except ValueError:
    df_customer = pd.read_excel(file_path, sheet_name='List Initial', usecols="E:G", header=1)

try:
    df_proptek = pd.read_excel(file_path, sheet_name='Proptek Modified')
except ValueError:
    df_proptek = pd.read_excel(file_path, sheet_name='Proptek')

df_customer = df_customer.dropna(how='all')
df_proptek = df_proptek.dropna(how='all')

# =======================================================
# 2. MAPPING INISIAL CUSTOMER
# =======================================================
kolom_target = next((col for col in df_customer.columns if 'Inisial' in str(col)), None)
nama_col = next((col for col in df_customer.columns if 'Nama' in str(col)), df_customer.columns[1] if len(df_customer.columns) > 1 else None)

if kolom_target and nama_col:
    customer_map = dict(zip(
        df_customer[kolom_target].dropna().astype(str).str.strip().str.upper(), 
        df_customer[nama_col].dropna().astype(str).str.strip()
    ))
else:
    customer_map = {}

if 'Customer Initial' in df_proptek.columns:
    df_proptek['Customer Initial'] = df_proptek['Customer Initial'].astype(str).str.strip().str.upper().map(customer_map).fillna(df_proptek['Customer Initial'])

# =======================================================
# 3. KONEKSI DATABASE POSTGRESQL
# =======================================================
DB_USER = os.getenv("DB_USER")
DB_PASS = os.getenv("DB_PASS")
DB_HOST = os.getenv("DB_HOST", "localhost")
DB_PORT = os.getenv("DB_PORT", "5432")
DB_NAME = os.getenv("DB_NAME")

engine = create_engine(f'postgresql://{DB_USER}:{DB_PASS}@{DB_HOST}:{DB_PORT}/{DB_NAME}')
print("Terkoneksi ke PostgreSQL. Memulai proses transfer...")

# =======================================================
# 4. HELPER & PROSES SINKRONISASI DATA MASTER
# =======================================================
def get_unique_list(series):
    items = {}
    for val in series.dropna().astype(str):
        for p in re.split(r',|&', val):
            clean = p.strip()
            if clean and clean.lower() not in ['nan', 'none', 'null', '-', '']:
                items[clean.lower()] = clean # Simpan dengan format aslinya
    return list(items.values())

with engine.begin() as conn:
    print("Memeriksa & Sync Data Master (Customer, SA, Principal)...")
    
    # Sync Customer
    for name in get_unique_list(df_proptek.get('Customer Initial', pd.Series([]))):
        conn.execute(
            text("INSERT INTO customers (customer_name, sector) SELECT :name, '-' WHERE NOT EXISTS (SELECT 1 FROM customers WHERE LOWER(customer_name) = LOWER(:name))"), 
            {"name": name}
        )
    
    # Sync Solution Architect (CEK DAN UPDATE JIKA EMAIL NULL)
    for name in get_unique_list(df_proptek.get('Solution Architect', pd.Series([]))):
        email = f"{re.sub(r'[^a-z0-9]', '', str(name).lower())}@dummy.com"
        
        # Cek apakah SA sudah ada di database
        existing_sa = conn.execute(
            text("SELECT id, email FROM solution_architects WHERE LOWER(name) = LOWER(:name) LIMIT 1"),
            {"name": name}
        ).fetchone()

        if existing_sa:
            # Jika sudah ada, tapi email-nya null atau kosong, maka UPDATE emailnya
            if existing_sa[1] is None or str(existing_sa[1]).strip() == "":
                conn.execute(
                    text("UPDATE solution_architects SET email = :email WHERE id = :id"),
                    {"email": email, "id": existing_sa[0]}
                )
        else:
            # Jika belum ada, langsung INSERT beserta emailnya
            conn.execute(
                text("INSERT INTO solution_architects (name, email) VALUES (:name, :email)"),
                {"name": name, "email": email}
            )
    
    # Sync Principal
    for name in get_unique_list(df_proptek.get('Principal', pd.Series([]))):
        conn.execute(
            text("INSERT INTO principals (principal_name) SELECT :name WHERE NOT EXISTS (SELECT 1 FROM principals WHERE LOWER(principal_name) = LOWER(:name))"), 
            {"name": name}
        )

# =======================================================
# 5. TRANSAKSI (UPSERT LOGIC DENGAN RETURNING ID)
# =======================================================
def get_str(val):
    v = str(val).strip()
    return "-" if not v or v.lower() in ['nan', 'none', 'null', ''] else v

# PERBAIKAN: Menghilangkan .0 jika angka berupa bilangan bulat (integer)
def get_num(val):
    v = str(val).strip()
    if not v or v.lower() in ['nan', 'none', 'null', '-', '']: 
        return 0
    try: 
        f = float(v)
        return int(f) if f.is_integer() else f
    except ValueError: 
        return 0

print("Memproses data Pipeline & Proposal...")

with engine.connect() as conn:
    # Mengambil ID Master Data dengan format lower() agar kebal terhadap Case-Sensitivity
    cust_map = {row[1].lower(): row[0] for row in conn.execute(text("SELECT id, customer_name FROM customers")).fetchall()}
    sa_map = {row[1].lower(): row[0] for row in conn.execute(text("SELECT id, name FROM solution_architects")).fetchall()}
    prin_map = {row[1].lower(): row[0] for row in conn.execute(text("SELECT id, principal_name FROM principals")).fetchall()}

    proposal_counter = 1

    for index, row in df_proptek.iterrows():
        try:
            cust_val = str(row.get('Customer Initial', '')).strip().lower()
            sa_val = str(row.get('Solution Architect', '')).strip().lower()
            
            customer_id = cust_map.get(cust_val)
            sa_id = sa_map.get(sa_val)

            if not customer_id:
                # Lewati jika baris ini tidak memiliki Customer yang valid
                continue

            # Ambil semua data dari Excel
            judul_proposal = get_str(row.get('Judul Proposal', row.get('Title', '')))
            solution_type = get_str(row.get('Solution', '-'))
            project_revenue = get_num(row.get('Project Revenue', row.get('project_revenue')))
            next_action_plan = get_str(row.get('Next Action Plan', row.get('next_action_plan')))
            est_time_line = get_str(row.get('Est Timeline', row.get('est_time_line')))
            current_status = get_str(row.get('Statu', row.get('Status', row.get('Current Status', '-'))))
            last_update = get_str(row.get('Last Update', row.get('last_update', '-')))
            sales_stage = get_str(row.get('Sales Stage', row.get('sales_stage', '-')))
            sales_type = get_str(row.get('Sales Type', row.get('sales_type', '-')))
            
            # Format nama project unik dengan solution type (Opsional)
            project_name = f"{judul_proposal} ({solution_type})" if solution_type != "-" else judul_proposal

            with conn.begin(): # Transaksi per baris agar aman
                # 1. Cek Pipeline (Apakah sudah ada?)
                existing_pipeline = conn.execute(
                    text("SELECT id FROM pipelines WHERE project_name = :p AND customer_id = :c LIMIT 1"),
                    {"p": project_name, "c": customer_id}
                ).scalar()

                if existing_pipeline:
                    # --- UPDATE PIPELINE EXISTING ---
                    pipeline_id = str(existing_pipeline)
                    conn.execute(
                        text("""
                            UPDATE pipelines 
                            SET solution_type = :st, solution_architect_id = :sa, project_revenue = :rev,
                                next_action_plan = :nap, est_time_line = :etl, current_status = :cs,
                                last_update = :lu, sales_stage = :ss, sales_type = :stype, updated_at = now()
                            WHERE id = :id
                        """),
                        {
                            "id": pipeline_id, "st": solution_type, "sa": sa_id, "rev": project_revenue,
                            "nap": next_action_plan, "etl": est_time_line, "cs": current_status,
                            "lu": last_update, "ss": sales_stage, "stype": sales_type
                        }
                    )
                    
                    conn.execute(text("UPDATE proposals SET title = :t, updated_at = now() WHERE pipeline_id = :pid"), {"t": judul_proposal, "pid": pipeline_id})
                    conn.execute(text("DELETE FROM principal_pipelines WHERE pipeline_id = :pid"), {"pid": pipeline_id})
                
                else:
                    # --- INSERT PIPELINE BARU (Generate ID by DB) ---
                    res = conn.execute(
                        text("""
                            INSERT INTO pipelines (
                                project_name, solution_type, customer_id, solution_architect_id, 
                                project_revenue, next_action_plan, est_time_line, current_status, 
                                last_update, sales_stage, sales_type
                            ) VALUES (
                                :p, :st, :c, :sa, :rev, :nap, :etl, :cs, :lu, :ss, :stype
                            ) RETURNING id
                        """),
                        {
                            "p": project_name, "st": solution_type, "c": customer_id, "sa": sa_id, "rev": project_revenue,
                            "nap": next_action_plan, "etl": est_time_line, "cs": current_status,
                            "lu": last_update, "ss": sales_stage, "stype": sales_type
                        }
                    )
                    pipeline_id = res.scalar()

                    # Penomoran Proposal
                    no_prop = str(row.get('Nomor Proposal', '')).strip()
                    if no_prop.lower() in ['nan', 'none', 'null', '', '-']:
                        no_prop = f"PROP-{proposal_counter:04d}"
                        proposal_counter += 1

                    # INSERT PROPOSAL TEMP
                    conn.execute(
                        text("INSERT INTO proposal_temps (no_proposal, status, pipeline_id) VALUES (:np, 'pending', :pid)"),
                        {"np": no_prop, "pid": pipeline_id}
                    )

                    # INSERT PROPOSAL
                    conn.execute(
                        text("""
                            INSERT INTO proposals (title, no_proposal, is_deleted, pipeline_id) 
                            VALUES (:t, :np, false, :pid)
                        """),
                        {"t": judul_proposal, "np": no_prop, "pid": pipeline_id}
                    )

                # --- INSERT RELASI MANY-TO-MANY (Principal Pipeline) ---
                principal_raw = str(row.get('Principal', '')).strip()
                if principal_raw and principal_raw.lower() not in ['nan', 'none', 'null', '', '-']:
                    for p_name in [p.strip() for p in re.split(r',|&', principal_raw) if p.strip()]:
                        p_id = prin_map.get(p_name.lower())
                        if p_id:
                            conn.execute(
                                text("INSERT INTO principal_pipelines (project_name, principal_id, pipeline_id) VALUES (:pn, :pid, :pipe_id)"),
                                {"pn": project_name, "pid": p_id, "pipe_id": pipeline_id}
                            )

        except Exception as row_err:
            print(f"Gagal memproses baris Excel ke-{index + 2}: {row_err}")

print("Berhasil! Seluruh data terhubung dan disinkronkan ke PostgreSQL.")