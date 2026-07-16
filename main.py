import os
import urllib.parse
import pandas as pd
from sqlalchemy import create_engine, text
from dotenv import load_dotenv
import bcrypt
import uuid

# =======================================================
# 1. KONFIGURASI DATABASE
# =======================================================
load_dotenv()
DB_USER = os.getenv("DB_USER", "postgres")
DB_PASS = os.getenv("DB_PASS", "password")
DB_HOST = os.getenv("DB_HOST", "localhost")
DB_PORT = os.getenv("DB_PORT", "5432")
DB_NAME = os.getenv("DB_NAME", "DB_Pipeline")

FILE_INPUT = os.getenv("EXCEL_SOURCE_PATH", "Data_Baru_Mapped.xlsx")

if not os.path.exists(FILE_INPUT):
    raise ValueError(f"File '{FILE_INPUT}' tidak ditemukan!")

# =======================================================
# 2. LOAD DATA_BARU_MAPPED
# =======================================================
print(f"Membaca file utama relasional: {FILE_INPUT}...")
df_cust = pd.read_excel(FILE_INPUT, sheet_name='Customers')
df_sa = pd.read_excel(FILE_INPUT, sheet_name='SolutionArchitects')
df_pr = pd.read_excel(FILE_INPUT, sheet_name='Principals')
df_pipe = pd.read_excel(FILE_INPUT, sheet_name='Pipelines')
df_pp = pd.read_excel(FILE_INPUT, sheet_name='PrincipalPipelines')
df_prop = pd.read_excel(FILE_INPUT, sheet_name='Proposals')
df_pt = pd.read_excel(FILE_INPUT, sheet_name='ProposalTemps')

# =======================================================
# 3. PROSES INSERT KE POSTGRESQL 
# =======================================================
encoded_password = urllib.parse.quote_plus(DB_PASS)
engine = create_engine(f'postgresql://{DB_USER}:{encoded_password}@{DB_HOST}:{DB_PORT}/{DB_NAME}')
print("Terkoneksi ke PostgreSQL. Memulai proses Insert...")

# Menyimpan mapping: Excel ID -> DB ID
cust_id_map, sa_id_map, pr_id_map, pipe_id_map = {}, {}, {}, {}
# Mapping tambahan: DB SA ID -> DB User ID (Untuk digunakan di Pipeline)
sa_db_to_user_db = {}

def clean(val, default='-'):
    if pd.isna(val) or str(val).strip().lower() in ['nan', 'none', 'null', '']:
        return default
    return str(val).strip()

def hash_password(password_str):
    return bcrypt.hashpw(password_str.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')

default_password_hashed = hash_password("123456")

with engine.begin() as conn:
    # --- A. PASTIKAN ROLE 'PRESALES' ADA ---
    print("Mengecek Role 'Presales'...")
    role_res = conn.execute(text("SELECT id FROM roles WHERE LOWER(name) = 'presales'")).scalar()
    if not role_res:
        presales_role_id = conn.execute(
            text("INSERT INTO roles (id, name) VALUES (:id, :name) RETURNING id"),
            {"id": str(uuid.uuid4()), "name": "presales"}
        ).scalar()
        print("-> Role 'Presales' dibuat.")
    else:
        presales_role_id = role_res

    # --- B. CUSTOMERS ---
    print("Memasukkan data Customers...")
    for _, row in df_cust.iterrows():
        excel_id, c_name = str(row['id']).strip(), str(row['customer_name']).strip()
        res = conn.execute(text("SELECT id FROM customers WHERE LOWER(customer_name) = LOWER(:name)"), {"name": c_name}).scalar()
        if res:
            cust_id_map[excel_id] = res
        else:
            new_id = conn.execute(
                text("INSERT INTO customers (customer_name, sector) VALUES (:name, :sector) RETURNING id"),
                {"name": c_name, "sector": clean(row.get('sector', '-'))}
            ).scalar()
            cust_id_map[excel_id] = new_id

    # --- C. SOLUTION ARCHITECTS & USERS ---
    print("Memasukkan data Solution Architects & Membuat Usernya...")
    for _, row in df_sa.iterrows():
        excel_id, sa_name, sa_mail = str(row['id']).strip(), str(row['name']).strip(), str(row['email']).strip()
        
        # Cek apakah SA sudah ada
        sa_db_id = conn.execute(text("SELECT id FROM solution_architects WHERE LOWER(name) = LOWER(:name)"), {"name": sa_name}).scalar()
        
        if not sa_db_id:
            # Insert SA baru
            sa_db_id = conn.execute(
                text("INSERT INTO solution_architects (name, email) VALUES (:name, :email) RETURNING id"),
                {"name": sa_name, "email": sa_mail}
            ).scalar()
        sa_id_map[excel_id] = sa_db_id

        # Cek apakah User dengan email tersebut sudah ada
        user_db_id = conn.execute(text("SELECT id FROM users WHERE LOWER(email) = LOWER(:e)"), {"e": sa_mail}).scalar()
        
        if not user_db_id:
            # Generate username dari email (sebelum @)
            username = sa_mail.split('@')[0]
            # Insert User baru untuk SA ini
            user_db_id = conn.execute(
                text("""
                    INSERT INTO users (id, user_name, email, password, role, sa_id, is_active, failed_login_attempts) 
                    VALUES (:uid, :uname, :email, :pw, :role, :said, true, 0) RETURNING id
                """),
                {
                    "uid": str(uuid.uuid4()), "uname": username, "email": sa_mail, 
                    "pw": '', "role": presales_role_id, "said": sa_db_id
                }
            ).scalar()
        else:
            # Jika user sudah ada, pastikan sa_id-nya terhubung
            conn.execute(text("UPDATE users SET sa_id = :said WHERE id = :uid"), {"said": sa_db_id, "uid": user_db_id})

        # Simpan mapping SA ke User untuk digunakan di Pipeline
        sa_db_to_user_db[sa_db_id] = user_db_id


    # --- D. PRINCIPALS ---
    print("Memasukkan data Principals...")
    for _, row in df_pr.iterrows():
        excel_id, p_name = str(row['id']).strip(), str(row['principal_name']).strip()
        res = conn.execute(text("SELECT id FROM principals WHERE LOWER(principal_name) = LOWER(:name)"), {"name": p_name}).scalar()
        if res:
            pr_id_map[excel_id] = res
        else:
            new_id = conn.execute(
                text("INSERT INTO principals (principal_name) VALUES (:name) RETURNING id"),
                {"name": p_name}
            ).scalar()
            pr_id_map[excel_id] = new_id

    # --- E. PIPELINES ---
    print("Memasukkan data Pipelines...")
    for _, row in df_pipe.iterrows():
        excel_id, p_name = str(row['id']).strip(), str(row['project_name']).strip()
        db_cust_id = cust_id_map.get(str(row.get('customer_id')).strip())
        db_sa_id = sa_id_map.get(str(row.get('solution_architect_id')).strip())
        
        # AMBIL USER_ID DARI SA_ID YANG MENGHANDLE PROJECT INI
        db_user_id = sa_db_to_user_db.get(db_sa_id) if db_sa_id else None
        
        if not db_cust_id: continue
            
        res = conn.execute(text("SELECT id FROM pipelines WHERE project_name = :p AND customer_id = :c"), {"p": p_name, "c": db_cust_id}).scalar()
        
        if res:
            pipe_id_map[excel_id] = res
            # Update user_id just in case
            if db_user_id:
                conn.execute(text("UPDATE pipelines SET user_id = :uid WHERE id = :pid"), {"uid": db_user_id, "pid": res})
        else:
            rev = row.get('project_revenue', 0)
            rev = 0 if pd.isna(rev) else rev
            new_id = conn.execute(
                text("""
                    INSERT INTO pipelines (
                        project_name, solution_type, project_revenue, next_action_plan, 
                        est_time_line, current_status, last_update, sales_stage, sales_type, 
                        customer_id, solution_architect_id, user_id
                    ) VALUES (
                        :p, :st, :rev, :nap, :etl, :cs, :lu, :ss, :stype, :cid, :said, :uid
                    ) RETURNING id
                """),
                {
                    "p": p_name, "st": clean(row.get('solution_type')), "rev": 0, "nap": clean(row.get('next_action_plan')),
                    "etl": clean(row.get('est_time_line')), "cs": clean(row.get('current_status')), "lu": clean(row.get('last_update')),
                    "ss": clean(row.get('sales_stage')), "stype": clean(row.get('sales_type')), 
                    "cid": db_cust_id, "said": db_sa_id, "uid": db_user_id
                }
            ).scalar()
            pipe_id_map[excel_id] = new_id

    # --- F. PRINCIPAL PIPELINES ---
    print("Memasukkan data relasi Principal & Pipelines...")
    for _, row in df_pp.iterrows():
        db_pipe_id = pipe_id_map.get(str(row.get('pipeline_id')).strip())
        db_prin_id = pr_id_map.get(str(row.get('principal_id')).strip())
        
        if db_pipe_id and db_prin_id:
            res = conn.execute(text("SELECT id FROM principal_pipelines WHERE pipeline_id = :pipe_id AND principal_id = :prin_id"), {"pipe_id": db_pipe_id, "prin_id": db_prin_id}).scalar()
            if not res:
                conn.execute(
                    text("INSERT INTO principal_pipelines (project_name, principal_id, pipeline_id) VALUES (:pn, :prin_id, :pipe_id)"),
                    {"pn": clean(row.get('project_name')), "prin_id": db_prin_id, "pipe_id": db_pipe_id}
                )

    # --- G. PROPOSALS ---
    print("Memasukkan data Proposals...")
    for _, row in df_prop.iterrows():
        db_pipe_id = pipe_id_map.get(str(row.get('pipeline_id')).strip())
        no_prop = str(row.get('no_proposal')).strip()
        
        if db_pipe_id and no_prop and no_prop.lower() not in ['nan', '']:
            res = conn.execute(text("SELECT id FROM proposals WHERE pipeline_id = :pid AND no_proposal = :np"), {"pid": db_pipe_id, "np": no_prop}).scalar()
            if not res:
                conn.execute(
                    text("""
                        INSERT INTO proposals (title, no_proposal, is_deleted, pipeline_id, file, rfa, rfi) 
                        VALUES (:t, :np, :is_del, :pid, :file, :rfa, :rfi)
                    """),
                    {
                        "t": clean(row.get('title')), "np": no_prop, "is_del": bool(row.get('is_deleted', False)), "pid": db_pipe_id,
                        "file": '', "rfa":'', "rfi": ''
                    }
                )

    # --- H. PROPOSAL TEMPS ---
    print("Memasukkan data Proposal Temps...")
    for _, row in df_pt.iterrows():
        db_pipe_id = pipe_id_map.get(str(row.get('pipeline_id')).strip())
        no_prop = str(row.get('no_proposal')).strip()
        
        if db_pipe_id and no_prop and no_prop.lower() not in ['nan', '']:
            res = conn.execute(text("SELECT id FROM proposal_temps WHERE pipeline_id = :pid AND no_proposal = :np"), {"pid": db_pipe_id, "np": no_prop}).scalar()
            if not res:
                conn.execute(
                    text("INSERT INTO proposal_temps (no_proposal, status, pipeline_id) VALUES (:np, :status, :pid)"),
                    {"np": no_prop, "status": "active", "pid": db_pipe_id}
                )

print("Berhasil! Seluruh data baru telah dimasukkan beserta pembuatan User untuk Presales.")