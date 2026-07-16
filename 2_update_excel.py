import pandas as pd
import re
import uuid
from datetime import datetime

FILE_MAPPED = "Data_Baru_Mapped.xlsx"
FILE_TEAM = "Team Members.xlsx" # NAMA FILE SUDAH DIPERBAIKI

def get_official_sa_info(raw_name, df_team, name_col, email_col):
    if not raw_name or pd.isna(raw_name) or str(raw_name).strip() == '':
        return None, "unknown@dummy.com"
        
    raw_name_lower = str(raw_name).lower().strip()
    
    for _, row in df_team.iterrows():
        full_name = str(row[name_col]).strip()
        email = str(row[email_col]).strip()
        if raw_name_lower == full_name.lower(): return full_name, email
            
    raw_words = raw_name_lower.split()
    for _, row in df_team.iterrows():
        full_name = str(row[name_col]).strip()
        email = str(row[email_col]).strip()
        if all(word in full_name.lower() for word in raw_words): return full_name, email
            
    for _, row in df_team.iterrows():
        full_name = str(row[name_col]).strip()
        email = str(row[email_col]).strip()
        if raw_name_lower in full_name.lower(): return full_name, email
            
    return str(raw_name).strip(), f"{re.sub(r'[^a-z0-9]', '', raw_name_lower)}@dummy.com"

print(f"Membaca {FILE_TEAM}...")
df_team = pd.read_excel(FILE_TEAM)
sa_name_col = [c for c in df_team.columns if 'name' in c.lower() or 'nama' in c.lower()][0]
sa_email_col = [c for c in df_team.columns if 'email' in c.lower()][0]

print(f"Membaca semua sheet dari {FILE_MAPPED}...")
sheets = pd.read_excel(FILE_MAPPED, sheet_name=None)
df_sa = sheets['SolutionArchitects']

print("Memperbarui nama SA yang sudah ada...")
existing_sa_names = set()
for idx, row in df_sa.iterrows():
    current_name = row['name']
    official_name, official_email = get_official_sa_info(current_name, df_team, sa_name_col, sa_email_col)
    
    df_sa.at[idx, 'name'] = official_name
    df_sa.at[idx, 'email'] = official_email
    existing_sa_names.add(official_name.lower())

print("Memasukkan seluruh data sisa dari Team Members ke dalam Excel...")
new_rows = []
current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

# Menambahkan anggota dari Team Members yang belum ada di df_sa
for _, row in df_team.iterrows():
    name = str(row[sa_name_col]).strip()
    email = str(row[sa_email_col]).strip()
    
    # Jika nama ini belum ada di data Mapped, tambahkan ke list baru
    if name.lower() not in existing_sa_names and name.lower() not in ['nan', 'none', '']:
        new_rows.append({
            'id': str(uuid.uuid4()),  # Buat ID dummy/sementara untuk format excel
            'name': name,
            'email': email,
            'created_at': current_time,
            'updated_at': current_time
        })
        existing_sa_names.add(name.lower())

# Gabungkan data lama dengan data tim yang baru ditambahkan
if new_rows:
    df_new = pd.DataFrame(new_rows)
    df_sa = pd.concat([df_sa, df_new], ignore_index=True)
    print(f"Berhasil menambahkan {len(new_rows)} anggota tim baru ke dalam daftar.")

sheets['SolutionArchitects'] = df_sa

print(f"Menyimpan perubahan ke {FILE_MAPPED}...")
with pd.ExcelWriter(FILE_MAPPED, engine='openpyxl') as writer:
    for sheet_name, df in sheets.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        
print("Update Excel selesai! Seluruh anggota tim sudah masuk ke Data_Baru_Mapped.xlsx.")