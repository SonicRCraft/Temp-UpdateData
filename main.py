import pandas as pd
from openpyxl.styles import Font
from sqlalchemy import create_engine
from dotenv import load_dotenv

load_dotenv()

file_path = os.getenv("EXCEL_SOURCE_PATH")
output_path = os.getenv("EXCEL_OUTPUT_PATH")

# 1. Membaca & Membersihkan Data
df_karyawan = pd.read_excel(file_path, sheet_name='List Initial', usecols="A:C", header=1)
df_customer = pd.read_excel(file_path, sheet_name='List Initial', usecols="E:G", header=1)
df_proptek = pd.read_excel(file_path, sheet_name='Proptek')

df_karyawan = df_karyawan.dropna(how='all')
df_customer = df_customer.dropna(how='all')
df_proptek = df_proptek.dropna(how='all')

# =======================================================
# 2. PROSES MEMASUKKAN DATA KE POSTGRESQL
# =======================================================

# Konfigurasi Database (Ganti dengan kredensial PostgreSQL kamu)
DB_USER = os.getenv("DB_USER")
DB_PASS = os.getenv("DB_PASS")
DB_HOST = os.getenv("DB_HOST")
DB_PORT = os.getenv("DB_PORT")
DB_NAME = os.getenv("DB_NAME")

# Membuat jembatan koneksi (engine)
engine = create_engine(f'postgresql://{DB_USER}:{DB_PASS}@{DB_HOST}:{DB_PORT}/{DB_NAME}')

print("Memulai proses transfer ke PostgreSQL...")
try:
    # Memasukkan masing-masing dataframe ke tabel terpisah
    # if_exists='replace' -> akan menghapus tabel lama dan menggantinya dengan yang baru
    # if_exists='append'  -> akan menambahkan data baru ke bawah data lama yang sudah ada di tabel
    
    df_karyawan.to_sql('inisial_karyawan', engine, if_exists='replace', index=False)
    df_customer.to_sql('inisial_customer', engine, if_exists='replace', index=False)
    df_proptek.to_sql('data_proptek', engine, if_exists='replace', index=False)
    
    print("Berhasil! Semua data telah masuk ke database PostgreSQL.")
except Exception as e:
    print(f"Gagal memasukkan data ke database. Error: {e}")

# =======================================================
# 3. PROSES MENYIMPAN & MERAPIKAN EXCEL (Tetap dipertahankan)   
# =======================================================
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    df_karyawan.to_excel(writer, sheet_name='Inisial Karyawan', index=False)
    df_customer.to_excel(writer, sheet_name='Inisial Customer', index=False)
    df_proptek.to_excel(writer, sheet_name='Proptek', index=False)
    
    for sheetname in writer.sheets:
        worksheet = writer.sheets[sheetname]
        for column_cells in worksheet.columns:
            max_length = 0
            column_letter = column_cells[0].column_letter 
            for cell in column_cells:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            worksheet.column_dimensions[column_letter].width = max_length + 2
            
        for cell in worksheet[1]:
            cell.font = Font(bold=True)

print("Berhasil! File Excel telah diperbarui dan dirapikan.")