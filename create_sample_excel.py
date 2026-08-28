"""
create_sample_excel.py
----------------------
Script pembantu untuk membuat file data_email.xlsx contoh.
Jalankan sekali: python create_sample_excel.py

Script ini BUKAN bagian dari program utama.
Tujuannya hanya untuk memudahkan setup awal.
"""

from pathlib import Path
import pandas as pd

DATA_DIR = Path(__file__).parent / "data"
DATA_DIR.mkdir(exist_ok=True)

sample_data = {
    "Nama":         ["Andi Saputra", "Budi Santoso", "Citra Dewi", "Dian Pratama"],
    "Email":        ["andi@example.com", "budi@example.com", "email-tidak-valid", "dian@example.com"],
    "Subject":      ["Undangan Rapat", "Laporan Bulanan", "Info Penting", "Konfirmasi Data"],
    "Message":      [
        "Dengan hormat, kami mengundang Anda untuk hadir dalam rapat bulanan.",
        "Terlampir laporan bulanan untuk bulan Juni 2026.",
        "Berikut informasi penting yang perlu Anda ketahui.",
        "Mohon konfirmasi data Anda sebelum tanggal 5 Juli 2026.",
    ],
    "Attachment":   ["", "", "", ""],
    "Status":       ["", "", "", ""],
    "Error":        ["", "", "", ""],
    "Tanggal Kirim": ["", "", "", ""],
}

df = pd.DataFrame(sample_data)
output_path = DATA_DIR / "data_email.xlsx"
df.to_excel(output_path, index=False, engine="openpyxl")

print(f"✅ File contoh berhasil dibuat: {output_path}")
print(f"   Total baris: {len(df)}")
print("\nCatatan: 'email-tidak-valid' sengaja dibuat salah untuk menguji validasi.")
