"""
config.py
---------
Tanggung jawab: Membaca semua konfigurasi dari file .env dan mendefinisikan
konstanta yang digunakan di seluruh project.

Mengapa file ini dibuat terpisah?
- Satu tempat untuk semua konfigurasi (Single Source of Truth).
- Jika ada perubahan konfigurasi, cukup ubah di sini.
- File lain tidak perlu tahu dari mana konfigurasi berasal.
- Memudahkan testing (bisa mock konfigurasi).

Hubungan dengan file lain:
- Diimpor oleh email_sender.py untuk mendapatkan kredensial SMTP.
- Diimpor oleh excel_reader.py untuk mendapatkan path file Excel.
- Diimpor oleh logger.py untuk mendapatkan path folder logs.
"""

import os
from pathlib import Path
from dotenv import load_dotenv

# =============================================================================
# LOAD FILE .ENV
# =============================================================================
# Mengapa load_dotenv()?
# Fungsi ini membaca file .env dan memasukkan isinya ke dalam os.environ
# sehingga bisa diakses dengan os.getenv() di mana saja dalam program.

# Path ke root directory project (folder Email_Automation/)
BASE_DIR: Path = Path(__file__).parent

# Load file .env dari root directory
load_dotenv(BASE_DIR / ".env")


# =============================================================================
# KONFIGURASI SMTP (Email)
# =============================================================================

EMAIL_ADDRESS: str = os.getenv("EMAIL_ADDRESS", "")
EMAIL_PASSWORD: str = os.getenv("EMAIL_PASSWORD", " ")
SMTP_SERVER: str = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT: int = int(os.getenv("SMTP_PORT", "587"))


# =============================================================================
# PATH FOLDER & FILE
# =============================================================================

DATA_DIR: Path = BASE_DIR / "data"
LOGS_DIR: Path = BASE_DIR / "logs"
ATTACHMENTS_DIR: Path = BASE_DIR / "attachments"
EXCEL_FILE: Path = DATA_DIR / "data_email.xlsx"


# =============================================================================
# KOLOM WAJIB EXCEL
# =============================================================================
# Didefinisikan di sini agar mudah diubah jika nama kolom berubah.
# excel_reader.py akan memvalidasi berdasarkan daftar ini.

REQUIRED_COLUMNS: list = [
    "Nama",
    "Email",
    "Subject",
    "Message",
    "Attachment",
    "Status",
    "Error",
    "Tanggal Kirim",
]


# =============================================================================
# KONSTANTA STATUS
# =============================================================================
# Menggunakan konstanta menghindari typo saat menulis string berulang kali.
# Contoh: lebih aman menulis STATUS_SENT daripada "Terkirim" berulang kali.

STATUS_SENT: str = "Terkirim"
STATUS_FAILED: str = "Gagal"
STATUS_SKIPPED: str = "Dilewati"


# =============================================================================
# VALIDASI KONFIGURASI
# =============================================================================

def validate_config() -> None:
    """
    Memvalidasi bahwa semua konfigurasi wajib sudah diisi.

    Tujuan:
        Memberikan pesan error yang jelas jika file .env belum diisi,
        sehingga program tidak gagal di tengah jalan dengan pesan yang membingungkan.

    Raises:
        ValueError: Jika EMAIL_ADDRESS atau EMAIL_PASSWORD kosong.
    """
    if not EMAIL_ADDRESS:
        raise ValueError(
            "EMAIL_ADDRESS belum diisi di file .env\n"
            "Salin .env.example menjadi .env lalu isi dengan data Anda."
        )
    if not EMAIL_PASSWORD:
        raise ValueError(
            "EMAIL_PASSWORD belum diisi di file .env\n"
            "Gunakan App Password dari Google, bukan password akun biasa."
        )
