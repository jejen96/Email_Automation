"""
logger.py
---------
Tanggung jawab: Menyiapkan dan mengelola sistem logging untuk seluruh project.

Mengapa file ini dibuat terpisah?
- Agar konfigurasi logging terpusat di satu tempat.
- File lain cukup memanggil get_logger() tanpa perlu tahu detail konfigurasinya.
- Memudahkan perubahan format log di masa depan (cukup ubah di sini).

Hubungan dengan file lain:
- Diimpor oleh semua file lain yang membutuhkan logging.
- Membaca LOGS_DIR dari config.py untuk menentukan lokasi file log.
"""

import logging
import sys
from datetime import datetime
from pathlib import Path

from config import LOGS_DIR


# =============================================================================
# CUSTOM LOG LEVEL: SUCCESS
# =============================================================================
# Python secara default hanya punya: DEBUG, INFO, WARNING, ERROR, CRITICAL
# Kita tambahkan level SUCCESS (antara INFO dan WARNING) untuk mencatat
# pengiriman email yang berhasil dengan lebih jelas.

SUCCESS_LEVEL: int = 25  # INFO=20, WARNING=30, jadi SUCCESS ada di antaranya
logging.addLevelName(SUCCESS_LEVEL, "SUCCESS")


def success(self: logging.Logger, message: str, *args, **kwargs) -> None:
    """Method tambahan agar bisa menulis logger.success('...')."""
    if self.isEnabledFor(SUCCESS_LEVEL):
        self._log(SUCCESS_LEVEL, message, args, **kwargs)


# Inject method success ke class Logger
logging.Logger.success = success  # type: ignore[attr-defined]


# =============================================================================
# SETUP LOGGER
# =============================================================================

def get_logger(name: str = "email_automation") -> logging.Logger:
    """
    Membuat dan mengembalikan logger yang sudah dikonfigurasi.

    Logger ini akan:
    - Menulis log ke file di folder logs/ (dengan nama berdasarkan tanggal).
    - Menulis log ke terminal (console) secara bersamaan.

    Args:
        name: Nama logger. Default "email_automation".

    Returns:
        logging.Logger: Logger yang siap digunakan.

    Cara pakai di file lain:
        from logger import get_logger
        logger = get_logger()
        logger.info("Program dimulai")
        logger.success("Email berhasil dikirim")
        logger.error("Terjadi kesalahan")
    """
    # Hindari duplikasi handler jika get_logger dipanggil lebih dari sekali
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger

    logger.setLevel(logging.DEBUG)

    # -------------------------------------------------------------------------
    # FORMAT LOG
    # Contoh output: 2026-06-30 09:10:01 | INFO     | Membaca file Excel
    # -------------------------------------------------------------------------
    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)-8s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # -------------------------------------------------------------------------
    # HANDLER 1: File Log
    # Setiap kali program dijalankan di hari berbeda, file log baru dibuat.
    # Format nama file: email_automation_2026-06-30.log
    # -------------------------------------------------------------------------
    LOGS_DIR.mkdir(parents=True, exist_ok=True)  # Buat folder logs/ jika belum ada
    log_filename: Path = LOGS_DIR / f"email_automation_{datetime.now().strftime('%Y-%m-%d')}.log"

    file_handler = logging.FileHandler(log_filename, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)

    # -------------------------------------------------------------------------
    # HANDLER 2: Console (Terminal)
    # Menampilkan log langsung di terminal saat program berjalan.
    # -------------------------------------------------------------------------
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)

    # Daftarkan kedua handler ke logger
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger
