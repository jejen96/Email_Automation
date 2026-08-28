"""
excel_reader.py
---------------
Tanggung jawab: Semua urusan baca dan tulis file Excel.
File ini TIDAK tahu cara mengirim email — tugasnya hanya mengelola data.

Mengapa file ini dibuat terpisah?
- Jika suatu hari sumber data berubah (dari Excel ke database/CSV/API),
  kita hanya perlu mengubah file ini tanpa menyentuh logika pengiriman email.
- Prinsip Separation of Concerns: pisahkan "mengambil data" dari "memproses data".

Hubungan dengan file lain:
- Diimpor oleh main.py untuk membaca data dan menyimpan hasil.
- Menggunakan konstanta dari config.py.
- Menggunakan fungsi dari utils.py.
- Menggunakan logger dari logger.py.
"""

from pathlib import Path
from typing import Generator, Optional, TypedDict

import pandas as pd

from attachment_validator import AttachmentValidationResult, validate_attachment, log_validation_error
from config import (
    EXCEL_FILE,
    REQUIRED_COLUMNS,
    STATUS_SENT,
)
from logger import get_logger
from utils import get_current_timestamp

logger = get_logger()


# =============================================================================
# TIPE DATA KUSTOM
# =============================================================================
# Menggunakan TypedDict bukan class biasa karena kita hanya butuh
# struktur data sederhana tanpa method. Lebih ringan dan lebih Pythonic.


class EmailRow(TypedDict):
    """
    Representasi satu baris data dari Excel.

    Mengapa TypedDict?
    Memberikan type hints yang jelas sehingga IDE bisa membantu autocomplete
    dan mendeteksi kesalahan sebelum program dijalankan.
    """
    index: int            # Nomor baris di DataFrame (untuk update nanti)
    nama: str             # Nama penerima
    email: str            # Alamat email penerima
    subject: str          # Subject email
    message: str          # Isi pesan
    attachment: Optional[Path]  # Path ke file lampiran (None jika tidak ada)


# =============================================================================
# FUNGSI UTAMA
# =============================================================================

def load_excel() -> pd.DataFrame:
    """
    Membaca file Excel dan mengembalikannya sebagai DataFrame pandas.

    DataFrame adalah seperti "tabel" di dalam Python — baris dan kolom,
    mirip spreadsheet, tapi bisa kita manipulasi dengan kode.

    Returns:
        pd.DataFrame: Data Excel yang sudah dimuat.

    Raises:
        FileNotFoundError: Jika file Excel tidak ditemukan.
        ValueError: Jika kolom wajib tidak lengkap.
    """
    logger.info(f"Membaca file Excel: {EXCEL_FILE}")

    # Cek apakah file ada sebelum mencoba membukanya
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"File Excel tidak ditemukan: {EXCEL_FILE}\n"
            f"Pastikan file data_email.xlsx ada di folder data/"
        )

    # Baca file Excel menggunakan pandas
    # engine="openpyxl" diperlukan untuk file .xlsx modern
    df = pd.read_excel(EXCEL_FILE, engine="openpyxl", dtype=str)

    # Bersihkan spasi di nama kolom (antisipasi typo di Excel)
    df.columns = df.columns.str.strip()

    logger.info(f"Berhasil membaca {len(df)} baris data")

    # Validasi kolom setelah membaca
    _validate_columns(df)

    return df


def _validate_columns(df: pd.DataFrame) -> None:
    """
    Memastikan semua kolom wajib tersedia di DataFrame.

    Mengapa fungsi terpisah?
    Agar load_excel() tetap ringkas dan mudah dibaca.
    Fungsi yang diawali _ (underscore) artinya "private" — hanya untuk
    digunakan di dalam file ini.

    Args:
        df: DataFrame yang akan divalidasi.

    Raises:
        ValueError: Jika ada kolom wajib yang hilang.
    """
    missing_columns = [col for col in REQUIRED_COLUMNS if col not in df.columns]

    if missing_columns:
        raise ValueError(
            f"Kolom wajib tidak ditemukan di file Excel: {missing_columns}\n"
            f"Pastikan file Excel memiliki kolom: {REQUIRED_COLUMNS}"
        )

    logger.info("Validasi kolom Excel: OK")


def iter_pending_rows(df: pd.DataFrame) -> Generator[EmailRow, None, None]:
    """
    Melakukan iterasi baris yang belum dikirim (Status bukan "Terkirim").

    Mengapa Generator?
    Generator menggunakan keyword 'yield' dan hanya memproses satu baris
    per iterasi — lebih hemat memori dibanding memuat semua baris sekaligus.
    Bayangkan seperti membaca buku halaman per halaman, bukan langsung
    menelan seluruh buku.

    Args:
        df: DataFrame yang sudah dimuat dari Excel.

    Yields:
        EmailRow: Data satu baris yang siap diproses.
    """
    for index, row in df.iterrows():
        # Lewati baris yang sudah berstatus Terkirim
        current_status = str(row.get("Status", "")).strip()
        if current_status == STATUS_SENT:
            logger.info(f"Baris {index + 1} ({row.get('Email', '')}) dilewati — sudah Terkirim")
            continue

        # Ambil data dari setiap kolom, ganti NaN dengan string kosong
        nama = _safe_str(row.get("Nama"))
        email = _safe_str(row.get("Email"))
        subject = _safe_str(row.get("Subject"))
        message = _safe_str(row.get("Message"))
        attachment_filename = _safe_str(row.get("Attachment"))

        # ------------------------------------------------------------------
        # VALIDASI ATTACHMENT — aturan bisnis utama:
        #
        # - Jika Attachment KOSONG  → lanjut, email dikirim tanpa lampiran.
        # - Jika Attachment DIISI   → file WAJIB ada.
        #   - File ada    → lampirkan dan kirim.
        #   - File tidak ada → BATALKAN email, catat error, lanjut ke baris berikut.
        #
        # Mengapa keputusan ini dibuat di sini (excel_reader) dan bukan di main?
        # Karena excel_reader yang paling tahu kondisi setiap baris data.
        # main.py cukup menerima baris yang sudah siap diproses atau
        # baris yang sudah ditandai gagal — tidak perlu tahu detailnya.
        # ------------------------------------------------------------------
        validation: AttachmentValidationResult = validate_attachment(attachment_filename)

        if not validation.is_valid:
            # File lampiran diminta tapi tidak ditemukan — batalkan email ini
            log_validation_error(logger, email, validation)
            # Update status langsung di DataFrame (tidak melalui main.py)
            _mark_attachment_error(df, int(index), validation.error_message)  # type: ignore[arg-type]
            # Lewati baris ini — TIDAK di-yield ke main.py
            continue

        yield EmailRow(
            index=int(index),  # type: ignore[arg-type]
            nama=nama,
            email=email,
            subject=subject,
            message=message,
            attachment=validation.file_path,  # None jika kosong, Path jika ada
        )


def update_row_success(df: pd.DataFrame, index: int) -> None:
    """
    Mengupdate baris di DataFrame setelah email berhasil dikirim.

    Args:
        df: DataFrame yang sedang diproses.
        index: Nomor baris yang akan diupdate.
    """
    df.at[index, "Status"] = "Terkirim"
    df.at[index, "Error"] = ""
    df.at[index, "Tanggal Kirim"] = get_current_timestamp()


def update_row_failed(df: pd.DataFrame, index: int, error_message: str) -> None:
    """
    Mengupdate baris di DataFrame setelah email gagal dikirim.

    Args:
        df: DataFrame yang sedang diproses.
        index: Nomor baris yang akan diupdate.
        error_message: Pesan error yang akan disimpan.
    """
    df.at[index, "Status"] = "Gagal"
    df.at[index, "Error"] = error_message
    df.at[index, "Tanggal Kirim"] = get_current_timestamp()


def save_excel(df: pd.DataFrame) -> None:
    """
    Menyimpan DataFrame kembali ke file Excel.

    Dipanggil setelah semua baris selesai diproses untuk menyimpan
    perubahan Status, Error, dan Tanggal Kirim.

    Args:
        df: DataFrame yang akan disimpan.
    """
    logger.info(f"Menyimpan perubahan ke file Excel: {EXCEL_FILE}")

    # index=False agar nomor baris pandas tidak ikut tersimpan sebagai kolom
    df.to_excel(EXCEL_FILE, index=False, engine="openpyxl")

    logger.info("File Excel berhasil disimpan")


# =============================================================================
# FUNGSI HELPER INTERNAL
# =============================================================================

def _safe_str(value) -> str:
    """
    Mengkonversi nilai dari pandas (termasuk NaN) menjadi string bersih.

    Pandas merepresentasikan sel kosong sebagai NaN (Not a Number),
    yang jika langsung dikonversi ke str menjadi "nan" — kita tidak mau itu.

    Args:
        value: Nilai dari sel Excel (bisa str, float NaN, None, dll).

    Returns:
        str: String bersih, atau "" jika nilai kosong/NaN.
    """
    if pd.isna(value):
        return ""
    return str(value).strip()


def _mark_attachment_error(
    df: pd.DataFrame,
    index: int,
    error_message: str,
) -> None:
    """
    Menandai baris sebagai Gagal karena attachment bermasalah,
    SEBELUM email dicoba dikirim.

    Dipanggil dari iter_pending_rows() saat validasi attachment gagal.
    Baris ini tidak akan di-yield ke main.py — langsung ditandai Gagal.

    Args:
        df           : DataFrame yang sedang diproses.
        index        : Nomor baris yang akan diupdate.
        error_message: Pesan error dari AttachmentValidationResult.
    """
    df.at[index, "Status"] = "Gagal"
    df.at[index, "Error"] = error_message
    df.at[index, "Tanggal Kirim"] = get_current_timestamp()
