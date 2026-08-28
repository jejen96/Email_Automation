"""
main.py
-------
Tanggung jawab: Titik masuk (entry point) program. Mengorkestrasi semua
modul lain untuk menjalankan alur program dari awal sampai selesai.

Analogi: main.py adalah "sutradara" — dia tidak berakting sendiri,
tapi mengatur semua "pemain" (modul lain) untuk tampil di waktu yang tepat.

Mengapa main.py harus SESIMPLE mungkin?
- Mudah dibaca: siapapun bisa memahami alur program hanya dengan
  membaca main.py tanpa harus masuk ke detail setiap modul.
- Mudah diubah: jika alur berubah, cukup ubah main.py.
- Single Responsibility: main.py hanya bertugas mengatur alur,
  bukan mengerjakan detail teknis.

Hubungan dengan file lain:
- Mengimpor dan menggunakan SEMUA modul lain.
- File ini yang dijalankan pengguna: python main.py
"""

import sys

import pandas as pd

from config import validate_config
from email_sender import EmailSender
from excel_formatter import format_excel
from excel_reader import (
    iter_pending_rows,
    load_excel,
    save_excel,
    update_row_failed,
    update_row_success,
)
from logger import get_logger

logger = get_logger()


# =============================================================================
# DATACLASS UNTUK RINGKASAN
# =============================================================================

class Summary:
    """
    Melacak statistik hasil pengiriman email.

    Mengapa class dan bukan variabel biasa?
    Lebih rapi dan mudah dikembangkan — jika nanti butuh tambahan
    statistik, cukup tambah attribute di sini.
    """

    def __init__(self) -> None:
        self.sent: int = 0      # Jumlah email berhasil dikirim
        self.failed: int = 0    # Jumlah email gagal
        self.skipped: int = 0   # Jumlah baris yang dilewati

    @property
    def total_processed(self) -> int:
        """Total baris yang diproses (tidak termasuk yang dilewati)."""
        return self.sent + self.failed

    def print_report(self) -> None:
        """Menampilkan ringkasan hasil di terminal."""
        separator = "=" * 50
        print(f"\n{separator}")
        print("       RINGKASAN PENGIRIMAN EMAIL")
        print(separator)
        print(f"  ✅  Berhasil Terkirim : {self.sent}")
        print(f"  ❌  Gagal             : {self.failed}")
        print(f"  ⏭️   Dilewati          : {self.skipped}")
        print(f"  📊  Total Diproses    : {self.total_processed}")
        print(separator)

        if self.failed > 0:
            print(
                "\n  ⚠️  Ada email yang gagal. Lihat kolom 'Error' di file Excel\n"
                "     dan file log di folder logs/ untuk detail.\n"
            )
        elif self.total_processed > 0:
            print("\n  🎉 Semua email berhasil dikirim!\n")


# =============================================================================
# FUNGSI UTAMA
# =============================================================================

def run() -> None:
    """
    Fungsi utama yang menjalankan seluruh alur program.

    Alur:
    1. Validasi konfigurasi (.env)
    2. Baca file Excel
    3. Koneksi ke SMTP server
    4. Iterasi setiap baris yang belum dikirim
    5. Kirim email satu per satu
    6. Update status di DataFrame
    7. Simpan perubahan ke Excel
    8. Tampilkan ringkasan
    """
    logger.info("=" * 50)
    logger.info("PROGRAM EMAIL AUTOMATION DIMULAI")
    logger.info("=" * 50)

    summary = Summary()

    # -------------------------------------------------------------------------
    # LANGKAH 1: Validasi konfigurasi
    # -------------------------------------------------------------------------
    try:
        validate_config()
    except ValueError as e:
        logger.error(f"Konfigurasi tidak valid: {e}")
        sys.exit(1)  # Hentikan program — tidak ada gunanya lanjut tanpa config

    # -------------------------------------------------------------------------
    # LANGKAH 2: Baca file Excel
    # -------------------------------------------------------------------------
    try:
        df: pd.DataFrame = load_excel()
    except FileNotFoundError as e:
        logger.error(str(e))
        sys.exit(1)
    except ValueError as e:
        logger.error(str(e))
        sys.exit(1)

    # Hitung baris yang sudah Terkirim sebelumnya (akan jadi 'skipped')
    from config import STATUS_SENT
    summary.skipped = int((df["Status"].astype(str).str.strip() == STATUS_SENT).sum())

    # -------------------------------------------------------------------------
    # LANGKAH 3 - 6: Koneksi SMTP dan kirim email
    #
    # Mengapa 'with EmailSender() as sender'?
    # - Koneksi dibuka sekali di awal (efisien)
    # - Jika terjadi error di tengah, koneksi tetap ditutup dengan aman
    #   karena __exit__ selalu dipanggil (seperti finally)
    # -------------------------------------------------------------------------
    try:
        with EmailSender() as sender:
            for row in iter_pending_rows(df):
                _process_row(sender, df, row, summary)

        # Hitung baris yang ditandai Gagal oleh attachment validator
        # (baris ini tidak masuk ke _process_row karena tidak di-yield)
        from config import STATUS_FAILED
        total_failed_in_df = int(
            (df["Status"].astype(str).str.strip() == STATUS_FAILED).sum()
        )
        # Attachment errors = total gagal di DataFrame - gagal dari SMTP
        attachment_errors = total_failed_in_df - summary.failed
        if attachment_errors > 0:
            summary.failed += attachment_errors

    except smtplib.SMTPAuthenticationError:
        logger.error(
            "Gagal login ke SMTP. Pastikan EMAIL_ADDRESS dan EMAIL_PASSWORD "
            "di file .env sudah benar. Untuk Gmail, gunakan App Password."
        )
        _save_and_report(df, summary)
        sys.exit(1)

    except (smtplib.SMTPConnectError, ConnectionError, OSError) as e:
        logger.error(f"Tidak dapat terhubung ke SMTP server: {e}")
        logger.error("Periksa koneksi internet dan pengaturan SMTP_SERVER/SMTP_PORT.")
        _save_and_report(df, summary)
        sys.exit(1)

    except Exception as e:
        # Tangkap error tak terduga agar file Excel tetap tersimpan
        logger.error(f"Error tidak terduga: {e}")
        _save_and_report(df, summary)
        raise

    # -------------------------------------------------------------------------
    # LANGKAH 7 & 8: Simpan dan tampilkan ringkasan
    # -------------------------------------------------------------------------
    _save_and_report(df, summary)


def _process_row(
    sender: EmailSender,
    df: pd.DataFrame,
    row: dict,
    summary: Summary,
) -> None:
    """
    Memproses satu baris data: kirim email dan update status.

    Mengapa fungsi terpisah?
    Agar loop utama di run() tetap bersih dan mudah dibaca.
    Prinsip: satu fungsi, satu tanggung jawab.

    Program TIDAK berhenti jika satu baris gagal — error dicatat
    dan program melanjutkan ke baris berikutnya.

    Args:
        sender: Objek EmailSender yang sudah terkoneksi.
        df: DataFrame untuk diupdate statusnya.
        row: Data satu baris dari iter_pending_rows().
        summary: Objek Summary untuk mencatat statistik.
    """
    import smtplib  # Import di sini untuk diakses di except

    try:
        sender.send(
            to_email=row["email"],
            to_name=row["nama"],
            subject=row["subject"],
            message=row["message"],
            attachment=row["attachment"],
        )
        update_row_success(df, row["index"])
        summary.sent += 1

    except ValueError as e:
        # Error validasi (email tidak valid, kosong, dll)
        error_msg = str(e)
        logger.warning(f"VALIDASI GAGAL | {row.get('email', 'N/A')} | {error_msg}")
        update_row_failed(df, row["index"], error_msg)
        summary.failed += 1

    except smtplib.SMTPRecipientsRefused as e:
        # Server menolak alamat penerima
        error_msg = f"Email penerima ditolak: {e}"
        logger.error(f"GAGAL | {row.get('email', 'N/A')} | {error_msg}")
        update_row_failed(df, row["index"], error_msg)
        summary.failed += 1

    except smtplib.SMTPException as e:
        # Error SMTP umum
        error_msg = f"SMTP Error: {e}"
        logger.error(f"GAGAL | {row.get('email', 'N/A')} | {error_msg}")
        update_row_failed(df, row["index"], error_msg)
        summary.failed += 1

    except Exception as e:
        # Error tidak terduga — catat tapi jangan hentikan program
        error_msg = f"Error tidak terduga: {e}"
        logger.error(f"GAGAL | {row.get('email', 'N/A')} | {error_msg}")
        update_row_failed(df, row["index"], error_msg)
        summary.failed += 1


def _save_and_report(df: pd.DataFrame, summary: Summary) -> None:
    """
    Menyimpan perubahan ke Excel, menerapkan formatting, dan menampilkan ringkasan.

    Urutan yang benar:
    1. save_excel()   → simpan DATA via pandas (cepat dan andal)
    2. format_excel() → terapkan STYLE via openpyxl (buka ulang file)
    3. print_report() → tampilkan ringkasan di terminal

    Mengapa save dulu baru format?
    Jika format_excel() gagal, data tetap tersimpan dengan benar.
    Formatting adalah bonus tampilan — tidak boleh mengorbankan data.

    Args:
        df     : DataFrame yang akan disimpan.
        summary: Objek Summary untuk ditampilkan.
    """
    try:
        save_excel(df)
    except Exception as e:
        logger.error(f"Gagal menyimpan file Excel: {e}")

    # Terapkan formatting SETELAH data tersimpan
    try:
        format_excel()
    except Exception as e:
        logger.warning(f"Formatting Excel gagal (data tetap aman): {e}")

    summary.print_report()
    logger.info("PROGRAM SELESAI")
    logger.info("=" * 50)


# =============================================================================
# ENTRY POINT
# =============================================================================
# Mengapa 'if __name__ == "__main__"'?
# Kondisi ini memastikan run() hanya dipanggil ketika file ini dijalankan
# langsung (python main.py), bukan ketika diimpor oleh file lain.
# Ini adalah konvensi standar Python.

if __name__ == "__main__":
    import smtplib  # noqa: F811 - dibutuhkan di sini untuk exception handling
    run()
