"""
email_sender.py
---------------
Tanggung jawab: Semua urusan koneksi SMTP dan pengiriman email.
File ini TIDAK tahu dari mana data berasal — tugasnya hanya mengirim.

Mengapa menggunakan Class di sini?
- Koneksi SMTP perlu dibuka sekali dan digunakan berkali-kali (efisien).
- Dengan class, kita bisa menyimpan "state" (koneksi aktif) di antara
  pemanggilan method.
- Analogi: Class ini seperti "kurir" yang sudah siap berangkat —
  kita beri surat satu per satu, dia kirim satu per satu, tanpa
  perlu pulang ke kantor setiap kali.

Hubungan dengan file lain:
- Diimpor dan digunakan oleh main.py.
- Membaca konfigurasi SMTP dari config.py.
- Menggunakan logger dari logger.py.
"""

import smtplib
import ssl
from email.message import EmailMessage
from pathlib import Path
from typing import Optional

from config import (
    EMAIL_ADDRESS,
    EMAIL_PASSWORD,
    SMTP_PORT,
    SMTP_SERVER,
)
from logger import get_logger
from utils import is_valid_email, truncate_message

logger = get_logger()


class EmailSender:
    """
    Mengelola koneksi SMTP dan pengiriman email.

    Mengapa class dan bukan fungsi biasa?
    Koneksi ke SMTP server adalah operasi mahal (butuh waktu dan resource).
    Dengan class, kita buka koneksi SATU KALI, kirim semua email,
    lalu tutup koneksi. Jauh lebih efisien daripada buka-tutup
    koneksi untuk setiap email.

    Penggunaan dengan 'with' statement (context manager):
        with EmailSender() as sender:
            sender.send(...)
            sender.send(...)
        # Koneksi otomatis ditutup di sini

    Attributes:
        _smtp: Objek koneksi SMTP (None sebelum connect dipanggil).
    """

    def __init__(self) -> None:
        """Inisialisasi tanpa membuka koneksi — koneksi dibuka saat connect()."""
        self._smtp: Optional[smtplib.SMTP] = None

    def connect(self) -> None:
        """
        Membuka koneksi ke SMTP server dan melakukan login.

        Menggunakan STARTTLS: koneksi dimulai tidak terenkripsi,
        lalu di-upgrade ke TLS (Transport Layer Security).
        Ini adalah standar keamanan modern untuk email.

        Raises:
            smtplib.SMTPAuthenticationError: Jika email/password salah.
            smtplib.SMTPConnectError: Jika tidak bisa terhubung ke server.
            ConnectionError: Jika internet terputus.
        """
        logger.info(f"Menghubungkan ke SMTP server {SMTP_SERVER}:{SMTP_PORT}...")

        # ssl.create_default_context() membuat konteks SSL yang aman
        # dengan sertifikat yang sudah terverifikasi
        context = ssl.create_default_context()

        self._smtp = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        self._smtp.ehlo()          # Perkenalan ke server SMTP
        self._smtp.starttls(context=context)  # Upgrade ke koneksi terenkripsi
        self._smtp.ehlo()          # Perkenalan ulang setelah TLS
        self._smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

        logger.info(f"Berhasil login sebagai {EMAIL_ADDRESS}")

    def disconnect(self) -> None:
        """
        Menutup koneksi SMTP dengan aman.
        Selalu dipanggil di bagian 'finally' atau '__exit__'.
        """
        if self._smtp:
            try:
                self._smtp.quit()
                logger.info("Koneksi SMTP ditutup")
            except Exception:
                # Abaikan error saat menutup — program tetap harus selesai
                pass
            finally:
                self._smtp = None

    def send(
        self,
        to_email: str,
        to_name: str,
        subject: str,
        message: str,
        attachment: Optional[Path] = None,
    ) -> None:
        """
        Mengirim satu email ke satu penerima.

        Urutan kerja:
        1. Validasi format email
        2. Buat objek EmailMessage
        3. Isi header (From, To, Subject)
        4. Tambahkan body pesan
        5. Tambahkan lampiran jika ada
        6. Kirim via SMTP

        Args:
            to_email: Alamat email penerima.
            to_name: Nama penerima (untuk logging dan header email).
            subject: Subject email.
            message: Isi pesan email (plain text).
            attachment: Path ke file lampiran, atau None jika tidak ada.

        Raises:
            ValueError: Jika format email tidak valid atau email kosong.
            RuntimeError: Jika connect() belum dipanggil.
            smtplib.SMTPException: Jika terjadi error saat pengiriman.
        """
        # Pastikan koneksi sudah aktif
        if not self._smtp:
            raise RuntimeError(
                "Koneksi SMTP belum dibuka. Panggil connect() terlebih dahulu."
            )

        # Validasi email
        if not to_email:
            raise ValueError("Alamat email penerima tidak boleh kosong")

        if not is_valid_email(to_email):
            raise ValueError(f"Format email tidak valid: '{to_email}'")

        # Bangun objek email
        msg = self._build_message(to_email, to_name, subject, message, attachment)

        # Kirim email
        logger.info(f"Mengirim email ke {to_email} | Subject: {truncate_message(subject)}")
        self._smtp.sendmail(EMAIL_ADDRESS, to_email, msg.as_string())
        logger.success(f"BERHASIL | {to_email} | {to_name}")  # type: ignore[attr-defined]

    def _build_message(
        self,
        to_email: str,
        to_name: str,
        subject: str,
        message: str,
        attachment: Optional[Path],
    ) -> EmailMessage:
        """
        Merakit objek EmailMessage yang siap dikirim.

        Dipisah dari send() agar send() tetap ringkas dan
        fungsi ini bisa ditest secara mandiri.

        Args:
            to_email: Alamat email penerima.
            to_name: Nama penerima.
            subject: Subject email.
            message: Isi pesan.
            attachment: Path lampiran atau None.

        Returns:
            EmailMessage: Objek email yang sudah lengkap.
        """
        msg = EmailMessage()

        # Header email
        msg["From"] = f"{EMAIL_ADDRESS}"
        msg["To"] = to_email
        msg["Subject"] = subject

        # Body pesan — plain text
        # Personalisasi: sisipkan nama penerima di awal pesan
        personalized_message = f"Kepada Yth. {to_name},\n\n{message}"
        msg.set_content(personalized_message)

        # Tambahkan lampiran jika ada
        if attachment and attachment.exists():
            self._attach_file(msg, attachment)
        elif attachment:
            # Lampiran diminta tapi file tidak ditemukan — log warning
            logger.warning(
                f"File lampiran tidak ditemukan: {attachment} "
                f"— email dikirim tanpa lampiran"
            )

        return msg

    def _attach_file(self, msg: EmailMessage, file_path: Path) -> None:
        """
        Menambahkan file sebagai lampiran ke objek EmailMessage.

        Cara kerja: File dibaca sebagai binary (bytes), lalu di-encode
        ke format MIME yang bisa dikirim lewat email.

        Args:
            msg: Objek EmailMessage yang akan diberi lampiran.
            file_path: Path ke file yang akan dilampirkan.
        """
        # Deteksi tipe MIME berdasarkan ekstensi file
        # Contoh: .pdf → application/pdf, .jpg → image/jpeg
        import mimetypes

        mime_type, _ = mimetypes.guess_type(str(file_path))
        if mime_type:
            maintype, subtype = mime_type.split("/", 1)
        else:
            maintype, subtype = "application", "octet-stream"

        with open(file_path, "rb") as f:
            file_data = f.read()

        msg.add_attachment(
            file_data,
            maintype=maintype,
            subtype=subtype,
            filename=file_path.name,  # Nama file yang tampil di email
        )

        logger.info(f"Lampiran ditambahkan: {file_path.name}")

    # -------------------------------------------------------------------------
    # CONTEXT MANAGER
    # Memungkinkan penggunaan: with EmailSender() as sender:
    # __enter__ dipanggil saat masuk blok 'with'
    # __exit__ dipanggil saat keluar blok 'with' (bahkan jika ada exception)
    # -------------------------------------------------------------------------

    def __enter__(self) -> "EmailSender":
        """Dipanggil saat masuk blok 'with'. Membuka koneksi SMTP."""
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        """Dipanggil saat keluar blok 'with'. Menutup koneksi SMTP."""
        self.disconnect()
        # Return False/None agar exception (jika ada) tetap di-propagate
        return None
