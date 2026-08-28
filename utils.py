"""
utils.py
--------
Tanggung jawab: Menyediakan fungsi-fungsi utilitas kecil yang digunakan
oleh berbagai file lain dalam project.

Mengapa file ini dibuat terpisah?
- Fungsi-fungsi kecil yang tidak secara spesifik milik satu modul
  dikumpulkan di sini agar tidak mengotori file lain.
- Prinsip DRY (Don't Repeat Yourself): tulis sekali, pakai di mana saja.
- Mudah di-unit test karena fungsinya sederhana dan tidak punya side effect.

Hubungan dengan file lain:
- Diimpor oleh excel_reader.py dan email_sender.py sesuai kebutuhan.
"""

import re
from datetime import datetime
from pathlib import Path
from typing import Optional


def is_valid_email(email: str) -> bool:
    """
    Memvalidasi apakah sebuah string adalah format email yang valid.

    Mengapa perlu validasi?
    Mengirim ke alamat email yang salah format akan menyebabkan SMTP error
    yang membingungkan. Lebih baik kita deteksi lebih awal.

    Regex yang digunakan adalah pola standar untuk email:
    - Bagian lokal: huruf, angka, titik, underscore, plus, minus
    - Tanda @
    - Domain: huruf, angka, titik, minus
    - TLD: minimal 2 karakter (seperti .com, .id, .co.id)

    Args:
        email: String yang akan divalidasi.

    Returns:
        bool: True jika valid, False jika tidak.

    Contoh:
        >>> is_valid_email("andi@gmail.com")   # True
        >>> is_valid_email("bukan-email")      # False
        >>> is_valid_email("")                 # False
    """
    if not email or not isinstance(email, str):
        return False

    pattern = r"^[\w\.\+\-]+@[\w\-]+\.[\w\.\-]{2,}$"
    return bool(re.match(pattern, email.strip()))


def get_current_timestamp() -> str:
    """
    Mengembalikan timestamp saat ini dalam format yang mudah dibaca.

    Format: YYYY-MM-DD HH:MM:SS
    Contoh: 2026-06-30 09:10:05

    Returns:
        str: Timestamp sebagai string.
    """
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def resolve_attachment_path(filename: str) -> Optional[Path]:
    """
    Mencari file lampiran secara otomatis di beberapa lokasi umum.

    Mengapa fungsi ini ditingkatkan?
    Sebelumnya pengguna harus menaruh file di folder attachments/ saja.
    Sekarang program mencari sendiri di lokasi yang paling umum digunakan,
    sehingga pengguna cukup menulis nama file di Excel tanpa perlu
    memikirkan di mana harus menaruh file tersebut.

    URUTAN PENCARIAN:
    1. Jika filename adalah absolute path (misal D:\\Dokumen\\file.pdf)
       -> gunakan langsung tanpa perlu mencari
    2. folder attachments/ di dalam project
    3. Desktop pengguna
    4. Documents / Dokumen pengguna
    5. Downloads / Unduhan pengguna

    Mengapa urutan ini?
    - Absolute path paling prioritas karena pengguna sudah spesifik.
    - attachments/ adalah folder resmi project, paling diutamakan.
    - Desktop, Documents, Downloads adalah lokasi paling sering dipakai.

    Args:
        filename: Nama file atau absolute path dari kolom Excel.
                  Contoh: "laporan.pdf" atau "D:\\Dokumen\\laporan.pdf"

    Returns:
        Optional[Path]: Path lengkap ke file jika ditemukan,
                        None jika tidak ditemukan atau filename kosong.
    """
    # ------------------------------------------------------------------
    # GUARD CLAUSE: Jika filename kosong, langsung kembalikan None.
    # Kolom Attachment memang boleh kosong — ini bukan error.
    # ------------------------------------------------------------------
    if not filename or not isinstance(filename, str) or filename.strip() == "":
        return None

    # Bersihkan spasi di awal/akhir yang mungkin tidak sengaja diketik
    filename = filename.strip()

    # ------------------------------------------------------------------
    # LANGKAH 1: Cek apakah filename adalah absolute path.
    #
    # Contoh absolute path: D:\Dokumen\laporan.pdf
    # Contoh bukan absolute: laporan.pdf
    #
    # Jika absolute, cek langsung — tidak perlu cari ke folder lain.
    # ------------------------------------------------------------------
    candidate = Path(filename)
    if candidate.is_absolute():
        if candidate.exists():
            return candidate
        return None  # Path diberikan tapi file tidak ada

    # ------------------------------------------------------------------
    # LANGKAH 2: Bangun daftar folder pencarian secara berurutan.
    #
    # Path.home() otomatis menemukan folder home pengguna:
    # Windows -> C:\Users\NamaUser
    #
    # Path(__file__).parent -> folder tempat utils.py berada (root project)
    # ------------------------------------------------------------------
    home = Path.home()
    project_attachments = Path(__file__).parent / "attachments"

    search_locations = [
        project_attachments,     # 1. attachments/ di dalam project
        home / "Desktop",        # 2. Desktop
        home / "Documents",      # 3. Documents (Windows English)
        home / "Downloads",      # 4. Downloads (Windows English)
        home / "Dokumen",        # 5. Dokumen (Windows Indonesia)
        home / "Unduhan",        # 6. Unduhan (Windows Indonesia)
    ]

    # ------------------------------------------------------------------
    # LANGKAH 3: Cari file di setiap folder secara berurutan.
    # Berhenti dan kembalikan path begitu file pertama ditemukan.
    # ------------------------------------------------------------------
    for folder in search_locations:
        candidate = folder / filename
        if candidate.exists():
            return candidate

    # ------------------------------------------------------------------
    # LANGKAH 4: File tidak ditemukan di semua lokasi -> kembalikan None.
    # ------------------------------------------------------------------
    return None


def get_search_locations() -> list:
    """
    Mengembalikan daftar folder yang dicari oleh resolve_attachment_path.

    Berguna untuk ditampilkan di pesan warning/error agar pengguna
    tahu di mana seharusnya menaruh file lampiran.

    Returns:
        list: Daftar objek Path folder pencarian.
    """
    home = Path.home()
    project_attachments = Path(__file__).parent / "attachments"

    return [
        project_attachments,
        home / "Desktop",
        home / "Documents",
        home / "Downloads",
        home / "Dokumen",
        home / "Unduhan",
    ]


def truncate_message(message: str, max_length: int = 50) -> str:
    """
    Memotong pesan panjang untuk keperluan logging agar tidak memenuhi layar.

    Args:
        message: Pesan yang akan dipotong.
        max_length: Panjang maksimal sebelum dipotong. Default 50.

    Returns:
        str: Pesan yang sudah dipotong dengan "..." jika melebihi batas.

    Contoh:
        >>> truncate_message("Ini adalah pesan yang sangat panjang sekali", 20)
        "Ini adalah pesan yan..."
    """
    if len(message) <= max_length:
        return message
    return message[:max_length] + "..."
