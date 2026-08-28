"""
attachment_validator.py
-----------------------
Tanggung jawab: Validasi keberadaan file lampiran sebelum email dikirim,
termasuk saran perbaikan jika nama file salah ketik (typo).

Mengapa modul ini dibuat terpisah dari utils.py?
- utils.py berisi fungsi-fungsi utilitas umum (validasi email, timestamp, dll).
- Logika validasi attachment sudah cukup kompleks — ada pencarian file,
  pencocokan nama mirip, dan pembentukan pesan error.
  Jika digabung ke utils.py, file itu akan terlalu besar dan sulit dibaca.
- Prinsip Single Responsibility: satu modul, satu tanggung jawab.
- Mudah dikembangkan: jika aturan validasi berubah, cukup ubah file ini.

Mengapa tidak mengubah EmailSender?
- EmailSender hanya bertugas mengirim — dia tidak perlu tahu urusan
  validasi file. Pisahkan concern agar kode tetap bersih.
- Validasi dilakukan SEBELUM EmailSender dipanggil, di lapisan excel_reader.

Hubungan dengan file lain:
- Diimpor oleh excel_reader.py di fungsi iter_pending_rows().
- Menggunakan get_search_locations() dari utils.py.
- TIDAK mengubah email_sender.py sama sekali.

Library yang digunakan:
- difflib (bawaan Python): untuk mencari nama file yang mirip.
- pathlib (bawaan Python): untuk operasi path file.
"""

import difflib
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional

from utils import get_search_locations, resolve_attachment_path


# =============================================================================
# DATACLASS HASIL VALIDASI
# =============================================================================

@dataclass
class AttachmentValidationResult:
    """
    Menyimpan hasil validasi satu file lampiran.

    Mengapa dataclass dan bukan dict biasa?
    - Lebih aman: IDE bisa mendeteksi jika kita salah tulis nama field.
    - Lebih jelas: pembaca kode langsung tahu field apa saja yang ada.
    - Lebih mudah dikembangkan: tinggal tambah field baru jika perlu.

    Attributes:
        is_valid     : True jika file ditemukan dan siap dilampirkan.
        file_path    : Path lengkap ke file jika ditemukan, None jika tidak.
        error_message: Pesan error lengkap untuk disimpan ke kolom Error Excel.
        suggestions  : Daftar nama file yang mirip (saran untuk operator).
        available_files: Semua file yang tersedia di folder pencarian.
    """
    is_valid: bool
    file_path: Optional[Path] = None
    error_message: str = ""
    suggestions: List[str] = field(default_factory=list)
    available_files: List[str] = field(default_factory=list)


# =============================================================================
# FUNGSI UTAMA: VALIDASI
# =============================================================================

def validate_attachment(filename: str) -> AttachmentValidationResult:
    """
    Memvalidasi file lampiran secara lengkap.

    Alur kerja:
    1. Jika filename kosong → valid (tidak ada lampiran, bukan error).
    2. Cari file menggunakan resolve_attachment_path() dari utils.py.
    3. Jika file ditemukan → valid, kembalikan path-nya.
    4. Jika file tidak ditemukan:
       a. Kumpulkan semua file dari semua folder pencarian.
       b. Cari nama yang mirip menggunakan difflib.
       c. Susun pesan error yang informatif.
       d. Kembalikan hasil invalid.

    Mengapa logika ini dipusatkan di sini?
    Agar excel_reader.py tidak perlu tahu detail cara validasi.
    excel_reader.py cukup memanggil validate_attachment() dan
    memeriksa is_valid-nya.

    Args:
        filename: Nama file atau path dari kolom Attachment di Excel.

    Returns:
        AttachmentValidationResult: Hasil validasi lengkap.

    Contoh:
        result = validate_attachment("lapora.pdf")
        if not result.is_valid:
            print(result.error_message)
            # Output:
            # Attachment tidak ditemukan: lapora.pdf
            # Saran:
            #   Apakah yang dimaksud:
            #   - laporan.pdf
            # Email dibatalkan.
    """
    # ------------------------------------------------------------------
    # KASUS 1: Filename kosong — tidak ada lampiran, ini valid.
    # Email tetap dikirim tanpa lampiran.
    # ------------------------------------------------------------------
    if not filename or not isinstance(filename, str) or filename.strip() == "":
        return AttachmentValidationResult(is_valid=True, file_path=None)

    filename = filename.strip()

    # ------------------------------------------------------------------
    # KASUS 2: Coba temukan file menggunakan fungsi pencarian dari utils.
    # Jika ketemu, langsung kembalikan hasil valid.
    # ------------------------------------------------------------------
    found_path = resolve_attachment_path(filename)
    if found_path is not None:
        return AttachmentValidationResult(is_valid=True, file_path=found_path)

    # ------------------------------------------------------------------
    # KASUS 3: File tidak ditemukan.
    # Mulai proses pencarian saran dan pembentukan pesan error.
    # ------------------------------------------------------------------

    # Kumpulkan semua file yang tersedia di semua folder pencarian
    all_available = _collect_available_files()

    # Cari nama yang mirip dengan yang diketik operator
    suggestions = _find_similar_names(filename, all_available)

    # Susun pesan error yang informatif
    error_message = _build_error_message(filename, suggestions, all_available)

    return AttachmentValidationResult(
        is_valid=False,
        file_path=None,
        error_message=error_message,
        suggestions=suggestions,
        available_files=all_available,
    )


# =============================================================================
# FUNGSI HELPER INTERNAL
# =============================================================================

def _collect_available_files() -> List[str]:
    """
    Mengumpulkan nama semua file yang ada di semua folder pencarian.

    Mengapa hanya nama file (bukan path lengkap)?
    - difflib.get_close_matches() bekerja dengan membandingkan string.
    - Operator mengetik nama file, bukan path lengkap.
    - Lebih mudah dibaca di pesan error.

    Returns:
        List[str]: Daftar nama file (tanpa path) dari semua lokasi pencarian.
                   Duplikat dihilangkan. Diurutkan alfabetis.
    """
    available: List[str] = []

    for folder in get_search_locations():
        # Lewati folder yang tidak ada di komputer ini
        if not folder.exists():
            continue

        # Ambil semua file (bukan subfolder) di folder ini
        # iterdir() menghasilkan semua item, is_file() menyaring hanya file
        for item in folder.iterdir():
            if item.is_file():
                # Tambahkan hanya jika belum ada (hindari duplikat)
                if item.name not in available:
                    available.append(item.name)

    # Urutkan agar mudah dibaca di log
    return sorted(available)


def _find_similar_names(
    filename: str,
    available_files: List[str],
    max_suggestions: int = 3,
    cutoff: float = 0.6,
) -> List[str]:
    """
    Mencari nama file yang mirip menggunakan difflib.get_close_matches().

    Cara kerja difflib:
    - Menghitung "similarity ratio" antara dua string (0.0 sampai 1.0).
    - Contoh: "lapora.pdf" vs "laporan.pdf" → ratio ~0.95 (sangat mirip)
    - Contoh: "laporan.pdf" vs "gambar.jpg" → ratio ~0.3 (tidak mirip)
    - Parameter cutoff: hanya tampilkan jika ratio >= cutoff.
    - Parameter n: maksimal berapa saran yang ditampilkan.

    Mengapa cutoff=0.6?
    Nilai ini cukup ketat untuk menghindari saran yang tidak relevan,
    tapi cukup longgar untuk menangkap typo umum (huruf kurang/lebih,
    typo satu huruf, dll).

    Args:
        filename       : Nama file yang dicari (kemungkinan typo).
        available_files: Daftar nama file yang tersedia.
        max_suggestions: Maksimal berapa saran ditampilkan.
        cutoff         : Minimum similarity (0.0 - 1.0).

    Returns:
        List[str]: Daftar nama file yang mirip, diurutkan dari paling mirip.
    """
    if not available_files:
        return []

    return difflib.get_close_matches(
        word=filename,
        possibilities=available_files,
        n=max_suggestions,
        cutoff=cutoff,
    )


def _build_error_message(
    filename: str,
    suggestions: List[str],
    available_files: List[str],
) -> str:
    """
    Menyusun pesan error yang informatif dan mudah dipahami operator.

    Pesan disusun dalam format yang konsisten:
    - Baris 1: nama file yang tidak ditemukan
    - Baris 2+: saran jika ada nama mirip
    - Baris terakhir: instruksi tindakan

    Mengapa pesan error harus informatif?
    Operator bukan programmer. Pesan "File not found" tidak membantu.
    Pesan yang baik langsung memberi tahu apa yang salah dan apa yang harus
    dilakukan — mengurangi kebingungan dan mempercepat perbaikan.

    Args:
        filename       : Nama file yang tidak ditemukan.
        suggestions    : Daftar nama file yang mirip (bisa kosong).
        available_files: Semua file tersedia (untuk fallback jika tidak ada saran).

    Returns:
        str: Pesan error yang siap ditampilkan di log dan disimpan ke Excel.
    """
    lines = []

    # Baris 1: Identifikasi masalah
    lines.append(f"Attachment tidak ditemukan: {filename}")

    if suggestions:
        # Ada nama yang mirip — tampilkan sebagai saran
        lines.append("Saran — Apakah yang dimaksud:")
        for suggestion in suggestions:
            lines.append(f"  - {suggestion}")

    elif available_files:
        # Tidak ada yang mirip tapi ada file lain — tampilkan semua
        # Batasi 10 file agar pesan tidak terlalu panjang
        lines.append("File yang tersedia di folder pencarian:")
        for f in available_files[:10]:
            lines.append(f"  - {f}")
        if len(available_files) > 10:
            lines.append(f"  ... dan {len(available_files) - 10} file lainnya")

    else:
        # Tidak ada file sama sekali di semua folder pencarian
        lines.append("Tidak ada file di folder attachments/.")
        lines.append("Taruh file lampiran di folder: attachments/")

    # Baris terakhir: instruksi
    lines.append("Email dibatalkan. Perbaiki nama file di Excel lalu jalankan ulang.")

    return "\n".join(lines)


def log_validation_error(
    logger,
    email: str,
    result: AttachmentValidationResult,
) -> None:
    """
    Menulis hasil validasi gagal ke logger dengan format yang mudah dibaca.

    Mengapa fungsi terpisah untuk logging?
    - Agar format log konsisten di seluruh program.
    - excel_reader.py tidak perlu tahu detail format log.
    - Jika format log perlu diubah, cukup ubah di sini.

    Args:
        logger: Objek logger dari modul logger.py.
        email : Alamat email penerima (untuk konteks di log).
        result: Hasil validasi dari validate_attachment().
    """
    # Tampilkan setiap baris pesan error sebagai baris log terpisah
    # agar mudah dibaca di terminal maupun di file log
    separator = "-" * 50
    logger.error(separator)
    logger.error(f"ATTACHMENT ERROR | {email}")
    for line in result.error_message.splitlines():
        logger.error(f"  {line}")
    logger.error(separator)
