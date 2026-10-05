"""Antarmuka baris perintah pipeline tutor online Universitas Terbuka.

Pemakaian biasa:

    python main.py

Perintah itu menampilkan daftar mata kuliah, menanyakan nomor sesi, lalu
menjalankan seluruh tahap sampai berkas .docx selesai.

Bisa juga langsung tanpa tanya jawab:

    python main.py --sesi 2
    python main.py --daftar
"""

from __future__ import annotations

import argparse
import re
import sys

import config
import courses
import moodle
from pipeline import Pipeline, PipelineGagal

GARIS = "=" * 68


def _pilih_dari_daftar(judul: str, opsi: list[str]) -> int | None:
    """Tampilkan daftar pilihan dan baca nomor dari pengguna."""
    print(f"\n{judul}")
    for i, teks in enumerate(opsi, 1):
        print(f"  {i:>2}. {teks}")
    print()
    while True:
        jawab = input("Masukkan nomor (atau ketik 'batal'): ").strip()
        if not jawab:
            continue
        if jawab.lower() in {"batal", "b", "q", "quit", "exit", "0"}:
            return None
        if jawab.isdigit() and 1 <= int(jawab) <= len(opsi):
            return int(jawab) - 1
        print("Nomor di luar daftar.")


def _cari_mata_kuliah(
    daftar: list[courses.MataKuliah], teks: str
) -> courses.MataKuliah | None:
    """Temukan satu mata kuliah dari kode, kelas, atau sebagian nama."""
    t = teks.strip().lower()
    if not t:
        return None
    persis = [
        m for m in daftar
        if m.kode.lower() == t or f"{m.kode}.{m.kelas}".lower() == t
    ]
    if len(persis) == 1:
        return persis[0]
    cocok = [m for m in daftar if t in m.nama.lower() or t in m.slug.lower()]
    if len(cocok) == 1:
        return cocok[0]
    return None


def _daftar_pilihan_sesi(
    klien: moodle.Moodle, matkul: courses.MataKuliah
) -> list[tuple[int, int]]:
    """Section mana yang benar-benar memuat soal, untuk ditampilkan."""
    return courses.sesi_berisi_soal(klien, matkul)


def tampilkan_daftar(klien: moodle.Moodle) -> list[courses.MataKuliah]:
    """Tampilkan seluruh mata kuliah beserta sesi yang punya pekerjaan."""
    daftar = courses.daftar_mata_kuliah_lengkap(klien)
    print(GARIS)
    print("MATA KULIAH")
    print(GARIS)
    for i, mk in enumerate(daftar, 1):
        kode = f"{mk.kode}.{mk.kelas}" if mk.kode and mk.kelas else (mk.kode or "-")
        print(f"  {i:>2}. {mk.nama}")
        print(f"      kode {kode}   id {mk.id}")
        try:
            sesi = _daftar_pilihan_sesi(klien, mk)
        except moodle.MoodleError as exc:
            print(f"      (gagal membaca section: {exc})")
            continue
        ada = [s for s, n in sesi if n > 0 and s > 0]
        if ada:
            rincian = []
            for s, n in sesi:
                if s > 0 and n > 0:
                    rincian.append(f"{s}({n})")
            print(f"      sesi dengan soal: {', '.join(rincian)}")
        else:
            print("      sesi dengan soal: belum ada")
    print(GARIS)
    return daftar


def main() -> int:
    ap = argparse.ArgumentParser(
        prog="main.py",
        description=(
            "Mengambil soal satu sesi tutorial online Universitas Terbuka, "
            "mengerjakannya dengan agent AI, lalu membuat berkas .docx."
        ),
    )
    ap.add_argument("--sesi", type=int, help="Nomor sesi, tanpa bertanya.")
    ap.add_argument(
        "--matkul",
        help="Kode, kelas, atau sebagian nama mata kuliah. Tanpa ini "
             "pengguna memilih dari daftar.",
    )
    ap.add_argument(
        "--daftar", action="store_true",
        help="Tampilkan daftar mata kuliah dan sesi yang punya soal, lalu berhenti.",
    )
    ap.add_argument(
        "--tanpa-docx", action="store_true",
        help="Berhenti setelah jawaban tertulis; jangan membuat dokumen.",
    )
    ap.add_argument(
        "--tanpa-gambar", action="store_true",
        help="Lewati pembacaan gambar lampiran oleh model penglihatan.",
    )
    ap.add_argument("--model", help="Model agent utama.")
    ap.add_argument("--model-mata", help="Model penglihatan.")
    args = ap.parse_args()

    config.ensure_dirs()

    klien = moodle.Moodle()
    masuk, pesan = klien.cek_login()
    if not masuk:
        print(f"\nGagal masuk: {pesan}\n", file=sys.stderr)
        return 2

    if args.daftar:
        tampilkan_daftar(klien)
        return 0

    daftar = courses.daftar_mata_kuliah_lengkap(klien)

    matkul: courses.MataKuliah | None = None
    if args.matkul:
        matkul = _cari_mata_kuliah(daftar, args.matkul)
        if matkul is None:
            print(
                f"\nMata kuliah {args.matkul!r} tidak ditemukan. "
                "Jalankan `python main.py --daftar` untuk melihat daftarnya.\n",
                file=sys.stderr,
            )
            return 2
        print(f"Mata kuliah: {matkul.label}")
    else:
        print("\nPilih mata kuliah:")
        for i, mk in enumerate(daftar, 1):
            kode = f" [{mk.kode}.{mk.kelas}]" if mk.kode and mk.kelas else ""
            print(f"  {i:>2}. {mk.nama}{kode}")

        jawab = input("\nMasukkan nomor mata kuliah: ").strip()
        if not jawab.isdigit() or not 1 <= int(jawab) <= len(daftar):
            print("Nomor di luar daftar.", file=sys.stderr)
            return 2
        matkul = daftar[int(jawab) - 1]
        print(f"\nMata kuliah: {matkul.label}")

    nomor = args.sesi
    if nomor is None:
        try:
            sesi_tersedia = [
                (s, n) for s, n in _daftar_pilihan_sesi(klien, matkul) if s > 0
            ]
        except moodle.MoodleError as exc:
            print(f"Gagal membaca section: {exc}", file=sys.stderr)
            return 2

        opsi = [
            f"Sesi {s} - {n} soal" + ("" if n else "  (tidak ada soal)")
            for s, n in sesi_tersedia
        ]
        if not opsi:
            print("Course ini tidak punya section selain halaman depan.", file=sys.stderr)
            return 2

        pilih = _pilih_dari_daftar(f"Pilih sesi untuk {matkul.nama}:", opsi)
        if pilih is None:
            print("Dibatalkan.")
            return 0
        nomor = sesi_tersedia[pilih][0]

    try:
        pipe = Pipeline(
            matkul,
            nomor,
            model_utama=args.model,
            model_mata=args.model_mata,
            dengan_gambar=not args.tanpa_gambar,
            tanpa_docx=args.tanpa_docx,
        )
        hasil = pipe.jalankan()
    except PipelineGagal as exc:
        print(f"\nPipeline berhenti: {exc}\n", file=sys.stderr)
        return 1
    except moodle.LoginMati as exc:
        print(f"\nSesi Moodle berakhir: {exc}\n", file=sys.stderr)
        return 2
    except KeyboardInterrupt:
        print("\nDihentikan pengguna.", file=sys.stderr)
        return 130

    print()
    if hasil.gagal:
        print("Catatan selama pipeline:")
        for g in hasil.gagal:
            print(f"  - {g}")
    if hasil.docx:
        print(f"Berkas jawaban: {hasil.docx}")
    if args.tanpa_docx and hasil.jawaban:
        print(f"Berkas markdown: {hasil.jawaban}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.exit(main())