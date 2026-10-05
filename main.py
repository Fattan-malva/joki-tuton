"""Antarmuka baris perintah pipeline tutor online Universitas Terbuka.

Pemakaian biasa:

    python main.py

Perintah itu menampilkan daftar mata kuliah, menanyakan nomor sesi, lalu
menjalankan seluruh tahap sampai berkas .docx selesai.

Bisa juga langsung tanpa tanya jawab:

    python main.py --sesi 2
    python main.py --daftar

Untuk memproses seluruh mata kuliah pada satu sesi, berurutan:

    python main.py --semua --sesi 4

Nomor sesi yang tersedia berbeda tiap mata kuliah, jadi `--semua` tanpa
`--sesi` menampilkan sesi mana saja yang punya soal di satu atau lebih
mata kuliah, lengkap dengan mata kuliah mana yang akan ikut.
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

# Kata yang sama-sama berarti "kerjakan semuanya".
_KATA_SEMUA = {"semua", "all", "*"}


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


# ------------------------------------------------------------- banyak matkul


def _sesi_per_mata_kuliah(
    klien: moodle.Moodle, daftar: list[courses.MataKuliah]
) -> dict[str, set[int]]:
    """Section yang punya soal, per mata kuliah.

    Nilainya dipakai dua kali: untuk menampilkan pilihan sesi yang benar saat
    memproses banyak mata kuliah, dan untuk melewati mata kuliah yang memang
    tidak punya sesi itu. Pengecekan di awal penting -- pipeline yang baru tahu
    sesinya kosong setelah semua agent dipanggil adalah pemborosan yang mahal.
    """
    hasil: dict[str, set[int]] = {}
    for mk in daftar:
        try:
            hasil[mk.id] = {
                s for s in courses.daftar_section(klien, mk) if s > 0
            }
        except moodle.MoodleError as exc:
            print(f"  ! Gagal membaca section {mk.nama}: {exc}")
            # Kosongkan supaya sesi ini tidak dipilih dari daftar. Melewati
            # lebih baik daripada membuat satu sesi-batch gagal diam-diam.
            hasil[mk.id] = set()
    return hasil


def _pilih_sesi_batch(
    klien: moodle.Moodle, daftar: list[courses.MataKuliah]
) -> tuple[int, dict[str, set[int]]] | None:
    """Tanyakan satu nomor sesi untuk semua mata kuliah.

    Yang ditampilkan hanya sesi yang ada di minimal satu mata kuliah, karena
    sesi yang tidak ada di mana pun hanya menambah satu tahap yang pasti
    terlewat. Jumlah mata kuliah per sesi ikut ditulis supaya pengguna tahu
    berapa banyak yang akan dijalankan.
    """
    per_mk = _sesi_per_mata_kuliah(klien, daftar)
    gabungan = sorted({s for sesi in per_mk.values() for s in sesi})
    if not gabungan:
        print("Tidak ada section dengan soal di mata kuliah mana pun.", file=sys.stderr)
        return None

    print("\nSesi yang punya soal:")
    for s in gabungan:
        pemilik = [mk for mk in daftar if s in per_mk.get(mk.id, ())]
        print(f"  Sesi {s:<3} -- {len(pemilik)} dari {len(daftar)} mata kuliah")

    pilih = _pilih_dari_daftar(
        "Pilih sesi yang akan diproses untuk semua mata kuliah:",
        [f"Sesi {s}" for s in gabungan],
    )
    if pilih is None:
        return None
    return gabungan[pilih], per_mk


def _jalankan_satu(
    matkul: courses.MataKuliah, nomor: int, args: argparse.Namespace
) -> tuple[bool, object, str]:
    """Jalankan pipeline untuk satu (mata kuliah, sesi).

    Semua kesalahan dibungkus jadi nilai balik, bukan exception yang lolos ke
    pemanggil. Dalam mode banyak mata kuliah, satu course yang gagal tidak
    boleh menghentikan course yang belum dijalankan.
    """
    try:
        pipe = Pipeline(
            matkul,
            nomor,
            model_utama=args.model,
            model_mata=args.model_mata,
            dengan_gambar=not args.tanpa_gambar,
            tanpa_docx=args.tanpa_docx,
        )
        return True, pipe.jalankan(), ""
    except PipelineGagal as exc:
        return False, None, f"pipeline berhenti: {exc}"
    except moodle.MoodleError as exc:
        return False, None, f"Moodle: {exc}"
    except OSError as exc:
        return False, None, f"berkas: {exc}"


def _cetak_ringkasan(hasil: object) -> None:
    """Cetak catatan dan berkas hasil satu sesi."""
    for g in getattr(hasil, "gagal", []) or []:
        print(f"    ! {g}")
    docx = getattr(hasil, "docx", None)
    jawaban = getattr(hasil, "jawaban", None)
    if docx:
        print(f"    Berkas jawaban: {docx}")
    if not docx and jawaban:
        print(f"    Berkas markdown: {jawaban}")


def _jalan_semua(klien: moodle.Moodle, daftar: list[courses.MataKuliah], args) -> int:
    """Proses seluruh matakuliah berurutan pada satu sesi.

    Tiga status, bukan dua. Mata kuliah yang tidak punya sesi itu **dilewati**,
    bukan gagal: tidak ada yang salah dan tidak ada agent yang terpanggil.
    Melabelinya "GAGAL" membuat ringkasan terlihat seperti ada yang rusak,
    padahal yang terjadi hanya course itu memang punya lebih sedikit section.
    """
    print(GARIS)
    print(f"MEMPROSES {len(daftar)} MATA KULIAH")
    print(GARIS)

    if args.sesi is not None:
        nomor = args.sesi
        # Section tetap dibaca meski sesinya sudah ditentukan. Satu panggilan
        # Moodle per course jauh lebih murah daripada menjalankan pipeline
        # penuh -- yang berarti beberapa agent -- untuk sesi yang memang
        # tidak ada di sana.
        per_mk = _sesi_per_mata_kuliah(klien, daftar)
        punya = [mk for mk in daftar if nomor in per_mk.get(mk.id, ())]
        print(f"Sesi {nomor} (dari --sesi)")
        if not punya:
            print(
                f"Peringatan: tidak ada satu pun mata kuliah yang punya sesi "
                f"{nomor}. Semua akan dilewati."
            )
        else:
            print(f"  {len(punya)} dari {len(daftar)} mata kuliah punya sesi {nomor}")
    else:
        pilihan = _pilih_sesi_batch(klien, daftar)
        if pilihan is None:
            print("Dibatalkan.")
            return 0
        nomor, per_mk = pilihan

    baris: list[tuple[str, str, str]] = []
    n_selesai = n_dilewati = n_gagal = 0
    for i, mk in enumerate(daftar, 1):
        if per_mk and mk.id in per_mk and nomor not in per_mk[mk.id]:
            baris.append((mk.nama, "dilewati", f"course ini tidak punya sesi {nomor}"))
            n_dilewati += 1
            print(f"\n--- [{i}/{len(daftar)}] {mk.nama}: dilewati "
                  f"(tidak ada sesi {nomor}) ---", flush=True)
            continue

        print(f"\n--- [{i}/{len(daftar)}] {mk.nama} ---", flush=True)
        ok, hasil, pesan = _jalankan_satu(mk, nomor, args)
        if ok:
            n_selesai += 1
            baris.append((mk.nama, "selesai", ""))
            _cetak_ringkasan(hasil)
        else:
            n_gagal += 1
            baris.append((mk.nama, "gagal", pesan))
            print(f"    GAGAL -- {pesan}", flush=True)

    print()
    print(GARIS)
    print(f"RINGKASAN SESI {nomor}")
    print(GARIS)
    tanda = {"selesai": "selesai  ", "dilewati": "dilewati", "gagal": "GAGAL    "}
    for nama, status, pesan in baris:
        print(f"  {tanda[status]} {nama}")
        if pesan:
            print(f"              {pesan}")
    bagian = [f"{n_selesai} selesai"]
    if n_dilewati:
        bagian.append(f"{n_dilewati} dilewati (tidak ada sesi {nomor})")
    if n_gagal:
        bagian.append(f"{n_gagal} gagal")
    print(f"\n{', '.join(bagian)} dari {len(daftar)} mata kuliah.")
    print(GARIS)
    # Yang dilewati bukan kegagalan: course itu memang tidak punya sesi itu.
    # Exit non-nol hanya kalau ada yang benar-benar gagal.
    return 1 if n_gagal else 0


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
    ap.add_argument(
        "--semua", action="store_true",
        help=(
            "Proses seluruh mata kuliah satu per satu pada sesi yang sama. "
            "Mata kuliah yang gagal tidak menghentikan yang berikutnya."
        ),
    )
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

    semua = args.semua
    matkul: courses.MataKuliah | None = None
    if args.matkul:
        if args.matkul.strip().lower() in _KATA_SEMUA:
            semua = True
        else:
            matkul = _cari_mata_kuliah(daftar, args.matkul)
            if matkul is None:
                print(
                    f"\nMata kuliah {args.matkul!r} tidak ditemukan. "
                    "Jalankan `python main.py --daftar` untuk melihat "
                    "daftarnya, atau pakai `--semua` untuk semuanya.\n",
                    file=sys.stderr,
                )
                return 2
            print(f"Mata kuliah: {matkul.label}")
    elif semua:
        print(f"\nDiproses semua mata kuliah ({len(daftar)} buah), berurutan.")
    else:
        print("\nPilih mata kuliah:")
        for i, mk in enumerate(daftar, 1):
            kode = f" [{mk.kode}.{mk.kelas}]" if mk.kode and mk.kelas else ""
            print(f"  {i:>2}. {mk.nama}{kode}")
        semua_no = len(daftar) + 1
        print(f"  {semua_no:>2}. Semua mata kuliah (diproses berurutan)")

        jawab = input("\nMasukkan nomor mata kuliah: ").strip()
        if jawab.lower() in {"batal", "b", "q", "quit", "exit"}:
            print("Dibatalkan.")
            return 0
        if not jawab.isdigit() or not 1 <= int(jawab) <= semua_no:
            print("Nomor di luar daftar.", file=sys.stderr)
            return 2
        if int(jawab) == semua_no:
            semua = True
        else:
            matkul = daftar[int(jawab) - 1]
            print(f"\nMata kuliah: {matkul.label}")

    if semua:
        try:
            return _jalan_semua(klien, daftar, args)
        except moodle.LoginMati as exc:
            print(f"\nSesi Moodle berakhir: {exc}\n", file=sys.stderr)
            return 2
        except KeyboardInterrupt:
            print("\nDihentikan pengguna.", file=sys.stderr)
            return 130

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