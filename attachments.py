"""Pengelolaan lampiran: unduhan, teks PDF, dan gambar halaman PDF.

Lampiran UT punya dua bentuk. Ada berkas yang benar-benar teks, dan ada PDF
bahan ajar yang sebenarnya hasil ekspor PowerPoint. PDF hasil ekspor itu
jarang sekali memuat teks yang cukup: Materi Inisiasi 1 Logika Informatika
punya 46 halaman tapi hanya sekitar 12 ribu karakter teks, sedangkan 56
gambarnya justru yang membawa isi. Membaca PDF seperti teks biasa akan
mendapatkan daftar nama butir tanpa satu pun penjelasannya.

Karena itu PDF dipecah dua jalur. Teksnya diekstrak supaya agent bisa
menganye(search), dan halamannya dirender menjadi PNG supaya model
penglihatan bisa membaca apa yang tertulis dan tergambar di tiap slide.
"""

from __future__ import annotations

import re
import shutil
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import unquote, urlparse

import config
import html2md
import moodle

# Ekstensi yang isinya sudah berupa teks.
TEKST_DATANG = {
    ".md", ".txt", ".csv", ".json", ".xml", ".html", ".htm",
    ".rtf", ".docx", ".doc", ".odt", ".xlsx", ".pptx",
}

# Lebar render halaman PDF, dalam piksel. 1600 cukup untuk membaca teks
# slide tanpa membuat berkas gambar terlalu besar untuk dikirim ke model.
LEBAR_RENDER = 1600
BATAS_HALAMAN_GAMBAR = 40

# Halaman PDF digabung beberapa halaman per gambar. Flag `--file` milik
# `opencode run` hanya menerima satu berkas, jadi menggabungkan halaman
# memangkas jumlah panggilan vision dari satu per halaman menjadi satu per
# beberapa halaman.
PER_GABUNGAN = 4
KOLOM_GABUNGAN = 2
SEL_LEBAR = 1200


@dataclass
class HasilBerkas:
    """Hasil pemrosesan satu lampiran."""

    url: str
    nama: str
    lokasi: Path | None = None
    teks: str = ""
    gambar: list[Path] = field(default_factory=list)
    catatan: list[str] = field(default_factory=list)
    gagal: bool = False

    @property
    def ada_isi(self) -> bool:
        return bool(self.teks.strip() or self.gambar)

    def ringkas(self) -> str:
        """Satu baris deskripsi untuk ditampilkan ke pengguna."""
        bagian = [self.nama]
        if self.lokasi and self.lokasi.is_file():
            ukuran = self.lokasi.stat().st_size
            bagian.append(f"{ukuran / 1024:.0f} KB")
        if self.teks.strip():
            bagian.append(f"{len(self.teks.split())} kata")
        if self.gambar:
            bagian.append(f"{len(self.gambar)} halaman gambar")
        if self.catatan:
            bagian.extend(self.catatan)
        return " | ".join(bagian)


def _pakai_pymupdf():
    """Impor PyMuPDF dengan nama modul yang benar untuk versinya.

    `fitz` sudah tidak direkomendasikan sejak PyMuPDF 1.24 dan Removal
    pemanggilannya memunculkan peringatan yang mengotori keluaran.
    """
    try:
        import pymupdf  # noqa: PLC0415

        return pymupdf
    except ImportError:
        import fitz  # noqa: PLC0415

        return fitz


def nama_berkas(url: str, fallback: str = "lampiran") -> str:
    """Ambil nama berkas akhir dari URL, lalu amankan untuk Windows."""
    path = unquote(urlparse(url).path)
    nama = path.rstrip("/").rsplit("/", 1)[-1] or fallback
    nama = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", nama).strip(" .")
    return nama[:120] or fallback


def _url_berkas_dari_halaman_resource(
    klien: moodle.Moodle, url_resource: str
) -> str:
    """Temukan URL `pluginfile.php` di balik halaman sebuah resource.

    `mod/resource/view.php` untuk berkas langsung mengirim 303 ke
    `pluginfile.php`, dan pada sesi UT 303 itu berakhir di halaman
    login. Meminta `&forceview=1` memaksa Moodle merender halaman, dan
    tautan berkasnya diambil dari sana.
    """
    if "/mod/resource/" not in url_resource:
        return url_resource

    if "forceview=1" not in url_resource:
        url_resource = (
            url_resource + ("&" if "?" in url_resource else "?") + "forceview=1"
        )

    from bs4 import BeautifulSoup  # noqa: PLC0415

    halaman = klien.ambil(url_resource)
    soup = BeautifulSoup(halaman.html, "lxml")

    for sel in (
        "div.resourceworkaround a[href*='pluginfile']",
        "div.urlworkaround a[href*='pluginfile']",
        "a[href*='mod_resource']",
        "a[href*='pluginfile']",
    ):
        for tautan in soup.select(sel):
            href = tautan.get("href") or ""
            if "pluginfile" in href and "/user/icon/" not in href:
                return klien.absolut(href)

    return ""


def unduh_berkas(
    klien: moodle.Moodle,
    url: str,
    folder: Path,
    *,
    label: str = "",
) -> HasilBerkas:
    """Unduh satu lampiran dan simpan ke `folder`."""
    folder.mkdir(parents=True, exist_ok=True)
    url_target = _url_berkas_dari_halaman_resource(klien, url)

    if not url_target:
        return HasilBerkas(
            url=url,
            nama=label or "resource",
            gagal=True,
            catatan=["Tidak ditemukan tautan berkas di halaman resource."],
        )

    nama = nama_berkas(url_target, config.slugify(label, "lampiran"))
    # Nama yang sama bisa muncul dua kali dalam satu sesi, jadi, hindari
    # penimpa diam-diam.
    tujuan = folder / nama
    n = 2
    while tujuan.exists():
        tujuan = folder / f"{tujuan.stem}_{n}{tujuan.suffix}"
        n += 1

    try:
        klien.unduh(url_target, tujuan)
    except (moodle.MoodleError, OSError) as exc:
        return HasilBerkas(
            url=url_target, nama=tujuan.name, gagal=True, catatan=[str(exc)[:200]]
        )

    return HasilBerkas(url=url_target, nama=tujuan.name, lokasi=tujuan)


# --------------------------------------------------------------- teks berkas

def _teks_docx(path: Path) -> str:
    """Baca teks dari berkas .docx tanpa dependensi python-docx."""
    import zipfile  # noqa: PLC0415
    import xml.etree.ElementTree as ET  # noqa: PLC0415

    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    keluar: list[str] = []
    try:
        with zipfile.ZipFile(path) as arsip:
            with arsip.open("word/document.xml") as fh:
                akar = ET.parse(fh).getroot()
    except (OSError, KeyError, ET.ParseError) as exc:
        return f"(berkas .docx tidak bisa dibaca: {exc})"

    for paragraf in akar.iter(f"{{{ns['w']}}}p"):
        bagian = [
            node.text or ""
            for node in paragraf.iter(f"{{{ns['w']}}}t")
        ]
        if bagian:
            keluar.append("".join(bagian))
    return "\n\n".join(keluar)


def _teks_pptx(path: Path) -> str:
    """Baca teks tiap slide dari .pptx."""
    import zipfile  # noqa: PLC0415
    import xml.etree.ElementTree as ET  # noqa: PLC0415
    import re as _re  # noqa: PLC0415

    a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    keluar: list[str] = []
    try:
        with zipfile.ZipFile(path) as arsip:
            slide = sorted(
                (n for n in arsip.namelist()
                 if _re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
                key=lambda n: int(_re.search(r"\d+", n.rsplit("/", 1)[-1]).group()),
            )
            for nomor, nama in enumerate(slide, 1):
                with arsip.open(nama) as fh:
                    akar = ET.parse(fh).getroot()
                baris = [
                    (node.text or "").strip()
                    for node in akar.iter(f"{{{a}}}t")
                    if (node.text or "").strip()
                ]
                if baris:
                    keluar.append(f"## Slide {nomor}\n" + "\n".join(baris))
    except (OSError, KeyError, ET.ParseError) as exc:
        return f"(berkas .pptx tidak bisa dibaca: {exc})"

    return "\n\n".join(keluar)


def teks_dari_berkas(
    klien: moodle.Moodle,
    url: str,
    *,
    label: str = "",
    folder_sementara: Path | None = None,
    with_gambar: bool = True,
) -> HasilBerkas:
    """Ambil teks dari berkas yang berada di balik URL Moodle.

    Ini jalur yang dipakai endpoint `/berkas` milik Reader. Berkas diunduh
    ke folder sementara, dibaca, lalu dihapus supaya Reader tidak
    menjadi tempat/anime penampungan disk.
    """
    sementara = Path(tempfile.mkdtemp(prefix="jkt-berkas-"))
    try:
        hasil = unduh_berkas(klien, url, sementara, label=label)
        if hasil.gagal or not hasil.lokasi:
            return hasil

        isi, catatan = _baca_isi(hasil.lokasi, with_gambar=with_gambar,
                                 folder_gambar=sementara / "halaman")
        hasil.teks = isi
        hasil.catatan.extend(catatan)
        if with_gambar:
            hasil.gambar = sorted((sementara / "halaman").glob("*.png"))
            # Gam bar harus ikut hidup selama pemanggil masih membutuhkannya.
            if hasil.gambar:
                hasil.lokasi = hasil.lokasi  # tetap pakai berkas asli
        return hasil
    finally:
        shutil.rmtree(sementara, ignore_errors=True)


def _baca_isi(
    path: Path, *, with_gambar: bool = True, folder_gambar: Path | None = None
) -> tuple[str, list[str]]:
    """Baca teks satu berkas dan, untuk PDF, render halaman jadi gambar."""
    suffix = path.suffix.lower()
    catatan: list[str] = []

    if suffix == ".pdf":
        return _baca_pdf(path, with_gambar=with_gambar,
                         folder_gambar=folder_gambar, catatan=catatan)

    if suffix == ".docx":
        return _teks_docx(path), catatan
    if suffix == ".pptx":
        return _teks_pptx(path), catatan

    try:
        isi = path.read_text(encoding="utf-8", errors="replace")
    except OSError as exc:
        return f"(berkas {path.name} tidak bisa dibaca: {exc})", catatan

    if suffix in (".html", ".htm"):
        return html2md.ke_markdown(isi, base_url=""), catatan
    if suffix == ".rtf":
        return _teks_rtf(isi), catatan
    if suffix in (".doc", ".xls", ".ppt"):
        catatan.append(
            f"{path.name} adalah format lama yang tidak bisa dibaca teks "
            "secara langsung; isinya dilewati."
        )
    return isi, catatan


def _teks_rtf(isi: str) -> str:
    """Ambil teks dari RTF dengan membuang kendali RTF."""
    tanpa_kendali = re.sub(r"\\[a-zA-Z]+-?\d*\s?", "", isi)
    tanpa_kendali = re.sub(r"[{}]", "", tanpa_kendali)
    return tanpa_kendali


def _baca_pdf(
    path: Path, *, with_gambar: bool, folder_gambar: Path | None,
    catatan: list[str],
) -> tuple[str, list[str]]:
    """Ekstrak teks PDF dan, bila diminta, render halamannya jadi PNG."""
    pymupdf = _pakai_pymupdf()
    try:
        dokumen = pymupdf.open(path)
    except Exception as exc:  # noqa: BLE001 - pustaka pihak ketiga
        return f"(PDF tidak bisa dibuka: {exc})", catatan

    try:
        halaman_n = dokumen.page_count
        potongan: list[str] = []
        for i in range(halaman_n):
            teks = dokumen[i].get_text().strip()
            if teks:
                potongan.append(f"### Halaman {i + 1}\n{teks}")
        teks_penuh = "\n\n".join(potongan)

        # Rasio teks terhadap gambar menentukan apakah model penglihatan
        # perlu turun tangan.
        gambar_total = sum(
            len(dokumen[i].get_images(full=True)) for i in range(halaman_n)
        )
        if not teks_penuh:
            catatan.append("PDF tidak memuat teks sama sekali; isinya ada di gambar.")
        elif gambar_total > halaman_n and len(teks_penuh) < 20000:
            catatan.append(
                "PDF ini hasil ekspor slide: teksnya sedikit dan isinya "
                "sebagian besar berupa gambar. Halaman sudah dirender, "
                "bacalah juga setiap gambarnya."
            )

        if not with_gambar or folder_gambar is None:
            return teks_penuh, catatan

        folder_gambar.mkdir(parents=True, exist_ok=True)
        n_render = min(halaman_n, BATAS_HALAMAN_GAMBAR)
        for i in range(n_render):
            try:
                pixmap = dokumen[i].get_pixmap(dpi=150)
            except Exception as exc:  # noqa: BLE001
                catatan.append(f"Halaman {i + 1} gagal dirender: {exc}")
                continue
            pixmap.save(folder_gambar / f"hal{i + 1:03d}.png")

        if halaman_n > n_render:
            catatan.append(
                f"Hanya {n_render} dari {halaman_n} halaman yang dirender."
            )
        return teks_penuh, catatan
    finally:
        dokumen.close()


# Gambar kecil yang bukan bagian soal: avatar, smiley, ikon tombol, logo.
POLA_GAMBAR_BUKAN_SOAL = re.compile(
    r"/user/icon/|/theme_|\bsmiley\b|/pix/|/amd/",
    re.I,
)


def kumpulkan_gambar_soal(
    klien: moodle.Moodle, url: str, folder: Path
) -> list[Path]:
    """Unduh gambar yang tertanam di dalam halaman soal.

    Sebagian soal Universitas Terbuka memuat gambar soal, diagram, atau
    tabelbesar yang harus diinsert apa adanya ke dokumen. Mengetiknya
    ulang dari transkripsi berisiko salahnomor, dan dokumen yang dinilai
    tutor harus sama persis dengan yang tertulis di halaman.
    """
    try:
        halaman = klien.ambil(url)
    except moodle.MoodleError:
        return []

    import html2md as _html2md  # noqa: PLC0415
    from bs4 import BeautifulSoup  # noqa: PLC0415

    soup = BeautifulSoup(halaman.html, "lxml")
    _html2md._selaraskan_mathml(soup)
    _html2md._selaraskan_script_mathjax(soup)
    if "/mod/assign/" in url:
        _html2md.buang_jawaban_mahasiswa(soup, elas=True)
    else:
        _html2md.buang_jawaban_mahasiswa(soup, elas=False)

    # Batasi pencarian ke post soal atau deskripsi tugas, bukan seluruh
    # halaman: sidebar dan kepala halaman penuh avatar.
    akar = soup.select_one("div.forumpost") or soup.select_one(
        "[data-region='main'], #region-main, main"
    )
    if akar is None:
        return []

    folder.mkdir(parents=True, exist_ok=True)
    keluar: list[Path] = []

    for i, tag in enumerate(akar.find_all("img"), 1):
        src = (tag.get("src") or tag.get("data-src") or "").strip()
        if not src or POLA_GAMBAR_BUKAN_SOAL.search(src):
            continue

        # URL inline base64 (`data:image/png;base64,...`) muncul kalau
        # dosen menempel gambar langsung di deskripsi section tanpa upload
        # file terpisah. Pola `endswith((...))` di bawah akan menolaknya
        # karena stringnya tidak punya ekstensi, jadi ditangani khusus di sini.
        if src.startswith("data:image/"):
            import base64 as _b64

            try:
                header, koma, data = src.partition(",")
                mime = header.split(";")[0].split(":", 1)[1].lower()
                ekst = {"image/png": ".png", "image/jpeg": ".jpg",
                        "image/gif": ".gif", "image/webp": ".webp"}.get(mime)
                if ekst is None or not data:
                    continue
                nama = f"{config.slugify(Path(folder).name, 'soal')}-{i:02d}{ekst}"
                tujuan = folder / nama
                tujuan.write_bytes(_b64.b64decode(data, validate=False))
            except Exception:  # noqa: BLE001
                continue
            if tujuan.is_file() and tujuan.stat().st_size > 1024:
                keluar.append(tujuan)
            continue

        penuh = klien.absolut(src)
        if not penuh.lower().split("?")[0].endswith((".png", ".jpg", ".jpeg", ".gif")):
            continue
        nama = f"{config.slugify(Path(folder).name, 'soal')}-{i:02d}" + \
            Path(unquote(urlparse(penuh).path)).suffix.lower()
        tujuan = folder / nama
        try:
            klien.unduh(penuh, tujuan)
        except (moodle.MoodleError, OSError):
            continue
        if tujuan.is_file() and tujuan.stat().st_size > 1024:
            keluar.append(tujuan)

    return keluar


def gabung_halaman(
    daftar_png: list[Path],
    folder_tujuan: Path,
    *,
    per_gabungan: int = PER_GABUNGAN,
    kolom: int = KOLOM_GABUNGAN,
    sel_lebar: int = SEL_LEBAR,
) -> list[Path]:
    """Gabungkan beberapa halaman PNG menjadi satu gambar montage.

    Montage memakai nomor halaman di sudut supaya model penglihatan tahu
    halaman mana yang sedang dibaca, dan supaya transkripnya bisa dirujuk
    kembali ke halaman asalnya.
    """
    if not daftar_png:
        return []

    try:
        from PIL import Image, ImageDraw  # noqa: PLC0415
    except ImportError:
        # Tanpa Pillow,PNGD yang sudah dirender dikirim satu per satu.
        return list(daftar_png)

    folder_tujuan.mkdir(parents=True, exist_ok=True)
    keluar: list[Path] = []
    per_gambar = max(1, per_gabungan)
    baris = max(1, per_gambar // max(1, kolom))

    for mulai in range(0, len(daftar_png), per_gambar):
        kelompok = daftar_png[mulai : mulai + per_gabungan]
        if len(kelompok) == 1:
            keluar.append(kelompok[0])
            continue

        sel = []
        for p in kelompok:
            try:
                with Image.open(p) as im:
                    im = im.convert("RGB")
                    tinggi = max(1, round(im.height * sel_lebar / im.width))
                    sel.append(im.resize((sel_lebar, tinggi)))
            except OSError:
                continue
        if not sel:
            continue

        tinggi_sel = max(im.height for im in sel)
        kanvas = Image.new(
            "RGB",
            (sel_lebar * kolom, tinggi_sel * baris),
            "white",
        )
        gambar = ImageDraw.Draw(kanvas)
        for i, im in enumerate(sel):
            x = (i % kolom) * sel_lebar
            y = (i // kolom) * tinggi_sel
            kanvas.paste(im, (x, y))
            gambar.rectangle(
                [x, y, x + sel_lebar - 1, y + tinggi_sel - 1],
                outline="#bbbbbb",
                width=2,
            )
            nomor = mulai + i + 1
            gambar.rectangle([x + 4, y + 4, x + 76, y + 34], fill="black")
            gambar.text((x + 14, y + 12), f"Hal {nomor}", fill="white")

        tujuan = folder_tujuan / f"montase{mulai // per_gambar + 1:02d}.png"
        kanvas.save(tujuan, optimize=True)
        keluar.append(tujuan)

    return keluar


def pilih_halaman_gambar(
    teks_per_halaman: list[str],
    jumlah_maks: int = 24,
) -> list[int]:
    """Pilih nomor halaman yang paling perlu dibaca model penglihatan.

    Urutannya: halaman dengan teks paling sedikit lebih dulu, karena di
    hasil ekspor slide halaman seperti itulah yang isinya berupa gambar dan
   diagram. Kalau jumlah halaman melebihi batas, pemotongan dilakukan dari
    belakang supaya sampul dan daftar isi tetap ikut terbaca.
    """
    n = len(teks_per_halaman)
    if n <= jumlah_maks:
        return list(range(n))

    urut = sorted(range(n), key=lambda i: (len(teks_per_halaman[i]), i))
    dipilih = set(urut[:jumlah_maks])
    # Halaman awal selalu dipertahankan.
    dipilih.update(range(min(3, n)))
    return sorted(dipilih)


def cari_lampiran_forum(klien: moodle.Moodle, url: str) -> list[tuple[str, str]]:
    """Ekstrak tautan lampiran forum dari halaman posting Diskusi.

    Forum Diskusi Universitas Terbuka sering menyimpan soal -- misalnya
    `Diskusi 4 - Model Linear Terapan.pdf` -- sebagai lampiran pada post
    pertama, bukan sebagai resource di halaman section. Oleh karena itu
    `_kumpulkan_gambar_soal` (yang hanya mencari <img>) tidak menemukannya,
    dan `kumpulkan_lampiran` (yang hanya diberi URL aktivitas tipe resource)
    tidak akan pernah melihatnya. Lampiran seperti ini harus ditemukan
    secara eksplisit.
    """
    try:
        halaman = klien.ambil(url)
    except moodle.MoodleError:
        return []
    hasil: list[tuple[str, str]] = []
    for m in re.finditer(
        r'href="([^"]*pluginfile\.php/[^"]*/mod_forum/attachment/[^"]+)"',
        halaman.html,
    ):
        url_abs = klien.absolut(m.group(1))
        nama = Path(unquote(urlparse(url_abs).path)).name
        hasil.append((url_abs, nama))
    return hasil


def kumpulkan_lampiran(
    klien: moodle.Moodle,
    urls: list[tuple[str, str]],
    folder: Path,
    *,
    with_gambar: bool = True,
) -> list[HasilBerkas]:
    """Unduh dan baca beberapa lampiran sekaligus.

    `urls` berisi pasangan (url, label). Lampiran yang gagal tidak
    menggagalkan proses: bahan ajar yang gagal diunduh hanya berarti ada
    sumber yang kurang, bukan tidak ada jawaban.
    """
    folder.mkdir(parents=True, exist_ok=True)
    hasil: list[HasilBerkas] = []

    for url, label in urls:
        if not url:
            continue
        unduhan = unduh_berkas(klien, url, folder / "berkas", label=label)
        if unduhan.gagal or not unduhan.lokasi:
            hasil.append(unduhan)
            continue

        isi, catatan = _baca_isi(
            unduhan.lokasi,
            with_gambar=with_gambar,
            folder_gambar=folder / "halaman" / config.slugify(unduhan.nama, "x"),
        )
        unduhan.teks = isi
        unduhan.catatan.extend(catatan)
        unduhan.gambar = sorted((folder / "halaman" / config.slugify(unduhan.nama, "x")).glob("*.png"))
        hasil.append(unduhan)

    return hasil