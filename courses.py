"""Pemetaan mata kuliah dan sesi dari Moodle Universitas Terbuka.

Pengamatan yang membentuk modul ini, semuanya diambil dari halaman
sebenarnya bukan dari tebakan:

1. Halaman `/my/` memuat daftar mata kuliah pada menu `mycourses`. Atribut
   `title` berisi nama lengkap, sedangkan teks tautannya berisi
   `KODE.KELAS`, misalnya `STSI4106.90`.

2. Semua course memakai format `onetopic`, dan isinya dirender satu
   section per permintaan. Mengambil `course/view.php?id=X` tanpa
   parameter section hanya menghasilkan halaman depan. Isi sesi harus
   diambil dari `&section=N`.

3. Nomor section adalah nomor sesi. Ini berlaku di keempat course yang
   diproses, termasuk saat nomor activity berbeda: di Basis Data section 3
   memuat `Forum Diskusi.2` sekaligus `Tugas.1`, jadi membaca nomor dari
   activity akan salah answered.

4. Penomoran pada `Materi Inisiasi N` menghitung materi, bukan sesi. Di
   Struktur Data, `Materi Inisiasi 2` berada di section 1. Karena itu
   nomor dari nama dipakai hanya sebagai pelengkap, bukan penentu.

5. `mod/resource/view.php` untuk berkas langsung mengirim 303 ke
   `pluginfile.php`, dan pada sesi UT perjalanan itu berakhir di halaman
   login. Halaman resource selalu diminta dengan `&forceview=1`.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from bs4 import BeautifulSoup

import config
import moodle

# --------------------------------------------------------------- pola nama

# Nomor sesi yang tersirat di nama activity. Nomor sebelum titik pada
# "Materi 1.2" adalah sesi, sedangkan setelah titik adalah nomor materi.
RE_NOMOR_NAMA = re.compile(
    r"(?:diskusi|tugas|sesi|materi\s*inisiasi|inisisasi|materi|inisiasi|"
    r"bab|topik|activity|lesson)\s*[.:]?\s*(\d{1,2})",
    re.I,
)
# "Materi 1.2" -> sesi 1.
RE_NOMOR_MATERI = re.compile(r"\bmateri\s+(\d{1,2})[.\-_]\d{1,2}\b", re.I)

# Nama activity yang bukan soal, apa pun jenisnya.
POLA_BUKAN_SOAL = re.compile(
    r"perkenalan|announcement|pengumuman|tata tertib|ethika|raw|rat\b|"
    r"kendala|saran|kehadiran|absen|feedback|umpan balik|nilai|"
    r"silabus|kalender|jadual|jadwal|informasi mata kuliah|sosialisasi|"
    r"eksplorasi|pemanfaatan|struktur tcp|webinar|praktikum mandiri",
    re.I,
)

# modtype yang isinya soal yang harus dijawab.
MODTIPE_SOAL = {"forum", "assign", "quiz"}
# modtype yang isinya bahan ajar atau petunjuk.
MODTIPE_MATERI = {"resource", "page", "folder", "url", "book", "lesson"}

# Nama forum yang bukan soal meski bertipe forum.
POLA_FORUM_BUKAN_SOAL = re.compile(
    r"perkenalan|announcement|pengumuman|feedback|umpan balik|"
    r"penilaian|nilai|diskusi umum|konsultasi|bantuan",
    re.I,
)

# Berkas pendukung yang nomornya cocok dengan sesi, tapi isinya bukan bahan
# ajar. "Template Tugas 1" bernomor 1, padahal formatnya untuk Tugas.1 di
# section lain; kalau ikut diambil, sesi 1 memperoleh lampiran yang tidak
# relevan.
POLA_BUKAN_BAHAN_AJAR = re.compile(
    r"\btemplate\b|\bformat\b|\blaporan\b|\bcontoh\b|\bsampel\b|"
    r"\bpedoman\b|\bguidelines?\b|\brubrik?\b|\brubric\b|"
    r"\blembar kerja\b|\bkuliah kosong\b|\bblank\b",
    re.I,
)


def _tarik_nama(cm) -> str:
    node = cm.select_one(".activityname, .instancename")
    if node is None:
        return ""
    # Nama activity pada mb2iq diikuti penanda jenis, misalnya
    # "Forum Diskusi.1 Forum" atau "Materi inisiasi 1 File". Bagian
    # terakhir itu bukan bagian dari nama.
    teks = re.sub(r"\s+", " ", node.get_text(" ", strip=True))
    return re.sub(
        r"\s+(Forum|Assignment|File|Page|URL|Lesson|Quiz|Label|Attendance|"
        r"Workshop|Book|Folder)$",
        "",
        teks,
        flags=re.I,
    ).strip()


def _tarik_modtype(cm) -> str:
    for kelas in cm.get("class") or []:
        if kelas.startswith("modtype_"):
            return kelas[len("modtype_") :]
    return ""


def nomor_dari_nama(nama: str) -> int | None:
    """Baca nomor sesi atau materi yang tersirat di nama activity."""
    if not nama:
        return None
    m = RE_NOMOR_MATERI.search(nama)
    if m:
        return int(m.group(1))
    m = RE_NOMOR_NAMA.search(nama)
    if m:
        return int(m.group(1))
    return None


# --------------------------------------------------------------- dataclass


@dataclass
class Aktivitas:
    """Satu activity di dalam course."""

    id: str
    modtype: str
    nama: str
    url: str
    section: int
    deskripsi: str = ""

    @property
    def nomor(self) -> int | None:
        return nomor_dari_nama(self.nama)

    @property
    def boleh_ke_soal(self) -> bool:
        """Activity ini benar-benar soal yang harus dijawab?"""
        if self.modtype not in MODTIPE_SOAL:
            return False
        if POLA_BUKAN_SOAL.search(self.nama):
            return False
        if self.modtype == "forum" and POLA_FORUM_BUKAN_SOAL.search(self.nama):
            return False
        return True

    @property
    def boleh_ke_materi(self) -> bool:
        """Activity ini bahan ajar atau petunjuk pendukung?"""
        return (
            self.modtype in MODTIPE_MATERI
            and self.modtype != "lesson"      # lesson berisi absensi
            and not POLA_BUKAN_SOAL.search(self.nama)
        )

    @property
    def boleh_ke_pelengkap(self) -> bool:
        """Boleh masuk sebagai pelengkap lintas section?

        Template dan format laporan tidak boleh. Nomor pada namanya cocok
        dengan nomor sesi, tapi isinya bukan bahasan sesi itu.
        """
        return (
            self.boleh_ke_materi
            and not POLA_BUKAN_BAHAN_AJAR.search(self.nama)
        )

    @property
    def jenis_soal(self) -> str:
        return {
            "forum": "Diskusi",
            "assign": "Tugas",
            "quiz": "Kuis",
        }.get(self.modtype, self.modtype)


@dataclass
class Sesi:
    """Satu sesi yang sudah dipetakan."""

    nomor: int
    section: int
    soal: list[Aktivitas] = field(default_factory=list)
    materi: list[Aktivitas] = field(default_factory=list)
    catatan: list[str] = field(default_factory=list)

    @property
    def punya_soal(self) -> bool:
        return bool(self.soal)

    def ringkas(self) -> str:
        jenis = sorted({a.jenis_soal for a in self.soal})
        bagian = [f"Sesi {self.nomor}"]
        bagian.append(f"section {self.section}")
        bagian.append(f"soal: {len(self.soal)} ({', '.join(jenis) or 'tidak ada'})")
        bagian.append(f"materi: {len(self.materi)}")
        return " | ".join(bagian)


@dataclass
class MataKuliah:
    """Satu course milik pengguna."""

    id: str
    nama: str
    kode: str = ""
    kelas: str = ""
    slug: str = ""
    jumlah_section: int = 0
    sesi_tersedia: list[int] = field(default_factory=list)

    @property
    def label(self) -> str:
        """Nama yang ditampilkan di daftar pilihan."""
        bagian = [self.nama]
        if self.kelas:
            bagian.append(f"kelas {self.kelas}")
        if self.kode:
            bagian.append(f"({self.kode})")
        return " ".join(bagian)


# --------------------------------------------------------------- scraping


def daftar_mata_kuliah(klien: moodle.Moodle) -> list[MataKuliah]:
    """Baca seluruh mata kuliah yang sedang diambil.

    Sumbernya menu `mycourses` di `/my/`, bukan halaman "My courses",
    karena halaman `/my/` sudah memuat nama lengkap dan kode kelas dalam
    satu permintaan.
    """
    halaman = klien.ambil(f"{klien.base_url}/my/")
    soup = BeautifulSoup(halaman.html, "lxml")

    hasil: dict[str, MataKuliah] = {}

    for tautan in soup.find_all("a", href=True):
        href = tautan["href"]
        m = re.search(r"course/view\.php\?id=(\d+)", href)
        if not m:
            continue
        cid = m.group(1)
        if cid in hasil:
            continue

        # `title` memuat nama lengkap; teks tautannya `KODE.KELAS`.
        nama = (tautan.get("title") or "").strip()
        teks = re.sub(r"\s+", " ", tautan.get_text(" ", strip=True))
        if not nama or not teks:
            continue

        kode, kelas = "", ""
        cocok = re.match(r"^([A-Z]{4}\d{4})\.(\d{1,3})$", teks)
        if cocok:
            kode, kelas = cocok.group(1), cocok.group(2)
        else:
            nama, kode = config.pisah_kode_matkul(nama)
            m_kelas = config.RE_KELAS.search(nama)
            if m_kelas:
                kelas = m_kelas.group(1)

        hasil[cid] = MataKuliah(
            id=cid,
            nama=nama.strip(),
            kode=kode,
            kelas=kelas,
            slug=config.slugify(nama.strip() or f"matkul-{cid}", f"matkul-{cid}"),
        )

    if not hasil:
        raise moodle.MoodleError(
            "Tidak ada mata kuliah yang ditemukan di /my/. "
            "Pastikan cookie masih berlaku dan akun punya kelas aktif."
        )

    urut = sorted(hasil.values(), key=lambda m: (m.nama.lower(), m.kelas))
    return urut


def daftar_mata_kuliah_lengkap(klien: moodle.Moodle) -> list[MataKuliah]:
    """Daftar mata kuliah yang kode dan kelasnya sudah terisi."""
    urut = daftar_mata_kuliah(klien)
    for matkul in urut:
        lengkapi_kode(klien, matkul)
    return urut


def daftar_section(klien: moodle.Moodle, matkul: MataKuliah) -> list[int]:
    """Daftar nomor section milik satu course."""
    halaman = klien.ambil(f"{klien.base_url}/course/view.php?id={matkul.id}")
    soup = BeautifulSoup(halaman.html, "lxml")

    nomor: set[int] = set()
    for opsi in soup.select("select[name=jump] option"):
        m = re.search(r"section=(\d+)", opsi.get("value") or "")
        if m:
            nomor.add(int(m.group(1)))
    if not nomor:
        # Tanpa dropdown, ambil dari tab onetopic.
        for tautan in soup.select("a.nav-link, ul.tabs li a"):
            m = re.search(r"section=(\d+)", tautan.get("href") or "")
            if m:
                nomor.add(int(m.group(1)))
    return sorted(nomor)


def lengkapi_kode(klien: moodle.Moodle, matkul: MataKuliah) -> MataKuliah:
    """Isi kode dan kelas mata kuliah dari `<title>` halaman course.

    Halaman `/my/` hanya memberi nama panjang dan teks tautannya yang juga
    nama panjang. Kode `STSI4106.90` baru muncul di judul halaman course,
    ditulis sebagai `STSI4106.90: Logika Informatika 90`. Tanpa kode ini,
    kop dokumen tidak bisa menampilkan kode mata kuliah.
    """
    if matkul.kode:
        return matkul

    try:
        halaman = klien.ambil(f"{klien.base_url}/course/view.php?id={matkul.id}")
    except moodle.MoodleError:
        return matkul

    soup = BeautifulSoup(halaman.html, "lxml")
    judul = soup.title.get_text(strip=True) if soup.title else ""

    # Bentuk paling jelas: judul halaman `KODE.KELAS: Nama | ElearningUT`.
    m = re.match(r"^\s*([A-Z]{4}\d{4})\.(\d{1,3})\s*:\s*(.+?)\s*\|", judul)
    if m:
        matkul.kode = m.group(1)
        if not matkul.kelas:
            matkul.kelas = m.group(2)
        nama_resmi = m.group(3).strip()
        if nama_resmi:
            matkul.nama = nama_resmi
        return matkul

    # Bentuk kedua: kode muncul di sidebar halaman course, pada teks
    # "My courses STSI4106.90". Judul halaman sendiri sering tidak memuatnya.
    teks = re.sub(r"\s+", " ", soup.get_text(" ", strip=True))
    m = re.search(r"\b([A-Z]{4}\d{4})\.(\d{1,3})\b", teks)
    if m:
        matkul.kode = m.group(1)
        if not matkul.kelas:
            matkul.kelas = m.group(2)
        return matkul

    # Bentuk terakhir: course tanpa kode di halaman mana pun.
    return matkul


def daftar_aktivitas(
    klien: moodle.Moodle, matkul: MataKuliah, section: int
) -> list[Aktivitas]:
    """Semua activity di satu section, apa pun jenisnya."""
    halaman = klien.ambil(
        f"{klien.base_url}/course/view.php?id={matkul.id}&section={section}"
    )
    soup = BeautifulSoup(halaman.html, "lxml")

    keluar: list[Aktivitas] = []
    for cm in soup.select("[data-for=cmitem]"):
        modtype = _tarik_modtype(cm)
        nama = _tarik_nama(cm)
        if modtype == "label" or not nama:
            continue

        node = cm.select_one(".activityname, .instancename")
        tautan = node.find("a", href=True) if node else None
        if tautan is None:
            tautan = cm.find("a", href=True)

        deskripsi = ""
        ket = cm.select_one(".activity-altcontent, .description")
        if ket:
            deskripsi = re.sub(r"\s+", " ", ket.get_text(" ", strip=True))[:300]

        keluar.append(
            Aktivitas(
                id=cm.get("data-id") or "",
                modtype=modtype,
                nama=nama,
                url=klien.absolut(tautan["href"]) if tautan else "",
                section=section,
                deskripsi=deskripsi,
            )
        )
    return keluar


def petakan_sesi(
    klien: moodle.Moodle, matkul: MataKuliah, nomor: int
) -> Sesi:
    """Rakit satu sesi: soal yang harus dijawab dan bahan ajarnya.

    Section adalah penentu utama. Bila section tersebut tidak memuat soal,
    nomor sesi dicari dari nama activity di seluruh course -- ini yang
   Diperlukan ketika admin MOOC meletakkan soal di section tetangga.
    """
    sections = daftar_section(klien, matkul)
    if nomor not in sections:
        raise moodle.MoodleError(
            f"Sesi {nomor} tidak ada di {matkul.nama}. "
            f"Sesi yang tersedia: {', '.join(str(s) for s in sections if s)}"
        )

    semua: list[Aktivitas] = []
    for s in sections:
        semua.extend(daftar_aktivitas(klien, matkul, s))

    di_section = [a for a in semua if a.section == nomor]
    soal = [a for a in di_section if a.boleh_ke_soal]
    catatan: list[str] = []

    if not soal:
        # Cari berdasarkan nomor pada nama, tapi hanya dari section yang
        # tidak punya soal sendiri. Kalau soalnya sudah menjadi milik
        # section lain,section itu yang berhak atasnya.
        section_dengan_soal = {a.section for a in semua if a.boleh_ke_soal}
        soal = [
            a for a in semua
            if a.boleh_ke_soal
            and a.nomor == nomor
            and a.section != nomor
            and a.section not in section_dengan_soal
        ]
        if soal:
            catatan.append(
                f"Soal sesi {nomor} tidak ada di section {nomor}, "
                "tetapi ada di section "
                f"{soal[0].section}."
            )

    materi_terpilih: dict[str, Aktivitas] = {}
    for a in di_section:
        if a.boleh_ke_materi:
            materi_terpilih[a.url or a.id] = a

    # Pelengkap: bahan ajar bernomor sama yang berada di section lain.
    # Ini yang menutup halaman depan, tempat beberapa course menaruh
    # "Materi Inisiasi 1".
    for a in semua:
        kunci = a.url or a.id
        if a.boleh_ke_pelengkap and a.nomor == nomor and kunci not in materi_terpilih:
            materi_terpilih[kunci] = a
            catatan.append(f"Bahan ajar diambil dari section {a.section}: {a.nama}")

    if not soal:
        ada_soal = sorted({a.section for a in semua if a.boleh_ke_soal})
        catatan.append(
            f"Section {nomor} tidak memuat soal. Section yang punya soal: "
            + ", ".join(str(s) for s in ada_soal if s)
            + ". Tanyakan ke tutor bila sesi ini memang seharusnya ada tugasnya."
        )

    return Sesi(
        nomor=nomor,
        section=nomor,
        soal=soal,
        materi=list(materi_terpilih.values()),
        catatan=catatan,
    )


def sesi_berisi_soal(
    klien: moodle.Moodle, matkul: MataKuliah
) -> list[tuple[int, int]]:
    """Pasangan (section, jumlah soal) untuk semua section.

    Dipakai saat pengguna belum memilih sesi, untuk menampilkan pilihan
    yang memang punya pekerjaan.
    """
    sections = daftar_section(klien, matkul)
    hasil: list[tuple[int, int]] = []
    for s in sections:
        if s == 0:
            continue
        aktivitas = daftar_aktivitas(klien, matkul, s)
        hasil.append((s, sum(1 for a in aktivitas if a.boleh_ke_soal)))
    return hasil