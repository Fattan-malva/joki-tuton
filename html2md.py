"""Konversi HTML Moodle menjadi Markdown untuk dibaca agent.

Tiga syarat yang membuat modul ini berdiri sendiri:

1. LaTeX harus utuh. Materi UT banyak persamaan, dan `tools/docx.py`
   mengubah LaTeX menjadi OMML (equation Word). Kalau `\\(...\\)` pecah di
   tahap konversi, persamaan hilang di dokumen akhir dan tidak ada tempat
   untuk memperbaikinya.

2. Boilerplate harus dibuang. Halaman UT membawa menu, sidebar, blok
   rekomendasi, notifikasi, dan JS. Kalau ikut diteruskan, agent membaca
   ribuan baris navigational sebelum menyentuh soalnya.

3. Jawaban mahasiswa lain harus dibuang sebelum model pernah melihatnya.
   Spec agent sudah melarang membacanya, tetapi larangan di prompt bukan
   jaminan. Di sini post mahasiswa dihapus dari dokumen, jadi isinya tidak
   pernah sampai ke model sama sekali.
"""

from __future__ import annotations

import re
from typing import Callable
from urllib.parse import unquote, urljoin

from bs4 import BeautifulSoup, NavigableString, Tag

# --------------------------------------------------------------- boilerplate

# Elemen yang isinya murni navigasi atau perilaku, bukan materi.
BUANG_ELEMEN = [
    "script:not([type='math/tex']):not([type='math/mml'])",
    "style",
    "noscript",
    "nav",
    "footer",
    "form",
    "button",
    "select",
    "input",
    "textarea",
    "link",
    "meta",
    "svg",
    "template",
]

# Kelas-kelas kerangka UT dan Moodle.
BUANG_KELAS = [
    "navbar",
    "navigation",
    "breadcrumb",
    "sidemenu",
    "drawer",
    "usermenu",
    "usernavigation",
    "footer",
    "mb2iq-footer",
    "block-footer",
    "region-top-collaborator-menu",
    "region-top-collaborator-menu-usernavigation",
    "collapsible-header",
    "mform",
    "formulation",
    "editing-toggle",
    "singlebutton",
    "action-buttons",
    "popover",
    "tooltip",
    "sr-only",
    "visually-hidden",
    "accesshide",
    "printbutton",
    "collapsible",
    "js-hidden",
    "datafilters",
    "searchbox",
    "langmenu",
    "logininfo",
]

# Region yang bukan isiermateri.
BUANG_ATTR = [
    {"role": "navigation"},
    {"role": "banner"},
    {"role": "contentinfo"},
    {"role": "search"},
    {"role": "button"},
    {"role": "menu"},
    {"hidden": "true"},
]

# Pemilih untuk memangkas bagian halaman yang bukan isiermateri.
# Diurut dari paling spesifik.
PILIH_ISI = [
    "#region-main .region-main-inner",
    "[data-region='main']",
    "#region-main",
    "main",
    "#maincontent",
    "div.forumpost",          # fallback untuk halaman forum
    "div.contents",
    "div.topic",
]

# --------------------------------------------------------LaTeX / MathJax

# Script MathJax: isinya LaTeX asli dan harus dipertahankan.
_SCRIPT_TEX = re.compile(r"script[^>]*type\s*=\s*['\"]math/tex['\"]", re.I)

# Elemen MathML. LaTeX-nya ada di <annotation encoding="application/x-tex">.
_ANNOTASI_TEX = re.compile(
    r"annotation[^>]*encoding\s*=\s*['\"](?:application/x-tex|text/tex)['\"]", re.I
)


def _selaraskan_script_mathjax(soup: BeautifulSoup) -> None:
    """Ganti pasangan span+script MathJax dengan LaTeX tunggal.

    MathJax 3 menuliskan hasil render di `<span class="MathJax">` lalu
    menyisipkan sumber aslinya di `<script type="math/tex">`. Kalau keduanya
    dibiarkan, agent melihat teks render dua kali, dan yang salah adalah
    penampungan yang tidak bisa diurai.

    Spandengan class `MathJax_Preview` adalah pratinjau yang sengaja
    disembunyikan dari mata manusia tapi tetap ada di HTML, jadi harus
    dibuang.
    """
    for script in list(soup.find_all("script")):
        if not _SCRIPT_TEX.search(str(script.get("type", "")) + str(script)):
            continue

        latex = script.get_text()
        if not latex.strip():
            script.decompose()
            continue

        # Buang span hasil render dan pratinjau tepat sebelum script.
        for nama_kelas in ("MathJax_Preview", "MathJax", "mwe-math-element"):
            for span in script.find_all_previous(class_=nama_kelas, limit=1):
                # Hanya buang kalau memang saudara langsung, supaya span
                # accordion yang kebetulan bernama sama tidak ikut hilang.
                if span.parent is script.parent:
                    span.decompose()

        script.replace_with(NavigableString(latex))


def _selaraskan_mathml(soup: BeautifulSoup) -> None:
    """Ganti blok `<math>` dengan LaTeX dari annotation-nya.

    Hanya Matematika yang punya annotation yang diganti. MathML tanpa
    annotation tidak bisa diterjemahkan tanpa pustaka LaTeX ke MathML, jadi
    isinya dibiarkan sebagai teks agar informasi tidak hilang.
    """
    for math in list(soup.find_all("math")):
        annotasi = math.find("annotation", encoding=re.compile(r"tex", re.I))
        if annotasi is None:
            annotasi = next(
                (a for a in math.find_all("annotation")
                 if _ANNOTASI_TEX.search(str(a))),
                None,
            )
        if annotasi is None:
            continue
        latex = annotasi.get_text().strip()
        if latex:
            math.replace_with(NavigableString(latex))
        else:
            math.decompose()


def _lindungi_math_teks(teks: str) -> str:
    """Kembalikan teks yang dipastikan tidak merusak blok LaTeX.

    Pemotongan baris pada `\\[ ... \\]` dan `$$ ... $$` akan menghasilkan
    LaTeX rusak, jadi spasi turun satu baris diganti dengan spasi biasa
    hanya di dalam blok tersebut.
    """
    pola = re.compile(r"(\\\[[\s\S]*?\\\]|\$\$[\s\S]*?\$\$|\\\([\s\S]*?\\\))")
    hasil = []
    akhir = 0
    for cocok in pola.finditer(teks):
        hasil.append(teks[akhir:cocok.start()])
        hasil.append(re.sub(r"\s*\n\s*", " ", cocok.group(1)))
        akhir = cocok.end()
    hasil.append(teks[akhir:])
    return "".join(hasil)


def selaraskan_pemisah_math(teks: str) -> str:
    """Ubah pemisah `\\(...\\)` dan `\\[...\\]` jadi `$...$` dan `$$...$$`.

    `tools/docx.py` hanya mengenali dua bentuk itu. Bentuk `\\(...\\)` yang
    dipakai MathJax akan lolos ke dokumen sebagai teks biasa, jadi `\\(P\\)`
    tampil sebagai tulisan `\\(P\\)`, bukan persamaan.

    Yang diproses hanya pemisah, isi LaTeX di dalamnya tidak diubah, dan
    blok `$...$` yang sudah ada tidak disentuh supaya tidak jadi dua kali.
    """
    if "\\(" not in teks and "\\[" not in teks:
        return teks

    # Ruas `$...$` yang sudah ada dikunci dulu supaya `\\[` di dalam teks
    # biasa tidak ikut ditafsirkan.
    terkunci: list[str] = []

    def kunci(m: re.Match[str]) -> str:
        terkunci.append(m.group(0))
        return f"\x00{len(terkunci) - 1}\x00"

    teks = re.sub(r"\$\$.+?\$\$|\$.+?\$", kunci, teks, flags=re.S)

    def blok(m: re.Match[str]) -> str:
        isi = m.group(1).strip()
        if not isi:
            return ""
        return f"\n\n$$\n{isi}\n$$\n\n"

    def sebaris(m: re.Match[str]) -> str:
        isi = m.group(1).strip()
        if not isi:
            return ""
        return f"${isi}$"

    # Blok dulu: `\\[` tidak boleh tertangkap pola `\\(`, dan `\\[` selalu
    # diikuti `\\]` jadi tidak ada risksalah silang.
    teks = re.sub(r"\\\[(.*?)\\\]", blok, teks, flags=re.S)
    teks = re.sub(r"\\\((.*?)\\\)", sebaris, teks, flags=re.S)

    for nomor, asli in enumerate(terkunci):
        teks = teks.replace(f"\x00{nomor}\x00", asli)

    return teks


def _spasi_latex(teks: str) -> str:
    """Cegah `\\(` berubah jadi `(` dan `$` berubah jadi markdown.

    Blok yang sudah utuh dari `_lindungi_math_teks` tidak boleh disentuh,
    karena `\\,` (ruang tipis LaTeX) dan `\\;` (ruang sedang) memang
    diawali backslash.
    """
    pola = re.compile(r"(\\\[[\s\S]*?\\\]|\$\$[\s\S]*?\$\$|\\\([\s\S]*?\\\))")

    def ganti(cocok):
        return cocok.group(1)

    return pola.sub(ganti, teks)


# ------------------------------------------------------------- kebersihan

_SPASI_GANDA = re.compile(r"[ \t]+")
_BARIS_KOSONG_GANDA = re.compile(r"\n{3,}")

# Bagian yang isinya tidak boleh diratakan spasi: blok kode, daftar
# bersarang, dan LaTeX. Tanpa placeholder, `_SPASI_GANDA` akan mengubah
# empat spasi indentasi kode menjadi satu, lalu dokumen akhir memuat kode
# Python yang salah letak.
_BLOK_UTUH = re.compile(
    r"```[\s\S]*?```"                      # blok kode berpagar
    r"|```[\s\S]*$"                        # blok kode tanpa penutup
    r"|\\\[[\s\S]*?\\\]"                   # LaTeX blok
    r"|\$\$[\s\S]*?\$\$"                   # LaTeX display
    r"|\\\([\s\S]*?\\\)"                   # LaTeX sebaris
    r"|(?m:(?:^[ \t]*(?:[-*+]|\d+[.)])[ \t]+[^\n]*\n?)+)"  # deretan butir daftar
)

# Penanda yang tidak mungkin muncul di teks hasil konversi ini.
_PLACEHOLDER_BLOK = "\x00BLOK\x00"


def bersihkan_teks(teks: str) -> str:
    """Rapatkan spasi tanpa merusak kode, daftar, dan LaTeX.

    Spasi diratakan hanya pada teks biasa. Indentasi blok kode dan butir
    daftar bersarang ikut diratakan kalau tidak disingkirkan lebih dulu,
    lalu hitungan penanda Markdown menjadi salah dan daftar kehilangan
    susunannya.
    """
    simpan: list[str] = []

    def simpan_blok(cocok: re.Match[str]) -> str:
        simpan.append(_lindungi_math_teks(cocok.group(0)))
        return f"{_PLACEHOLDER_BLOK}{len(simpan) - 1}\x00"

    teks = _BLOK_UTUH.sub(simpan_blok, teks)

    teks = _SPASI_GANDA.sub(" ", teks)
    teks = "\n".join(baris.rstrip() for baris in teks.split("\n"))
    teks = _BARIS_KOSONG_GANDA.sub("\n\n", teks)

    # Kembalikan blok yang tadi disingkirkan.
    for i, isi in enumerate(simpan):
        teks = teks.replace(f"{_PLACEHOLDER_BLOK}{i}\x00", isi)
    return teks.strip()


# -------------------------------------------------------------- Reddit

def _teks_dalam(node: Tag, separator: str = " ") -> str:
    return node.get_text(separator, strip=True)


def _tambah_kelas(tag: Tag, kelas: str) -> None:
    kelas_ada = tag.get("class") or []
    kelas_ada.append(kelas)
    tag["class"] = kelas_ada


def buang_kerangka(soup: BeautifulSoup) -> BeautifulSoup:
    """Buang navigasi, JS, notifikasi, dan kerangka tema."""
    for sel in BUANG_ELEMEN:
        for el in soup.select(sel):
            el.decompose()

    for kelas in BUANG_KELAS:
        for el in soup.find_all(class_=kelas):
            el.decompose()

    for atribut in BUANG_ATTR:
        for el in soup.find_all(attrs=atribut):
            el.decompose()

    return soup


def _pilih_isi(soup: BeautifulSoup) -> Tag:
    """Ambil elemen yang paling mungkin merupakan isi utama halaman."""
    # Post yang sudah dibersihkan harus dipertahankan utuh, termasuk
    # pembungkusnya, supaya penanda kelas `jkt-reader-bersih` tidak hilang
    # sebelum sempat dibaca.
    for tag in soup.find_all(class_=KELAS_DIBERSIHKAN):
        induk = tag.parent
        while induk is not None and induk.name != "body":
            if "region-main" in " ".join(induk.get("class") or []) or induk.name == "main":
                return induk
            induk = induk.parent
        return tag

    for sel in PILIH_ISI:
        found = soup.select(sel)
        if found:
            # `div.forumpost` sebagai fallback dipakai hanya kalau tidak ada
            # region main, jadi ambil yang paling besar.
            if len(found) > 1:
                return max(found, key=lambda e: len(e.get_text()))
            return found[0]
    return soup


# ------------------------------------------------------ buang post mahasiswa

# Header post forum pada tema mb2iq adalah tag `<header>` biasa, bukan
# `.forum-post-header`. Pola inilah yang membedakan naskah soal dari jawaban
# mahasiswa: post soal tidak punya nama+NIM penulis, sedangkan tiap balasan
# mahasiswa selalu diawali "by NAMA NIM".
POLA_PENULIS = re.compile(r"\bby\s+\S", re.I)

# Post soal ditandai Moodle dengan kelas `firstpost` dan `starter`.
KELAS_POST_SOAL = ("firstpost", "starter")

# Post yang sudah dipastikan bersih, untuk keperluan pemeriksaan ulang.
KELAS_DIBERSIHKAN = "jkt-reader-bersih"


def _teks_header_post(post: Tag) -> str:
    """Ambil teks header post, tempat nama penulis berada."""
    header = post.find("header")
    if header is None:
        header = post.select_one(
            ".forum-post-header, .postheader, .author, .post-top, .top"
        )
    if header is not None:
        return header.get_text(" ", strip=True)
    # Tanpa elemen header, ambil teks sebelum badan post. Pada mb2iq anak
    # langsung post adalah satu div berisi header dan badan, jadi 300 karakter
    # pertama sudah pasti Emoji judul dan penulis.
    return post.get_text(" ", strip=True)[:300]


def _adalah_post_soal(post: Tag) -> bool:
    """Benarkah ini naskah soal yang ditulis dosen?"""
    kelas = " ".join(post.get("class") or [])
    if any(k in kelas for k in KELAS_POST_SOAL):
        return True
    return not POLA_PENULIS.search(_teks_header_post(post))


def _pembungkus_post(post: Tag) -> Tag | None:
    """Naik ke `article.forum-post-container` milik sebuah post.

    Post balasan dibungkus article; menghapus article-nya sekali saja
    ikut membuang semua balasan di bawahnya, sehingga tidak perlu memindai
    setiap post bersarang.
    """
    node = post
    while node is not None:
        if node.name == "article" and "forum-post-container" in " ".join(
            node.get("class") or []
        ):
            return node
        node = node.parent
    return None


def _tblah_antarmuka_dalam_post(post: Tag) -> None:
    """Buang tombol kendali dan baris judul forum dari dalam post soal.

    Baris judul pada tema mb2iq memuat nama topik, tanggal, dan "Number of
    replies". Semuanya kerangka, bukan bagian soal.
    """
    # Judul topik biasanya duplikat dengan halaman, jadi cukup judul.
    for h in post.find_all(["h2", "h3", "h4", "h5"]):
        h.decompose()
    for header in post.find_all("header"):
        header.decompose()
    for sel in (
        ".forum-post-controls", ".commands", ".post-controls",
        ".btn-group", ".postbody .link", ".forumpost .link",
        "a[href*='discuss.php']", "span.forum-post-author",
    ):
        for el in post.select(sel):
            el.decompose()


def buang_jawaban_mahasiswa(soup: BeautifulSoup, *, elas: bool = False) -> None:
    """Hapus seluruh kiriman mahasiswa dari halaman soal.

    Forum: yang dipertahankan hanya post yang ditulis dosen. Semua thread
    balasan dibuang. Pemotongan dilakukan di lapisan data, bukan hanya lewat
    larangan di prompt agent, karena jawaban mahasiswa tidak boleh masuk ke
    context model sama sekali.

    Tugas: blok kiriman dan blok umpan balik tutor ikut dibuang. Umpan balik
    itu berisi analisis tutor atas jawaban mahasiswa itu sendiri, jadi
    membiarkannya masuk berarti pipeline menyalin pekerjaan yang sudah
    dinilai, bukan menjawab sendiri.
    """
    if elas:
        for sel in (
            "div.plugincontentsummary", "div.assignsubmission",
            "div.assignfeedback", "div.feedback", "div.submission",
            "div.grade", "table.grades", "div.attempts",
            "div.assignment_submission", "div.assign_plugin_supplementary",
            "div.submissionsummary",
        ):
            for blok in soup.select(sel):
                blok.decompose()

        # Tabel status memuat tautan ke berkas kiriman.
        for tabel in soup.select("table.generaltable"):
            isi = tabel.get_text(" ", strip=True).lower()
            if any(k in isi for k in ("submission status", "file submissions",
                                      "nilai", "grade", "last modified")):
                tabel.decompose()

        # Judul bagian yang tertinggal setelah tabelnya dibuang.
        for judul in soup.find_all(["h2", "h3", "h4"]):
            isi = judul.get_text(" ", strip=True).lower()
            if any(k in isi for k in ("submission status", "submission summary",
                                      "file submissions", "grade", "nilai")):
                judul.decompose()
        return

    posts = soup.select("div.forumpost")
    if not posts:
        return

    # Post soal hampir selalu yang pertama dan ditandai `firstpost`.
    soal = next((p for p in posts if _adalah_post_soal(p)), None)
    if soal is None:
        # Tidak ada post tanpa penulis. Ambil yang pertama supaya isinya
        # tidak hilang seluruhnya, dan beri tanda agar pemeriksaan di
        # `pipeline.py` bisa memperingatkan pengguna.
        soal = posts[0]
        _tambah_kelas(soal, "jkt-post-tanpa-penulis")

    for post in posts:
        if post is soal:
            continue
        target = _pembungkus_post(post) or post
        # Jangan bongkar elemen yang memuat post soal, walau struktur forum
        # ternyata membuat post soal menjadi leluhur post lain.
        if target is not soal and not any(d is soal for d in target.descendants):
            target.decompose()

    # Post soal dibersihkan dari kendali forum.
    if soal is not None and soal.parent is not None:
        _tblah_antarmuka_dalam_post(soal)
        _tambah_kelas(soal, KELAS_DIBERSIHKAN)


def buang_antarmuka_forum(soup: BeautifulSoup) -> None:
    """Buang kerangka forum: sidebar, statistik, filter, dan navigasi topik."""
    for sel in (
        "div.forumheader", "div.forumsearch", "div.forumtoolbar",
        "div.discussionlist", "div.forumsubscription",
        "ul.forumcontrols", "div.popover", "div.usermenu",
        "div.recent-users", "div.forum-user-summary",
        "div.col-md-3", "div.col-3", "nav.breadcrumb",
    ):
        for el in soup.select(sel):
            # Jangan buang pembungkus yang memuat post soal.
            if el.select("div.forumpost"):
                continue
            el.decompose()


# -------------------------------------------------------------- konversi

_ESCAPE_MD = str.maketrans({
    "\\": "\\\\",
    "`": "\\`",
    "*": "\\*",
    "_": "\\_",
})


def _teks_md(node: NavigableString) -> str:
    """Ubah teks biasa menjadi Markdown tanpa merusak LaTeX."""
    teks = str(node)
    if not teks.strip():
        return " "
    # LaTeX di dalam teks biasa dibiarkan apa adanya.
    if "\\(" in teks or "\\[" in teks or "$$" in teks:
        return teks
    teks = teks.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
    teks = teks.replace("&quot;", '"').replace("&#39;", "'").replace("&nbsp;", " ")
    # Karakter markdown di teks biasa di-escape, tapi hanya yang berarti.
    teks = teks.translate(_ESCAPE_MD)
    return teks


class _Konverter:
    """Pengubah elemen HTML menjadi Markdown."""

    def __init__(self, rewrite: Callable[[str], str] | None, base_url: str) -> None:
        self.rewrite = rewrite
        self.base_url = base_url
        self.daftar_nomor: list[int] = []

    def url(self, nilai: str) -> str:
        penuh = urljoin(self.base_url, nilai) if self.base_url else nilai
        if self.rewrite:
            return self.rewrite(penuh)
        return penuh

    # -------------------------------------------------------------- masuk

    def blok(self, node: Tag, indentasi: int = 0) -> str:
        """Rekursi utama untuk elemen lalu anak-anaknya."""
        bagian: list[str] = []
        for anak in node.children:
            if isinstance(anak, NavigableString):
                bagian.append(_teks_md(anak))
                continue
            if not isinstance(anak, Tag):
                continue
            bagian.append(self.elemen(anak, indentasi))
        return "".join(bagian)

    def elemen(self, tag: Tag, indentasi: int = 0) -> str:
        nama = tag.name.lower()

        if nama in ("script", "style"):
            return ""
        if nama in ("h1", "h2", "h3", "h4", "h5", "h6"):
            return self.judul(tag, indentasi)
        if nama == "p":
            return self.paragraf(tag, indentasi)
        if nama == "br":
            return "\n"
        if nama in ("ul", "ol"):
            return self.daftar(tag, indentasi)
        if nama == "li":
            return self.konversi_li(tag, indentasi)
        if nama in ("table",):
            return self.tabel(tag)
        if nama == "pre":
            return self.blok_kode(tag)
        if nama == "code" and tag.parent and tag.parent.name != "pre":
            return f"`{tag.get_text(strip=True)}`"
        if nama == "a":
            return self.tautan(tag)
        if nama == "img":
            return self.gambar(tag)
        if nama in ("strong", "b"):
            isi = self.blok(tag, indentasi).strip()
            return f"**{isi}**" if isi else ""
        if nama in ("em", "i"):
            isi = self.blok(tag, indentasi).strip()
            return f"*{isi}*" if isi else ""
        if nama == "del" or nama == "s":
            isi = self.blok(tag, indentasi).strip()
            return f"~~{isi}~~" if isi else ""
        if nama == "sup":
            isi = tag.get_text(strip=True)
            return f"^{isi}" if isi else ""
        if nama == "sub":
            isi = tag.get_text(strip=True)
            return f"_{isi}" if isi else ""
        if nama == "blockquote":
            isi = bersihkan_teks(self.blok(tag, indentasi))
            return "\n".join(f"> {b}" for b in isi.split("\n"))
        if nama == "hr":
            return "\n---\n"
        if nama == "dl":
            return self.definisi(tag, indentasi)
        if nama in ("dd", "dt"):
            return self.blok(tag, indentasi)
        if nama in ("figure", "figcaption"):
            return self.blok(tag, indentasi)
        if nama in ("span", "small", "font", "bdi", "bdo", "mark", "time",
                    "label", "section", "article", "main", "body", "html"):
            return self.blok(tag, indentasi)
        if nama in ("div", "center", "details", "summary"):
            isi = self.blok(tag, indentasi)
            # Div kosong yang hanya pembungkus tidak perlu baris baru.
            if not isi.strip():
                return ""
            return f"\n{isi}\n" if not isi.strip().startswith(("\n", "-", "#")) else isi
        return self.blok(tag, indentasi)

    # -------------------------------------------------------------- Utils

    def judul(self, tag: Tag, indentasi: int = 0) -> str:
        tingkat = int(tag.name[1])
        isi = bersihkan_teks(self.blok(tag, indentasi))
        if not isi:
            return ""
        return f"\n{'#' * tingkat} {isi}\n"

    def paragraf(self, tag: Tag, indentasi: int = 0) -> str:
        isi = bersihkan_teks(self.blok(tag, indentasi))
        if not isi:
            return ""
        # Paragraf yang hanya berisi gambar atau tautan tidak perlu baris kosong.
        return f"\n{isi}\n\n"

    def tautan(self, tag: Tag) -> str:
        teks = bersihkan_teks(tag.get_text(" ", strip=True))
        href = (tag.get("href") or "").strip()
        if not href or href.startswith("#"):
            return teks
        tujuan = self.url(href)
        if not teks:
            teks = tujuan
        return f"[{teks}]({tujuan})"

    def gambar(self, tag: Tag) -> str:
        src = (tag.get("src") or tag.get("data-src") or "").strip()
        if not src:
            return ""
        alt = (tag.get("alt") or "").strip()
        return f"\n![{alt}]({self.url(src)})\n"

    def definisi(self, tag: Tag, indentasi: int) -> str:
        keluar = []
        for anak in tag.find_all(["dt", "dd"], recursive=False):
            isi = bersihkan_teks(self.blok(anak, indentasi))
            if not isi:
                continue
            keluar.append(isi if anak.name == "dt" else f": {isi}")
        return "\n".join(keluar) + "\n"

    def konversi_li(self, tag: Tag, indentasi: int) -> str:
        # Isi langsung li dipisah dari daftar anak supaya penomoran tingkat
        # anak tidak menggeser penomoran induk.
        langsung: list[str] = []
        anak_daftar: list[Tag] = []
        for anak in tag.children:
            if isinstance(anak, NavigableString):
                langsung.append(_teks_md(anak))
            elif isinstance(anak, Tag):
                if anak.name in ("ul", "ol"):
                    anak_daftar.append(anak)
                else:
                    langsung.append(self.elemen(anak, indentasi))

        isi = bersihkan_teks("".join(langsung))
        if not isi and not anak_daftar:
            return ""

        pad = "    " * indentasi
        # Counter bernilai 0 menandai `ul`, yang selalu memakai bullet dan
        # tidak pernah dinomori. Untuk `ol` counter mulai dari `start`.
        if self.daftar_nomor and self.daftar_nomor[-1] > 0:
            penanda = f"{self.daftar_nomor[-1]}."
            self.daftar_nomor[-1] += 1
        else:
            penanda = "-"

        baris = [f"{pad}{penanda} {isi}" if isi else f"{pad}{penanda}"]
        for anak in anak_daftar:
            baris.append(self.daftar(anak, indentasi + 1))
        return "\n".join(baris) + "\n"

    def daftar(self, tag: Tag, indentasi: int = 0) -> str:
        if tag.name == "ol":
            awal = tag.get("start")
            self.daftar_nomor.append(int(awal) if awal and awal.isdigit() else 1)
        else:
            self.daftar_nomor.append(0)
        keluar = []
        for li in tag.find_all("li", recursive=False):
            keluar.append(self.konversi_li(li, indentasi))
        self.daftar_nomor.pop()
        return "\n".join(k for k in keluar if k)

    def blok_kode(self, tag: Tag) -> str:
        isi = tag.get_text().strip("\n")
        bahasa = ""
        kode = tag.find("code")
        if kode is not None:
            kelas = " ".join(kode.get("class") or [])
            m = re.search(r"(?:language-|lang-)([\w+#-]+)", kelas)
            if m:
                bahasa = m.group(1)
        # Indentasi isi blok kode tidak ditambahkan di sini: penambah spasi akan
        # hilang saat penanda Markdown dihitung ulang, dan Word menampilkan
        # isi apa adanya.
        isi = "\n".join(b.rstrip() for b in isi.split("\n"))
        return f"\n```{bahasa}\n{isi}\n```\n"

    # ------------------------------------------------------------- tabel

    def tabel(self, tag: Tag) -> str:
        baris = []
        for tr in tag.find_all("tr"):
            sel = tr.find_all(["th", "td"], recursive=False) or tr.find_all(["th", "td"])
            if not sel:
                continue
            isi = [bersihkan_teks(self.blok(s).replace("\n", " ")).replace("|", "\\|")
                   for s in sel]
            baris.append(isi)

        if not baris:
            return ""

        lebar = max(len(b) for b in baris)
        baris = [b + [""] * (lebar - len(b)) for b in baris]

        keluar = ["| " + " | ".join(baris[0]) + " |",
                  "|" + "|".join([" --- "] * lebar) + "|"]
        for b in baris[1:]:
            keluar.append("| " + " | ".join(b) + " |")
        return "\n" + "\n".join(keluar) + "\n"


def ke_markdown(
    html: str,
    *,
    base_url: str = "",
    rewrite: Callable[[str], str] | None = None,
    elas: bool = False,
    buang_mahasiswa: bool = True,
    ambil_isi_saja: bool = True,
    judul: str = "",
) -> str:
    """Ubah satu halaman Moodle menjadi Markdown.

    Args:
        html: HTML mentah dari Moodle.
        base_url: URL halaman, dipakai untuk menyelesaikan tautan relatif.
        rewrite: Fungsi yang mengubah URL Moodle menjadi URL Reader.
        elas: True kalau halaman yang dimuat adalah halaman Tugas.
        buang_mahasiswa: Buang kiriman dan jawaban mahasiswa.
        ambil_isi_saja: Batasi hasil ke region konten utama.
        judul: Judul opsional yang ditaruh di baris paling atas.
    """
    soup = BeautifulSoup(html or "", "lxml")

    # MathJax dan MathML lebih dulu, sebelum tag-nya dibuang.
    _selaraskan_mathml(soup)
    _selaraskan_script_mathjax(soup)

    if buang_mahasiswa:
        buang_antarmuka_forum(soup)
        buang_jawaban_mahasiswa(soup, elas=elas)

    if ambil_isi_saja:
        # Region main dibuang terpisah supaya post forum yang sudah bersih
        # ikut terbawa. Kalau tidak ada region main,`_pilih_isi` akan
        # memakai `div.forumpost`.
        soup = _pilih_isi(soup)

    buang_kerangka(soup)

    konverter = _Konverter(rewrite, base_url)
    md = konverter.blok(soup)
    md = bersihkan_teks(md)

    # Setelah baris kosong dirapikan. `\\[` jadi `$$` butuh baris kosong
    # di sekitarnya supaya `tools/docx.py` membacanya sebagai blok
    # persamaan, bukan paragraf biasa.
    md = bersihkan_teks(selaraskan_pemisah_math(md))

    if judul:
        md = f"# {judul}\n\n{md}".strip()

    return md + "\n" if md else ""


def bersihkan_nama_berkas(url: str) -> str:
    """Ambil nama berkas akhir dari URL pluginfile."""
    path = unquote(url.split("?")[0].split("#")[0])
    nama = path.rstrip("/").rsplit("/", 1)[-1]
    return nama or "lampiran"