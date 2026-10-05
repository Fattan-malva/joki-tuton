from __future__ import annotations

import copy
import os
import re
from pathlib import Path

from datetime import datetime

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

# Placeholder "referensi tidak terverifikasi" (lihat main.py). Bukan entri
# daftar pustaka, jadi docx merendernya berbeda dari butir bersitasi.
#
# Polanya sengaja longgar. Yang muncul di jawaban bukan hanya kalimat versi
# agent ("tidak ada referensi yang terverifikasi") tapi juga kalimat karangan
# penulis sendiri seperti "Tidak ada referensi tambahan yang diperlukan."
# Keduanya bukan rujukan dan tidak boleh dicetak seperti entri APA sungguhan.
_NO_REFERENSI_RE = re.compile(
    r"tidak\s+ada\s+(?:referensi|sumber)\b|"
    r"tidak\s+perlu\s+rujukan|tidak\s+ada\s+rujukan|"
    r"tidak\s+diperlukan\s+rujukan",
    re.IGNORECASE,
)

# Hanging indent APA 7, dalam twentieths of a point (unit yang dipakai OOXML).
# 720 = 720/1440 inci = 0,5 inci, persis angka yang diminta APA 7.
_APA_HANGING_TWIP = 720

# Tahun (atau `n.d.`) di dalam kurung: kunci pemecah entri referensi menjadi
# penulis, tahun, dan judul.
_TAHUN_RE = re.compile(r"\((?:n\.d\.|\d{4}[a-z]?)\)")

# Batas bagian yang dicetak miring pada ekor entri jurnal. Di APA 7 yang
# miring adalah nama majalah beserta nomornya; nomor issue, halaman, dan
# penerbit tetap tegak.
#
# Pola sebelumnya hanya "`koma + angka`", dan itu kelewat longgar: rentang
# halaman di conference proceedings (`, 55-59`) ikut cocok sehingga seluruh
# ekor entri -- termasuk tautan DOI -- ikut miring. Aturannya sekarang
# dipersempit: koma itu harus dekat dengan nomor issue (`(1)`), kata
# `Article`, atau kata kunci volume eksplisit.
_BATAS_VOLUME_RE = re.compile(
    r",\s*\d+(?=\s*\(|\s*(?:Article|art\.?|hlm|p|pp)\b|\s*[.,;]|\s*$)"
    r"|,\s*(?:Article|art\.?)\s*\d+"
    r"|\bvol\.?\s*\d+",
    re.IGNORECASE,
)

# Entri prosiding dimulai dengan "In ...". Yang miring di sana adalah judul
# prosidingnya, jadi batasnya dicari dari hlm./pp./halaman, bukan dari volume.
_AWAL_IN_RE = re.compile(r"^in\s+", re.IGNORECASE)
_EDITOR_PROSIDING_RE = re.compile(
    r"^[^,]{0,80}?\((?:Ed|Edt|Eds)\.?\)\s*,?\s*", re.IGNORECASE
)
_BATAS_PROSIDING_RE = re.compile(
    r",\s*(?:hlm\.?|hal\.?|pp\.?|p\.)\s*\d+"
    r"|,\s*\d+\s*[-–]\s*\d+\s*\.?\s*$"
    r"|\s*\.\s+(?=[A-Z])",
    re.IGNORECASE,
)

# Tautan DOI/URL di ujung entri. Dilarik keluar lebih dulu supaya tidak
# pernah ikut miring dan tidak ikut terjerat di dalam `ekor`.
_TAUTAN_RE = re.compile(r"https?://\S+|\bdoi\s*[:：]?\s*10\.\S+", re.IGNORECASE)
# "diakses 12 Oktober 2026 dari" menggantung setelah URL-nya dilepas.
_EKOR_DIAKSES_RE = re.compile(
    r"\s*[,;.]?\s*(?:di\s*akses|diakses|accessed|retrieved)\b[^,]{0,48}?\bdari\s*$",
    re.IGNORECASE,
)

# Singkatan yang titik akhirnya bukan akhir kalimat. Tanpa daftar ini
# `... (Ed.), Seminar ..., hlm. 55` terpotong tepat di "hlm." dan judul
# prosiding ikut miring.
_BUKAN_AKHIR_KALIMAT_RE = re.compile(
    r"(?:^|[\s(\[])[A-Z]\.$"  # inisial orang: "J. R. R."
    r"|\b(?:e\.g|i\.e|dkk|dsb|a\.n\.s|h\.a|hlm|hal|pp|ed|eds|vol|nos|no|jil"
    r"|cet|thn|Prof|Drs|Dr|Ir|St|Mrs|Mr)\.\s*$",
    re.IGNORECASE,
)

# Alamat Reader lokal. Agent research sering menuliskan URL `/berkas?u=...&k=...`
# sebagai "tautan" rujukan, dan `k=` itu kunci akses yang masih hidup selama
# pipeline berjalan.
#
# Nama punya akhiran `_ENTRI` karena `_URL_READER_RE` di bawah dipakai untuk
# teks soal dan polannya lebih longgar (dia mau memakai seluruh `\S*`).
_URL_READER_ENTRI_RE = re.compile(
    r"https?://(?:127\.0\.0\.1|localhost)(?::\d+)?[^\s,;)\]]*"
)
_MARKAH_TAUTAN_RE = re.compile(r"!?\[[^\]\n]*\]\([^)\n]*\)")
_TOKEN_READER_RE = re.compile(r"[?&]k=[A-Za-z0-9_\-]+")

# Baris penutup yang diminta tiap spec agent, dan sisa laporan yang tidak
# pernah boleh masuk daftar pustaka.
_BARIS_BUKAN_ENTRI_RE = re.compile(
    r"^\s*(?:SELESAI|DRAFT|BEFORE|AFTER)\s*$|"
    r"^\s*[-*]?\s*(?:Draft|Before|After|Remaining patterns)\s*[:\-]?\s*$",
    re.IGNORECASE,
)

# Ukuran judul "Daftar Pustaka". Di template penanda ini bukan heading, tapi
# `Normal` bold TNR 14. Dicocokkan di sini supaya bagian yang paling dilihat
# tutor tampil sama persis dengan contoh.
_REFERENSI_JUDUK_PT = 14

_BOLD_RE = re.compile(r"\*\*(.+?)\*\*")
_ITALIC_RE = re.compile(r"(?<!\*)\*([^*\n]+?)\*(?!\*)")
_INLINE_CODE_RE = re.compile(r"`([^`\n]+)`")
_MATRIX_INLINE_RE = re.compile(r"\[\[.*?\]\]")
_CODE_FENCE_RE = re.compile(r"^\s*(```+|~~~+)")
# Karakter kontrol yang tidak sah di XML 1.0 (lihat _strip_control_chars).
_CONTROL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")
# Tanggal tetap untuk docProps/core.xml. Dipakai secara tetap, bukan
# datetime.now(), supaya dokumen tidak membocorkan kapan file dibuat.
_NEUTRAL_DOC_TIMESTAMP = datetime(2024, 1, 1, 0, 0, 0)
_LATEX_REPLACEMENTS = {
    r"\rightarrow": "→",
    r"\to": "→",
    r"\leftarrow": "←",
    r"\leftrightarrow": "↔",
    r"\times": "×",
    r"\cdot": "·",
    r"\leq": "≤",
    r"\geq": "≥",
    r"\neq": "≠",
    r"\pm": "±",
    r"\infty": "∞",
    r"\ldots": "...",
}

# OMML (Word math) namespace
_MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

# \mathbb{R} dan kawan-kawan → karakter matematika double-struck
_MATHBB_MAP = {
    "R": "ℝ", "Z": "ℤ", "N": "ℕ", "Q": "ℚ", "C": "ℂ", "H": "ℍ",
    "P": "ℙ", "E": "𝔼", "F": "𝔽", "A": "𝔸", "B": "𝔹", "D": "𝔻",
}


def _esc(t: str) -> str:
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _omml_run(text: str) -> str:
    return f"<m:r><m:t>{_esc(text)}</m:t></m:r>"


def _omml_sup(base: str, exp: str) -> str:
    return f"<m:sSup><m:e>{base}</m:e><m:sup>{exp}</m:sup></m:sSup>"


def _omml_sub(base: str, sub: str) -> str:
    return f"<m:sSub><m:e>{base}</m:e><m:sub>{sub}</m:sub></m:sSub>"


def _omml_sqrt(inner: str) -> str:
    return (
        '<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr>'
        f"<m:deg/><m:e>{inner}</m:e></m:rad>"
    )


def _omml_matrix(cell_rows: list[list[str]]) -> str:
    body = "".join(
        "<m:mr>" + "".join(f"<m:e>{cell}</m:e>" for cell in row) + "</m:mr>"
        for row in cell_rows
    )
    matrix = f"<m:m>{body}</m:m>"
    return (
        '<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr>'
        f"<m:e>{matrix}</m:e></m:d>"
    )


# ---------------------------------------------------------------- environment
#
# LaTeX environment yang jadi kisi `m:m`, bukan teks.
#
# Versi lama hanya mengenali nama environment yang berakhiran "matrix"
# (`matrix`, `pmatrix`, `bmatrix`, `vmatrix`). Sisanya jatuh ke cabang
# `nodes.append(_omml_run(content))`, yaitu isinya ditempel apa adanya sebagai
# TEKS. Gejalanya persis yang dikeluhkan tutor: di dalam kotak equation Word
# muncul tulisan `\begin{cases} 2x + y - z = 3 \\ -x + 3y + 2z = 7`, termasuk
# `&`, `\\`, dan `\begin{array}{ccc|c}`. Yang paling sering muncul di soal
# Aljabar Linear Elementer justru yang tidak dikenali: `cases` (sistem
# persamaan linear) dan `array` (matriks augmentasi).
_ENV_KISI = frozenset({
    "matrix", "smallmatrix", "subarray", "array",
    "pmatrix", "bmatrix", "Bmatrix", "vmatrix", "Vmatrix",
    "cases", "dcases",
})

# Environment yang jadi baris rata (`m:eqArr`): `aligned`, `align`, `split`, dan
# `gather`. Bedanya dengan kisi: `&` di sini adalah titik rata-rata, bukan
# pemisah kolom.
_ENV_RATA = frozenset({
    "aligned", "align", "alignedat", "split",
    "gather", "gathered", "eqnarray",
})

# Penghapus environment -> (pembuka, penutup). `None` berarti tanpa penghapus;
# string kosong berarti disembunyikan (`cases` hanya punya `{` di kiri).
_DELIM_ENV = {
    "pmatrix": ("(", ")"),
    "bmatrix": ("[", "]"),
    "Bmatrix": ("{", "}"),
    "vmatrix": ("|", "|"),
    "Vmatrix": ("‖", "‖"),
    "cases": ("{", ""),
    "dcases": ("{", ""),
}

# Posisi kolom dari spec `\begin{array}{ccc|c}`.
_JC_KOLOM = {"l": "left", "c": "center", "r": "right"}

# Pasangan penghapus yang boleh dipakai `\left ... \right` untuk membungkus
# `matrix`/`array` yang tidak punya penghapus sendiri.
_PASANGAN_PENUTUP = {"[": "]", "(": ")", "{": "}", "|": "|"}

# Argumen opsional setelah pemisah baris: `\\[4pt]`, `\\[1.5ex]`, `\\[2mm]`.
# Tanpa pola ini, `[4pt]` ikut jadi isi sel dan muncul di dokumen.
_SPASI_OPT_RE = re.compile(
    r"^\s*\[\s*\d{0,3}(?:\.\d+)?\s*(?:pt|em|ex|mu|mm|cm|in|px|dd)?\s*\]"
)


def _bungkus_delim(dalam: str, pembuka: str | None, penutup: str | None) -> str:
    """Bungkus satu simpul OMML dengan penghapus `m:d`."""
    if not pembuka and not penutup:
        return dalam
    b = f'<m:begChr m:val="{pembuka}"/>' if pembuka else '<m:begChr m:val=""/>'
    e = f'<m:endChr m:val="{penutup}"/>' if penutup else '<m:endChr m:val=""/>'
    return f"<m:d><m:dPr>{b}{e}<m:ctrlPr/></m:dPr><m:e>{dalam}</m:e></m:d>"


def _bangun_matriks(
    sel_per_baris: list[list[str]], kolom: list[str]
) -> str:
    """OMML `m:m` dari sel yang sudah diurai, plus kolom pemisah dari spec.

    `kolom` berasal dari spec `\begin{array}` dan boleh memuat `|`. OMML tidak
    punya garis pemisah kolom seperti `|` di LaTeX, jadi `|` dijadikan kolom
    sendiri berisi `│` di tengah. Hasilnya bukan garis mati yang menempel,
    tapi matriks augmentasi tetap terbaca apa adanya -- dan itu jauh lebih
    baik daripada menampilkan `\begin{array}{ccc|c}` mentah-mentah.
    """
    nkol = max([len(r) for r in sel_per_baris] or [0]) or 1

    if kolom:
        # Spec boleh menambah kolom (setiap `|` menambah satu kolom pemisah),
        # jadi baris harus disamakan ke lebar spec, bukan hanya ke lebar
        # terpanjang dari isi.
        lebar = max(nkol, len(kolom))
        sel_bersih: list[list[str]] = []
        for row in sel_per_baris:
            row = list(row) + [""] * max(0, lebar - len(row))
            sel_bersih.append([
                _omml_run("│") if spec == "|" else row[k]
                for k, spec in enumerate(kolom)
            ])
    else:
        sel_bersih = [
            list(row) + [_omml_run("")] * max(0, nkol - len(row))
            for row in sel_per_baris
        ]

    body = "".join(
        "<m:mr>" + "".join(f"<m:e>{sel}</m:e>" for sel in row) + "</m:mr>"
        for row in sel_bersih
    )

    mpr = ""
    if kolom:
        # Per kolom kunci ini yang menentukan rata-atas/rata-tengah/rata-kiri.
        mc = "".join(
            f'<m:mc><m:mcPr><m:count m:val="1"/>'
            f'<m:mcJc m:val="{_JC_KOLOM.get(spec, "center")}"/>'
            f"</m:mcPr></m:mc>"
            for spec in kolom
        )
        mpr = f"<m:mPr><m:mcs>{mc}</m:mcs><m:ctrlPr/></m:mPr>"
    return f"<m:m>{mpr}{body}</m:m>"


def _bangun_rata(sel_per_baris: list[list[str]]) -> str:
    """OMML `m:eqArr` untuk `aligned`/`gather` dan sejenisnya.

    `&` di environment ini adalah titik rata-rata, jadi tiap baris boleh punya
    lebih dari satu `m:e`; Word menyelaraskan kolom-kolomnya sendiri.
    """
    body = "".join(
        "<m:mr>" + "".join(f"<m:e>{sel}</m:e>" for sel in row) + "</m:mr>"
        for row in sel_per_baris
    )
    return (
        '<m:eqArr><m:eqArrPr><m:maxDist m:val="0"/>'
        f"<m:objDist m:val=\"1\"/><m:ctrlPr/></m:eqArrPr>{body}</m:eqArr>"
    )


def _pecah_baris(konten: str) -> list[str]:
    """Pecah isi environment per `\\`, buang argumen spasi opsional."""
    return [_SPASI_OPT_RE.sub("", b) for b in re.split(r"\\\\", konten)]


def _pecah_sel(baris: str) -> list[str]:
    """Pecah satu baris per `&` lalu urai tiap sel jadi OMML."""
    return [
        "".join(_parse_latex_math(sel)) or _omml_run("")
        for sel in baris.split("&")
    ]


def _ada_isi(sel: list[str]) -> bool:
    return any(re.sub(r"<[^>]+>", "", s).strip() for s in sel)


def _kolom_dari_spek(spek: str) -> list[str]:
    """`ccc|c` -> `['c', 'c', 'c', '|', 'c']`.

    `|` disisipkan sebagai kolom tersendiri supaya jadi garis di antara kolom
    ke-3 dan ke-4 pada matriks augmentasi.
    """
    kolom: list[str] = []
    i, n = 0, len(spek)
    while i < n:
        c = spek[i]
        if c == "@":
            j = spek.find("}", i)
            i = (j + 1) if j != -1 else n
            continue
        if c in "lcr|":
            kolom.append(c)
        i += 1
    # `|` di spec menandai batas SEBELUM kolom berikutnya, jadi ia harus punya
    # satu karakter spec di sebelah kanannya.
    out: list[str] = []
    for k, c in enumerate(kolom):
        if c == "|":
            out.append("|")
            out.append("c")
        else:
            out.append(c)
    return out


def _ambil_spek(konten: str) -> tuple[str, str]:
    """Ambil argumen wajib `\begin{array}{spec}` dari awal konten."""
    teks = konten.lstrip()
    if not teks.startswith("{"):
        return "", konten
    depth = 0
    for i, ch in enumerate(teks):
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return teks[1:i], teks[i + 1 :]
    return "", konten


def _render_environment(env: str, konten: str) -> str:
    r"""OMML untuk satu `\begin{env} ... \end{env}`."""
    if env in _ENV_RATA:
        baris = [_pecah_sel(b) for b in _pecah_baris(konten)]
        baris = [b for b in baris if _ada_isi(b)]
        return _bangun_rata(baris) if baris else ""

    spek = ""
    if env in ("array", "subarray"):
        # Hanya `array`/`subarray` yang mewajibkan spec kolom. `cases` tidak,
        # dan baris pertamanya boleh saja diawali kurung kurawal.
        spek, konten = _ambil_spek(konten)

    baris = [_pecah_sel(b) for b in _pecah_baris(konten)]
    baris = [b for b in baris if _ada_isi(b)]
    if not baris:
        return ""

    kisi = _bangun_matriks(baris, _kolom_dari_spek(spek))
    pembuka, penutup = _DELIM_ENV.get(env, (None, None))
    return _bungkus_delim(kisi, pembuka, penutup)


def _cocok_penghapus(nodes: list[str], simpul: str, s: str, i: int) -> str:
    r"""Pasangkan `[`/`]` yang sudah tercetak dengan grid di dalamnya.

    `\left[\begin{array}...\end{array}\right]` menghasilkan node `[` lalu grid
    lalu node `]`. Tanpa langkah ini ketiganya jadi tiga run biasa, sehingga
    kurung spon tidak ikut meninggi bersama matriksnya. Di sini node `[` diambil
    kembali, `]` dilewati, dan keduanya diganti penghapus `m:d` yang benar.
    """
    if not nodes:
        return simpul
    pembuka = re.sub(r"<[^>]+>", "", nodes[-1])
    if pembuka not in _PASANGAN_PENUTUP:
        return simpul

    j = i
    while j < len(s) and s[j] == " ":
        j += 1
    # `\right` ada di `_CMD_MAP` dan jadi string kosong, jadi harus dilewati
    # di sini supaya penutupnya sendiri yang dibaca.
    if s.startswith("\\right", j):
        j += len("\\right")
        while j < len(s) and s[j] == " ":
            j += 1
    if j < len(s) and s[j] in _PASANGAN_PENUTUP.values():
        nodes.pop()
        return _bungkus_delim(simpul, pembuka, s[j])
    return simpul


def _omml_fraction(num: str, den: str, *, bar: bool = True) -> str:
    num_xml = "".join(_parse_latex_math(num)) or _omml_run("")
    den_xml = "".join(_parse_latex_math(den)) or _omml_run("")
    kind = "bar" if bar else "noBar"
    return (
        f"<m:f><m:fPr><m:type m:val=\"{kind}\"/></m:fPr>"
        f"<m:num>{num_xml}</m:num><m:den>{den_xml}</m:den></m:f>"
    )


def _omml_sqrt_n(inner: str, deg: str = "") -> str:
    if deg:
        return (
            '<m:rad><m:radPr><m:degHide m:val="0"/></m:radPr>'
            f'<m:deg>{deg}</m:deg><m:e>{inner}</m:e></m:rad>'
        )
    return _omml_sqrt(inner)


def _omml_math_accent(cmd: str, inner: str) -> str:
    _ACCENT_CHR = {
        "hat": "̂", "widehat": "̂", "bar": "̄", "overline": "̄",
        "underline": "̲", "vec": "⃗", "dot": "̇", "ddot": "̈",
        "tilde": "̃", "widetilde": "̃", "check": "̌",
        "acute": "́", "grave": "̀", "breve": "̆",
    }
    inner_xml = "".join(_parse_latex_math(inner)) or _omml_run("")
    chr_ = _ACCENT_CHR.get(cmd, "")
    return (
        f'<m:acc><m:accPr><m:chr m:val="{chr_}"/></m:accPr>'
        f"<m:e>{inner_xml}</m:e></m:acc>"
    )


def _omml_nary(chr_: str, sub: str, sup: str, inner: str) -> str:
    sub_xml = "".join(_parse_latex_math(sub)) or _omml_run("")
    sup_xml = "".join(_parse_latex_math(sup)) or _omml_run("")
    inner_xml = "".join(_parse_latex_math(inner)) or _omml_run("")
    return (
        "<m:nary><m:naryPr>"
        f'<m:chr m:val="{chr_}"/><m:limLoc m:val="undOvr"/>'
        f'<m:subHide m:val="{"0" if sub else "1"}"/>'
        f'<m:supHide m:val="{"0" if sup else "1"}"/>'
        "</m:naryPr>"
        f"<m:sub>{sub_xml}</m:sub><m:sup>{sup_xml}</m:sup>"
        f"<m:e>{inner_xml}</m:e></m:nary>"
    )


def _omml_lim(base_cmd: str, limit: str, over: bool) -> str:
    """OMML untuk \\overset{}{} / \\underset{}{} / \\xrightarrow{}.

    ``base_cmd`` adalah tulisan LaTeX pada posisi utama (mis. \\rightarrow),
    ``limit`` adalah label di atas/bawah. Gunakan m:limUpp / m:limLow.
    """
    base_xml = "".join(_parse_latex_math(base_cmd)) or _omml_run("")
    limit_xml = "".join(_parse_latex_math(limit)) or _omml_run("")
    if over:
        return f"<m:limUpp><m:e>{base_xml}</m:e><m:lim>{limit_xml}</m:lim></m:limUpp>"
    return f"<m:limLow><m:e>{base_xml}</m:e><m:lim>{limit_xml}</m:lim></m:limLow>"


def _parse_latex_math(s: str) -> list[str]:
    """Parse LaTeX math string into OMML XML element list.

    Handles: \\begin{bmatrix}...\\end{bmatrix}, \\sqrt{}, x^{n},
    \\alpha, \\times, \\cdot, \\neq, \\leq, \\geq, \\text{}, and
    inline * as ×.  Unknown commands are rendered as plain text.
    """
    _ALPHA_MAP = {
        "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ",
        "epsilon": "ε", "theta": "θ", "lambda": "λ", "mu": "μ",
        "pi": "π", "sigma": "σ", "phi": "φ", "omega": "ω",
        "rho": "ρ", "tau": "τ", "psi": "ψ", "chi": "χ",
        "eta": "η", "kappa": "κ", "xi": "ξ", "zeta": "ζ",
        "Delta": "Δ", "Sigma": "Σ", "Omega": "Ω", "Pi": "Π",
        "Phi": "Φ", "Psi": "Ψ", "Theta": "Θ", "Lambda": "Λ",
    }
    _CMD_MAP = {
        "times": "×", "cdot": "·", "pm": "±", "mp": "∓",
        "leq": "≤", "geq": "≥", "neq": "≠", "approx": "≈",
        "equiv": "≡", "sim": "∼", "propto": "∝", "mid": "|",
        "le": "≤", "ge": "≥", "ne": "≠", "lt": "<", "gt": ">",
        "rightarrow": "→", "leftarrow": "←", "leftrightarrow": "↔",
        "Rightarrow": "⇒", "Leftarrow": "⇐", "Leftrightarrow": "⇔",
        "to": "→", "mapsto": "↦", "longrightarrow": "⟶",
        "longleftarrow": "⟵", "Longrightarrow": "⟹", "Longleftarrow": "⟸",
        "Longleftrightarrow": "⟺", "iff": "⟺", "%": "%",
        "Longmapsto": "⟼", "hookrightarrow": "↪", "hookleftarrow": "↩",
        "rightharpoonup": "⇀", "leftharpoonup": "↼", "rightharpoondown": "⇁",
        "leftharpoondown": "↽", "rightsquigarrow": "⇝", "uparrow": "↑",
        "downarrow": "↓", "updownarrow": "↕", "Uparrow": "⇑", "Downarrow": "⇓",
        "Updownarrow": "⇕", "nearrow": "↗", "searrow": "↘", "nwarrow": "↖",
        "swarrow": "↙", "rightleftharpoons": "⇌", "dashrightarrow": "⇢",
        "infty": "∞", "partial": "∂", "nabla": "∇",
        "forall": "∀", "exists": "∃", "in": "∈", "notin": "∉",
        "subset": "⊂", "supset": "⊃", "subseteq": "⊆", "supseteq": "⊇",
        "cup": "∪", "cap": "∩",
        "land": "∧", "wedge": "∧", "bigwedge": "⋀",
        "lor": "∨", "vee": "∨", "bigvee": "⋁",
        "lnot": "¬", "neg": "¬",
        "top": "⊤", "bot": "⊥", "vdash": "⊢", "vDash": "⊨",
        "models": "⊨", "implies": "⇒", "impliedby": "⇐",
        "oplus": "⊕", "otimes": "⊗", "ominus": "⊖", "odot": "⊙",
        "Box": "◻", "Diamond": "◇", "triangle": "△", "wr": "≀",
        "dots": "…", "dotsc": "…", "vdots": "⋮",
        "emptyset": "∅", "ldots": "…", "cdots": "⋯", "vdots": "⋮",
        "ddots": "⋱", "hbar": "ℏ", "ell": "ℓ",
        "langle": "⟨", "rangle": "⟩", "surd": "√", "backslash": "∖",
        "circ": "∘", "bullet": "∙", "cdotp": "·", "colon": ":",
        "sum": "∑", "prod": "∏", "int": "∫", "iint": "∬", "iiint": "∭",
        "oint": "∮", "bigcup": "⋃", "bigcap": "⋂",
        "bigoplus": "⊕", "bigotimes": "⊗", "coprod": "∐",
        "lim": "lim", "limsup": "lim sup", "liminf": "lim inf",
        "left": "", "right": "", "newcommand": "", "operatorname": " ",
        "text": " ", "quad": " ", "qquad": "  ", "hline": "",
        # Perintah yang sering tampil bersama environment `cases`, `aligned`,
        # dan `array`: pengukur `\left`/`\right` dan pemisah `\vert`.
        "vert": "|", "Vert": "‖", "lvert": "|", "rvert": "|",
        "lVert": "‖", "rVert": "‖", "nvert": "∤",
        "lbrace": "{", "rbrace": "}", "lbrack": "[", "rbrack": "]",
        "displaystyle": "", "textstyle": "", "scriptstyle": "",
        "nonumber": "", "notag": "", "label": "",
        "bmod": " mod ", "pmod": " mod ", "mod": "mod",
        "because": "∵", "therefore": "∴",
        "ll": "≪", "gg": "≫",
        "leqslant": "⩽", "geqslant": "⩾",
        "coloncolon": "::",
    }

    nodes: list[str] = []
    i, n = 0, len(s)

    def _read_braced(pos: int) -> tuple[str, int]:
        while pos < n and s[pos] == " ":
            pos += 1
        if pos < n and s[pos] == "{":
            depth, start = 1, pos + 1
            pos += 1
            while pos < n and depth:
                if s[pos] == "{":
                    depth += 1
                elif s[pos] == "}":
                    depth -= 1
                pos += 1
            return s[start : pos - 1], pos
        if pos < n:
            return s[pos], pos + 1
        return "", pos

    while i < n:
        c = s[i]

        if c == " " or c == "\n":
            i += 1
            continue

        # LaTeX command
        if c == "\\":
            i += 1
            if i >= n:
                nodes.append(_omml_run("\\"))
                break
            nxt = s[i]
            # Escaped special: \{ \} \_ \^ \* \% \$
            if nxt in ("{", "}", "_", "^", "*", "%", "$", "|"):
                nodes.append(_omml_run(nxt))
                i += 1
                continue
            # Spacing commands: \  \, \; \: \! 
            if nxt == " ":
                i += 1
                continue
            if nxt in ",;:!":
                nodes.append(_omml_run(" " if nxt != "!" else ""))
                i += 1
                continue
            # Command name
            if nxt.isalpha():
                j = i
                while j < n and s[j].isalpha():
                    j += 1
                cmd = s[i:j]
                i = j
                if cmd in _ALPHA_MAP:
                    nodes.append(_omml_run(_ALPHA_MAP[cmd]))
                elif cmd in ("frac", "dfrac", "tfrac", "cfrac"):
                    num, i = _read_braced(i)
                    den, i = _read_braced(i)
                    nodes.append(_omml_fraction(num, den))
                elif cmd == "binom":
                    num, i = _read_braced(i)
                    den, i = _read_braced(i)
                    inner = _omml_fraction(num, den, bar=False)
                    nodes.append(
                        '<m:d><m:dPr><m:begChr m:val="("/>'
                        '<m:endChr m:val=")"/></m:dPr>'
                        f"<m:e>{inner}</m:e></m:d>"
                    )
                elif cmd == "sqrt":
                    deg = None
                    if i < n and s[i] == "[":
                        end = s.find("]", i)
                        if end != -1:
                            deg = s[i + 1:end]
                            i = end + 1
                    inner, i = _read_braced(i)
                    inner_xml = "".join(_parse_latex_math(inner))
                    deg_xml = "".join(_parse_latex_math(deg)) if deg else ""
                    nodes.append(_omml_sqrt_n(inner_xml, deg_xml))
                elif cmd == "text":
                    inner, i = _read_braced(i)
                    nodes.append(_omml_run(inner))
                elif cmd == "operatorname":
                    inner, i = _read_braced(i)
                    nodes.append(_omml_run(inner))
                elif cmd == "mathbb":
                    inner, i = _read_braced(i)
                    ch = inner.strip()
                    nodes.append(_omml_run(_MATHBB_MAP.get(ch, ch)))
                elif cmd in ("mathrm", "mathbf", "mathit", "mathcal",
                             "mathtt", "mathsf", "mathscr", "mathnormal"):
                    inner, i = _read_braced(i)
                    nodes.append(_omml_run(inner))
                elif cmd in ("hat", "widehat", "bar", "overline", "underline",
                             "vec", "dot", "ddot", "tilde", "widetilde",
                             "check", "acute", "grave", "breve"):
                    inner, i = _read_braced(i)
                    nodes.append(_omml_math_accent(cmd, inner))
                elif cmd in ("overset", "stackrel"):
                    over, i = _read_braced(i)
                    base, i = _read_braced(i)
                    nodes.append(_omml_lim(base, over, True))
                elif cmd == "underset":
                    under, i = _read_braced(i)
                    base, i = _read_braced(i)
                    nodes.append(_omml_lim(base, under, False))
                elif cmd in ("xrightarrow", "longrightarrow", "xRightarrow",
                             "xleftarrow", "xleftrightarrow", "xmapsto"):
                    arrow_chr = {
                        "xrightarrow": "→", "longrightarrow": "⟶",
                        "xRightarrow": "⇒", "xleftarrow": "←",
                        "xleftrightarrow": "↔", "xmapsto": "↦",
                    }[cmd]
                    label, i = _read_braced(i)
                    nodes.append(_omml_lim(arrow_chr, label, True))
                elif cmd == "begin":
                    env, i = _read_braced(i)
                    content, i = _read_until_end(s, i, env)
                    # `\begin{align*}` dan `\begin{equation*}` membawa bintang
                    # pada nama environment-nya.
                    nama_env = env.strip().rstrip("*")
                    if nama_env in _ENV_KISI or nama_env in _ENV_RATA:
                        simpul = _render_environment(nama_env, content)
                        # `matrix`/`array` tidak punya penghapus sendiri; kalau
                        # dikurung `\left[ ... \right]`, kurungnya dipindahkan
                        # jadi penghapus `m:d` yang benar.
                        if nama_env not in _DELIM_ENV:
                            simpul = _cocok_penghapus(nodes, simpul, s, i)
                        if simpul:
                            nodes.append(simpul)
                        else:
                            # Environment kosong: jangan sisakan kotak kosong.
                            nodes.append(_omml_run(content.strip()))
                    else:
                        nodes.append(_omml_run(content))
                elif cmd in _CMD_MAP:
                    rep = _CMD_MAP[cmd]
                    if rep:
                        nodes.append(_omml_run(rep))
                else:
                    nodes.append(_omml_run(cmd))
                continue
            # \ followed by non-alpha: treat as text
            nodes.append(_omml_run(nxt))
            i += 1
            continue

        # Superscript
        if c == "^":
            base = nodes.pop() if nodes else _omml_run("")
            i += 1
            exp, i = _read_braced(i)
            nodes.append(_omml_sup(base, "".join(_parse_latex_math(exp)) if exp else _omml_run("")))
            continue

        # Subscript
        if c == "_":
            base = nodes.pop() if nodes else _omml_run("")
            i += 1
            sub, i = _read_braced(i)
            sub_nodes = "".join(_parse_latex_math(sub)) if sub else _omml_run("")
            nodes.append(_omml_sub(base, sub_nodes))
            continue

        # Brace group
        if c == "{":
            depth, start = 1, i + 1
            i += 1
            while i < n and depth:
                if s[i] == "{":
                    depth += 1
                elif s[i] == "}":
                    depth -= 1
                i += 1
            inner = s[start : i - 1]
            nodes.extend(_parse_latex_math(inner))
            continue

        # Parenthesized group → <m:d> delimiter (so (AB)_{11} stays one base
        # AND inner LaTeX like \beta_0 is parsed recursively, not raw text).
        if c == "(":
            depth, j = 1, i + 1
            while j < n and depth:
                if s[j] == "(":
                    depth += 1
                elif s[j] == ")":
                    depth -= 1
                j += 1
            if depth == 0:
                inner = "".join(_parse_latex_math(s[i + 1 : j - 1])) or _omml_run("")
                nodes.append(
                    '<m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/>'
                    f"</m:dPr><m:e>{inner}</m:e></m:d>"
                )
                i = j
            else:
                nodes.append(_omml_run(s[i:]))
                i = n
            continue

        # ASCII matrix: [[a, b], [c, d]]
        if s.startswith("[[", i):
            j = s.find("]]", i)
            if j != -1:
                inner = s[i + 2 : j]
                rows: list[list[str]] = []
                for part in re.split(r"\]\s*,\s*\[", inner):
                    cells = [c.strip() for c in part.split(",") if c.strip()]
                    rows.append(
                        ["".join(_parse_latex_math(c)) if c else _omml_run("") for c in cells]
                    )
                nodes.append(_omml_matrix(rows))
                i = j + 2
                continue
            nodes.append(_omml_run(s[i]))
            i += 1
            continue

        # Numeric literal
        m = re.match(r"\d+(?:\.\d+)?", s[i:])
        if m:
            nodes.append(_omml_run(m.group()))
            i += m.end()
            continue

        # Alphabetic variable (multi-letter run)
        if c.isalpha():
            j = i
            while j < n and s[j].isalpha():
                j += 1
            nodes.append(_omml_run(s[i:j]))
            i = j
            continue

        # * outside \-command → ×
        if c == "*":
            nodes.append(_omml_run("×"))
            i += 1
            continue

        # ASCII arrows & relations → Unicode (jika AI menulis `A -> B`).
        #
        # URUTAN WAJIB: yang paling panjang dulu. Versi lama menaruh `<=` setelah
        # `<-`/`->`, dan `<-` memang tidak menabrak `<=`... tapi `<=` sendiri
        # dipetakan ke `⇐`, bukan `≤`. Hasilnya `a <= b` tampil sebagai
        # "a ⇐ b", yang lebih buruk daripada teks mentahnya: terlihat seperti
        # operator yang berbeda, dan pengajar bisa salah baca.
        _ASCII_REPLS = (
            ("<=>", "⟺"), ("<->", "↔"), ("->>", "⟶"), ("-->>", "⟶"),
            ("-->", "⟶"), ("<--", "⟵"), ("<<-", "⟵"), ("->", "→"), ("<-", "←"),
            ("=>", "⇒"), ("<=", "≤"), (">=", "≥"), ("!=", "≠"), ("<>", "≠"),
            ("|-", "∣"), ("|=", "≡"),
        )
        matched = None
        for ascii_seq, uni in _ASCII_REPLS:
            if s.startswith(ascii_seq, i):
                matched = (ascii_seq, uni)
                break
        if matched:
            nodes.append(_omml_run(matched[1]))
            i += len(matched[0])
            continue

        # Everything else: parentheses, operators, brackets, etc.
        nodes.append(_omml_run(c))
        i += 1

    return nodes


def _read_until_end(s: str, i: int, env: str) -> tuple[str, int]:
    r"""Read content until \\end{env}, returning (content, new_index).

    `\end{array*}` juga diterima untuk environment berversi bintang, karena
    `\begin{array*}` menulis `\end{array*}`.
    """
    tag = f"\\end{{{env}}}"
    j = s.find(tag, i)
    if j == -1 and env.rstrip("*"):
        tag = f"\\end{{{env.rstrip('*')}*}}"
        j = s.find(tag, i)
    if j == -1:
        return s[i:], len(s)
    return s[i:j], j + len(tag)


def _equation_omml(math_str: str) -> str:
    content = "".join(_parse_latex_math(math_str))
    return f'<m:oMath xmlns:m="{_MATH_NS}">{content}</m:oMath>'


def _append_equation(paragraph, math_str: str):
    """Tempel satu Word equation (OMML) ke paragraf.

    Kalau tidak ada yang bisa di-parse, JANGAN tempel apa pun. Sebelumnya
    jalur ini selalu menempel `<m:oMath><m:r><m:t/></m:r></m:oMath>` untuk
    input kosong, dan itu muncul sebagai baris kosong di Word -- tepat di
    sebelah teks yang sebenarnya sudah tercetak. Gejalanya: di preview terlihat
    "-", lalu di Word ada satu baris kosong tambahan setelah tiap blok matematika.
    """
    if not (math_str or "").strip():
        return
    if not _parse_latex_math(math_str):
        # Tidak ada satu pun simpul: lebih baik teks biasa daripada kotak kosong.
        _add_runs(paragraph, math_str.strip())
        return
    try:
        from lxml import etree

        el = etree.fromstring(_equation_omml(math_str).encode("utf-8"))
        paragraph._p.append(el)
    except Exception:  # noqa: BLE001
        paragraph.add_run(math_str)


def _add_runs(paragraph, text: str):
    """Terapkan **bold** dan *italic* sederhana ke satu paragraf."""
    text = _normalize_text(text)
    pos = 0
    for m in _BOLD_RE.finditer(text):
        if m.start() > pos:
            _add_italic_runs(paragraph, text[pos:m.start()])
        run = paragraph.add_run(m.group(1))
        run.bold = True
        pos = m.end()
    _add_italic_runs(paragraph, text[pos:])


def _add_italic_runs(paragraph, text: str):
    """Terapkan *italic* dan konversi matriks inline [[...]] jadi equation."""
    text = _normalize_text(text)
    pos = 0
    for m in _MATRIX_INLINE_RE.finditer(text):
        if m.start() > pos:
            _add_italic_text(paragraph, text[pos:m.start()])
        _append_equation(paragraph, m.group(0))
        pos = m.end()
    _add_italic_text(paragraph, text[pos:])


def _add_italic_text(paragraph, text: str):
    """Terapkan *italic* dan konversi ruas LaTeX inline (\\beta, \\sigma^2, ...)
    yang ditulis tanpa $...$ menjadi equation."""
    pos = 0
    for start, end in _find_math_spans(text):
        if start > pos:
            _add_italic_plain(paragraph, text[pos:start])
        _append_equation(paragraph, text[start:end])
        pos = end
    _add_italic_plain(paragraph, text[pos:])


def _add_italic_plain(paragraph, text: str):
    """Tambah run dengan dukungan *italic* dan inline ``code`` (monospace)."""
    pos = 0
    for m in _INLINE_CODE_RE.finditer(text):
        if m.start() > pos:
            _add_italic_plain_runs(paragraph, text[pos:m.start()])
        run = paragraph.add_run(m.group(1))
        run.font.name = "Consolas"
        run.font.size = Pt(10)
        pos = m.end()
    _add_italic_plain_runs(paragraph, text[pos:])


def _add_italic_plain_runs(paragraph, text: str):
    # Tanda perbandingan ASCII di prosa --> lambang yang benar. Tanpa ini
    # "a <= b" dan "a <> b" tercetak apa adanya. Di Times New Roman `<=`
    # terbaca seperti dua tanda yang berdempet, bukan satu operator, dan itu
    #itulah yang dikeluhkan tutor: tandanya tidak rapi dan ambigu. Cuma teks prosa --
    # inline code ditangani terpisah dan tidak boleh berubah.
    text = _normalisasi_operator(text)
    pos = 0
    for m in _ITALIC_RE.finditer(text):
        if m.start() > pos:
            paragraph.add_run(text[pos:m.start()])
        run = paragraph.add_run(m.group(1))
        run.italic = True
        pos = m.end()
    if pos < len(text):
        paragraph.add_run(text[pos:])


# Pasangan tanda perbandingan ASCII --> lambang. Urutan dari yang terpanjang.
_OPERATOR_PL = (
    ("<=>", "⟺"), ("<->", "↔"), ("<>", "≠"), ("<=", "≤"), (">=", "≥"),
    ("!=", "≠"), ("->", "→"), ("<-", "←"), ("=>", "⇒"),
)


def _normalisasi_operator(text: str) -> str:
    """Ganti tanda perbandingan ASCII dengan lambang matematika.

    Hanya untuk teks prosa. Kode inline dan blok kode tidak boleh tersentuh,
    karena `a <= b` di dalam contoh kode memang harus tetap seperti itu.
    """
    if not text or ("<" not in text and ">" not in text and "!" not in text):
        return text
    out = text
    for ascii_seq, uni in _OPERATOR_PL:
        if ascii_seq in out:
            out = out.replace(ascii_seq, uni)
    return out


def _strip_control_chars(text: str) -> str:
    """Buang karakter kontrol yang tidak sah di XML 1.0.

    lxml melempar "ValueError: All strings must be XML compatible" kalau ada
    NUL atau 0x01-0x1f di dalam run. Teks hasil scraping Moodle bisa memuat
    karakter seperti itu bila halaman resource/File mengirim binary mentah.
    """
    if not text:
        return ""
    return _CONTROL_RE.sub("", text)


def _normalize_text(text: str) -> str:
    """Hilangkan artefak LaTeX/Markdown yang tidak boleh tampil mentah di Word."""
    text = _strip_control_chars(text)
    for source, replacement in _LATEX_REPLACEMENTS.items():
        text = text.replace(source, replacement)
    text = re.sub(r"\\(?:left|right)\b", "", text)
    text = re.sub(r"\\text\{([^{}]*)\}", r"\1", text)
    text = re.sub(r"(?<![A-Za-z])rightarrow(?![A-Za-z])", "→", text)
    text = re.sub(r"(?<![A-Za-z])leftarrow(?![A-Za-z])", "←", text)
    return text


_LATEX_CMD_RE = re.compile(r"\\[A-Za-z]+")

# Kata yang sah di dalam ekspresi matematika (nama fungsi, kata LaTeX).
_MATH_WORDS = {
    "alpha", "beta", "gamma", "delta", "epsilon", "zeta", "eta", "theta",
    "iota", "kappa", "lambda", "mu", "nu", "xi", "omicron", "pi", "rho",
    "sigma", "tau", "upsilon", "phi", "chi", "psi", "omega",
    "Normal", "Bernoulli", "Binomial", "Poisson", "Exponential", "Uniform",
    "Gaussian", "sim", "mid", "given", "s.t", "of", "and", "or",
    "sin", "cos", "tan", "cot", "sec", "csc", "arcsin", "arccos", "arctan",
    "sinh", "cosh", "tanh", "log", "ln", "exp", "lim", "sup", "inf", "max",
    "min", "Pr", "P", "E", "Var", "Cov", "Corr", "det", "dim", "mod", "Ker",
    "Im", "Re", "adj", "sum", "prod", "int", "frac", "sqrt", "times", "cdot",
    "left", "right", "quad", "text", "begin", "end", "pmatrix", "bmatrix",
    "gather", "eqnarray", "matrix", "mid", "vert", "Vert",
}


def _tokenize_rough(line: str) -> list[str]:
    return re.findall(r"[A-Za-z]+|\d+(?:\.\d+)?|[^\sA-Za-z0-9]+", line)


def _line_is_display_math(line: str) -> bool:
    """True bila satu baris (tanpa $) sebagian besar adalah ekspresi LaTeX,
    mis. `YmidX;∼;Normal(\\beta_0 + \\beta_1 x_1,\\; \\sigma^2)`."""
    line = (line or "").strip()
    if not line or "$" in line:
        return False  # sudah ditangani jalur $...$ / $$...$$
    if not _LATEX_CMD_RE.search(line):
        return False  # tidak ada perintah LaTeX → teks biasa
    toks = _tokenize_rough(line)
    if not toks:
        return False
    prose = sum(
        1 for t in toks
        if t.isalpha() and len(t) >= 4 and t not in _MATH_WORDS
    )
    # Maksimal ~20% token adalah kata prosa murni: kalimat seperti
    # `Jadi nilai yang dicari adalah \beta = 5` jangan dijadikan equation.
    return prose * 5 <= len(toks)


def _scan_math_run_end(text: str, j: int) -> int:
    """Perluas dari posisi j sampai batas akhir ruas matematika yang masuk akal."""
    n = len(text)
    depth = 0
    while j < n:
        c = text[j]
        if c == " " and depth == 0:
            k = j
            while k < n and text[k] == " ":
                k += 1
            m = re.match(r"[\\A-Za-z0-9_.^]+", text[k:k + 10])
            nxt = m.group(0) if m else ""
            if nxt and "\\" not in nxt and re.fullmatch(r"[A-Za-z]+", nxt):
                if nxt.lower() not in _MATH_WORDS:
                    return j
                j = k
                continue
            j = k
            continue
        if c in "([{":
            depth += 1
        elif c in ")]}":
            depth -= 1
            if depth < 0:
                return j
        j += 1
    return n


def _find_math_spans(text: str) -> list[tuple[int, int]]:
    """Temukan ruas inline LaTeX tanpa $ di dalam teks → list (start, end).
    Ruas diperluas ke kurung pembuka/penutup yang mengelilinginya, jadi
    `Normal(\\beta_0 + \\beta_1, \\sigma^2)` ikut ditangkap mencakup parens."""
    spans: list[tuple[int, int]] = []
    i = 0
    n = len(text)
    while i < n:
        m = _LATEX_CMD_RE.search(text, i)
        if not m:
            break
        start = m.start()
        end = _scan_math_run_end(text, m.end())
        if end > start:
            s2, e2 = start, end
            for op, cl in (("(", ")"), ("[", "]"), ("{", "}")):
                k = s2 - 1
                while k >= 0 and text[k] == " ":
                    k -= 1
                t = e2
                while t < n and text[t] == " ":
                    t += 1
                if k >= 0 and text[k] == op and t < n and text[t] == cl:
                    s2, e2 = k, t + 1
                    break
            spans.append((s2, e2))
            i = max(e2, m.end())
    return spans


def _is_table_row(line: str) -> bool:
    return line.strip().startswith("|") and line.strip().endswith("|") and "|" in line[1:]


def _shade_paragraph(paragraph, fill: str = "F2F2F2"):
    """Beri latar abu-abu tipis pada paragraf (hasil: blok kode terlihat)."""
    ppr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    ppr.append(shd)


def _add_code_block(doc: Document, code: str):
    """Render blok kode (` ```sql ... ``` `) sebagai satu paragraf bergaya
    monospace dengan latar abu-abu dan indent, baris dipertahankan."""
    lines = code.replace("\r\n", "\n").split("\n")
    p = doc.add_paragraph()
    _shade_paragraph(p)
    pf = p.paragraph_format
    pf.left_indent = Cm(0.6)
    pf.right_indent = Cm(0.6)
    pf.space_before = Pt(6)
    pf.space_after = Pt(6)
    for idx, code_line in enumerate(lines):
        run = p.add_run(code_line if code_line else " ")
        run.font.name = "Consolas"
        run.font.size = Pt(9)
        if idx < len(lines) - 1:
            run.add_break()


# Label enumerator di awal baris: "1.", "2)", "a.", "b)", "(1)", "iv)".
# Dipakai untuk menggeser paragraf berikutnya agar rata dengan TEKS label,
# bukan dengan angkanya -- hasil yang sama dengan yang you'd dapat di Word
# saat mengetik daftar manual. Nomor yang diketik sebagai teks literal (bukan
# fitur List Number) tidak pernah membawa indentasi gantung ke paragraf
# berikutnya, sehingga jawaban mulai lurus di bawah angka.
_LABEL_RE = re.compile(r"^\s*(\(?\d{1,2}[.)]|[a-zA-Z][.)]|\d{1,2}\s*[-–])\s+")


def _label_indent_pt(prefix: str, font_pt: float) -> float:
    """Lebar indentasi (pt) untuk menyamakan baris dengan teks sebuah label.

    `prefix` harus menyertakan spasi setelah label, karena yang ingin disejajarkan
    adalah teks SETELAH "1. " -- bukan karakter terakhir labelnya.

    Perkiraan 0.5em per karakter. Untuk Times New Roman 12pt, "1. " (3 karakter)
    jadi 18pt = 0.25 inci -- sama persis dengan `w:ind w:left="360"
    w:hanging="360"` yang dipakai definisi List Number di template, jadi daftar
    bertingkat tidak meleset dari contoh.
    """
    return 0.5 * font_pt * len(prefix)


# Label pembuka baris daftar: "- ", "* ", "+ ", "1. ", "2) ".
# Berbeda dari `_LABEL_RE`, ini benar-benar menandai item daftar, bukan label
# paragraf biasa, dan dipakai sebelum pemecahan `$...$`.
_LABEL_DAFTAR_RE = re.compile(r"^\s*(?:[-*+]\s+|\d{1,2}[.)]\s+)")


def _tambah_inline(paragraph, teks: str) -> None:
    """Isi satu paragraf dengan teks yang berselang-seling `$...$` dan biasa."""
    for idx, bagian in enumerate(re.split(r"\$([^$]+)\$", teks)):
        if not bagian:
            continue
        if idx % 2:
            _append_equation(paragraph, bagian)
        else:
            _add_runs(paragraph, bagian)


def _butir_bernomor(doc: Document, num: str, daftar_num_id: str | None):
    """Buat paragraf butir bernomor asli Word.

    Return `(paragraf, num_id_baru)` kalau penomoran Word berhasil dipasang,
    atau `(paragraf, None)` kalau dokumen tanpa definisi penomoran -- pada
    kasus itu penomorannya diketik sebagai teks biasa supaya butirnya tetap
    terbaca dan bernomor.
    """
    try:
        p = doc.add_paragraph(style="List Number")
    except KeyError:
        p = doc.add_paragraph()
        p.add_run(f"{num}. ")
        return p, None

    if daftar_num_id is None:
        daftar_num_id = _numbering_restart(doc.part.numbering_part.element, int(num))
    if daftar_num_id:
        _pasang_nomor(p, daftar_num_id)
    return p, daftar_num_id


def _render_markdown(doc: Document, md: str):
    lines = md.splitlines()
    i = 0
    ordered_idx = 0
    # Label baris terakhir (mis. "1. Soal Satu") + lebar indentasi yang harus
    # dipakai paragraf-paragraf berikutnya sampai label/heading berikutnya.
    pending_label_pt: float | None = None
    body_font_pt = 12.0
    try:
        size = doc.styles["Normal"].font.size
        if size is not None:
            body_font_pt = size.pt
    except KeyError:
        pass
    # Setelah heading Daftar Pustaka, butir bernomor adalah referensi, bukan
    # butir jawaban -- dan harus dirender berbeda (lihat blok ordered list).
    in_references = False
    # `w:numId` daftar bernomor yang sedang berjalan. Satu daftar = satu
    # `w:num`; di-NULL-kan setiap kali daftar selesai (lihat reset di bawah),
    # supaya daftar berikutnya dimulai dari angka yang ditulis, bukan dari
    # kelanjutan butir sebelumnya.
    daftar_num_id: str | None = None
    while i < len(lines):
        line = lines[i].rstrip()

        # Fenced code block (```sql ... ```): render apa adanya, tanpa parsing
        # markdown/math di dalamnya.
        fence_m = _CODE_FENCE_RE.match(line)
        if fence_m:
            fence = fence_m.group(1)[0] * 3
            code_lines: list[str] = []
            i += 1
            while i < len(lines):
                if lines[i].strip().startswith(fence):
                    i += 1
                    break
                code_lines.append(lines[i])
                i += 1
            _add_code_block(doc, "\n".join(code_lines))
            ordered_idx = 0
            daftar_num_id = None
            continue

        # Inline math: convert one-line $...$ expressions to Word equations.
        #
        # Label daftar TETAP dibaca lebih dulu. Versi lama langsung memecah
        # baris jadi bagian `$...$` dan teks, sehingga baris seperti
        # `- **Pivot kedua**: bagi baris kedua dengan $\frac{5}{2}$, lalu ...`
        # kehilangan tanda `-`-nya dan tercetak sebagai
        # "- Pivot kedua: ...". Persis itu yang dikeluhkan tutor: bullet ikut
        # tampil mentah di depan paragraf.
        if line.count("$") >= 2 and "$$" not in line:
            label_m = _LABEL_DAFTAR_RE.match(line)
            if label_m and label_m.group(0).lstrip()[:1] in "-*+":
                p = doc.add_paragraph(style="List Bullet")
                _tambah_inline(p, line[label_m.end():])
                if pending_label_pt:
                    p.paragraph_format.left_indent = Cm(0) + Pt(pending_label_pt)
                ordered_idx = 0
                daftar_num_id = None
            elif label_m:
                num = label_m.group(0).strip().rstrip(".)")
                p = _butir_bernomor(doc, num, daftar_num_id)
                if isinstance(p, tuple):
                    p, daftar_num_id = p
                _tambah_inline(p, line[label_m.end():])
                if pending_label_pt:
                    p.paragraph_format.left_indent = Cm(0) + Pt(pending_label_pt)
                pending_label_pt = _label_indent_pt(f"{num}. ", body_font_pt)
                ordered_idx = 1
            else:
                p = doc.add_paragraph()
                _tambah_inline(p, line)
                if pending_label_pt:
                    p.paragraph_format.left_indent = Cm(0) + Pt(pending_label_pt)
                ordered_idx = 0
                daftar_num_id = None
            i += 1
            continue

        # Tabel markdown
        if _is_table_row(line) and i + 1 < len(lines) and re.match(
            r"^\s*\|[\s:|-]+\|\s*$", lines[i + 1]
        ):
            header_cells = [c.strip() for c in line.strip().strip("|").split("|")]
            rows = []
            j = i + 2
            while j < len(lines) and _is_table_row(lines[j]):
                rows.append([c.strip() for c in lines[j].strip().strip("|").split("|")])
                j += 1
            table = doc.add_table(rows=1, cols=len(header_cells))
            table.style = "Table Grid"
            for col, cell in enumerate(header_cells):
                table.rows[0].cells[col].text = cell
                table.rows[0].cells[col].paragraphs[0].runs[0].bold = True
            for row_data in rows:
                cells = table.add_row().cells
                for col in range(len(header_cells)):
                    if col < len(row_data):
                        cells[col].text = row_data[col]
                    else:
                        cells[col].text = ""
            doc.add_paragraph()
            i = j
            continue

        # Equation Word (OMML): blok $$...$$, boleh satu baris atau banyak baris.
        #
        # Bentuk banyak baris WAJIB didukung karena itu yang ditulis model:
        # `$$` sendiri di satu baris, isi di baris-baris berikutnya, `$$` lagi
        # di akhir. Versi lama memproses tiap baris terpisah, sehingga baris
        # penutup `$$` menghasilkan ruas kosong -- dan ruas kosong itulah yang
        # jadi `<m:oMath>` kosong di dokumen.
        if "$$" in line:
            # Kumpulkan sampai jumlah `$$` genap. Yang dihitung TOTAL, bukan
            # per baris: baris `a < b$$` hanya punya satu, jadi kalau yang
            # diperiksa cuma baris terakhir, blok berikutnya ikut tersedot dan
            # seluruh sisa dokumen berakhir jadi satu paragraf.
            blok = [line]
            j = i
            n_dolar = line.count("$$")
            while n_dolar < 2 and j + 1 < len(lines):
                j += 1
                blok.append(lines[j])
                n_dolar += lines[j].count("$$")
            # Tanpa pasangan: jangan menelan sisa dokumen. Cukup satu baris.
            if n_dolar < 2:
                blok = [line]
                j = i
            # `i = j + 1`, bukan `j`: tanpa itu baris penutup `$$` diproses
            # lagi sebagai blok baru, dan karena penutupnya tidak punya pasangan
            # lagi, blok itu terus berjalan sampai habis dan seluruh sisa
            # dokumen hilang.
            i = j + 1

            m = re.match(r"^([-*])\s+", blok[0])
            style = "List Bullet" if m else None
            p = doc.add_paragraph(style=style) if style else doc.add_paragraph()

            # Baris-baris digabung dulu supaya `\begin{matrix} ... \end{matrix}`
            # yang membentang beberapa baris tidak terpotong di batas baris.
            # PENTING: `$$` dipakai sebagai PEMBATAS, bukan dihapus -- kalau
            # delimiter-nya dibuang lebih dulu, seluruh isi blok jadi teks biasa
            # dan persis tidak terjadi render sebagai Word equation.
            gabungan = " ".join(blok)
            bagian = gabungan.split("$$")
            teks_di_luar = False
            for idx, seg in enumerate(bagian):
                if idx % 2:
                    _append_equation(p, seg)
                elif seg.strip():
                    teks_di_luar = True
                    _add_runs(p, seg)
            if not teks_di_luar:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            ordered_idx = 0
            daftar_num_id = None
            continue

        # Equation LaTeX telanjang (tanpa $): baris seperti
        # `YmidX;∼;Normal(\beta_0 + \beta_1 x_1,\; \sigma^2)`.
        if (
            re.match(r"^[-*>#|\d]", line.strip()) is None
            and _line_is_display_math(line)
        ):
            p = doc.add_paragraph()
            if len(line.strip()) < 60:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _append_equation(p, line.strip())
            i += 1
            continue

        # Heading -> style Heading 2/3 milik template, bukan ukuran font manual.
        # Menyetel size manual di sini membuat definisi template (ukuran, warna,
        # jarak) tidak berlaku dan hasil keluaran terlihat berbeda dari contoh.
        m = re.match(r"^(#{1,4})\s+(.*)$", line)
        if m:
            level = len(m.group(1))
            heading_text = m.group(2).strip()
            if heading_text.lower().startswith("daftar pustaka"):
                in_references = True
                p = doc.add_paragraph()
                run = p.add_run(heading_text)
                run.bold = True
                # "Daftar Pustaka" di template BUKAN heading: `Normal` bold
                # TNR 14, sama seperti penanda bagian lain. Memakainya sebagai
                # Heading 2 membuatnya tampil bergaya heading -- beda dari
                # contoh, dan ini bagian yang paling dilihat tutor.
                try:
                    p.style = doc.styles["Normal"]
                except KeyError:
                    pass
                # Bold hitam Times New Roman, ukuran dari template. Yang lama
                # hanya `bold` + ukuran, jadi font dan warna ikut bawaan style
                # dan tampilan "Daftar Pustaka" tidak sama dengan judul lain.
                _rapikan_heading(run, doc, "Normal", pt=_REFERENSI_JUDUK_PT)
                daftar_num_id = None
                i += 1
                continue
            style_name = _heading_style(doc, level)
            p = doc.add_paragraph()
            run = p.add_run(heading_text)
            if style_name != "Normal":
                p.style = doc.styles[style_name]
                # Ukuran diambil dari style template (bukan ditetapkan manual),
                # jadi mengubah ukuran di template tetap ikut terbawa.
                _rapikan_heading(run, doc, style_name)
            else:
                # Tanpa style Heading (dokumen tanpa template), tetap beri
                # pembedaan visual sesuai tingkatnya.
                bawaan = 14 if level == 2 else 12
                _rapikan_heading(run, doc, "Normal", pt=bawaan)
            # Heading berlabel ("### 1. Soal Satu"): judulnya boleh di margin,
            # tapi isi di bawahnya harus rata dengan teks setelah label, bukan
            # dengan angkanya. Catat lebarnya untuk paragraf berikutnya.
            lm = _LABEL_RE.match(heading_text)
            pending_label_pt = _label_indent_pt(lm.group(0), body_font_pt) if lm else None
            daftar_num_id = None
            i += 1
            continue

        # Daftar Pustaka: APA 7 murni. Tidak ada nomor di depan dan tidak ada bullet
        # -- penandanya adalah urutan alfabetis. Nomor otomatis Word justru
        # merusak di sini: ia menghitung butir jawaban yang mendahuluinya, jadi
        # referensi bisa mulai dari angka 4 dan angka yang disebut di dalam
        # teks tidak cocok. Nomor literal yang diketik manual juga dibuang,
        # karena yang diminta tutor adalah Daftar Pustaka gaya APA 7 yang bersih.
        #
        # Setiap entri juga dibersihkan dan diformat ulang di sini: penulis
        # dicetak tebal, judul miring, dan alamat Reader dibuang. Lihat
        # `_format_entri_referensi`.
        if in_references:
            teks = _format_entri_referensi(line)
            if not teks:
                i += 1
                continue
            p = doc.add_paragraph()
            _add_runs(p, teks)
            i += 1
            pending_label_pt = None
            continue

        # Bullet list
        if re.match(r"^[-*]\s+", line):
            p = doc.add_paragraph(style="List Bullet")
            _add_runs(p, re.sub(r"^[-*]\s+", "", line))
            if pending_label_pt:
                p.paragraph_format.left_indent = Cm(0) + Pt(pending_label_pt)
            i += 1
            ordered_idx = 0
            daftar_num_id = None
            continue

        # Ordered list -> penomoran asli Word ("List Number"), bukan angka yang
        # diketik manual. Versi lama menulis "1. " sebagai run biasa, jadi
        # penomoran tidak pernah menyesuaikan saat butir disisipkan di tengah,
        # dan style-nya tidak sama dengan template.
        #
        # Satu daftar = satu `w:num`, dipasang di semua paragraf butirnya.
        # Word menghitung per `w:num`, jadi `w:num` baru untuk tiap butir
        # membuat setiap butir tampil sebagai "1."; Sebaliknya, satu `w:num`
        # bersama membuat semua daftar di dokumen berbagi hitungan, sehingga
        # butir "1." di sesi kedua bisa tampil sebagai "4."
        #
        # Isi Daftar Pustaka tidak pernah sampai ke sini: blok `in_references`
        # di atas sudah mengambilnya lebih dulu.
        mo = re.match(r"^(\d+)[.)]\s+(.*)$", line)
        if mo:
            num = mo.group(1).lstrip("0") or "0"
            try:
                p = doc.add_paragraph(style="List Number")
            except KeyError:
                p = doc.add_paragraph()
                p.add_run(f"{num}. ")
                _add_runs(p, mo.group(2))
            else:
                if daftar_num_id is None:
                    daftar_num_id = _numbering_restart(
                        doc.part.numbering_part.element, int(num)
                    )
                if daftar_num_id:
                    _pasang_nomor(p, daftar_num_id)
                _add_runs(p, mo.group(2))
            if pending_label_pt:
                p.paragraph_format.left_indent = Cm(0) + Pt(pending_label_pt)
            # Butir bernomor itu sendiri jadi label untuk paragraf setelahnya.
            pending_label_pt = _label_indent_pt(f"{num}. ", body_font_pt)
            ordered_idx = 1
            i += 1
            continue

        ordered_idx = 0
        daftar_num_id = None
        # Empty line → skip; blockquote → plain paragraph
        if not line.strip():
            i += 1
            continue
        # Horizontal rule --- / *** → spacer paragraph
        if re.match(r"^[-*=_]{3,}\s*$", line.strip()):
            doc.add_paragraph()
            i += 1
            continue
        p = doc.add_paragraph()
        _add_runs(p, line.strip())
        # Paragraf ini bisa jadi label baru ("a. Poin ..." / "(1) Butir ...").
        # Kalau iya, baris ini sendiri boleh di margin (atau di indent label
        # sebelumnya), dan paragraf SESUDAHNYA yang harus rata dengan teksnya.
        lm = _LABEL_RE.match(line.strip())
        if lm:
            p.paragraph_format.left_indent = (
                Cm(0) + Pt(pending_label_pt) if pending_label_pt else None
            )
            pending_label_pt = _label_indent_pt(lm.group(0), body_font_pt)
        elif pending_label_pt:
            p.paragraph_format.left_indent = Cm(0) + Pt(pending_label_pt)
        i += 1


def _clean_soal(md: str) -> str:
    """Buang metadata lampiran/transkripsi dari soal, pertahankan isi transkripsi."""
    lines = md.splitlines()
    out: list[str] = []
    skip_attachments = False
    for line in lines:
        stripped = line.strip()
        if re.match(r"^##\s+Lampiran\s*:?", stripped):
            skip_attachments = True
            continue
        if re.match(r"^##\s+Isi\s+lampiran\b", stripped, re.IGNORECASE):
            continue
        if stripped == "# Transkrip Soal":
            continue
        if skip_attachments:
            if re.match(r"^[-*]\s+", stripped) or not stripped:
                continue
            skip_attachments = False
        out.append(line)
    return "\n".join(out)


def _load_base_document(template_path: Path | None):
    """Buka template sebagai dokumen dasar.

    Template adalah sumber kebenaran format. Dengan memuat filenya, kita mewarisi
    seluruh definisi yang tidak kita Listing-kan satu per satu: page setup
    (ukuran kertas, margin), theme, numbering.xml (penomoran butir asli Word),
    dan seluruh definisi style. Kalau nanti template diperbarui, hasil keluaran
    ikut berubah tanpa menyentuh kode.

    Kandungan teksnya dibuang; yang dipertahankan hanya kerangka. Kalau template
    hilang atau rusak, jatuh ke dokumen kosong dengan tipografi bawaan supaya
    pipeline tidak berhenti total hanya karena satu file.

    Return (doc, dari_template, font_nama, font_ukuran).
    """
    if template_path and Path(template_path).is_file():
        try:
            doc = Document(str(template_path))
            # PENTING: tipografi harus dibaca SEBELUM body dikosongkan. Setelah
            # dihapus tidak ada satu pun run tersisa untuk disampel, sehingga
            # font isi template tak akan pernah terdeteksi dan `Normal` akan
            # kembali ke default-nya sendiri (Arial, bukan Times New Roman).
            font_name, font_size = _template_body_font(doc)
            if not _siapkan_lembar(doc):
                # Bukan template format jawaban (tidak ada `Heading 2` sebagai
                # penanda contoh). Kosongkan total supaya isinya tidak bocor ke
                # dokumen hasil.
                _clear_body(doc)
            return doc, True, font_name, font_size
        except Exception as exc:  # noqa: BLE001 - template tidak boleh mematikan pipeline
            print(f"  ! template tidak bisa dibaca ({Path(template_path).name}): {exc}")
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)
    rpr = style.element.get_or_add_rPr()
    rf = rpr.get_or_add_rFonts()
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        q = qn(attr)
        if q in rf.attrib:
            del rf.attrib[q]
    rf.set(qn("w:eastAsia"), "Times New Roman")
    return doc, False, "Times New Roman", Pt(12)


def _numbering_restart(numbering, mulai: int) -> str | None:
    """Buat `w:num` yang menomori mulai dari `mulai`. Return `w:numId`.

    Word menyimpan penomoran di `numbering.xml`, bukan di paragraf. Style
    `List Number` hanya memberi arahan "pakai daftar nomor ini"; angka yang
    muncul di layar berasal dari urutan kemunculan `w:num` itu di seluruh
    dokumen. Akibatnya daftar kedua melanjutkan hitungan daftar pertama -- butir
    yang ditulis "1." tampil sebagai "4.".

    Solusinya bukan mengetik angka manual, karena itu mengembalikan masalah
    lama: angka tidak ikut menyesuaikan saat butir disisipkan di tengah. Yang
    benar adalah satu `w:num` per daftar, dengan `w:startOverride` berisi
    angka yang diminta. Format daftar -- jenis angka, indentasi gantung, font --
    tetap warisan template; hanya titik mulainya yang dipatah.

    Panggil sekali per daftar, lalu pasang `w:numId` hasilnya di SETIAP
    paragraf butir. Memanggil ulang untuk butir kedua memecah daftar jadi dua
    `w:num`, dan Word menghitung per `w:num` -- hasilnya butir kedua tampil
    sebagai "1." lagi.
    """
    abstract = _cari_abstract_desimal(numbering)
    if abstract is None:
        return None
    return _buat_num_dengan_start(numbering, abstract, mulai)


def _pasang_nomor(paragraf, num_id: str) -> None:
    """Arahkan satu paragraf ke `w:num` tertentu.

    `w:numPr` disisipkan tepat setelah `w:pStyle`, bukan di posisi nol. Urutan
    anak `w:pPr` ditentukan skema OOXML, dan `numPr` yang muncul sebelum
    `pStyle` membuat Word menolak membuka berkas.
    """
    pPr = paragraf._p.get_or_add_pPr()
    lama = pPr.find(qn("w:numPr"))
    if lama is not None:
        pPr.remove(lama)
    numPr = OxmlElement("w:numPr")
    numId_el = OxmlElement("w:numId")
    numId_el.set(qn("w:val"), num_id)
    numPr.append(numId_el)
    gaya = pPr.find(qn("w:pStyle"))
    if gaya is not None:
        gaya.addnext(numPr)
    else:
        pPr.insert(0, numPr)



def _cari_abstract_desimal(numbering) -> str | None:
    """`abstractNumId` pertama yang formatnya angka desimal, atau None."""
    for abstract in numbering.findall(qn("w:abstractNum")):
        lvl = abstract.find(qn("w:lvl"))
        if lvl is None:
            continue
        fmt = lvl.find(qn("w:numFmt"))
        if fmt is not None and fmt.get(qn("w:val")) == "decimal":
            return abstract.get(qn("w:abstractNumId"))
    return None


def _buat_num_dengan_start(numbering, abstract_id: str, mulai: int) -> str:
    """Buat `w:num` baru yang menunjuk `abstract_id` dengan `startOverride`."""
    dipakai = [
        int(n.get(qn("w:numId")) or 0)
        for n in numbering.findall(qn("w:num"))
        if (n.get(qn("w:numId")) or "").isdigit()
    ]
    num_id = str(max(dipakai or [0]) + 1)

    num_el = OxmlElement("w:num")
    num_el.set(qn("w:numId"), num_id)
    ref = OxmlElement("w:abstractNumId")
    ref.set(qn("w:val"), abstract_id)
    num_el.append(ref)
    override = OxmlElement("w:lvlOverride")
    override.set(qn("w:ilvl"), "0")
    start = OxmlElement("w:startOverride")
    start.set(qn("w:val"), str(mulai))
    override.append(start)
    num_el.append(override)

    # `w:num` harus ditulis setelah semua `w:abstractNum`; skema numbering.xml
    # menolaknya kalau muncul di tengah.
    terakhir_abstract = numbering.findall(qn("w:abstractNum"))
    if terakhir_abstract:
        terakhir_abstract[-1].addnext(num_el)
    else:
        numbering.insert(0, num_el)
    return num_id


def _clear_body(doc: Document) -> None:
    """Kosongkan isi dokumen, pertahankan sectPr (page setup) di akhir body.

    `sectPr` tidak boleh dihapus: itu yang membawa ukuran kertas, margin, dan
    jarak header/footer. Tanpa itu, Letter/margin template hilang dan hasilnya
    kembali ke default Word.
    """
    body = doc.element.body
    for child in list(body):
        if child.tag == qn("w:sectPr"):
            continue
        body.remove(child)


def _siapkan_lembar(doc: Document) -> bool:
    """Bersihkan template jadi lembar jawaban kosong. True kalau berhasil.

    Template `ContohFormatJawaban.docx` bukan sekadar kerangka format; isinya
    adalah satu contoh jawaban yang sudah jadi lengkap dengan Daftar Pustaka.
    Versi lama membuang seluruh isi (`_clear_body`) lalu membangun ulang dari
    nol, sehingga detail yang paling menentukan tampilan -- tinggi baris, jarak
    antarparagraf, ukuran heading, definisi tabel -- ikut hilang, dan hasilnya
    cuma "mirip" contoh.

    Yang benar adalah mengisi lembar itu: simpan judul di paling atas dan tabel
    identitas apa adanya, lalu buang sisanya -- rubrik soal, contoh jawaban, dan
    Daftar Pustaka contoh -- supaya yang ditulis pipeline masuk persis di tempat
    yang seharusnya. Titik potongnya adalah `Heading 2` pertama, karena di situ
    contoh sesi mulai.

    `sectPr` di akhir body tetap utuh. Return False kalau tidak ada `Heading 2`
    sama sekali (template lain, bukan format jawaban); pemanggil lalu jatuh ke
    `_clear_body` supaya tetap menghasilkan dokumen yang bisa dibuka.
    """
    body = doc.element.body
    potong = None
    for child in body:
        if child.tag != qn("w:p"):
            continue
        style = child.find(qn("w:pPr") + "/" + qn("w:pStyle"))
        nilai = style.get(qn("w:val")) if style is not None else None
        if nilai and nilai.lower().replace(" ", "") == "heading2":
            potong = child
            break
    if potong is None:
        return False

    # Sisanya dibuang: dari `Heading 2` pertama sampai sebelum `sectPr`.
    # Judul dan tabel identitas berada SEBELUM titik potong, jadi tetap utuh.
    lewat = False
    for child in list(body):
        if child is potong:
            lewat = True
        if lewat and child.tag != qn("w:sectPr"):
            body.remove(child)
    return True


def _template_body_font(doc: Document) -> tuple[str, Pt | None]:
    """Ambil font yang BENAR-BENAR dipakai isi template.

    Penting karena template bisa tidak konsisten: style `Normal` menunjuk satu
    font, tapi tiap run di dalamnya membawa override sendiri. Kalau kita cuma
    percaya style `Normal`, paragraf yang kita tambahkan akan tampil berbeda dari
    paragraf contoh -- persis kesalahan "tidak sama persis" yang harus dihindari.
    Jadi yang dicari adalah run yang paling sering muncul, bukan default style.
    """
    counts: dict[tuple[str, int], int] = {}
    for para in doc.paragraphs:
        for run in para.runs:
            name = run.font.name
            if not name:
                continue
            size = int(run.font.size.pt) if run.font.size is not None else 0
            counts[(name, size)] = counts.get((name, size), 0) + len(run.text.strip()) or 1
    if not counts:
        return "", None
    (name, size), _ = max(counts.items(), key=lambda kv: kv[1])
    return name, Pt(size) if size else None


def _apply_base_typography(doc: Document, font_name: str, font_size: Pt | None) -> None:
    """Samakan style paragraf dasar dengan font isi template."""
    if not font_name:
        return
    for style_name in ("Normal", "List Paragraph", "Body Text"):
        try:
            style = doc.styles[style_name]
        except KeyError:
            continue
        style.font.name = font_name
        if font_size is not None:
            style.font.size = font_size
        rpr = style.element.get_or_add_rPr()
        rf = rpr.get_or_add_rFonts()
        for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
            q = qn(attr)
            if q in rf.attrib:
                del rf.attrib[q]
        rf.set(qn("w:eastAsia"), font_name)


_IDENTITAS_LABELS = {
    "nama": "nama",
    "nim": "nim",
    "semester": "semester",
    "ut daerah": "ut_daerah",
    "program studi": "prodi",
}


def _ganti_teks_paragraf(paragraf, teks: str) -> None:
    """Tulis ulang teks paragraf tanpa kehilangan format run pertamanya.

    Format itu dibawa oleh run, bukan paragraf: template memberi judul
    `Times New Roman` 18pt bold lewat run-nya. `paragraf.text = ...` akan
    membuat python-docx menebus seluruh run dan menaruh format di paragraf,
    sehingga tampilan berubah -- dan perubahan sekecil apa pun di template
    akan hilang.
    """
    if not paragraf.runs:
        paragraf.add_run(teks)
        return
    paragraf.runs[0].text = teks
    for sisa in paragraf.runs[1:]:
        sisa.text = ""


def _judul_dokumen(meta: dict) -> str:
    """`Diskusi 4 - Aljabar Linear Elementer 95 (STMA4113)`.

    Nama matkul dan kodenya ikut karena berkas dari beberapa mata kuliah dikirim
    ke tutor yang sama. Tanpa nama dan kode, tumpukan berkas itu hanya bisa
    dicocokkan lewat isi -- dan pengajar tidak mungkin membaca isi setiap berkas.
    Kode diambil dari isi lampiran karena nama course di Moodle tidak memuatnya.
    """
    bagian = f"{meta.get('kind_label', '')} {meta.get('display_index', '')}".strip()
    matkul = str(meta.get("matkul") or "").strip()
    kode = str(meta.get("matkul_kode") or "").strip()
    if matkul:
        bagian = f"{bagian} - {matkul}".strip(" -")
    if kode:
        bagian = f"{bagian} ({kode})"
    return bagian


def _isi_judul(doc: Document, meta: dict) -> bool:
    """Isi paragraf judul yang sudah ada di template. True kalau ketemu.

    Ukuran font TIDAK ditulis ulang di sini. Template yang menentukan
    Ukuran, jadi menyalin `Pt(18)` ke kode hanya mengunci nilai sekarang:
    begitu tutor menggeser judul jadi 16pt di template, hasil pipeline tetap
    18pt -- persis ketidakcocokan yang harus dihindari. Yang ditulis hanya
    teksnya; format run template dibiarkan apa adanya.
    """
    judul = _judul_dokumen(meta)
    for paragraf in doc.paragraphs:
        if paragraf.text.strip():
            if judul:
                _ganti_teks_paragraf(paragraf, judul)
            return True
    return False


def _isi_tabel_identitas(doc: Document, meta: dict) -> bool:
    """Isi tabel identitas yang sudah ada di template.

    Dua aturan, dan keduanya soal apa yang terjadi pada baris yang `meta` tidak
    punya nilainya.

    Baris dengan nilai dari `meta` ditulis ulang: `.env` adalah sumber data
    yang dipakai pengguna, jadi dia yang menang atas isi template.

    Baris tanpa nilai dari `meta` DIHAPUS, bukan dibiarkan. Alasannya bukan
    soal tampilan: template `ContohFormatJawaban.docx` diisi dengan identitas
    mahasiswa yang menjadi contoh -- nama dan NIM orang sungguhan. Membiarkannya
    berarti berkas yang diserahkan ke tutor memuat identitas orang lain, dan
    dan itu baru ketahuan setelah berkas diserahkan. Baris yang dihapus
    juga mencegah label menggantung seperti "Semester |" yang terlihat seperti
    tabel gagal terisi.
    """
    if not doc.tables:
        return False
    table = doc.tables[0]

    # Template `ContohFormatJawaban.docx` hanya punya baris Nama, NIM,
    # Semester, dan UT Daerah. `PRODI` tetap ada di `.env` tapi tidak punya
    # baris untuk ditulis, jadi nilainya diam-diam hilang dari dokumen.
    # Baris disalin dari baris identitas terakhir supaya format, lebar kolom,
    # dan gaya paragrafnya sama persis dengan baris lain, lalu ditambahkan di
    # bawah. Menyalin elemen `w:tr` jauh lebih aman daripada membuat baris baru:
    # `add_row()` menghasilkan tabel tanpa style dan tanpa format run, dan itu
    # langsung kelihatan beda.
    _LABEL_TAMPIL = {
    "nama": "Nama",
    "nim": "NIM",
    "semester": "Semester",
    "ut daerah": "UT Daerah",
    "program studi": "Program Studi",
}

    # `_IDENTITAS_LABELS` dipetakan dari teks label ke kunci `meta`, jadi
    # urutannya label dulu, baru kunci. Teks label yang ditulis ke dokumen
    # memakai kapitalisasi seperti baris template ("Nama", "UT Daerah"), bukan
    # kunci huruf kecil yang dipakai untuk pencarian.
    label_tambahan = [
        (_LABEL_TAMPIL[label], kunci)
        for label, kunci in _IDENTITAS_LABELS.items()
        if kunci not in _label_ada_di_tabel(table)
        and str(meta.get(kunci) or "").strip()
    ]

    for row in list(table.rows):
        sel = row.cells
        if len(sel) < 2:
            continue
        kunci = _IDENTITAS_LABELS.get(sel[0].text.strip().lower().rstrip(":").strip())
        if not kunci:
            # Label yang tidak dikenali (mis. kolom tambahan milik pengguna)
            # dibiarkan: bukan urusan pipeline.
            continue
        nilai = str(meta.get(kunci) or "").strip()
        if not nilai:
            table._tbl.remove(row._tr)
            continue
        _ganti_teks_paragraf(sel[1].paragraphs[0], nilai)

    for label, kunci in label_tambahan:
        _tambah_baris_identitas(table, label, str(meta[kunci]).strip())
    return True


def _label_ada_di_tabel(table) -> set[str]:
    """Kunci identitas yang barisnya sudah ada di tabel."""
    ada: set[str] = set()
    for row in table.rows:
        sel = row.cells
        if len(sel) < 2:
            continue
        kunci = _IDENTITAS_LABELS.get(
            sel[0].text.strip().lower().rstrip(":").strip()
        )
        if kunci:
            ada.add(kunci)
    return ada


def _tambah_baris_identitas(table, label: str, nilai: str) -> None:
    """Salin baris terakhir tabel identitas lalu isi dengan label dan nilai.

    Baris disalin, bukan dibuat baru, supaya hasilnya tidak berbeda dari baris
    lain: template memberi format run lewat style tabel dan penataan paragraf,
    dan `add_row()` tidak membawa salah satunya.
    """
    baris = list(table.rows)
    if not baris:
        return

    # `deepcopy` sekali, dipakai dua kali: yang disalin keluar adalah sel
    # asli supaya format run-nya ikut, bukan sel kosong.
    tiru = copy.deepcopy(baris[-1]._tr)
    table._tbl.append(tiru)

    # Isi ulang lewat sel yang sudah diformat, bukan lewat paragraf baru.
    sel_baru = table.rows[-1].cells
    _ganti_teks_paragraf(sel_baru[0].paragraphs[0], label)
    _ganti_teks_paragraf(sel_baru[1].paragraphs[0], nilai)


def _identitas_table(doc: Document, meta: dict) -> None:
    """Tabel identitas 2 kolom: label | nilai.

    Dipakai hanya kalau template TIDAK punya tabel identitas (dokumen kosong
    atau template lain). Kalau template punya, tabel itu yang dipakai --
    lihat `_isi_tabel_identitas`.
    """
    rows = [
        ("Nama", meta.get("nama", "")),
        ("NIM", meta.get("nim", "")),
        ("Semester", meta.get("semester", "")),
        ("UT Daerah", meta.get("ut_daerah", "")),
    ]
    filled = [(label, str(val).strip()) for label, val in rows if str(val or "").strip()]
    if not filled:
        return
    table = doc.add_table(rows=0, cols=2)
    try:
        table.style = "Table Grid"
    except KeyError:
        pass
    for label, value in filled:
        cells = table.add_row().cells
        label_run = cells[0].paragraphs[0].add_run(label)
        label_run.bold = True
        cells[1].paragraphs[0].add_run(value)
    doc.add_paragraph()


def build_docx(
    *,
    jawaban_md: str,
    soal_text: str,
    meta: dict,
    out_docx: Path,
    include_soal: bool = True,
    template: Path | None = None,
    gambar_soal: list[Path] | None = None,
) -> Path:
    # Sanitasi di pintu masuk: meta, soal, dan jawaban bisa memuat karakter
    # kontrol dari scraping Moodle. Run header/judul tidak lewat _normalize_text
    # sehingga harus dibersihkan di sini juga.
    soal_text = _strip_control_chars(soal_text)
    jawaban_md = _strip_control_chars(jawaban_md)
    meta = {k: _strip_control_chars(str(v)) for k, v in meta.items()}

    if template is None:
        template = _default_template_path()
    doc, from_template, font_name, font_size = _load_base_document(
        Path(template) if template else None
    )
    if from_template:
        _apply_base_typography(doc, font_name, font_size)

    # Judul dan tabel identitas: template sudah menyediakan keduanya, jadi
    # isinya ditulis ke tempat itu. Baris `if not ...` hanya berlaku untuk
    # dokumen tanpa template, yang memang tidak punya apa pun untuk diisi.
    if not _isi_judul(doc, meta):
        # TNR 18 bold center, mengikuti template. Ukuran diambil dari style
        # `Title` kalau ada supaya perubahan di template ikut terbawa.
        title = _judul_dokumen(meta)
        tp = doc.add_paragraph()
        tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        tr = tp.add_run(title)
        tr.bold = True
        tr.font.size = Pt(18)

    if not _isi_tabel_identitas(doc, meta):
        _identitas_table(doc, meta)

    # Soal ditulis ulang di depan jawaban. Alasannya bukan estetika: berkas
    # yang dikumpulkan tutor harus bisa diperiksa utuh, dan tanpa soal di
    # dalamnya pengajar tidak bisa menilai apakah jawaban menjawab yang
    # ditanyakan.
    #
    # Dua bentuk, karena bentuk soal berbeda antara mata kuliah:
    #   - gambar (screenshot): DISISIPKAN apa adanya. Mengetik ulang dari
    #     transkripsi berisiko salah -- dan lebihCommerceStill, lampiran itu
    #     adalah naskah resmi yang harus dinilai persis seperti aslinya.
    #   - teks: ditempel apa adanya setelah metadata Moodle dibuang.
    if include_soal:
        _tulis_soal(doc, soal_text, gambar_soal=gambar_soal)

    # Jawab (dari markdown opencode)
    body = _bersihkan_tanda_humanizer(_drop_identity_echo(_answer_body(jawaban_md)))
    if body.strip():
        _render_markdown(doc, body)

    # Daftar Pustaka dirapikan setelah seluruh isi masuk: urutan alfabetis dan
    # indentasi gantung baru bisa dipastikan setelah paragraf terakhir ada.
    _urutkan_referensi_apa(doc)
    _style_referensi_apa(doc)
    _tandai_placeholder_referensi(doc)
    _neutralize_metadata(doc, meta)

    out_docx.parent.mkdir(parents=True, exist_ok=True)
    return _simpan_aman(doc, out_docx)


def _simpan_aman(doc: Document, out_docx: Path) -> Path:
    """Simpan dokumen ke `out_docx` tanpa kehilangan hasil kalau berkas terkunci.

    `doc.save()` langsung ke berkas tujuan gagal dengan `PermissionError` kalau
    berkas itu sedang dibuka Word, sedang disinkronkan OneDrive, atau masih
    dipegang antivirus. Kejadiannya nyata di mesin ini: satu sesi selesai
    seluruh tahap, lalu semua pekerjaan hilang karena `save()` melempar
    exception di baris terakhir.

    Dua lapis:

    1. Tulis ke berkas sementara lebih dulu, lalu `os.replace` ke tujuan.
       `os.replace` bersifat atomik dan lebih tahan terhadap proses lain yang
       menahan handle sebentar.
    2. Kalau `os.replace` tetap ditolak, simpan ke nama bernomor
       (`nama-1.docx`, `nama-2.docx`, ...) sehingga jawaban tetap tersimpan
       dan pemanggil bisa memberi tahu berkas mana yang jadi.

    Path yang benar-benar ditulis dikembalikan, jadi pemanggil tidak pernah
    melaporkan nama berkas yang gagal ditulis.
    """
    sementara = out_docx.with_name(f"{out_docx.name}.sedang-dibuat")
    sementara.unlink(missing_ok=True)
    doc.save(str(sementara))

    try:
        os.replace(sementara, out_docx)
        return out_docx
    except PermissionError:
        pass

    for nomor in range(1, 100):
        cadangan = out_docx.with_name(f"{out_docx.stem}-{nomor}{out_docx.suffix}")
        if cadangan.exists():
            continue
        try:
            os.replace(sementara, cadangan)
            return cadangan
        except PermissionError:
            continue

    sementara.unlink(missing_ok=True)
    raise PermissionError(
        f"Tidak bisa menyimpan {out_docx.name}: berkas sedang dipakai program lain "
        "dan tidak ada nama alternatif yang bisa ditulis."
    )


def _default_template_path() -> Path | None:
    """Template standar; `None` kalau folder template tidak ada."""
    from config import TEMPLATE_DIR

    if not TEMPLATE_DIR.is_dir():
        return None
    docx = sorted(TEMPLATE_DIR.glob("*.docx"))
    return docx[0] if docx else None


def _heading_style(doc: Document, level: int) -> str:
    """`##` -> Heading 2, `###` -> Heading 3, dengan fallback aman.

    Template menentukan ukuran heading lewat style-nya sendiri, jadi style
    dipakai apa adanya -- hanya UKURAN yang diambil dari sana. Warna dan font
    tidak ikut, karena keduanya bawaan Word (biru aksen + Calibri Light) dan
    sama sekali tidak cocok untuk berkas yang dikumpulkan ke tutor. Yang
    memperbaiki semuanya adalah `_rapikan_heading`.
    """
    wanted = f"Heading {min(max(level, 2), 3)}"
    try:
        doc.styles[wanted]
    except KeyError:
        return "Normal"
    return wanted


# Font dan warna heading. Ditetapkan per-run, bukan lewat style, karena style
# `Heading N` bawaan Word mewarisi warna aksen dan font tema -- dokumen hasil
# akan tampil biru dan Calibri Light, dan itu langsung terbaca sebagai "dibuat
# mesin", padahal yang salah hanya style bawaannya.
_HEADING_FONT = "Times New Roman"
_HEADING_COLOR = "000000"  # hitam


def _rapikan_heading(run, doc: Document, style_name: str, *, pt: int | None = None) -> None:
    """Jadikan satu run heading: bold, hitam, Times New Roman.

    Ukuran diambil dari style template kalau ada, supaya perubahan ukuran di
    template tetap ikut terbawa. Font dan warna dipasang LANGSUNG di run, bukan
    di style: begitu tutor mengubah font lewat style, hasil pipeline tetap
    konsisten dengan yang diminta di sini.
    """
    run.bold = True
    run.font.name = _HEADING_FONT
    run.font.color.rgb = RGBColor.from_string(_HEADING_COLOR)
    # `eastAsia`/`cs` wajib diisi juga; kalau tidak, teks dengan karakter
    # non-Latin akan jatuh ke font tema dan terlihat tidak serasi.
    rpr = run._r.get_or_add_rPr()
    rf = rpr.get_or_add_rFonts()
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        q = qn(attr)
        if q in rf.attrib:
            del rf.attrib[q]
    rf.set(qn("w:ascii"), _HEADING_FONT)
    rf.set(qn("w:hAnsi"), _HEADING_FONT)
    rf.set(qn("w:eastAsia"), _HEADING_FONT)
    rf.set(qn("w:cs"), _HEADING_FONT)

    if pt is None:
        pt = _heading_pt(doc, style_name, 0)
    if pt:
        run.font.size = Pt(pt)


def _heading_pt(doc: Document, style_name: str, bawaan: int) -> int:
    """Ukuran heading dari style template, atau `bawaan` kalau tidak ada."""
    if style_name == "Normal":
        return bawaan
    try:
        ukuran = doc.styles[style_name].font.size
    except KeyError:
        return bawaan
    return int(ukuran.pt) if ukuran is not None else bawaan


def _answer_body(jawaban_md: str) -> str:
    """Ambil bagian jawaban saja, buang sisa sebelum heading pertama."""
    for marker in ("## jawaban mahasiswa", "## jawab", "## jawaban"):
        idx = jawaban_md.lower().find(marker)
        if idx != -1:
            return jawaban_md[idx:]
    return jawaban_md


# Baris identitas yang sering ditulis ulang agen di dalam "Jawaban Mahasiswa"
# (Nama:, NIM:, Prodi:, Mata kuliah:). `_identitas_table` sudah menaruh data
# yang sama di tabel kop, jadi lewatannya berduplikasi di dokumen -- dan
# telegram-style ini tidak ada di template sama sekali.
_IDENTITY_LINE_RE = re.compile(
    r"^\s*(?:[-*]\s*)?(?:\*\*)?"
    r"(?:nama|nim|prodi|program\s+studi|mata\s+kuliah|semester|ut\s+daerah|fakultas)"
    r"\s*(?:\*\*)?\s*:",
    re.IGNORECASE,
)


def _drop_identity_echo(md: str) -> str:
    """Buang blok identitas yang ditulis ulang agen.

    Baris identitas muncul TEPAT SESUDAH heading pembungkus ("## Jawaban
    Mahasiswa") dan SEBELUM sub-heading pertama ("### ..."), jadi penyaringannya
    hanya berlaku di rentang itu. Setelah ada `###`, isi jawaban dimulai dan
    tidak ada lagi yang boleh dipotong -- kalimat biasa yang kebetulan diawali
    "Mata kuliah: ..." di tengah paragraf harus utuh.
    """
    out: list[str] = []
    seen_subheading = False
    for line in md.splitlines():
        stripped = line.lstrip()
        if stripped.startswith("#"):
            # `###` ke bawah = isi jawaban benar-benar dimulai. Setelah itu
            # tidak ada lagi baris identitas yang boleh dipotong.
            if stripped.startswith("###"):
                seen_subheading = True
            out.append(line)
            continue
        if not seen_subheading and _IDENTITY_LINE_RE.match(line):
            continue
        out.append(line)
    return "\n".join(out)


def _neutralize_metadata(doc: Document, meta: dict) -> None:
    """Sembunyikan sidik jari generator di properti dokumen.

    python-docx menuliskan `dc:creator = python-docx` dan
    `dc:description = generated by python-docx` di docProps/core.xml. Kedua
    metadata itu dibaca plagiarism checker/cek kesamaan dokumen, jadi diganti
    dengan identitas mahasiswa dan tanggal tetap (bukan waktu asli pembuatan,
    yang justru membuka kapan file ini dibuat).
    """
    try:
        core = doc.core_properties
    except Exception:  # noqa: BLE001
        return
    try:
        author = str(meta.get("nama", "")).strip() or "python-docx"
        core.author = author
        core.last_modified_by = author
        core.description = ""
        core.category = ""
        core.comments = ""
        core.identifier = ""
        core.language = "id-ID"
        core.title = (
            f"{meta.get('kind_label', '')} {meta.get('display_index', '')} "
            f"{meta.get('matkul', '')}"
        ).strip()
        core.subject = str(meta.get("matkul", ""))
        core.created = _NEUTRAL_DOC_TIMESTAMP
        core.modified = _NEUTRAL_DOC_TIMESTAMP
        core.revision = 1
    except Exception:  # noqa: BLE001
        pass


def _pisah_kalimat(teks: str) -> tuple[str, str]:
    """Pisahkan `(judul). (ekor)` pada titik pertama yang bukan inisial.

    Titik setelah singkatan (`hlm.`, `pp.`, `Ed.`) bukan akhir kalimat. Kalau
    tidak dibedakan, judul prosiding ikut terpotong dan sisa berikutnya
    salah dianggap sebagai nama majalah.
    """
    for m in re.finditer(r"\.\s+", teks):
        sebelum = teks[: m.start() + 1]
        # Inisial seperti "J. R. R." dan singkatan seperti "hlm." tidak boleh
        # dipotong sebagai akhir kalimat.
        if _BUKAN_AKHIR_KALIMAT_RE.search(sebelum):
            continue
        return sebelum, teks[m.end():]
    return teks, ""


def _bagian_miring(ekor: str) -> tuple[int, str]:
    """Bagian `ekor` yang menurut APA 7 harus dicetak miring.

    Dikembalikan sebagai `(posisi, teks)` karena posisinya tidak selalu nol:
    pada entri prosiding, yang miring adalah judul prosiding, sementara
    `In R. Hidayat (Ed.),` di depannya tetap tegak. Kalau pemanggil hanya
    memotong berdasarkan panjang teks, Separuh ekor akan terpotong di tempat
    yang salah.

    Dua bentuk yang muncul di Daftar Pustaka tutorial:

    * artikel jurnal -- miring dari posisi 0 sampai nama majalah dan
      nomornya, lalu `(1), 45-59` tetap tegak;
    * prosiding -- `In A. W. Saputra (Ed.), Seminar ..., hlm. 55-59`, yang
      miring hanya judul prosidingnya, bukan nama editornya.

    Kalau tidak ada pola yang cocok, dikembalikan `(0, "")` supaya ekor tetap
    tegak. Jangan menebaknya: memiringkan satu ekor entri yang salah lebih
    buruk daripada membiarkannya tegak.
    """
    if not ekor:
        return 0, ""

    m_in = _AWAL_IN_RE.match(ekor)
    if m_in:
        # posisi dihitung terhadap `ekor` utuh, bukan terhadap potongan yang
        # sudah digeser -- kalau tidak, potongannya menunjuk teks yang salah.
        posisi = m_in.end()
        sisa = ekor[posisi:]
        m_ed = _EDITOR_PROSIDING_RE.match(sisa)
        if m_ed:
            posisi += m_ed.end()
            sisa = sisa[m_ed.end():]
        m_batas = _BATAS_PROSIDING_RE.search(sisa)
        if not m_batas:
            return 0, ""
        return posisi, sisa[: m_batas.start()].rstrip(" ,;:")

    m_vol = _BATAS_VOLUME_RE.search(ekor)
    if not m_vol:
        return 0, ""
    return 0, ekor[: m_vol.end()].rstrip(" ,;:")


def _format_entri_referensi(baris: str) -> str:
    """Rapikan satu baris Daftar Pustaka menjadi markah APA yang diminta.

    Empat hal dikerjakan di sini, dan keempatnya gagal diam-diam kalau
    dilewati:

    1. **Penulis tebal, judul miring.** Tutor meminta format itu eksplisit.
       Dulu format itu diserahkan sepenuhnya ke agen: kalau agen lupa menulis
       `*Judul*`, dokumen keluar tanpa miring sama sekali dan tidak ada yang
       memperbaikinya. Di sini miring pada badan entri dibuang lebih dulu,
       lalu ditambahkan kembali dari struktur baris: `**penulis**`,
       `*judul*`, dan untuk artikel jurnal nama majalahnya juga miring.
    2. **Alamat Reader dibuang.** Referensi bahan ajar sering diisi URL
       `http://127.0.0.1:PORT/berkas?u=...&k=<token>`. URL itu mati begitu
       Reader berhenti, dan `k=`-nya adalah kunci akses akun yang masih hidup.
       Mencetaknya di dokumen sama saja menempelkan kunci itu ke berkas yang
       diserahkan ke tutor.
    3. **Baris bukan entri dibuang.** `SELESAI`, `TIDAK ADA REFERENSI YANG
       TERVERIFIKASI`, dan kalimat karangan seperti "Tidak ada referensi
       tambahan yang diperlukan." bukan rujukan.
    4. **Nomor dan bullet dibuang.** Penanda urutannya adalah alfabetis.

    Tautan DOI/URL ditarik keluar lebih dulu, sebelum miring dihitung.
    Kalau tidak, tautan ikut masuk ke bagian miring, dan DOI yang ikut
    tercetak miring adalah kegagalan yang paling kelihatan di seluruh
    Daftar Pustaka.

    Return string ber-markah, atau `""` kalau barisnya memang bukan entri.
    """
    teks = (baris or "").strip()
    if not teks or _BARIS_BUKAN_ENTRI_RE.match(teks) or _NO_REFERENSI_RE.search(teks):
        return ""

    teks = re.sub(r"^(?:[-*+]|\[?\d+\]?[.)]?)\s+", "", teks).strip()

    # Alamat Reader beserta kunci aksesnya, lalu markdown tautan yang
    # membungkusnya. Urutan itu penting: kuncinya dibuang lebih dulu supaya
    # sisa URL tidak menyisakan `?&` yang menggantung di akhir entri.
    teks = _TOKEN_READER_RE.sub("", teks)
    teks = _URL_READER_ENTRI_RE.sub("", teks)
    teks = _MARKAH_TAUTAN_RE.sub("", teks)

    # Tautan asli dipisah dari badan entri. Tinggalkan di ujung, tidak pernah
    # ikut miring.
    tautan = ""
    m_tautan = None
    for m in _TAUTAN_RE.finditer(teks):
        m_tautan = m
    if m_tautan:
        tautan = m_tautan.group(0).rstrip(".,;)")
        teks = (teks[: m_tautan.start()] + " " + teks[m_tautan.end():])
    teks = _EKOR_DIAKSES_RE.sub("", teks)

    teks = re.sub(r"\s+([.,;])", r"\1", teks)
    teks = re.sub(r"\s{2,}", " ", teks).strip(" ;,")
    if not teks:
        return tautan

    # Markah lama dibuang supaya tidak jadi dobel: agen boleh sudah menulis
    # `**Rosen, K. H.**` atau `*Discrete Mathematics*`.
    teks = _BOLD_RE.sub(r"\1", teks)
    teks = _ITALIC_RE.sub(r"\1", teks)

    m_tahun = _TAHUN_RE.search(teks)
    if not m_tahun:
        # Tanpa tahun tidak ada cara memisahkan penulis dari judul. Entri
        # seperti ini tetap dicetak apa adanya daripada ditebak-tebak.
        return _sambung_tautan(teks, tautan)

    penulis = teks[: m_tahun.start()].strip().rstrip(",").strip()
    tahun = m_tahun.group(0)
    sisa = teks[m_tahun.end():]
    if sisa.startswith("."):
        sisa = sisa[1:]
    sisa = sisa.lstrip()

    if not penulis:
        return _sambung_tautan(teks, tautan)

    judul, ekor = _pisah_kalimat(sisa)
    judul, ekor = judul.strip(), ekor.strip()

    hasil = f"**{penulis}** {tahun}."
    if judul:
        hasil += f" *{judul}*"
    if ekor:
        # Pada artikel jurnal, nama majalah beserta nomornya juga miring.
        # `_bagian_miring` yang memutuskan; kalau tidak ada pola yang cocok,
        # ekornya dibiarkan tegak daripada nebak.
        posisi, miring = _bagian_miring(ekor)
        if miring:
            if posisi:
                hasil += " " + ekor[:posisi].rstrip()
            hasil += f" *{miring}*"
            sisa_ekor = ekor[posisi + len(miring):].strip()
            if sisa_ekor:
                # Ekor yang dimulai dengan tanda baca menyambung langsung,
                # supaya tidak lahir "`*Judul* . Jakarta`".
                hasil += (
                    sisa_ekor if sisa_ekor[0] in ".,;:"
                    else f" {sisa_ekor}"
                )
        else:
            hasil += f" {ekor}"
    return _sambung_tautan(hasil, tautan).strip()


def _sambung_tautan(teks: str, tautan: str) -> str:
    """Tempelkan tautan di ujung entri, sekali saja dan tanpa miring."""
    if not tautan:
        return teks.strip()
    return f"{teks.strip()} {tautan}".strip()


def _bersihkan_tanda_humanizer(md: str) -> str:
    """Buang sisa label laporan skill `humanizer` dari draf jawaban.

    Pipeline menyuruh worker memakai skill itu dalam mode File, dan mode itu
    memang menulis teks akhir saja ke berkas. Tapi kalau worker salah pilih
    mode pasted, yang ikut tertulis adalah daftar polanya: blok `**Before:**`,
    `Draft:`, `Remaining patterns`, dan pemisah `---` di sekitarnya. Label
    seperti itu bukan cuma tidak rapi -- penilai langsung tahu teksnya lewat
    alat, bukan ditulis mahasiswa.
    """
    if not md:
        return md

    label = re.compile(
        r"^\s*(?:[-*>#]+\s*)?\**"
        r"(?:before|after|draft|remaining patterns?|"
        r"pola (?:yang )?tersisa|catatan (?:humanizer|skill))"
        r"\**\s*(?:[-*:]\s*)?.*$",
        re.IGNORECASE,
    )

    keluar: list[str] = []
    for b in md.splitlines():
        if label.match(b):
            # Pemisah `---` tepat sebelum blok label ikut dibuang supaya
            # tidak tertinggal garis sendiri di dokumen.
            if keluar and keluar[-1].strip() in ("---", "***", "___"):
                keluar.pop()
            continue
        keluar.append(b)

    return re.sub(r"\n{3,}", "\n\n", "\n".join(keluar))


def _paragraf_referensi(doc: Document) -> list:
    """Paragraf entri Daftar Pustaka, dalam urutan dokumen.

    Placeholder "tidak ada referensi yang terverifikasi" bukan entri, jadi
    tidak ikut; ia ditangani terpisah sebagai catatan.
    """
    found = False
    entries = []
    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt.lower().startswith("daftar pustaka"):
            found = True
            continue
        if not found or not txt or _NO_REFERENSI_RE.search(txt):
            continue
        entries.append(p)
    return entries


def _kunci_urut_referensi(teks: str) -> str:
    """Kunci pengurutan alfabetis APA untuk satu entri.

    APA 7 mengurutkan berdasarkan nama belakang penulis. Tanda baca di awal
    entri harus diabaikan, supaya entri yang diawali kutipan tidak mendahului
    entri yang diawali huruf biasa.
    """
    bersih = re.sub(r"^\d+[.)]\s+", "", teks)
    bersih = re.sub(r"^[^\w\s]+", "", bersih)
    bersih = bersih.lstrip("\"'“”‘’ ")
    return bersih.casefold()


def _urutkan_referensi_apa(doc: Document) -> None:
    """Urutkan entri Daftar Pustaka secara alfabetis (A-Z), di dalam dokumen.

    Karena Daftar Pustaka tidak lagi memakai nomor, urutan alfabetis itu satu-
   -satunya penanda urutannya. Tanpa langkah ini, urutannya persis seperti
    urutan penemuan agen -- dan itu langsung terlihat, karena entri pertama
    yang dimulai huruf besar bisa jatuh di tengah daftar.

    Pengurutan dilakukan dengan memindahkan elemen XML, bukan menulis ulang
    paragraf. Run di dalam entri membawa format -- italic untuk judul buku,
    tautan, nomor halaman -- dan semuanya ikut hilang kalau teksnya disalin
    ulang ke paragraf baru.
    """
    entries = _paragraf_referensi(doc)
    if len(entries) < 2:
        return
    urut = sorted(entries, key=lambda p: _kunci_urut_referensi(p.text.strip()))
    if urut == entries:
        return
    anchor = entries[0]._p
    for p in urut:
        anchor.addnext(p._p)
        anchor = p._p


def _style_referensi_apa(doc: Document) -> None:
    """Referensi diberi hanging indent APA 7: `w:ind left=720 hanging=720`.

    Ini membetulkan angka yang lama dipakai. Versi sebelumnya memakai
    `Cm(0.63)`, yang kira-kira 0,25 inci, karena meniru template. Tapi APA 7
    menetapkan 0,5 inci, dan itulah yang diminta tutor. Dengan 0,25 inci, baris
    kedua rujukan nyaris rata dengan teks biasa sehingga Daftar Pustaka terlihat
    seperti paragraf biasa, bukan daftar. 720 twentieths of a point =
    720/1440 inci = 0,5 inci persis.

    Nilainya ditulis langsung ke `w:ind`, bukan lewat `paragraph_format`,
    supaya style tidak menimpanya belakangan. `List Paragraph` di template
    membawa `w:ind left="720"` sendiri, dan itulah yang membuat indentasi
    gantung hilang begitu paragraf diklik di Word.
    """
    for p in _paragraf_referensi(doc):
        pPr = p._p.get_or_add_pPr()
        ind = pPr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind")
            pPr.append(ind)
        ind.set(qn("w:left"), str(_APA_HANGING_TWIP))
        ind.set(qn("w:hanging"), str(_APA_HANGING_TWIP))
        # Style bawaan bisa membawa indentasi lain lewat `w:pStyle`; lepaskan
        # supaya nilai di atas benar-benar yang dipakai.
        for gaya in pPr.findall(qn("w:pStyle")):
            nilai = (gaya.get(qn("w:val")) or "").replace(" ", "")
            if nilai in {"ListParagraph", "ListNumber", "ListBullet"}:
                pPr.remove(gaya)


def _tandai_placeholder_referensi(doc: Document) -> None:
    """Placeholder referensi dicetak miring dan rata tengah.

    Placeholder itu jawaban yang jujur, bukan sitasi. Bentuknya dibedakan agar
    tutor langsung tahu tidak ada sumber yang bisa diperiksa, bukan mengira
    dokumen ini punya Daftar Pustaka sungguhan.
    """
    for p in doc.paragraphs:
        if _NO_REFERENSI_RE.search(p.text):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.italic = True
            return


def _convert_to_doc(docx_path: Path, doc_path: Path) -> bool:
    try:
        import win32com.client  # type: ignore

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        try:
            wd = word.Documents.Open(str(docx_path.resolve()))
            wd.SaveAs2(str(doc_path.resolve()), FileFormat=0)  # 0 = Word 97-2003 .doc
            wd.Close(False)
        finally:
            word.Quit()
        return True
    except Exception as exc:  # noqa: BLE001
        print(f"  ! Konversi ke .doc dilewati ({exc}). File .docx dipertahankan.")
        return False


def rangkas_dokumen(docx_path: Path) -> list[str]:
    """Beri tahu apa yang benar-benar keluar dari renderer.

    Tanpa ini, tahap docx.py tidak terlihat di log sama sekali: pipeline
    hanya mencetak baris "Dokumen: <path>", jadi ketika persamaan keluar
    sebagai LaTeX mentah tidak ada yang bisa bilang itu terjadi di sini,
    bukan di model. Baris-baris ini dibaca ulang dari berkas yang benar-benar
    ditulis, jadi angkanya mencerminkan isi dokumen, bukan niat renderer.
    """
    try:
        doc = Document(str(docx_path))
    except Exception as exc:  # noqa: BLE001 - ringkasan tidak boleh gagal diam
        return [f"dokumen tidak bisa dibaca ulang untuk ringkasan ({exc})"]

    n_equation = len(doc.element.body.findall(f".//{{{_MATH_NS}}}oMath"))
    # Equation yang masih membawa `\` berarti ada environment LaTeX yang
    # tidak punya OMML-nya dan jatuh ke teks mentah.
    n_mentah = sum(
        1
        for eq in doc.element.body.findall(f".//{{{_MATH_NS}}}oMath")
        for t in eq.iter(f"{{{_MATH_NS}}}t")
        if (t.text or "").count("\\") >= 2
    )
    n_kisi = len(doc.element.body.findall(f".//{{{_MATH_NS}}}m"))
    n_rata = len(doc.element.body.findall(f".//{{{_MATH_NS}}}eqArr"))

    referensi = _paragraf_referensi(doc)
    n_miring = sum(
        1 for p in referensi for r in p.runs if r.italic and (r.text or "").strip()
    )
    n_tebal = sum(
        1 for p in referensi for r in p.runs if r.bold and (r.text or "").strip()
    )

    out = [
        f"{len(doc.paragraphs)} paragraf, {n_equation} equation Word"
        f" ({n_kisi} matriks, {n_rata} baris rata)",
        f"{len(referensi)} entri Daftar Pustaka"
        f" ({n_tebal} bagian tebal, {n_miring} bagian miring)",
    ]
    if n_mentah:
        out.append(
            f"! {n_mentah} equation masih memuat LaTeX mentah "
            "(environment yang belum punya OMML)"
        )
    return out


def save_doc(
    jawaban_md: str,
    soal_text: str,
    meta: dict,
    out_dir: Path,
    *,
    template: Path | None = None,
    include_soal: bool = True,
    gambar_soal: list[Path] | None = None,
) -> tuple[Path, Path]:
    """Simpan jawaban sebagai satu file .docx final (equation OMML asli).

    Equation asli (Word Equation) hanya dapat disimpan dalam format OOXML
    (.docx); konversi ke .doc lama akan mengubah equation menjadi gambar.
    Return (docx_path, docx_path) agar pemanggil tetap API tuple.

    `include_soal` (`False` = jangan tulis bagian soal) dan `gambar_soal`
    (gambar soal yang harus disisipkan apa adanya) lihat `build_docx`.
    """
    base = meta.get("file_base", "jawaban")
    docx_path = out_dir / f"{base}.docx"
    doc_path = out_dir / f"{base}.doc"
    doc_path.unlink(missing_ok=True)

    if template is not None and not Path(template).is_file():
        print(f"  ! format jawaban tidak ditemukan ({Path(template).name}), pakai template standar.")
        template = None

    # Path hasil `build_docx` yang dikembalikan, bukan `docx_path` tebakan.
    # Kalau berkas tujuan sedang terkunci, `_simpan_aman` menulis ke nama
    # bernomor; kalau pipeline tetap melaporkan `docx_path`, yang disebut ke
    # pengguna adalah berkas lama yang isinya belum pernah ditulis ulang.
    ditulis = build_docx(
        jawaban_md=jawaban_md,
        soal_text=soal_text,
        meta=meta,
        out_docx=docx_path,
        # Soal tetap dirender: di depan jawaban, dengan bentuk yang sama
        # seperti sumbernya (gambar disisipkan, teks ditempel).
        include_soal=include_soal,
        template=template,
        gambar_soal=gambar_soal,
    )
    return ditulis, ditulis


# Metadata Moodle yang ikut ter-render tapi tidak ada artinya di berkas yang
# dikumpulkan tutor. Semuanya satu baris "Label: nilai" yang ditambahkan Reader
# supaya jejaknya jelas saat agent membaca -- bukan isi soal.
_SOAL_BUANG_RE = re.compile(
    r"^\s*[-*]?\s*(?:lokasi|di buka|opened|opened by|due|due date|deadline|"
    r"jenis konten|url sumber|halaman ini|last modified|terakhir diubah|"
    r"kelas|kursus|subject|group|grup|tag|assessment|penilaian|nilai|"
    r"jumlah|attempt|percobaan)\s*[:\-]\s*.*$",
    re.IGNORECASE,
)

# Baris "Tugas 1" / "Forum Diskusi.4" sebagai label -- bukan isi soal.
_SOAL_LABEL_RE = re.compile(
    r"^\s*#{0,3}\s*(?:soal|tugas|diskusi|pertanyaan|tasks?|forum)\b[\s.:0-9-]*$",
    re.IGNORECASE,
)

# Lebar gambar yang masih muat di antara margin template (A4/Letter, margin
# kiri-kanan ~1.9 cm). 15 cm aman untuk keduanya.
_GAMBAR_LEBAR_CM = 15.0


def _soal_bersih(soal_text: str) -> str:
    """Buang kerangka Moodle dari teks soal, sisakan instruksi yang sebenarnya."""
    lines = (soal_text or "").splitlines()
    out: list[str] = []
    for line in lines:
        if _SOAL_BUANG_RE.match(line):
            continue
        if _SOAL_LABEL_RE.match(line):
            continue
        out.append(line)
    # Runtuhkan blok kosong berlebih yang ditinggalkan oleh penyaringan.
    teks = "\n".join(out).strip()
    teks = re.sub(r"\n{3,}", "\n\n", teks)
    return _buang_markah_reader(teks)


# Penjaga terakhir terhadap kebocoran token Reader.
#
# `_bersihkan_markah_lokal` di `html2md.py` sudah membuang tautan Reader saat
# teks soal dibaca dari Moodle, dan tahap `_tahap_gambar_soal` mengganti gambar
# dengan transkripsi. Dua-duanya berada di lapisan data, jadi bisa dilewati
# kalau ada pemanggil lain yang mengisi `soal_text` sendiri -- tes, atau
# perubahan di kemudian hari. `127.0.0.1/soal?u=` adalah URL bertoken yang
# hanya berlaku selama pipeline hidup: begitu keluar ke berkas yang
# dikumpulkan tutor, tautan itu mati dan, lebih buruk, memuat kunci Reader
# yang sedang aktif.
_MARKAH_GAMBAR_RE = re.compile(r"!?\[[^\]\n]*\]\([^)\n]*\)")
_URL_READER_RE = re.compile(r"https?://(?:127\.0\.0\.1|localhost)(?::\d+)?\S*")


def _buang_markah_reader(teks: str) -> str:
    """Buang markdown gambar dan URL Reader dari teks soal."""
    bersih = _MARKAH_GAMBAR_RE.sub("", teks)
    bersih = _URL_READER_RE.sub("[tautan dihapus]", bersih)
    return re.sub(r"[ \t]{2,}", " ", bersih).strip()


def _sisip_gambar(doc: Document, path: Path) -> bool:
    """Sisipkan satu gambar ke paragraf baru. True kalau berhasil."""
    try:
        from docx.shared import Cm as _Cm

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(str(path), width=_Cm(_GAMBAR_LEBAR_CM))
        return True
    except Exception as exc:  # noqa: BLE001 - gambar rusak tidak boleh mematikan docx
        print(f"  ! gambar soal gagal disisipkan ({Path(path).name}): {exc}")
        return False


def _tulis_soal(
    doc: Document,
    soal_text: str,
    *,
    gambar_soal: list[Path] | None = None,
) -> None:
    """Tulis bagian `## Soal` di depan jawaban.

    Dua sumber, keduanya dipertahankan apa adanya karena keduanya adalah naskah
    resmi: gambar lampiran disisipkan, teks soal ditempel. Yang dibuang hanya
    kerangka Moodle (baris "Lokasi:", "URL sumber:", dan sejenisnya) yang tidak
    pernah menjadi bagian soal.

    Kalau soalnya berupa gambar, teks hasil transkripsi tidak boleh menggantikan
    gambar: transkripsi dibuat model dan bisa keliru, sedangkan yang dipegang
    pengajar adalah naskah aslinya. Kalau teksnya ada DAN berarti (bukan cuma
    "halaman ini tidak memuat teks"), teksnya ikut dicetak sebagai alat bantu
    baca.
    """
    gambar = [Path(p) for p in (gambar_soal or []) if Path(p).is_file()]
    teks = _soal_bersih(soal_text)

    if not gambar and not teks:
        return

    # Judul "Soal" selalu bold hitam TNR, sama seperti heading lain.
    sh = doc.add_paragraph(style=_heading_style(doc, 2))
    run = sh.add_run("Soal")
    _rapikan_heading(run, doc, _heading_style(doc, 2))

    sisip_ok = False
    for path in gambar[:3]:
        sisip_ok = _sisip_gambar(doc, path) or sisip_ok

    # Teks ikut dicetak kalau salah satu: gambar gagal disisipkan, atau teksnya
    # berarti lebih dari sekadar pesan "halaman ini tidak memuat teks".
    if teks and len(teks) > 120:
        _render_markdown(doc, teks)
    elif teks and not sisip_ok:
        _render_markdown(doc, teks)

    doc.add_paragraph()
