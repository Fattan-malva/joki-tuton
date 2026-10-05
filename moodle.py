"""Klien HTTP untuk Moodle Universitas Terbuka.

Tugas modul ini sempit dan penting: menjalankan satu sesi Moodle yang sah,
memverifikasi bahwa cookie benar-benar masih hidup, dan memberi cache HTML
supaya pipeline tidak menarik halaman yang sama berulang kali.

Kenapa tidak memakai `requests` langsung di modul lain: penyuntikan cookie
harus konsisten di mana-mana, dan verifikasi login harus punya satu jawaban
yang sama. Kalau tiap modul cek sendiri-sendiri, satu tempat akan lupa dan
menerjemahkan "cookie expired" menjadi "matkul tidak ditemukan".
"""

from __future__ import annotations

import hashlib
import re
import time
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import urljoin, urlparse

import requests

import config

_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/126.0 Safari/537.36"
)


class MoodleError(RuntimeError):
    """Kegagalan yang perlu dijelaskan ke pengguna, bukan traceback kosong."""


class LoginMati(MoodleError):
    """Cookie sesi tidak lagi berlaku di server."""


@dataclass
class Halaman:
    """Satu halaman Moodle yang sudah diambil."""

    url: str
    html: str
    status: int
    dari_cache: bool = False
    headers: dict = field(default_factory=dict)


def _cookie_header() -> str:
    """Susun header Cookie dari nilai yang ada di `.env`.

    Nilai di `.env` hanya berisi nilai cookie tanpa nama. Kalau ternyata sudah
    berbentuk `Nama=Nilai; Nama2=Nilai2`, string itu dipakai langsung -- jadi
    cookie yang lebih dari satu tidak perlu kode tambahan.
    """
    raw = config.COOKIE_MOODLE
    if not raw:
        raise MoodleError(
            "COOKIE_MOODLE kosong di .env. Salin cookie MoodleSession dari "
            "peramban lalu isikan."
        )
    if "=" in raw:
        return raw
    return f"{config.MOODLE_COOKIE_NAME}={raw}"


class Moodle:
    """Sesi HTTP terautentikasi untuk satu host Moodle."""

    def __init__(
        self,
        base_url: str | None = None,
        cookie: str | None = None,
        cache_dir: Path | None = None,
    ) -> None:
        self.base_url = (base_url or config.URL_MOODLE).rstrip("/")
        if not self.base_url:
            raise MoodleError("URL_MOODLE kosong di .env.")
        self._cookie = cookie or _cookie_header()
        self.cache_dir = Path(cache_dir or config.CACHE_DIR)
        self.cache_dir.mkdir(parents=True, exist_ok=True)

        self.sesi = requests.Session()
        self.sesi.headers.update(
            {
                "User-Agent": _USER_AGENT,
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Accept-Language": "id-ID,id;q=0.9",
                "Cookie": self._cookie,
            }
        )
        self.sesskey: str = ""

    # ------------------------------------------------------------------ cache

    def _kunci_cache(self, url: str) -> Path:
        digest = hashlib.sha256(url.encode("utf-8")).hexdigest()[:24]
        return self.cache_dir / f"{digest}.html"

    def _baca_cache(self, url: str) -> str | None:
        path = self._kunci_cache(url)
        if not path.is_file():
            return None
        if time.time() - path.stat().st_mtime > config.TTL_CACHE_DETIK:
            return None
        try:
            return path.read_text(encoding="utf-8", errors="replace")
        except OSError:
            return None

    def _tulis_cache(self, url: str, html: str) -> None:
        try:
            self._kunci_cache(url).write_text(html, encoding="utf-8")
        except OSError:
            # Cache bukan hal penting. Kalau disk penuh atau file terkunci,
            # pipeline harus tetap jalan.
            pass

    # ------------------------------------------------------------------ ambil

    def absolut(self, url: str) -> str:
        """Ubah URL relatif dari halaman Moodle menjadi URL lengkap."""
        if not url:
            return ""
        if url.startswith(("http://", "https://")):
            return url
        return urljoin(self.base_url + "/", url.lstrip("/"))

    def ambil(
        self,
        url: str,
        *,
        pakai_cache: bool = True,
        timeout: int = 45,
    ) -> Halaman:
        """Ambil satu halaman. Melempar `LoginMati` kalau cookie sudah tidak berlaku."""
        url = self.absolut(url)
        if pakai_cache:
            cached = self._baca_cache(url)
            if cached is not None:
                return Halaman(url=url, html=cached, status=200, dari_cache=True)

        try:
            resp = self.sesi.get(url, timeout=timeout, allow_redirects=True)
        except requests.RequestException as exc:
            raise MoodleError(f"Gagal menghubungi Moodle: {exc}") from exc

        # Halaman login adalah bentuk paling umum dari cookie yang mati. Cache
        # tidak boleh dipakai untuk url ini: isinya hasil sebelum sesi
        # kedaluwarsa akan membuat pipeline melaporkan matkul ada padahal
        # server sudah menolaknya.
        if self._adalah_halaman_login(resp.text, resp.url):
            if pakai_cache:
                self._kunci_cache(url).unlink(missing_ok=True)
            raise LoginMati(
                "Cookie Moodle tidak berlaku lagi. Buka "
                f"{self.base_url} di peramban, login, lalu salin ulang "
                "nilai cookie MoodleSession ke .env."
            )

        html = resp.text
        if resp.ok and pakai_cache:
            self._tulis_cache(url, html)
        self._tangkap_sesskey(html)
        return Halaman(
            url=resp.url,
            html=html,
            status=resp.status_code,
            headers=dict(resp.headers),
        )

    @staticmethod
    def _adalah_halaman_login(html: str, url_akhir: str) -> bool:
        """Tentukan apakah respons sebenarnya halaman login.

        Urutan pemeriksaan penting. Halaman login UT memuat rantai CSS dan JS
        yang panjang, sehingga penanda `notloggedin` baru muncul di karakter
        ke-4000-an -- memindai 6000 karakter awal tidak cukup dan, lebih buruk,
        memindai seluruh dokumen bisa salah tangkap dari string JavaScript di
        halaman yang sah. Karena itu penanda dicari di tag `<body>` saja.
        """
        url = (url_akhir or "").split("?")[0].split("#")[0]
        if url.endswith("/login/index.php") or "/login/index.php" in url:
            return True
        if "/webservice/" in url or url.endswith("/token.php"):
            return True

        # Periksa tag pembuka <body> saja, bukan seluruh dokumen.
        m = re.search(r"<body\b[^>]*>", html or "", re.I)
        if not m:
            # Tanpa tag <body> tidak ada yang bisa diperiksa. Kembalikan False
            # supaya deteksi ini tidak menebak-nebak.
            return False
        body_tag = m.group(0)
        if re.search(r"\bnotloggedin\b", body_tag):
            return True
        if re.search(r'id=["\']page-login-index["\']', body_tag):
            return True
        return False

    def _tangkap_sesskey(self, html: str) -> None:
        """Ambil `sesskey` dari config JS Moodle.

        Reader butuh ini untuk membuat tautan yang bisa dibuka di sisi server,
        dan beberapa halaman Moodle memintanya pada parameter GET.
        """
        if self.sesskey:
            return
        for penanda in ('"sesskey":"', 'sesskey=&#039;', 'name="sesskey" value="'):
            idx = html.find(penanda)
            if idx == -1:
                continue
            sisa = html[idx + len(penanda) :]
            if penanda == '"sesskey":"':
                akhir = sisa.find('"')
                if akhir > 0:
                    self.sesskey = sisa[:akhir]
                    return
            else:
                awal = 0 if penanda.startswith("sesskey=") else 0
                akhir = min(
                    (p for p in (sisa.find("'"), sisa.find('"')) if p > 0),
                    default=-1,
                )
                if akhir > 0:
                    self.sesskey = sisa[awal:akhir]
                    return

    # ------------------------------------------------------------------ cek login

    def cek_login(self) -> tuple[bool, str]:
        """Cek apakah cookie masih hidup.

        Returns `(berhasil, pesan)`. Pesan selalu diisi supaya pemanggil bisa
        langsung menampilkannya tanpa menyusun kalimat sendiri.
        """
        try:
            halaman = self.ambil(f"{self.base_url}/my/", pakai_cache=False, timeout=30)
        except LoginMati as exc:
            return False, str(exc)
        except MoodleError as exc:
            return False, str(exc)

        html = halaman.html
        # Penanda sudah masuk: di halaman `/my/` Moodle selalu memuat tautan ke
        # Dashboard dan ke daftar "My courses". Kalau keduanya hilang berarti
        # server sudah mengeluarkan sesi ini.
        masuk = all(t in html for t in ("my/courses.php", "user/profile"))
        if masuk:
            return True, "Cookie valid."
        if "notloggedin" in html[:20000] or "page-login-index" in html[:20000]:
            return False, (
                "Moodle masih menampilkan halaman untuk tamu. Cookie "
                "MoodleSession kemungkinan sudah kedaluwarsa."
            )
        return False, (
            "Halaman /my/ terbuka tapi tidak memuat penanda DASBOR. "
            "Bisa jadi sesi sudah berakhir di server."
        )

    def unduh(self, url: str, dest: Path, *, timeout: int = 120) -> Path:
        """Unduh satu berkas lampiran ke `dest`.

        Nama berkas diambil dari URL supaya nama asli dari Moodle ikut terbawa --
        nama itu sering jadi petunjuk apakah lampiran milik dosen atau kiriman
        mahasiswa.
        """
        url = self.absolut(url)
        dest.parent.mkdir(parents=True, exist_ok=True)

        if self._adalah_halaman_login("", url):
            raise LoginMati("URL lampiran mengarah ke halaman login.")

        try:
            with self.sesi.get(url, timeout=timeout, stream=True, allow_redirects=True) as resp:
                if self._adalah_halaman_login("", resp.url):
                    raise LoginMati("Cookie habis saat mengunduh lampiran.")
                if not resp.ok:
                    raise MoodleError(f"Unduhan gagal ({resp.status_code}): {url}")

                ctype = (resp.headers.get("Content-Type") or "").lower()
                if "text/html" in ctype:
                    # pluginfile.php yang gagal membalas halaman login dengan
                    # status 200. Menyimpan HTML itu sebagai PDF akan
                    # menghasilkan bahan ajar berisi kata "Masuk ke situs",
                    # jadi lebih baik gagal keras.
                    raise LoginMati(
                        "Moodle membalas halaman login, bukan berkas. "
                        "Cookie kemungkinan sudah kedaluwarsa."
                    )

                nama = Path(urlparse(resp.url).path).name or "lampiran"
                # Header Content-Disposition kadang memuat nama asli.
                cd = resp.headers.get("Content-Disposition", "")
                if "filename=" in cd:
                    kandidat = cd.split("filename=")[-1].strip("\"' ")
                    if kandidat:
                        nama = kandidat

                batas = config.MAX_UNDUH_MB * 1024 * 1024
                total = 0
                with dest.open("wb") as fh:
                    for potongan in resp.iter_content(chunk_size=65536):
                        total += len(potongan)
                        if total > batas:
                            raise MoodleError(
                                f"Lampiran lebih dari {config.MAX_UNDUH_MB} MB: {nama}"
                            )
                        fh.write(potongan)
        except requests.RequestException as exc:
            raise MoodleError(f"Gagal mengunduh lampiran: {exc}") from exc

        return dest