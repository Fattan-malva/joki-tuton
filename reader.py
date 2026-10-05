"""Server HTTP lokal yang menjadi jembatan menuju Moodle.

Agent diberi URL `http://127.0.0.1:PORT/soal?u=...` supaya bisa memakai
`webfetch`. Alasannya `webfetch` milik opencode tidak bisa mengirim header
`Cookie`, sedangkan seluruh isi Moodle hanya terbuka bagi sesi yang sudah
login. Tanpa server ini, agent hanya melihat halaman untuk tamu.

Tiga tanggung jawab server ini:

1. Menyuntikkan cookie Moodle ke setiap permintaan.
2. Mengubah HTML menjadi Markdown, karena keluaran `webfetch` berupa teks
   dan agent lebih tepat membaca Markdown daripada HTML mentah.
3. Menerjemahkan ulang tautan. Tautan di dalam Markdown harus tetap bisa
   dibuka, jadi URL Moodle ditulis ulang menjadi URL Reader.

Berjalan di dalam proses pipeline sendiri, hanya mendengarkan di
`127.0.0.1`, dan tidak menerima permintaan dari jaringan luar.
"""

from __future__ import annotations

import base64
import re
import secrets
import sys
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Callable
from urllib.parse import parse_qs, quote, unquote, urlparse

import config
import html2md
import moodle


def _kode(url: str) -> str:
    """Kodekan URL agar muat di parameter query tanpa merusak karakter."""
    return base64.urlsafe_b64encode(url.encode("utf-8")).decode("ascii").rstrip("=")


def _buka(kode: str) -> str:
    """Kebalikan dari `_kode`, dengan padding yang lupa oleh direstore."""
    if not kode:
        return ""
    padding = "=" * (-len(kode) % 4)
    try:
        return base64.urlsafe_b64decode(kode + padding).decode("utf-8")
    except (ValueError, UnicodeDecodeError):
        return ""


_RE_RESOURCE_VIEW = re.compile(r"/mod/resource/view\.php", re.I)


def _perbaiki_resource(url: str) -> str:
    """Paksa `forceview=1` pada halaman resource Moodle.

    Tanpa flag itu, `mod/resource/view.php` untuk berkas langsung membalas 303
    ke `pluginfile.php`, dan pada sesi UT perjalanan itu berakhir di halaman
    login. Agent pemetaan lalu menerima 401 dan mencatat bahannya "tidak bisa
    dibuka", padahal pipeline sendiri membacanya tanpa masalah.

    Dengan `forceview=1` Moodle merender halaman resource yang memuat tautan
    ke `pluginfile.php`, dan tautan itu bisa diikuti Reader seperti biasa.
    """
    if not url or not _RE_RESOURCE_VIEW.search(url):
        return url
    if "forceview=" in url.lower():
        return url
    return url + ("&" if "?" in url else "?") + "forceview=1"


class Reader:
    """Jembatan lokal menuju Moodle untuk dibaca agent."""

    def __init__(self, klien: moodle.Moodle) -> None:
        self.klien = klien
        self.token = secrets.token_urlsafe(18)
        self._server: ThreadingHTTPServer | None = None
        self._thread: threading.Thread | None = None
        self.port: int = 0

    # ------------------------------------------------------------- alamat

    @property
    def akar(self) -> str:
        return f"http://{config.READER_HOST}:{self.port}"

    def soal(self, url: str, label: str = "") -> str:
        """Alamat Reader untuk satu halaman Moodle."""
        return (
            f"{self.akar}/{config.TARGET_SOAL}"
            f"?u={_kode(_perbaiki_resource(url))}&k={self.token}&t={quote(label, safe='')}"
        )

    def materi(self, url: str, label: str = "") -> str:
        """Alamat Reader untuk halaman bahan ajar."""
        return (
            f"{self.akar}/{config.TARGET_MATERI}"
            f"?u={_kode(_perbaiki_resource(url))}&k={self.token}&t={quote(label, safe='')}"
        )

    def berkas(self, url: str, label: str = "") -> str:
        """Alamat Reader untuk berkas biner seperti PDF atau gambar.

        Agent memakai `read` untuk berkas gambar, dan `webfetch` tidak
        berguna di sini karena isinya bukan teks. Yang dilayani adalah
        teks hasil ekstraksi PDF, bukan PDF mentahnya, supaya `read` tidak
        perlu menangani biner melalui perantara HTTP.
        """
        return (
            f"{self.akar}/berkas"
            f"?u={_kode(url)}&k={self.token}&t={quote(label, safe='')}"
        )

    def ulang(self, nilai: str) -> str:
        """Terjemahkan URL Moodle menjadi URL Reader, atau biarkan bila luar.

        Hanya tautan ke host Moodle yang diterjemahkan. Tautan ke luar,
        misalnya artikel jurnal, dibiarkan apa adanya supaya agent bisa
        membacanya langsung lewat `webfetch`.
        """
        if not nilai:
            return nilai
        if nilai.startswith("data:"):
            return nilai
        if config.READER_HOST in nilai:
            return nilai

        host = urlparse(nilai).netloc.lower()
        host_moodle = urlparse(config.URL_MOODLE).netloc.lower()
        if host != host_moodle:
            return nilai

        # `pluginfile.php` melayani berkas biner, sedangkan `mod/...` melayani
        # halaman HTML. Menblind keduanya ke jalur yang sama akan membuat
        # `webfetch` menerima PDF mentah yang tidak bisa dibaca.
        if "/pluginfile.php/" in nilai:
            return self.berkas(nilai, html2md.bersihkan_nama_berkas(nilai))
        return self.soal(nilai)

    def pengalir(self) -> Callable[[str], str]:
        """Fungsi rewrite untuk `html2md.ke_markdown`."""
        return self.ulang

    # -------------------------------------------------------------- hidup

    def mulai(self) -> str:
        """Nyalakan server dan kembalikan alamatnya."""
        if self._server is not None:
            return self.akar

        reader = self

        class _Handler(BaseHTTPRequestHandler):
            protocol_version = "HTTP/1.1"
            server_version = "JktReader/1.0"

            def log_message(self, *_args) -> None:  # noqa: D102
                # Taksonomi request akan memenuhi terminal dengan baris
                # URL lengkap yang memuat cookie dalam bentuk dasar64.
                pass

            def do_GET(self) -> None:  # noqa: N802
                reader._layani(self)

        class _Server(ThreadingHTTPServer):
            # Agent menutup koneksi tanpa memberi tahu saat henti membaca
            # halaman atau saat `webfetch` sudah cukup isinya. Tanpa
            # penangan ini, `socketserver` mencetak seluruh traceback
            # ConnectionResetError ke stderr dan bercampur dengan keluaran
            # pipeline. Koneksi terputus bukan kegagalan Reader: halamannya
            # sudah terlayani atau memang tidak pernah diminta.
            def handle_error(self, request, client_address) -> None:
                exc = sys.exc_info()[1]
                if isinstance(exc, (ConnectionResetError, BrokenPipeError,
                                    ConnectionAbortedError, TimeoutError)):
                    return
                super().handle_error(request, client_address)

        self._server = _Server(
            (config.READER_HOST, config.READER_PORT), _Handler
        )
        self.port = self._server.server_address[1]
        self._thread = threading.Thread(
            target=self._server.serve_forever, name="jkt-reader", daemon=True
        )
        self._thread.start()
        return self.akar

    def berhenti(self) -> None:
        if self._server is not None:
            self._server.shutdown()
            self._server.server_close()
            self._server = None
        self._thread = None

    def __enter__(self) -> "Reader":
        self.mulai()
        return self

    def __exit__(self, *_exc) -> None:
        self.berhenti()

    # -------------------------------------------------------------- melayani

    def _layani(self, handler: BaseHTTPRequestHandler) -> None:
        bagian = urlparse(handler.path)
        param = parse_qs(bagian.query)

        token = (param.get("k") or [""])[0]
        if not secrets.compare_digest(token, self.token):
            self._balas(handler, 403, "Token Reader salah atau kedaluwarsa.")
            return

        url_asal = _buka((param.get("u") or [""])[0])
        if not url_asal:
            self._balas(handler, 400, "Parameter u tidak berisi URL yang sah.")
            return

        label = unquote((param.get("t") or [""])[0])
        jalur = bagian.path.strip("/")

        try:
            if jalur == config.TARGET_SOAL:
                self._layani_soal(handler, url_asal, label)
            elif jalur == config.TARGET_MATERI:
                self._layani_materi(handler, url_asal, label)
            elif jalur == "berkas":
                self._layani_berkas(handler, url_asal, label)
            else:
                self._balas(
                    handler, 404,
                    f"Jalur {jalur!r} tidak dikenal. Gunakan "
                    f"/{config.TARGET_SOAL}, /{config.TARGET_MATERI}, atau /berkas.",
                )
        except moodle.LoginMati as exc:
            self._balas(handler, 401, f"Sesi Moodle berakhir. {exc}")
        except moodle.MoodleError as exc:
            self._balas(handler, 502, f"Gagal mengambil {url_asal}\n\n{exc}")

    def _layani_soal(self, handler, url: str, label: str) -> None:
        """Halaman soal: forum atau tugas, tanpa jawaban mahasiswa."""
        halaman = self.klien.ambil(url)
        elas = "/mod/assign/" in url or "/mod/quiz/" in url
        md = html2md.ke_markdown(
            halaman.html,
            base_url=halaman.url,
            rewrite=self.ulang,
            elas=elas,
            judul=label or None,
        )
        self._balas(handler, 200, self._berkoridor(md, url))

    def _layani_materi(self, handler, url: str, label: str) -> None:
        """Halaman bahan ajar: deskripsi dan tautan lampiran."""
        halaman = self.klien.ambil(url)
        md = html2md.ke_markdown(
            halaman.html,
            base_url=halaman.url,
            rewrite=self.ulang,
            buang_mahasiswa=False,
            judul=label or None,
        )
        self._balas(handler, 200, self._berkoridor(md, url))

    def _layani_berkas(self, handler, url: str, label: str) -> None:
        """Isi berkas biner, diubah menjadi teks sebelum dikirim."""
        from attachments import teks_dari_berkas

        hasil = teks_dari_berkas(self.klien, url, label=label)
        if hasil.gagal:
            self._balas(handler, 502, hasil.teks or "Berkas gagal diambil.")
            return
        catatan = " | ".join(hasil.catatan) if hasil.catatan else ""
        self._balas(
            handler, 200,
            self._berkoridor(hasil.teks, url, catatan=catatan, nama=hasil.nama),
        )

    @staticmethod
    def _berkoridor(
        isi: str, url: str, *, catatan: str = "", nama: str = ""
    ) -> str:
        """Bungkus hasil dengan penanda asal dan aturan pembacaan."""
        kepala = ["<!--", f"Sumber: {url}"]
        if nama:
            kepala.append(f"Berkas: {nama}")
        if catatan:
            kepala.append(f"Catatan: {catatan}")
        kepala.append(
            "Halaman ini sudah dibersihkan: navigasi situs, notifikasi, dan "
            "jawaban mahasiswa lain tidak disertakan. Tautan bertanda "
            f"{config.READER_HOST} mengarah ke halaman Moodle lain dan boleh "
            "dibuka langsung dengan webfetch."
        )
        kepala.append("-->")
        return "\n".join(kepala) + "\n\n" + (isi or "(tidak ada teks yang bisa dibaca)")

    @staticmethod
    def _balas(handler: BaseHTTPRequestHandler, status: int, isi: str) -> None:
        data = isi.encode("utf-8", errors="replace")
        handler.send_response(status)
        handler.send_header("Content-Type", "text/plain; charset=utf-8")
        handler.send_header("Content-Length", str(len(data)))
        handler.end_headers()
        handler.wfile.write(data)