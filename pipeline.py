"""Orkestrator: dari satu (mata kuliah, sesi) menjadi satu berkas .docx.

Rangkaiannya:

    petakan sesi -> Reader hidup -> peta soal -> bahan ajar -> transkrip
    -> referensi -> jawaban -> .docx

Pemetaan sesi, pembacaan HTML, dan unduhan lampiran dikerjakan Python,
bukan agent. Agent hanya dipakai untuk empat hal yang memang butuh bahasa:
memahami soal, mentranskripsi gambar, mencari referensi, dan menulis
jawaban. Agent tidak pernah memegang cookie Moodle; semua akses lewat
`reader.py` yang berjalan di dalam proses ini.

Urutan ini penting. Peta soal ditulis lebih dulu karena isinya menentukan
lampiran mana yang perlu diunduh, dan daftar pustaka menulis lebih dulu
karena `worker` harus menyalinnya apa adanya, bukan mencari sendiri.
"""

from __future__ import annotations

import importlib.util
import json
import queue
import re
import subprocess
import sys
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path

import attachments
import config
import courses
import html2md
import moodle
import reader as reader_mod
from courses import MataKuliah, Sesi


class PipelineGagal(Exception):
    """Kesalahan yang layak menghentikan seluruh proses."""


# ------------------------------------------------------------ pemanggilan agent


@dataclass
class HasilAgent:
    """Keluaran satu pemanggilan `opencode run`."""

    perintah: list[str]
    selesai: bool
    stdout: str
    stderr: str
    durasi: float
    kode: int

    @property
    def ringkas(self) -> str:
        return (
            f"{self.kode} dalam {self.durasi:.0f} d, "
            f"{len(self.stdout)} karakter keluaran"
        )


# ------------------------------------------------- penampil proses agent


_ALIRAN_KELUARAN = True

# Panjang maksimum satu baris proses. Baris yang lebih panjang dipotong supaya
# terminal tetap bisa dibaca dan log tidak membengkak.
_BARIS_MAKS = 110

# Kunci Reader pada URL Reader berbentuk `k=<base64url>`. Nilai itu membuat
# siapa pun yang memegangnya bisa membuka setiap halaman akun tersebut selama
# Reader masih hidup, jadi tidak boleh muncul di terminal maupun di log.
# `([?&])` ikut ditangkap supaya tanda pemisah paramsnya tidak ikut hilang saat
# nilai k-nya diganti.
_TOKEN_URL_RE = re.compile(r"([?&])k=[A-Za-z0-9_\-]+")


def _saring_token(teks: str) -> str:
    """Ganti nilai `k=` pada URL Reader dengan penanda. Berlaku untuk log juga."""
    return _TOKEN_URL_RE.sub(r"\1***", teks)


def _ringkas_singkat(nilai: object) -> str:
    """Nama berkas atau path URL tanpa bagian sensitifnya.

    `webfetch` menerima URL Reader lengkap, termasuk `?u=...&k=<token>`.
    Token itu adalah kunci yang masih hidup selama pipeline berjalan, jadi
    jangan pernah dicetak ke terminal: terminal sering dialihkan ke berkas
    log, dan berkas log sering ikut dikirim atau ditempel di tempat lain. Yang
    ditampilkan cukup bentuk jalannya: nama berkas, atau segmen terakhir path
    tanpa query string.
    """
    teks = str(nilai).split("?", 1)[0].rstrip("&")
    if "://" in teks:
        return teks.rsplit("/", 1)[-1] or teks.rsplit("/", 2)[-2]
    return Path(teks).name or Path(teks).parent.name


def _alir(event: dict) -> None:
    """Cetak satu baris ringkas soal apa yang sedang dikerjakan agent.

    Agent `worker` satu sesi bisa berjalan beberapa menit tanpa mengetik
    apa pun ke terminal. Tanpa baris seperti ini, tidak ada cara membedakan
    "masih bekerja" dari "macet". Yang ditampilkan bukan isi jawaban --
    teksnya sendiri sudah masuk ke berkas hasil dan nanti masuk dokumen --
    tapi jenis pekerjaannya: langkah ke berapa, tool apa yang dipanggil, dan
    token yang terpakai.

    Setiap baris disaring `_saring_token` lebih dulu. Model kadang menulis
    ulang URL yang baru saja dia buka, dan URL Reader memuat kunci akses
    (`?u=...&k=...`) yang masih hidup selama pipeline berjalan.
    """
    if not _ALIRAN_KELUARAN:
        return

    jenis = event.get("type")
    part = event.get("part")
    if not isinstance(part, dict):
        return

    if jenis == "text":
        teks = part.get("text") or ""
        baris = (teks.strip().splitlines() or [""])[0].strip()
        if not baris:
            return
        print(f"      . {_saring_token(baris)[:_BARIS_MAKS]}", file=sys.stderr, flush=True)

    elif jenis == "tool_use":
        alat = part.get("tool") or "?"
        keadaan = part.get("state") or {}
        status = ""
        if isinstance(keadaan, dict):
            status = str(keadaan.get("status") or "")
            muatan = keadaan.get("input") or {}
            if isinstance(muatan, dict):
                ringkas = (
                    muatan.get("filePath")
                    or muatan.get("path")
                    or muatan.get("pattern")
                    or muatan.get("url")
                )
                if ringkas:
                    status = f"{status} {_ringkas_singkat(ringkas)}"
        print(f"      > {alat} {status}".rstrip(), file=sys.stderr, flush=True)

    elif jenis == "step_finish":
        usage = part.get("tokens") or {}
        if isinstance(usage, dict):
            masuk = usage.get("input") or 0
            keluar = usage.get("output") or 0
            print(
                f"      = langkah selesai ({masuk} in / {keluar} out token)",
                file=sys.stderr, flush=True,
            )


def panggil_agent(
    *,
    nama: str,
    prompt: str,
    model: str,
    berkas: list[Path] | None = None,
    cwd: Path | None = None,
    waktu_maks: int | None = None,
    diam: bool = False,
) -> HasilAgent:
    """Jalankan satu agent lewat `opencode run` dan kumpulkan keluarannya.

    `--auto` sengaja tidak dipakai. Flag itu menyetujui setiap permintaan
    izin tanpa batas, termasuk yang tidak diizinkan spec agent. Semua
    agent di `.opencode/agents/` sudah mendeklarasikan izin yang mereka
    perlukan dengan `effect: allow`, jadi tidak ada permintaan izin yang
    perlu dijawab otomatis.

    `--standalone` dipakai supaya proses tidak bergantung pada service
    latar yang mungkin tidak berjalan saat pipeline dipanggil.
    """
    executable = config.cari_opencode()
    if executable is None:
        raise PipelineGagal(
            "opencode tidak ditemukan di PATH. Pasang opencode "
            "(npm install -g opencode-ai) lalu jalankan ulang."
        )

    perintah = [
        str(executable), "run",
        "--standalone",
        "--agent", nama,
        "--model", model,
        "--format", "json",
    ]
    for b in berkas or []:
        perintah += ["--file", str(b)]

    # Prompt dikirim lewat stdin, bukan sebagai argumen.
    #
    # Argumen `message` di `opencode run` bersifat opsional dan dibaca dari
    # stdin kalau tidak diberikan. Prompt pipeline berisi baris baru, `&`,
    # `|`, `>`, kutip, dan `\frac{a}{b}`. Semua itu akan diurai ulang oleh
    # ParseCommandLine Windows dan, kalau executable-nya shim `.cmd`, sekali
    # lagi oleh cmd.exe. Pengujian menunjukkan prompt sampai dalam keadaan
    # terpotong atau ditafsirkan sebagai operator. stdin tidak melewati
    # parse argumen sama sekali, jadi bentuk teks apa pun aman.

    batas = waktu_maks or config.TIMEOUT_AGENT_DETIK
    mulai = time.monotonic()

    # stdout dibaca sambil mengalir, bukan dikumpulkan diam-diam.
    #
    # `--format json` menulis JSONL satu baris per event, dan event itu muncul
    # saat kejadiannya terjadi, bukan setelah model selesai. Membacanya langsung
    # dari pipa membuat proses agent terlihat di terminal. `subprocess.run`
    # dengan `capture_output` Instead menunggu sampai proses selesai, jadi tidak
    # ada yang bisa ditampilkan selama beberapa menit itu.
    #
    # stderr dikumpulkan lewat thread supaya tidak menumpuk di pipa dan
    # memblokir opencode ketika isinya banyak.
    try:
        proses = subprocess.Popen(  # noqa: S603
            perintah,
            cwd=str(cwd) if cwd else None,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
    except FileNotFoundError as exc:
        raise PipelineGagal(
            f"Executable `{executable}` tidak bisa dijalankan. "
            "Kalau ini shim npm, pastikan paket opencode terpasang penuh."
        ) from exc
    except OSError as exc:
        raise PipelineGagal(f"Gagal menjalankan opencode: {exc}") from exc

    kumpulan_stderr: list[str] = []

    def kumpulkan_stderr() -> None:
        if proses.stderr is not None:
            kumpulan_stderr.append(proses.stderr.read())

    benang = threading.Thread(target=kumpulkan_stderr, name="jkt-stderr", daemon=True)
    benang.start()

    baris_antre: queue.Queue[str | None] = queue.Queue()

    def baca_stdout() -> None:
        if proses.stdout is None:
            baris_antre.put(None)
            return
        try:
            for b in proses.stdout:
                baris_antre.put(b)
        finally:
            baris_antre.put(None)

    benang_stdout = threading.Thread(target=baca_stdout, name="jkt-stdout", daemon=True)
    benang_stdout.start()

    if proses.stdin is not None:
        try:
            proses.stdin.write(prompt)
            proses.stdin.close()
        except (BrokenPipeError, OSError):
            # opencode menolak prompt (mis. model tidak tersedia). Prosesnya
            # tetap dibaca sampai habis supaya stderr-nya jelas.
            pass

    potongan: list[str] = []
    while True:
        sisa = batas - (time.monotonic() - mulai)
        if sisa <= 0:
            proses.kill()
            proses.wait()
            return HasilAgent(
                perintah=perintah,
                selesai=False,
                stdout=_teks_dari_json("".join(potongan)),
                stderr="".join(kumpulan_stderr)
                + f"\nMelewati batas waktu {batas} detik.",
                durasi=time.monotonic() - mulai,
                kode=-1,
            )
        try:
            b = baris_antre.get(timeout=min(sisa, 1.0))
        except queue.Empty:
            continue
        if b is None:
            break
        potongan.append(b)
        if diam:
            continue
        b = b.strip()
        if b.startswith("{"):
            try:
                _alir(json.loads(b))
            except ValueError:
                pass

    proses.wait()
    benang.join(timeout=5)
    durasi = time.monotonic() - mulai
    keluaran = "".join(potongan)

    # `--format json` mengeluarkan JSONL per langkah. Yang dibutuhkan pipeline
    # adalah teks jawabannya, jadi JSON-nya dibongkar lebih dulu. Kalau tidak
    # ada yang cocok, keluaran mentah tetap dipakai supaya log tetap berguna.
    teks = _teks_dari_json(keluaran) or keluaran

    # `SELESAI` adalah penanda yang diminta tiap spec agent. Tanpa
    # penandanya, keluaran yang ada belum tentu utuh dan tidak boleh dipakai.
    selesai = "SELESAI" in teks.upper()

    return HasilAgent(
        perintah=perintah,
        selesai=selesai,
        stdout=teks,
        stderr="".join(kumpulan_stderr),
        durasi=durasi,
        kode=proses.returncode,
    )


def _isi_cukup(berkas: Path) -> bool:
    """True kalau berkas hasil sudah berisi, bukan sekadar kerangka."""
    try:
        return berkas.is_file() and berkas.stat().st_size >= config.MIN_ISI_HASIL
    except OSError:
        return False


def _teks_dari_json(keluaran: str) -> str:
    """Ambil teks assistant dari keluaran `--format json`.

    Keluaran opencode bukan satu objek JSON, melainkan JSONL: satu objek per
    baris, setiap langkah punya event sendiri. Teks jawaban berada di
    `part.text`, bukan di kunci atas. Parser lama hanya mencari kunci atas,
    jadi ia mengembalikan string kosong dan setiap pemanggilan agent dianggap
    gagal.

    Bentuk lain tetap ditangani sebagai cadangan, karena `--format json`
    berbeda versi bisa berbeda.
    """
    import json  # noqa: PLC0415

    potongan: list[str] = []

    for b in keluaran.splitlines():
        b = b.strip()
        if not b.startswith("{"):
            continue
        try:
            data = json.loads(b)
        except ValueError:
            continue
        if not isinstance(data, dict):
            continue

        # Bentuk JSONL opencode: {"type": "text", "part": {...,"text": ...}}
        part = data.get("part")
        if isinstance(part, dict):
            teks = part.get("text")
            if isinstance(teks, str) and teks.strip():
                potongan.append(teks)

            # Varian: blok alat butuh tool state, bukan teks jawaban.
            if not potongan and data.get("type") in ("tool", "tool_use"):
                continue

        # Cadangan bentuk datar.
        if not potongan:
            for kunci in ("text", "result", "output", "content"):
                nilai = data.get(kunci)
                if isinstance(nilai, str) and nilai.strip():
                    potongan.append(nilai)
                elif isinstance(nilai, list):
                    gabungan = "\n".join(
                        str(item.get("text", "")) if isinstance(item, dict) else str(item)
                        for item in nilai
                    ).strip()
                    if gabungan:
                        potongan.append(gabungan)

    return "\n".join(potongan).strip()


def panggil_agent_bertahap(
    *,
    nama: str,
    prompt: str,
    model: str,
    berkas: list[Path] | None = None,
    cwd: Path | None = None,
    hasil_wajib: Path | None = None,
    label: str = "",
    diam: bool = False,
) -> HasilAgent:
    """Panggil agent, ulang bila gagal atau berkas hasilnya tidak muncul.

    Kegagalan yang paling sering bukan error, melainkan model yang berhenti
    di tengah jalan. Karena itu berkas hasil ikut diperiksa: ada agent yang
    keluar dengan kode 0 tanpa menulis apa pun.

    Berkas yang sudah berisi dianggap cukup, walau model tidak mengakhiri
    keluarannya dengan `SELESAI`. Alasannya: `SELESAI` adalah penutup yang
    diminta spec, bukan syarat. Model yang pekerjaannya selesai tapi lupa
    menutupnya tidak menghasilkan pekerjaan lebih buruk kalau dijalankan lagi,
    sedangkan mengulang pemanggilan yang sudah berhasil itu mahal -- peta soal
    satu sesi memakan sekitar tiga menit. Percobaan kedua kadang juga berhenti
    lebih awal dari yang pertama dan menimpa berkas yang sudah benar dengan isi
    lebih pendek.

    Karena itu berkas yang ada dinilai cukup kalau isinya bukan sekadar
    kerangka. Ambang `config.MIN_ISI_HASIL` dipakai untuk membedakan "selesai"
    dari "sudah mulai tapi belum menghasilkan apa-apa".
    """
    percobaan = max(1, config.MAX_Coba_AGENT)
    terakhir: HasilAgent | None = None

    for ke in range(1, percobaan + 1):
        if ke > 1:
            time.sleep(3)
        hasil = panggil_agent(
            nama=nama, prompt=prompt, model=model, berkas=berkas, cwd=cwd, diam=diam
        )
        terakhir = hasil

        if hasil_wajib is None:
            lengkap = hasil.selesai
        elif hasil.selesai:
            lengkap = hasil_wajib.is_file()
        else:
            lengkap = _isi_cukup(hasil_wajib)

        if lengkap:
            if not hasil.selesai:
                print(
                    f"    ! {label or nama} tidak menutup dengan SELESAI, tapi "
                    f"{hasil_wajib.name} sudah berisi. Dipakai apa adanya.",
                    file=sys.stderr,
                )
            return hasil

        print(
            f"    ! {label or nama} percobaan {ke}/{percobaan} belum selesai "
            f"({hasil.ringkas})",
            file=sys.stderr,
        )
        if hasil.stderr.strip():
            baris = hasil.stderr.strip().splitlines()[-2:]
            for b in baris:
                print(f"      {b[:160]}", file=sys.stderr)

    assert terakhir is not None
    return terakhir


# ---------------------------------------------------------------- tahap pipeline


@dataclass
class HasilSesi:
    """Semua yang dihasilkan satu sesi."""

    matkul: MataKuliah
    sesi: Sesi
    docx: Path | None = None
    peta: Path | None = None
    jawaban: Path | None = None
    referensi: Path | None = None
    transkrip: list[Path] = field(default_factory=list)
    lampiran: list[attachments.HasilBerkas] = field(default_factory=list)
    gambar_soal: list[Path] = field(default_factory=list)
    catatan: list[str] = field(default_factory=list)
    gagal: list[str] = field(default_factory=list)


class Pipeline:
    """Satu proses untuk satu (mata kuliah, sesi)."""

    def __init__(
        self,
        matkul: MataKuliah,
        nomor_sesi: int,
        *,
        model_utama: str | None = None,
        model_mata: str | None = None,
        dengan_gambar: bool = True,
        tanpa_docx: bool = False,
    ) -> None:
        self.matkul = matkul
        self.nomor = nomor_sesi
        self.model_utama = model_utama or config.PRIMARY_MODEL
        self.model_mata = model_mata or config.VISION_MODEL
        self.dengan_gambar = dengan_gambar
        self.tanpa_docx = tanpa_docx

        self.klien = moodle.Moodle()
        self.dirs = config.work_dirs(matkul.slug, nomor_sesi)
        self.reader: reader_mod.Reader | None = None
        self.hasil = HasilSesi(matkul=matkul, sesi=Sesi(nomor_sesi, nomor_sesi))

    # -------------------------------------------------------------- utilitas

    def _siapkan(self) -> None:
        config.ensure_dirs()
        for path in self.dirs.values():
            path.mkdir(parents=True, exist_ok=True)
        if not config.URL_MOODLE or not config.COOKIE_MOODLE:
            raise PipelineGagal(
                "URL_MOODLE atau COOKIE_MOODLE belum diisi di .env. "
                "Isi keduanya lalu jalankan ulang."
            )

    def _log(self, pesan: str) -> None:
        print(pesan, flush=True)

    def _simpan_log(self, nama: str, isi: str) -> Path:
        tujuan = self.dirs["log"] / nama
        tujuan.parent.mkdir(parents=True, exist_ok=True)
        # Log agent penuh dengan URL Reader, jadi kunci aksesnya disaring dulu.
        # Log sering dibaca atau dikirim saat memperbaiki masalah, dan kunci
        # yang bocor di sana tetap hidup selama pipeline berjalan.
        tujuan.write_text(_saring_token(isi), encoding="utf-8", errors="replace")
        return tujuan

    # ------------------------------------------------------------------ run

    def jalankan(self) -> HasilSesi:
        self._siapkan()

        self._log(f"\n=== {self.matkul.label} - Sesi {self.nomor} ===")

        # 1. Cek login lebih dulu. Kegagalan di sini paling sering cookie
        #    kedaluwarsa, dan pesan error Moodle untuk kasus itu menyesatkan.
        masuk, pesan = self.klien.cek_login()
        if not masuk:
            raise PipelineGagal(pesan)
        self._log(f"  [1/8] Sesi Moodle: {pesan}")

        # 2. Petakan sesi: soal dan bahan ajar yang benar-benar milik sesi ini.
        sesi = courses.petakan_sesi(self.klien, self.matkul, self.nomor)
        self.hasil.sesi = sesi
        self.hasil.catatan.extend(sesi.catatan)
        self._log(f"  [2/8] {sesi.ringkas()}")
        for c in sesi.catatan:
            self._log(f"        {c}")

        # 3. Nyalakan Reader. Semua akses agent ke Moodle lewat sini.
        self.reader = reader_mod.Reader(self.klien)
        akar = self.reader.mulai()
        self._log(f"  [3/8] Reader lokal: {akar}")

        try:
            soal_teks, gambar_soal = self._kumpulkan_soal()
            self.hasil.gambar_soal = gambar_soal
            transkrip_soal = self._tahap_gambar_soal(gambar_soal)
            peta = self._tahap_peta_sesi(transkrip_soal, soal_teks)
            transkrip = transkrip_soal + self._tahap_bahan_ajar()
            self.hasil.transkrip = transkrip
            referensi = self._tahap_referensi()
            self._tahap_jawaban(referensi, soal_teks)
            if not self.tanpa_docx:
                self._tahap_docx(soal_teks)
        finally:
            self.reader.berhenti()

        return self.hasil

    # ------------------------------------------------------ 4. kumpulkan soal

    def _kumpulkan_soal(self) -> tuple[str, list[Path]]:
        """Baca seluruh halaman soal, hasilkan teks bersih dan gambar.

        Teksnya yang ditulis di depan jawaban pada dokumen, dan gambarnya
        disisipkan apa adanya supaya tetap sama dengan yang ada di halaman.
        """
        assert self.reader is not None
        bagian: list[str] = []
        gambar: list[Path] = []

        for a in self.hasil.sesi.soal:
            try:
                halaman = self.klien.ambil(a.url)
            except moodle.MoodleError as exc:
                self.hasil.gagal.append(f"Gagal membuka {a.nama}: {exc}")
                continue

            # Judul kosong: nama aktivitas ditulis sebagai baris biasa.
            # `tools/docx.py` sudah memasang label "Soal" sendiri, jadi
            # heading dari sini akan jadi dua judul berturut-turut.
            md = html2md.ke_markdown(
                halaman.html,
                base_url=halaman.url,
                rewrite=self.reader.ulang,
                elas=a.modtype == "assign",
                judul="",
            )
            md = _bersihkan_markah_lokal(md, self.reader.akar)
            bagian.append(f"**{a.nama}**\n\n{md}".strip())

            found = attachments.kumpulkan_gambar_soal(
                self.klien, a.url, self.dirs["berkas"] / "soal"
            )
            gambar.extend(found)
            if found:
                self._log(f"        {a.nama}: {len(found)} gambar soal diambil")

        return "\n\n---\n\n".join(b for b in bagian if b.strip()), gambar

    # ------------------------------------------------------- 5. peta soal

    def _tahap_gambar_soal(self, gambar: list[Path]) -> list[Path]:
        """Transkripsikan gambar soal dengan model penglihatan.

        Dijalankan sebelum peta soal, bukan sesudahnya, karena peta itulah yang
        harus menyatakan isi gambarnya. Kalau transkripnya belum ada, agent
        pemetaan hanya melihat gambar biner lewat Reader dan mencatat soalnya
        "belum dibaca" -- padahal gambarnya bisa dibaca dengan model penglihatan.
        """
        if not gambar:
            self._log("  Tidak ada gambar soal untuk ditranskripsikan")
            return []

        transkrip: list[Path] = []
        for nomor, berkas in enumerate(gambar, 1):
            t = self._tahap_transkrip_soal(berkas, nomor)
            if t:
                transkrip.append(t)

        if transkrip:
            self._log(f"  {len(transkrip)}/{len(gambar)} gambar soal ditranskripsikan")
        else:
            self.hasil.gagal.append(
                "Gambar soal gagal ditranskripsikan, soalnya kemungkinan "
                "hanya ada di dalam gambar. Lihat _log/vision-soal-*.log."
            )
        return transkrip

    def _tahap_peta_sesi(
        self, transkrip_soal: list[Path], soal_teks: str = ""
    ) -> Path | None:
        assert self.reader is not None
        tujuan = self.dirs["petak"] / f"sesi{self.nomor}.md"
        if tujuan.is_file():
            self._log(f"  [4/8] Peta soal sudah ada: {tujuan.name}")
            self.hasil.peta = tujuan
            return tujuan

        urls_soal = "\n".join(
            f"  - {a.jenis_soal} {a.nama}: {self.reader.soal(a.url, a.nama)}"
            for a in self.hasil.sesi.soal
        ) or "  - (tidak ada soal terdeteksi di section ini)"

        urls_materi = "\n".join(
            f"  - {a.nama}: {self.reader.materi(a.url, a.nama)}"
            for a in self.hasil.sesi.materi
        ) or "  - (tidak ada bahan ajar terdeteksi)"

        blok_gambar = _potong_soal_dari_gambar(transkrip_soal)
        blok_soal = _potong_teks(soal_teks, config.BATAS_TEKS_SOAL)

        prompt = f"""Kamu adalah pemetaan soal untuk SATU sesi mata kuliah.

Mata kuliah : {self.matkul.nama} ({self.matkul.kode or 'kode tidak disebut'})
Kelas       : {self.matkul.kelas or 'tidak disebut'}
Sesi        : {self.nomor}
Halaman section : {self.klien.base_url}/course/view.php?id={self.matkul.id}&section={self.nomor}

Batas eksplorasi untuk sesi ini: {config.BATAS_HALAMAN_SESI} halaman.

Tulis peta soal ke: {tujuan}

## Teks soal (sudah dibaca pipeline, inilah isinya)
Pipeline ini sudah membuka setiap halaman soal di bawah dan mengubahnya jadi
teks. Teks di bawah inilah rumusan soal yang resmi, jadi JANGAN buka URL soal
untuk membaca soalnya -- membukanya hanya membuang anggaran eksplorasimu dan
kadang justru mengembalikan halaman kosong.

{blok_soal}

## Soal di sesi ini (buka hanya untuk lampiran atau detail yang belum ada di teks)
{urls_soal}

## Bahan ajar di sesi ini (buka semuanya)
{urls_materi}

Semua URL di atas sudah berupa URL Reader lokal yang menyuntikkan sesi
Moodle. Buka dengan webfetch apa adanya. Tautan di dalam halaman hasil
webfetch juga sudah berupa URL Reader dan bisa diikuti langsung. Jangan
menyalin ulang URL itu secara manual -- kalau perlu, salin utuh dari blok di
atas, karena satu karakter base64 yang keliru membuat Reader menolaknya.

{blok_gambar}
Kerjakan sesuai spec-mu: tulis bagian `## Bahan ajar wajib sesi ini` dari teks
halaman sesi, lalu tiap soal beserta format jawaban, rubrik, dan lampirannya.
Jangan menempel isi jawaban mahasiswa lain ke peta mana pun.
"""

        hasil = panggil_agent_bertahap(
            nama=config.AGENT_SCRAPPER,
            prompt=prompt,
            model=self.model_utama,
            cwd=config.BASE_DIR,
            hasil_wajib=tujuan,
            label="peta soal",
        )
        self._simpan_log("scrapper.log", hasil.stdout + "\n\n--- STDERR ---\n" + hasil.stderr)

        if not tujuan.is_file():
            self.hasil.gagal.append(
                "Peta soal tidak terbentuk. Lihat _log/scrapper.log."
            )
            self._log(f"        ! peta soal gagal ({hasil.ringkas})")
            return None

        self.hasil.peta = tujuan
        self._log(f"  [4/8] Peta soal: {tujuan.stat().st_size} byte")
        return tujuan

    # -------------------------------------------------- 6. bahan ajar + transkrip

    def _tahap_bahan_ajar(self) -> list[Path]:
        """Unduh lampiran, lalu transkripsikan gambarnya lewat model vision.

        Bahan ajar Universitas Terbuka umumnya hasil ekspor PowerPoint ke
        PDF: teksnya sedikit dan sebagian besar isi slide berada di gambar.
        Membaca PDF sebagai teks menghasilkan daftar judul tanpa isi, jadi
        halaman yang sudah dirender harus dibaca model penglihatan.
        """
        pasangan = [
            (a.url, a.nama)
            for a in self.hasil.sesi.materi
            if a.modtype == "resource" and a.url
        ]
        if not pasangan:
            self._log("  [5/8] Tidak ada lampiran bahan ajar untuk diunduh")
            return []

        lampiran = attachments.kumpulkan_lampiran(
            self.klien, pasangan, self.dirs["bahan"],
            with_gambar=self.dengan_gambar,
        )
        self.hasil.lampiran = lampiran

        for h in lampiran:
            if h.gagal:
                self._log(f"        ! {h.nama} gagal: {' '.join(h.catatan)[:90]}")
                continue
            self._log(f"        {h.ringkas()}")

        transkrip: list[Path] = []
        for h in lampiran:
            if not h.gambar:
                continue
            grup = self._grup_gambar(h)
            for berkas in grup:
                t = self._tahap_transkrip(h, berkas)
                if t:
                    transkrip.append(t)

        if not lampiran:
            self._log("  [5/8] Semua lampiran gagal diunduh")
            return []
        self._log(
            f"  [5/8] {len(lampiran)} lampiran, {len(transkrip)} transkrip gambar"
        )
        return transkrip

    def _grup_gambar(self, h: attachments.HasilBerkas) -> list[Path]:
        """Render halaman PDF yang perlu dibaca, lalu gabung beberapa per berkas.

        Flag `--file` milik `opencode run` hanya menerima satu berkas, jadi
        montage diperlukan: tanpa itu, satu PDF 46 halaman jadi 46
        panggilan vision.
        """
        if not self.dengan_gambar:
            return []

        folder = h.lokasi.parent if h.lokasi else self.dirs["bahan"]
        tujuan = folder / f"montase-{h.nama}"
        grup = attachments.gabung_halaman(h.gambar, tujuan)

        # mount list dibatasi supaya biaya vision tetap terkendali. Sisa
        # halaman dicatat, bukan dihapus diam-diam.
        maks = config.BATAS_MONTASE_PER_LAMPIRAN
        if len(grup) > maks:
            grup = grup[:maks]
            self.hasil.catatan.append(
                f"{h.nama}: hanya {maks} dari {len(h.gambar)} halaman "
                "dikirim ke model penglihatan."
            )
        return grup

    def _tahap_transkrip(self, h: attachments.HasilBerkas, gambar: Path) -> Path | None:
        """Transkripsikan satu berkas gambar dengan agent vision."""
        return self._transkrip_gambar(
            gambar,
            slug=config.slugify(h.nama, "x"),
            judul=f"# Transkrip lampiran: {h.nama}",
            konteks=(
                f"URL Reader untuk halaman lampirannya (cadangan, boleh dibuka "
                f"dengan webfetch kalau perlu konteks tambahan):\n"
                f"  {self.reader.berkas(h.url, h.nama)}\n"
            ) if self.reader is not None else "",
        )

    def _tahap_transkrip_soal(self, gambar: Path, nomor: int) -> Path | None:
        """Transkripsikan gambar yang menempel pada halaman soal.

        Banyak soal UT tidak tertulis sebagai teks melainkan diketik sebagai
        gambar di dalam post forum. Gambarnya sudah diunduh untuk
        disisipkan ke dokumen, tapi tanpa transkripsi ada dua hal yang tidak
        bisa dikerjakan: peta soal akan mencatat soalnya "belum dibaca", dan
        dokumen keluarannya memuat gambar tanpa keterangan. Transkripsi inilah
        yang menutup kedua kekosongan itu.
        """
        return self._transkrip_gambar(
            gambar,
            slug=f"soal-{nomor:02d}",
            judul=f"# Transkripsi gambar soal {nomor}",
            konteks="",
            gabung_pilihan=False,
        )

    def _transkrip_gambar(
        self,
        gambar: Path,
        *,
        slug: str,
        judul: str,
        konteks: str,
        gabung_pilihan: bool = True,
    ) -> Path | None:
        """Transkripsikan satu berkas gambar dengan agent vision."""
        tujuan = self.dirs["bahan"] / f"transkrip-{slug}.md"
        stem = gambar.stem
        montage = gabung_pilihan and stem.startswith("montase")

        if montage:
            # Montase memuat beberapa halaman sekaligus, jadi transkripnya
            # ditambahkan ke berkas yang sama, bukan menimpanya.
            sementara = self.dirs["bahan"] / f".tmp-{stem}.md"
            sementara.write_text(
                f"{judul}\n\n"
                f"Berkas gambar: {gambar.name} (montase beberapa halaman)\n",
                encoding="utf-8",
            )
            file_tujuan = sementara
            gabung = True
        else:
            file_tujuan = tujuan
            gabung = False
            file_tujuan.write_text(
                f"{judul}\n\n"
                f"Berkas gambar: {gambar.name}\n",
                encoding="utf-8",
            )

        prompt = (
            f"Transkripsikan isi gambar lampiran berikut ke: {file_tujuan}\n\n"
            f"Nama lampiran: {judul.lstrip('# ')}\n"
            f"{konteks}\n"
            "Hanya menyalin teks yang tertulis. Persahankan angka, simbol, "
            "dan notasi matematika. Jangan menjawab, jangan meringkas."
        )
        if gabung:
            prompt += (
                "\n\nGambar ini montage beberapa halaman. Transkripsikan "
                "SEMUA halaman yang terlihat, pisahkan dengan baris "
                "`--- Halaman berikutnya ---`."
            )

        hasil = panggil_agent_bertahap(
            nama=config.AGENT_VISION,
            prompt=prompt,
            model=self.model_mata,
            berkas=[gambar],
            cwd=config.BASE_DIR,
            hasil_wajib=file_tujuan,
            label="transkrip",
        )
        self._simpan_log(
            f"vision-{stem}.log", hasil.stdout + "\n\n--- STDERR ---\n" + hasil.stderr
        )

        if not file_tujuan.is_file():
            return None

        if gabung:
            with tujuan.open("a", encoding="utf-8") as fh:
                fh.write("\n")
                fh.write(file_tujuan.read_text(encoding="utf-8"))
            file_tujuan.unlink(missing_ok=True)
            return tujuan
        return file_tujuan

    # ------------------------------------------------------ 7. daftar pustaka

    def _tahap_referensi(self) -> Path | None:
        """Siapkan daftar pustaka untuk sesi ini.

        Daftar pustaka selalu diminta ke agent `research`, walaupun peta soal
        sudah menyebut buku resmi. Alasannya bukan mengabaikan spec research,
        yang menyuruh pipeline menulis sendiri kalau bahasan sudah
        terverifikasi: buku UT umumnya terbit lebih dari sepuluh tahun lalu,
        sedangkan jawaban yang dikumpulkan harus menyertakan sumber yang
        bisa diperiksa dan terbit dalam sepuluh tahun terakhir. Satu daftar
        harus memuat keduanya, dan hanya agent yang bisa mencari sumber
        kedua itu.
        """
        tujuan = self.dirs["jawaban"] / f"referensi_sesi_{self.nomor}.md"
        if tujuan.is_file() and tujuan.stat().st_size > 40:
            self._log(f"  [6/8] Daftar pustaka sudah ada: {tujuan.name}")
            return tujuan

        if self.hasil.peta is None:
            # Peta gagal: jangan pernah mencari referensi untuk soal yang
            # tidak diketahui. Agent akan mengarang.
            self.hasil.gagal.append(
                "Daftar pustaka tidak dicari karena peta soal tidak terbentuk."
            )
            return None

        if not self.hasil.peta.is_file():
            self.hasil.gagal.append("Peta soal hilang sebelum dicari referensinya.")
            return None

        isi_peta = self.hasil.peta.read_text(encoding="utf-8", errors="replace")
        bagian_bahan = _potong_bagian(isi_peta, "## Bahan ajar wajib sesi ini")
        soal_ringkas = _potong_soal_dari_peta(isi_peta)

        prompt = f"""Cari daftar pustaka untuk satu soal mata kuliah.

Mata kuliah : {self.matkul.nama} ({self.matkul.kode or 'kode tidak disebut'})
Sesi        : {self.nomor}

Batas yang berlaku untukmu:
- MAKSIMAL {config.MAX_REFERENSI} referensi.
- Hanya terbitan {config.TAHUN_MIN} atau setelahnya.
- Setiap entri harus punya URL atau DOI yang benar-benar bisa dibuka. Ambil
  halaman katalog atau repository untuk memastikan, jangan menuliskan URL
  dari ingatan.

## Bahan ajar wajib sesi ini
{bagian_bahan or '(tidak ada di peta soal)'}

## Ringkasan soal
{soal_ringkas or '(tidak ada di peta soal)'}

Tulis daftar APA 7 ke: {tujuan}
Satu entri per baris, tanpa nomor, tanpa bullet, tanpa ISBN atau jumlah
halaman yang tidak kamu lihat sendiri di katalog resmi.
"""
        hasil = panggil_agent_bertahap(
            nama=config.AGENT_RESEARCH,
            prompt=prompt,
            model=self.model_utama,
            cwd=config.BASE_DIR,
            hasil_wajib=tujuan,
            label="daftar pustaka",
        )
        self._simpan_log(
            "research.log", hasil.stdout + "\n\n--- STDERR ---\n" + hasil.stderr
        )

        if not tujuan.is_file():
            self.hasil.gagal.append(
                "Daftar pustaka tidak terbentuk. Lihat _log/research.log."
            )
            self._log(f"        ! daftar pustaka gagal ({hasil.ringkas})")
            return None

        self.hasil.referensi = tujuan
        self._log(f"  [6/8] Daftar pustaka: {tujuan.stat().st_size} byte")
        return tujuan

    # ------------------------------------------------------- 8. tulis jawaban

    def _tahap_jawaban(
        self, referensi: Path | None, soal_teks: str = ""
    ) -> None:
        assert self.reader is not None
        tujuan = self.dirs["jawaban"] / "jawaban.md"

        if self.hasil.peta is None:
            raise PipelineGagal(
                "Peta soal tidak terbentuk, jadi jawaban tidak bisa ditulis. "
                "Lihat _log/scrapper.log untuk penyebabnya."
            )

        peta = self.hasil.peta
        daftar_referensi = referensi if referensi else None

        transkrip_teks = []
        hilang: list[str] = []
        for t in self.hasil.transkrip:
            # Daftar transkrip dibuat beberapa tahap lalu dibaca lagi di sini.
            # Kalau salah satunya hilang di antara waktu itu -- misalnya karena
            # proses lain membersihkan folder, atau ada yang menggeser berkasnya
            # -- pipeline harus tetap jalan dan memberitahu, bukan berhenti dengan
            # `FileNotFoundError` setelah semua agent lain selesai.
            if not t.is_file():
                hilang.append(t.name)
                continue
            isi = t.read_text(encoding="utf-8", errors="replace")
            if len(isi.strip()) > 30:
                transkrip_teks.append(f"### {t.name}\n{isi.strip()}")

        if hilang:
            self.hasil.gagal.append(
                f"Transkrip hilang sebelum dipakai: {', '.join(hilang)}. "
                "Jalankan ulang sesi ini bila bahan ajarnya memang dibutuhkan."
            )
            self._log(f"        ! {len(hilang)} transkrip hilang: {', '.join(hilang)}")

        blok_transkrip = "\n\n".join(transkrip_teks) or "(tidak ada transkrip gambar)"

        lampiran_teks = []
        for h in self.hasil.lampiran:
            if h.gagal:
                continue
            bagian = [f"### {h.nama}"]
            if h.lokasi:
                bagian.append(f"Berkas lokal: {h.lokasi}")
            if h.teks.strip():
                bagian.append("Teks hasil ekstraksi:\n" + h.teks.strip())
            else:
                bagian.append("(tidak ada teks yang bisa diekstrak)")
            lampiran_teks.append("\n".join(bagian))
        blok_lampiran = "\n\n".join(lampiran_teks) or "(tidak ada lampiran)"

        urls_cadangan = "\n".join(
            f"  - {a.jenis_soal} {a.nama}: {self.reader.soal(a.url, a.nama)}"
            for a in self.hasil.sesi.soal
        ) or "  - (tidak ada)"

        blok_soal = _potong_teks(soal_teks, config.BATAS_TEKS_SOAL)

        prompt = f"""Tulis jawaban tutorial online untuk satu sesi.

Mata kuliah : {self.matkul.nama} ({self.matkul.kode or 'kode tidak disebut'})
Kelas       : {self.matkul.kelas or 'tidak disebut'}
Sesi        : {self.nomor}

## Rumusan soal (sudah dibaca pipeline, inilah naskah resminya)
Teks di bawah inilah soal yang harus kamu jawab. Penomoran, urutan butir, dan
semua syarat format diambil dari sini. Jangan membukanya ulang lewat URL Reader
kecuali ada bagian yang benar-benar tidak ada di teks ini.

{blok_soal}

## Peta soal (WAJIB dibaca lengkap)
File: {peta}

## Daftar pustaka final (salin persis ke bagian Daftar Pustaka)
File: {daftar_referensi if daftar_referensi else '(tidak ada; tulis entri yang benar-benar kamu verifikasi)'}

## Transkrip lampiran
Berkas `transkrip-soal-*.md` adalah hasil bacaan gambar soal resmi, jadi
isi file itu ADALAH rumusan soal. Sisanya transkrip halaman bahan ajar.
{blok_transkrip}

## Lampiran dan teks bahan ajar
{blok_lampiran}

## URL Reader cadangan (hanya kalau peta tidak cukup)
{urls_cadangan}

Tulis jawaban ke: {tujuan}

Struktur jawaban yang diminta dokumen akhir:

- `## Jawaban` sebagai heading utama isi jawaban.

- **Soal ditulis ulang lebih dulu, baru dijawab.** Setiap nomor soal memakai
  heading `###` yang memuat nomor dan rumusan soalnya, misalnya
  `### 1. Tentukan 2A, AB, dan 2A + AB dari matriks A dan B berikut`.
  Rumusan itu diambil dari peta soal atau transkrip gambar resmi, bukan
  dicari ulang dari ingatan. Tulis ulang butir-butir dalam soal apa adanya
  tepat di bawah heading-nya, dengan penomoran yang sama seperti di soal.

- **Setiap butir dalam soal wajib dijawab sendiri.** Kalau soal nomor 1 punya
  tiga titik yang diminta dihitung, buat tiga subbagian bernomor dengan
  penomoran yang mengikuti butir aslinya (`#### 1.1`, `#### 1.2`, `#### 1.3`
  untuk butir `a`, `b`, `c`; atau `#### 1.a` dan seterusnya kalau soalnya
  memakai huruf). Jangan menggabungkan tiga butir itu ke dalam satu
  paragraf, dan jangan memecahnya menjadi subbagian yang tidak ada di soal.

- Penjelasan tambahan, catatan, atau pembuktian ditulis sebagai subbagian
  dengan nomor lanjutan yang jelas (`#### 1.4`) atau sebagai bagian tersendiri
  di bawah semua nomor (`### Penutup`). Penomoran di jawaban harus selalu bisa
  dilacak ke nomor butir di soal.

- `## Daftar Pustaka` di bagian paling akhir, isinya disalin utuh dari
  berkas daftar pustaka di atas.
- Persamaan ditulis dengan LaTeX memakai pemisah dolar: `$...$` untuk
  sebaris dan `$$` di baris sendiri untuk blok. Bentuk `\\(...\\)` dan
  `\\[...\\]` jangan dipakai karena renderer dokumen hanya mengenali
  bentuk dolar. Matriks memakai notasi `[[a, b], [c, d]]`.
- Tulis sebagai mahasiswa yang mengerjakan tugas, pakai kata "saya".
- Pakai skill `humanizer` dalam mode File pada berkas jawaban ini.
"""

        hasil = panggil_agent_bertahap(
            nama=config.AGENT_WORKER,
            prompt=prompt,
            model=self.model_utama,
            cwd=config.BASE_DIR,
            hasil_wajib=tujuan,
            label="jawaban",
        )
        self._simpan_log(
            "worker.log", hasil.stdout + "\n\n--- STDERR ---\n" + hasil.stderr
        )

        if not tujuan.is_file():
            raise PipelineGagal(
                "Jawaban tidak tertulis. Lihat _log/worker.log untuk penyebabnya."
            )

        self.hasil.jawaban = tujuan
        self._log(f"  [7/8] Jawaban: {tujuan.stat().st_size} byte")

    # -------------------------------------------------------------- 9. docx

    def _tahap_docx(self, soal_teks: str) -> None:
        if self.hasil.jawaban is None:
            raise PipelineGagal("Jawaban belum ada, dokumen tidak bisa dibuat.")

        jawaban = self.hasil.jawaban.read_text(encoding="utf-8", errors="replace")
        if not jawaban.strip():
            raise PipelineGagal("Berkas jawaban kosong, dokumen tidak dibuat.")

        # Identitas mahasiswa dibaca ulang dari `.env` saat render, bukan
        # saat pipeline mulai, supaya perbaikan `.env` tanpa mengulang
        # pemanggilan agent langsung terpakai.
        ident = config.identitas()
        kosong = [k for k, v in ident.items() if not v]
        if kosong:
            self.hasil.catatan.append(
                "Isi " + ", ".join(k.upper() for k in kosong) + " di .env "
                "agar baris tabel pada dokumen ikut terisi."
            )

        jenis = sorted({a.jenis_soal for a in self.hasil.sesi.soal})
        kind_label = jenis[0] if len(jenis) == 1 else "Tutor Online"
        nomor_label = str(self.nomor)
        if len(self.hasil.sesi.soal) > 1:
            nomor_label = f"{self.nomor} ({len(self.hasil.sesi.soal)} soal)"

        meta = {
            "file_base": f"{self.matkul.slug}-sesi-{self.nomor}",
            "kind_label": kind_label,
            "display_index": nomor_label,
            "matkul": self.matkul.nama,
            "matkul_kode": self.matkul.kode,
            **ident,
        }

        # Dimuat lewat jalur berkas, bukan `import tools.docx`. Folder `tools`
        # tidak punya `__init__.py`, jadi `from config import TEMPLATE_DIR` di
        # dalamnya hanya berhasil kalau root proyek ada di `sys.path`.
        sys.path.insert(0, str(config.BASE_DIR))
        spec = importlib.util.spec_from_file_location(
            "jkt_docx", config.TOOLS_DIR / "docx.py"
        )
        if spec is None or spec.loader is None:
            raise PipelineGagal("tools/docx.py tidak bisa dimuat.")
        docx_tool = importlib.util.module_from_spec(spec)
        sys.modules["jkt_docx"] = docx_tool
        spec.loader.exec_module(docx_tool)

        template = None
        if config.TEMPLATE_DIR.is_dir():
            kandidat = sorted(config.TEMPLATE_DIR.glob("*.docx"))
            template = kandidat[0] if kandidat else None

        out = config.output_dir(self.matkul.slug, self.nomor)
        out.mkdir(parents=True, exist_ok=True)
        docx_path, _ = docx_tool.save_doc(
            jawaban_md=jawaban,
            soal_text=soal_teks,
            meta=meta,
            out_dir=out,
            template=template,
            include_soal=True,
            gambar_soal=self.hasil.gambar_soal or None,
        )
        self.hasil.docx = docx_path
        # `_simpan_aman` di `tools/docx.py` berpindah ke nama bernomor kalau
        # berkas tujuan sedang terkunci Word/OneDrive. Kalau itu terjadi, nama
        # yang disebut di sini harus yang benar-benar ditulis, dan pengguna
        # perlu tahu berkas lamanya tidak ikut ter-update.
        if docx_path.name != f"{meta['file_base']}.docx":
            self._log(
                f"  ! {meta['file_base']}.docx sedang dipakai program lain "
                f"(tutup di Word/penjelajah berkas), dokumen ditulis sebagai "
                f"{docx_path.name}"
            )
        self._log(f"  [8/8] Dokumen: {docx_path}")


# ---------------------------------------------------------------- pemdongsoal


def _bersihkan_markah_lokal(teks: str, akar_reader: str) -> str:
    """Buang markah gambar dan alamat Reader dari teks soal.

    Dua masalah yang keduanya nyata di dokumen:

    1. **Token bocor.** `html2md` menulis ulang tautan Moodle menjadi URL
       Reader lokal, dan URL itu memuat token sesi_reader (`?u=...&k=...`).
       Markdown-nya ikut masuk ke dokumen apa adanya, jadi berkas yang
       diserahkan ke tutor ikut memuat token yang masih hidup selama
       pipeline berjalan. Token itu bukan hal yang boleh seenaknya: siapa pun
       yang memegangnya bisa membuka setiap halaman akun tersebut selama
       Reader masih hidup.

    2. **Gambar dobel.** `tools/docx.py` sudah menyisipkan berkas gambarnya
       sendiri lewat `gambar_soal`. Kalau teksnya juga masih memuat
       `![alt](...)`, dokumen memuat gambar dua kali, dan yang kedua tetap
       berbentuk teks mentah karena markdown gambar tidak dirender.

    Teks tautan biasa tetap dipertahankan; hanya alamatnya yang dibuang, karena
    label tautan sering berisi instruksi soal yang penting.
    """
    if not teks:
        return teks

    host = re.escape((akar_reader or "").rstrip("/"))

    # Markdown gambar dibuang seluruhnya: `tools/docx.py` menyisipkan berkas
    # gambarnya sendiri lewat `gambar_soal`, jadi menyisakan markahnya hanya
    # menambah gambar kedua yang tidak pernah dirender.
    teks = re.sub(r"!\[[^\]]*\]\([^)]*\)", "", teks)

    # Tautan ke Reader `[label](http://127.0.0.1:PORT/...)` jadi `label`.
    if host:
        teks = re.sub(r"\]\((?:" + host + r")/[^)]*\)", "]", teks)
        # Alamat Reader yang berdiri sendiri, tanpa tanda kurung markdown.
        teks = re.sub(re.escape((akar_reader or "").rstrip("/")) + r"/\S*", "", teks)

    # Baris yang tinggal kosong dibuang, termasuk sisa gambar.
    teks = re.sub(r"[ \t]+\n", "\n", teks)
    teks = re.sub(r"\n{3,}", "\n\n", teks)
    return teks.strip()


def _potong_bagian(isi: str, judul: str) -> str:
    """Ambil satu bagian berheading dari peta soal."""
    pola = re.compile(
        r"^#{1,6}\s*" + re.escape(judul) + r"\s*$(.*?)(?=^#{1,6}\s|\Z)",
        re.I | re.M | re.S,
    )
    m = pola.search(isi or "")
    return m.group(1).strip() if m else ""


def _potong_soal_dari_peta(isi: str, batas: int = 2500) -> str:
    """Ringkas peta soal jadi potongan yang cukup untuk mencari referensi."""
    bagian = re.sub(r"^\s*[-*]\s*Halaman:.*$", "", isi or "", flags=re.M)
    bagian = re.sub(r"^#{1,6}\s*Soal yang ditanyakan:\s*$", "", bagian, flags=re.M)
    return bagian.strip()[:batas]


def _potong_teks(teks: str, batas: int) -> str:
    """Potong teks panjang dengan penanda yang jelas di bagian yang dibuang.

    Pemotongan di tengah kalimat membuat agent mengira soalnya memang berhenti
    di situ, lalu menyimpulkan soalnya tidak lengkap. Karena itu bagian yang
    dibuang ditandai, bukan dipangkas diam-diam.
    """
    bersih = (teks or "").strip()
    if not bersih:
        return "(tidak ada teks soal; soal kemungkinan ada di dalam gambar)"
    if len(bersih) <= batas:
        return bersih
    return (
        bersih[:batas].rstrip()
        + "\n\n[... teks soal dipotong oleh pipeline di sini karena terlalu "
        "panjang. Buka URL soal kalau ada bagian yang benar-benar perlu ...]"
    )


def _potong_soal_dari_gambar(transkrip: list[Path], batas: int = 2500) -> str:
    """Blok transkripsi gambar soal untuk disisipkan ke prompt pemetaan.

    Soal UT sering diketik sebagai gambar di dalam post forum. Tanpa blok ini,
    pemetaan hanya melihat PNG biner lewat Reader dan mencatat soalnya
    "belum dibaca", padahal isinya sudah dibaca model penglihatan.
    """
    if not transkrip:
        return ""
    potongan: list[str] = []
    for t in transkrip:
        if not t.is_file():
            continue
        isi = t.read_text(encoding="utf-8", errors="replace").strip()
        if len(isi) > 20:
            potongan.append(isi[:batas])
    if not potongan:
        return ""
    return (
        "## Isi gambar soal (SUDAH DIBACA model penglihatan, pakai ini)\n\n"
        + "\n\n".join(potongan)
        + "\n\nGambar-gambar di atas adalah bagian dari soal resmi sesi ini.\n"
        + "Jadikan rumusan dan instruksi format di dalamnya sebagai isi soal,\n"
        + "bukan menulis `belum dibaca`.\n"
    )