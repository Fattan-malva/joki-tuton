# joki-tuton

Pipeline otomatis untuk mengerjakan soal tutorial online Universitas Terbuka
dan menghasilkan dokumen jawabannya (`.docx`).

Setiap sesi: ambil soal dari Moodle, tulis jawaban dengan agent AI, ambil
bahan ajar resmi sebagai referensi pertama, format persamaan jadi persamaan
Word asli, lalu rapikan Daftar Pustaka (penulis **tebal**, judul *miring*,
hanging indent).

---

## Persyaratan

- Python 3.11+ (`py` launcher di Windows, atau `python` langsung)
- [OpenCode v2](https://opencode.ai) terinstal dan perintah `opencode` ada di
  PATH — pipeline memanggil `opencode run` untuk semua tahap agent
- Akun Moodle UT (elearning.ut.ac.id) dengan cookie login yang masih hidup
- Paket Python: lihat `requirements.txt` (kalau belum ada, install dengan
  `pip install python-docx pdfplumber pdf2image pymupdf beautifulsoup4 lxml
  requests`)

## Instalasi OpenCode v2

1. Install dari situs resmi: `https://opencode.ai` atau jalankan:

   ```powershell
   npm install -g opencode-ai
   ```

2. Pastikan perintah `opencode` bisa dipanggil dari terminal mana pun:

   ```powershell
   opencode --version
   ```

   Kalau tidak dikenali, tambahkan folder global npm ke `PATH`
   (biasanya `%APPDATA%\npm`).

3. Susun `~/.config/opencode/opencode.json` minimal:

   ```json
   {
     "provider": {
       "opencode": {
         "models": {
           "zhipu/glm-4.6": {},
           "openai/gpt-5-mini": {}
         }
       }
     }
   }
   ```

   Pipeline memakai model utama untuk pemetaan/penjawaban/riset, dan model
   penglihatan untuk transkripsi gambar soal (sesuaikan dengan yang tersedia).
   Lihat `config.PRIMARY_MODEL` dan `config.VISION_MODEL` kalau perlu ganti.

## Setup project

```powershell
cd E:\DEVOPS\AI\joki-tuton
py -m pip install python-docx pdfplumber pdf2image pymupdf beautifulsoup4 lxml requests
copy .env.example .env   # kalau .env belum ada; kalau sudah, lewati
```

Isi `.env`:

```
URL_MOODLE=https://elearning.ut.ac.id
COOKIE_MOODLE=<cookie dari browser setelah login manual>
UT_NAMA=...
UT_NIM=...
UT_SEMESTER=...
UT_PRODI=...
```

Cookie didapat dari browser (DevTools → Application → Cookies untuk
`elearning.ut.ac.id`) setelah Anda login manual. Pipeline tidak meminta
password.

## Menjalankan

**Interaktif (pilih mata kuliah di menu):**

```powershell
py .\main.py
```

atau dobel klik `runner.bat`.

**Semua mata kuliah, satu sesi yang sama:**

```powershell
py .\main.py --semua --sesi 4
```

**Satu mata kuliah, satu sesi:**

```powershell
py .\main.py --matkul "Aljabar Linear Elementer" --sesi 4
```

**Lain-lain:**

```powershell
py .\main.py --daftar           # tampilkan semua course + sesi yang punya soal
py .\main.py --tanpa-docx       # hasilkan jawaban.md saja
py .\main.py --tanpa-gambar     # lewati transkripsi gambar oleh model vision
runner.bat --semua --sesi 4    # sama seperti di atas via batch
```

## Hasil

- Jawaban: `output/<slug>/sesi-<N>/<slug>-sesi-<N>.docx`
- Kerja: `_kerja/<slug>/sesi-<N>/` (peta soal, transkrip, lampiran, log)

Log tiap tahap ada di `_kerja/.../_log/`. Kalau satu course gagal di mode
`--semua`, yang berikutnya tetap berjalan dan ringkasan akhir menampilkannya.

## Struktur

| Folder/berkas      | Fungsi                                            |
|--------------------|---------------------------------------------------|
| `main.py`          | Antarmuka baris perintah                          |
| `pipeline.py`      | Orkestrasi 8 tahap: login → soal → peta → lampiran → ref → jawaban → docx |
| `config.py`        | Konstanta dan path                                |
| `courses.py`       | Pemetaan course/section/soal dari Moodle          |
| `reader.py`        | Proxy lokal untuk agent (semua URL Moodle→127.0.0.1) |
| `attachments.py`   | Unduh lampiran, render PDF→gambar, kumpulkan gambar soal |
| `tools/docx.py`    | Renderer DOCX: equation Word, daftar pustaka, template |
| `.opencode/agents/`| Spec agent: scrapper, worker, research, vision     |
| `runner.bat`       | Shortcut Windows                                  |

## Catatan keamanan

- Semua HTTP ke Moodle lewat proxy lokal `reader.py`; tidak ada cookie atau
  token yang ditulis ke berkas yang diserahkan ke tutor.
- Token Reader (`?k=`) disaring dari log dan peta soal.

## Kalau gagal

Ikuti urutan ini:

1. `_kerja/<slug>/sesi-<N>/_log/*.log` — cari baris `= langkah selesai` terakhir.
2. Kalau peta soal gagal: lihat `scrapper.log` — kemungkinan forum-nya
   lembar kosong (cek apakah soalnya di resource/lampiran forum).
3. Kalau persamaan di Word merah/mentah: jalankan ulang hanya tahap docx
   dengan menghapus `output/.../file.docx`.
4. Paksa pengerjaan ulang satu sesi: hapus `_kerja/<slug>/sesi-<N>/_jawaban/jawaban.md`
   dan berkas `.docx` tujuan.
