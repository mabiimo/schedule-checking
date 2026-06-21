# Cek Plan vs Actual — Versi Vercel

Versi aplikasi ini sudah direstrukturisasi agar bisa di-deploy ke
**Vercel (gratis, domain `*.vercel.app`)** sebagai Serverless Function,
menggantikan versi Docker/FastAPI penuh.

## Yang Berubah dari Versi Docker

| | Versi Docker | Versi Vercel |
|---|---|---|
| Server | FastAPI + Uvicorn (selalu hidup) | FastAPI sebagai Serverless Function (hidup saat dipanggil) |
| Halaman utama | Jinja2 Template (`templates/index.html`) | File statis (`public/index.html`) |
| CSS/JS | `/static/...` | `/...` (folder `public/` diserve langsung di root) |
| Struktur folder | `app/app/main.py` | `api/index.py` (konvensi wajib Vercel) |

Logika pencocokan data (`checker.py`) **tidak ada yang diubah** — hanya dipindah lokasinya.

## ⚠️ Batasan Penting di Free Tier (Hobby Plan)

Sebelum deploy, perlu Anda ketahui:

1. **Timeout 10 detik per request.** Kalau file Excel yang diupload sangat besar
   (ribuan baris, banyak brand), proses generate Excel/PPT bisa melebihi 10 detik
   dan akan gagal dengan error 504. Untuk file dengan puluhan-ratusan baris
   (seperti contoh kasus Anda), ini aman.
2. **Tidak boleh untuk komersial.** Hobby plan Vercel khusus untuk proyek pribadi/non-komersial.
3. **Ukuran bundle.** Total dependency (~80-90MB) masih jauh di bawah limit 500MB,
   jadi aman dari sisi ukuran.
4. **Tidak ada state/database** — sesuai kebutuhan Anda, aplikasi ini memang
   stateless (upload → proses → download), jadi cocok untuk serverless.
5. **Cold start.** Karena function "tidur" saat tidak dipakai, request pertama
   setelah idle lama mungkin terasa sedikit lebih lambat (beberapa detik) saat
   memuat ulang library pandas/openpyxl.

## Cara Deploy

### Opsi A — Lewat Dashboard Vercel (paling mudah, tanpa CLI)

1. Push folder ini ke repository GitHub (boleh public atau private).
2. Buka [vercel.com](https://vercel.com), login/daftar pakai akun GitHub.
3. Klik **"Add New" → "Project"**, pilih repository yang berisi folder ini.
4. Saat konfigurasi project:
   - **Framework Preset**: pilih "Other" (Vercel akan otomatis deteksi FastAPI dari `api/index.py`).
   - **Root Directory**: biarkan default (kosong) jika folder ini ada di root repo.
   - Tidak perlu isi Build Command / Output Directory secara manual.
5. Klik **Deploy**. Tunggu 1-3 menit.
6. Setelah selesai, Anda akan mendapat URL seperti `https://nama-project-anda.vercel.app`.

### Opsi B — Lewat Vercel CLI

```bash
npm i -g vercel
cd folder-project-ini
vercel login
vercel --prod
```

Ikuti instruksi interaktifnya (pilih scope/akun, nama project, dll).

## Struktur Project

```
.
├── api/
│   ├── __init__.py
│   ├── index.py        # Entrypoint FastAPI (wajib di folder api/ untuk Vercel)
│   └── checker.py       # Logika pencocokan plan vs actual (sama seperti versi Docker)
├── public/
│   ├── index.html       # Halaman utama (file statis, diserve langsung oleh Vercel CDN)
│   ├── style.css
│   └── app.js
├── requirements.txt
├── vercel.json           # Konfigurasi function (maxDuration, dll)
└── .vercelignore
```

## Testing Lokal Sebelum Deploy

Jika sudah install Vercel CLI:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
vercel dev
```

Buka `http://localhost:3000`.

Atau, untuk sekadar mengecek logika backend tanpa Vercel CLI:

```bash
pip install -r requirements.txt uvicorn
uvicorn api.index:app --reload --port 8001
```

(Catatan: cara ini tidak akan menyajikan `public/index.html` secara otomatis
seperti yang Vercel lakukan — hanya untuk testing endpoint `/api/...` saja.)

## Kapan Sebaiknya TIDAK Pakai Vercel

Jika ke depannya:
- File Excel yang diupload bisa sampai ribuan baris dan proses jadi lama (>10 detik), atau
- Aplikasi ini dipakai untuk keperluan komersial/perusahaan dengan trafik tinggi,

...maka lebih baik tetap pakai versi Docker (`docker compose up`) di VPS/server
sendiri atau platform seperti Railway/Render yang free tier-nya juga mendukung
container Docker secara langsung tanpa limit waktu eksekusi sekencang Vercel.
