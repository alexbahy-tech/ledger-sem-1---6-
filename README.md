# 📚 SMK Ledger Nilai XII — Panduan Deploy ke GitHub Pages

## 🚀 Langkah 1: Upload ke GitHub

1. **Buat repositori baru** di GitHub (misal: `smk-ledger-nilai`)
2. Upload file `index.html` ke root repositori
3. Pastikan nama file **persis** `index.html` (huruf kecil semua)

## 🌐 Langkah 2: Aktifkan GitHub Pages

1. Buka repositori → klik tab **Settings**
2. Di sidebar kiri, klik **Pages**
3. Di bagian **Source**, pilih:
   - Branch: `main`
   - Folder: `/ (root)`
4. Klik **Save**
5. Tunggu 1–3 menit, URL akan muncul:
   ```
   https://[username].github.io/[nama-repo]/
   ```

## 📊 Langkah 3: Konfigurasi Google Sheets

### Agar spreadsheet bisa dibaca:
1. Buka spreadsheet Google Sheets Anda
2. Klik tombol **Share** (pojok kanan atas)
3. Ubah ke **"Anyone with the link" → Viewer**
4. Klik **Done**

### Ambil Spreadsheet ID dan GID:
Dari URL spreadsheet:
```
https://docs.google.com/spreadsheets/d/[SHEET_ID]/edit#gid=[GID]
```
- `SHEET_ID` = string panjang setelah `/d/`
- `GID` = angka setelah `gid=`

### Set di dashboard:
1. Buka dashboard → klik **Pengaturan** di sidebar
2. Isi **Spreadsheet ID** dan **Sheet GID**
3. Klik **Simpan & Reload Data**

## ⚙️ Langkah 4: Deploy Google Apps Script (untuk Input Nilai)

1. Buka [script.google.com](https://script.google.com)
2. Buat project baru → paste isi `Code.gs`
3. Klik **Deploy → New deployment**
4. Pilih type: **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Klik **Deploy** → copy URL yang muncul
6. Di dashboard → **Pengaturan** → isi URL di kolom **URL Web App**
7. Klik **Tes Koneksi** untuk verifikasi

## ❓ Troubleshooting

| Masalah | Solusi |
|---------|--------|
| Halaman 404 | Pastikan file bernama `index.html` dan GitHub Pages sudah aktif |
| Data tidak muncul | Cek sharing spreadsheet (harus "Anyone with link → Viewer") |
| Input nilai tidak tersimpan | Pastikan URL GAS sudah diisi di Pengaturan |
| CORS error | Gunakan proxy — dashboard sudah otomatis mencoba beberapa proxy |

## 📁 Struktur File

```
[repo]/
└── index.html      ← File utama (wajib ada di root)
└── README.md       ← Panduan ini (opsional)
```

> **Catatan:** Hanya butuh satu file `index.html` — tidak perlu folder atau file lain!
