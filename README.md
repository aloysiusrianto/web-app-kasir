# Web App Kasir

POS toko: frontend HTML statis + backend Google Apps Script yang menulis ke Google Sheets.

- Kasir: cari produk, keranjang, tunai / QRIS / transfer, struk
- Produk: CRUD (hapus = nonaktif)
- Riwayat transaksi
- Pajak toko tersimpan di sheet Settings

Spreadsheet otomatis dibuat saat Web App pertama kali dibuka (`Kasir — Data Toko`) dengan sheet `Products`, `Sales`, `SaleItems`, `Settings`. Enam produk demo di-seed jika sheet kosong.

## Arsitektur

```
Vercel (index.html)  --POST text/plain JSON-->  Apps Script /exec  -->  Google Sheets
     atau
Apps Script Web App menyajikan index.html + google.script.run
```

`Content-Type: text/plain` dipakai agar browser tidak mengirim preflight CORS ke `script.google.com`.

## File

| File | Fungsi |
|---|---|
| `index.html` | UI kasir (tema gelap/terang) |
| `Code.gs` | Backend: produk, checkout ber-lock, stok, riwayat |
| `appsscript.json` | Manifest Apps Script (V8, akses anonim) |
| `vercel.json` | Header cache untuk deploy statis |
| `.claspignore` | Jangan unggah README/git ke Apps Script |

## 1. Deploy backend Google Apps Script

Cara cepat (editor web):

1. Buka [script.google.com](https://script.google.com) → **New project**.
2. Ganti `Code.gs` dengan isi file ini. Tambah file HTML bernama `index` (tanpa `.html`) dan tempel `index.html`.
3. Project Settings → tempel `appsscript.json` (atau set timezone `Asia/Jakarta`).
4. **Deploy → New deployment → Web app**
   - Execute as: **Me**
   - Who has access: **Anyone** (wajib jika frontend Vercel memanggil API)
5. Salin URL `https://script.google.com/macros/s/…/exec`.
6. Buka URL itu sekali di browser. Spreadsheet dan produk demo dibuat otomatis.

Cara clasp (opsional, dari folder ini):

```bash
npm i -g @google/clasp
clasp login
clasp create --type webapp --title "Web App Kasir" --rootDir .
clasp push
clasp deploy --description "v1"
```

Izin yang diminta: Spreadsheet (buat/baca/tulis) dan Properties.

## 2. Hubungkan frontend ke Vercel

1. Import repo GitHub ini di [vercel.com/new](https://vercel.com/new).
2. Framework Preset: **Other**. Root: `.`  Output: biarkan kosong (static `index.html`).
3. Deploy.
4. Buka situs Vercel → **Pengaturan** → tempel URL Web App Apps Script → Simpan.
5. Status di header harus **Terhubung**.

Tanpa Vercel, UI juga hidup di URL `/exec` Apps Script itu sendiri.

## API (POST ke `/exec`)

Body JSON (`Content-Type: text/plain`):

```json
{ "action": "bootstrap", "payload": {} }
```

Aksi: `health`, `bootstrap`, `listProducts`, `saveProduct`, `deleteProduct`, `checkout`, `listSales`, `getSale`, `saveSettings`.

Checkout menolak stok kurang dan uang tunai di bawah total. Stok di-update di dalam `LockService`.

## Pengembangan lokal

Buka `index.html` langsung. Tanpa URL API, banner konfigurasi muncul. Isi URL di Pengaturan (tersimpan di `localStorage` kunci `kasir.apiUrl`).
