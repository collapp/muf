# Collection Tracker

Aplikasi web untuk upload data Excel (write off / collection) dan melacak
status & catatan kunjungan per kontrak. Semua data tersimpan di browser
(`localStorage`) di perangkat yang dipakai — tidak ada server/database.

## 1. Coba di komputer sendiri dulu (opsional)

Butuh [Node.js](https://nodejs.org) terinstal (versi 18 ke atas).

```bash
npm install
npm run dev
```

Buka alamat yang muncul di terminal (biasanya `http://localhost:5173`).

## 2. Upload ke GitHub

```bash
git init
git add .
git commit -m "Collection tracker"
git branch -M main
git remote add origin https://github.com/USERNAME/NAMA-REPO.git
git push -u origin main
```

Ganti `USERNAME/NAMA-REPO` dengan repo GitHub Anda.

## 3. Deploy ke GitHub Pages

Cara termudah, pakai paket `gh-pages` yang sudah disiapkan di `package.json`:

```bash
npm run deploy
```

Perintah ini akan build project lalu mem-publish folder `dist` ke branch
`gh-pages`. Setelah itu:

1. Buka repo di GitHub → **Settings → Pages**
2. Pada **Branch**, pilih `gh-pages` / folder `root` → **Save**
3. Tunggu 1–2 menit, situs akan aktif di:
   `https://USERNAME.github.io/NAMA-REPO/`

## Catatan penting

- **Data tidak tersinkron antar perangkat/browser.** Karena datanya
  disimpan secara lokal (`localStorage`), data yang Anda upload di laptop
  tidak akan muncul kalau dibuka dari HP atau browser lain. Kalau nanti
  perlu data yang sama bisa diakses dari banyak perangkat/tim, aplikasi ini
  perlu ditambah backend/database — beri tahu saya kalau itu diperlukan.
- **Membersihkan cache browser akan menghapus data.** Ekspor/backup manual
  belum ada di versi ini; bisa ditambahkan kalau dibutuhkan.
- File Excel yang didukung: `.xlsx` / `.xls`, dengan sheet yang punya kolom
  seperti `CONTRACTNO`, `CUSTNAME`, `NOPOL`, `BAL_PRIN`, dst — sheet dengan
  kecocokan kolom terbanyak akan otomatis dipilih.
