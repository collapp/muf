# Setup Firestore Database (Wajib Sebelum Deploy)

Aplikasi sekarang menyimpan data langsung ke **Firestore** (database Firebase),
bukan lagi di HP masing-masing. Semua yang login lihat data yang sama, real-time.

Semua langkah ini lewat **browser HP**, tidak perlu terminal.

## 1. Aktifkan Firestore Database

1. Buka **console.firebase.google.com** → pilih project `datacollection-95473`
2. Menu kiri → **Build → Firestore Database** → **Create database**
3. Pilih lokasi server (misal `asia-southeast2` / Jakarta, biar cepat) → **Next**
4. Pilih **Start in production mode** → **Enable**

(Jangan pilih "test mode" — nanti datanya bisa dibaca siapa saja tanpa login.)

## 2. Pasang Security Rules

1. Masih di halaman **Firestore Database**, buka tab **Rules**
2. Hapus semua isi kotak rules, ganti dengan isi file **`firestore.rules`**
   (sudah saya siapkan di zip)
3. **Publish**

Rules ini artinya: **hanya akun yang sudah login** (yang Anda buat manual di
Authentication → Users) yang bisa baca/tulis data. Orang lain yang kebetulan
tahu URL web-nya tidak akan bisa akses data sama sekali.

## 3. Update file di GitHub

File yang perlu ditimpa/ditambah di repo Anda (semua isinya ada di zip):

| File | Status |
|---|---|
| `src/firebase.js` | **timpa** (nambah koneksi Firestore) |
| `src/firestore.js` | **file baru** — buat via "Add file → Create new file" |
| `src/App.jsx` | **timpa** (logika utama, sekarang pakai Firestore) |
| `firestore.rules` | **file baru** (opsional disimpan di repo untuk arsip — yang penting sudah di-publish di Firebase Console langkah 2) |

Setelah semua file ter-*commit*, Actions otomatis build & deploy ulang. Tunggu
sampai ✅ hijau, lalu buka lagi web-nya.

## 4. Uji coba

1. Buka web-nya, login pakai akun yang sudah dibuat
2. Upload file Excel bulanan seperti biasa — sekarang akan muncul progress
   "Menyimpan ke database… 1200/5839" (agak lama karena nulis ribuan data,
   normal, tunggu sampai selesai)
3. Coba buka dari **HP lain** (atau akun lain), login — datanya harus sudah
   langsung ada tanpa perlu upload ulang ✅

## Yang berubah dari versi sebelumnya

- **Data kini satu pusat** — semua yang login lihat & edit data yang sama,
  real-time (tidak perlu refresh manual).
- **Upload bulanan otomatis "merge"** — kontrak lama yang statusnya/catatannya
  sudah diisi tidak akan hilang; kontrak yang tidak ada lagi di file Excel
  terbaru otomatis pindah ke tab **Riwayat** (bukan dihapus).
- Tombol "Hapus semua data" di versi lama sudah dilepas — terlalu berisiko
  untuk database bersama tim.
- Backup ke Excel (tombol ↓ di header) tetap ada, sekarang menarik data
  langsung dari Firestore.

## Kalau ada error saat login/akses data

- **"Missing or insufficient permissions"** → cek lagi Rules di langkah 2,
  pastikan sudah di-**Publish**, dan pastikan akun yang dipakai memang sudah
  dibuat di Authentication → Users.
- **Data tidak muncul sama sekali** → cek tab **Actions** di GitHub, pastikan
  build terakhir ✅ (bukan ❌), dan pastikan `src/firestore.js` benar-benar
  sudah ter-upload (bukan cuma App.jsx).
