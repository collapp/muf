# Deploy Collection Tracker — GitHub + Firebase (Semua dari HP, Tanpa Terminal)

Caranya: file-file project dibuat langsung lewat **website GitHub** (bukan aplikasi/app,
buka **github.com** di browser HP), lalu GitHub otomatis build & upload ke **Firebase
Hosting** setiap kali ada perubahan. Tidak perlu install apa-apa di HP.

---

## Bagian 1 — Buat repo di GitHub

1. Buka **github.com** di browser HP, login.
2. Tombol **+** (kanan atas) → **New repository**.
3. Isi nama repo, misal `collection-tracker` → pilih **Public** atau **Private** →
   **Create repository**.

## Bagian 2 — Isi file project (satu per satu)

Trik untuk HP: di GitHub, saat bikin file baru, kalau nama filenya kamu tulis **lengkap
dengan folder** (misal `src/App.jsx`), GitHub otomatis bikin foldernya. Jadi tidak perlu
upload folder.

Di halaman repo → **Add file → Create new file**. Ulangi untuk tiap file di bawah:
ketik nama file di kolom nama, tempel (paste) isinya di kotak besar di bawahnya, lalu
scroll ke bawah → **Commit changes**.

Semua isi file ini juga sudah saya siapkan di file zip yang saya kirim sebelumnya
(`collection-app.zip`) — buka/extract di HP (aplikasi Files/My Files bisa extract zip),
lalu buka tiap file dengan text editor untuk copy isinya kalau mau lebih gampang
copy-paste daripada mengetik ulang dari sini.

Buat 8 file ini, dengan nama (path) persis seperti berikut:

1. `package.json`
2. `vite.config.js`
3. `index.html`
4. `.gitignore`
5. `firebase.json`
6. `.firebaserc`
7. `src/main.jsx`
8. `src/App.jsx`
9. `.github/workflows/firebase-hosting.yml`

Isinya ambil dari file dengan nama sama di dalam `collection-app.zip`. Yang **wajib
diedit isinya** sebelum commit:

- **`.firebaserc`** — ganti `GANTI-DENGAN-PROJECT-ID-FIREBASE-ANDA` dengan Project ID
  Firebase Anda (didapat di Bagian 3).
- **`.github/workflows/firebase-hosting.yml`** — ganti `GANTI-DENGAN-PROJECT-ID-FIREBASE-ANDA`
  di baris `projectId:` paling bawah, dengan Project ID yang sama.

(File `src/App.jsx` isinya panjang — pastikan ke-copy lengkap dari awal sampai akhir.)

---

## Bagian 3 — Buat project Firebase

1. Buka **console.firebase.google.com** di browser HP, login dengan akun Google.
2. **Add project** → beri nama → ikuti langkah sampai selesai (Google Analytics boleh
   dimatikan, tidak perlu).
3. Setelah masuk ke dashboard project, catat **Project ID**-nya (bukan nama project) —
   terlihat di **Project settings** (ikon gerigi ⚙️ di pojok kiri atas → **Project
   settings**). Project ID inilah yang dipakai untuk mengganti
   `GANTI-DENGAN-PROJECT-ID-FIREBASE-ANDA` di dua file tadi.
4. Di menu kiri, buka **Build → Hosting** → **Get started** → lewati saja langkah-langkah
   CLI-nya (klik **Next** terus sampai selesai) — kita tidak pakai CLI, cukup sampai
   Hosting-nya "aktif" di project.

## Bagian 4 — Ambil kunci Service Account (untuk izin GitHub deploy ke Firebase)

1. Masih di **Project settings** → tab **Service accounts**.
2. Klik **Generate new private key** → **Generate key**. File `.json` akan ke-download
   ke HP (biasanya masuk folder Download).
3. Buka file `.json` itu dengan text editor / browser (bisa lewat app Files → buka
   dengan → browser atau text editor), lalu **copy semua isinya** (seluruh teks, dari
   `{` sampai `}`).

## Bagian 5 — Simpan kunci itu sebagai Secret di GitHub

1. Buka repo GitHub Anda → **Settings** (tab repo, bukan settings akun) → menu kiri
   **Secrets and variables → Actions**.
2. **New repository secret**.
3. Name: `FIREBASE_SERVICE_ACCOUNT`
4. Secret: tempel (paste) seluruh isi file `.json` tadi.
5. **Add secret**.

## Bagian 6 — Deploy

Begitu file `.github/workflows/firebase-hosting.yml` ter-*commit* (Bagian 2) dan secret
sudah tersimpan (Bagian 5), GitHub akan otomatis jalan setiap ada perubahan/commit ke
branch `main`. Cek prosesnya di tab **Actions** pada repo Anda — tunggu sampai muncul
tanda centang hijau ✅ (biasanya 1-3 menit).

Setelah selesai, aplikasi Anda bisa diakses di:

```
https://PROJECT-ID-ANDA.web.app
```

(ganti `PROJECT-ID-ANDA` dengan Project ID dari Bagian 3).

---

## Kalau ada perubahan lagi nanti

Tinggal edit file yang mau diubah langsung di GitHub (buka file → ikon pensil ✏️ →
edit → **Commit changes**). Setiap commit ke `main` otomatis ter-deploy ulang, tidak
perlu ulangi langkah Firebase/secret.

## Catatan

- Data yang di-upload di aplikasi (dari Excel) tetap tersimpan di **localStorage**
  browser HP/perangkat masing-masing — tidak otomatis tersinkron ke Firebase database.
  Firebase di sini hanya dipakai sebagai **hosting** (tempat aplikasinya di-online-kan).
- Kalau nanti mau datanya benar-benar tersimpan terpusat (misal supaya bisa diakses dari
  banyak perangkat/anggota tim), itu perlu tambahan **Firestore Database** — beri tahu
  saya kalau itu yang dimaksud, karena bagian kodenya perlu ditambah.
