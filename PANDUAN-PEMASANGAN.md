# Panduan Pemasangan AMS untuk Pelanggan Baru (Jalur A: satu salinan per pelanggan)

Perkiraan waktu: 1–2 jam per pelanggan setelah template siap.
**Jangan pernah menjalankan langkah reset di spreadsheet produksi (Swadharma).**

## A. Siapkan template bersih (sekali saja)
1. Di akun Google Bapak, buka spreadsheet produksi → **Berkas → Buat salinan**. Beri nama `AMS Template`.
2. Di salinan itu buka **Ekstensi → Apps Script** (script ikut tersalin).
3. **Project Settings → Script Properties**: hapus semua properti lama (kunci Gemini, email, salt milik Swadharma), lalu tambah `IZIN_RESET_INSTALASI` = `YA`.
4. Jalankan fungsi `siapkanInstalasiBaru` dari editor. Hasil di **Execution log**: daftar yang dibersihkan dan **login awal (admin + password acak)**. Properti pengaman terhapus otomatis.
5. Periksa spreadsheet: Jadwal/Realisasi/DataKaryawan kosong, tidak ada sheet `Arsip_*`.
6. Jangan dipakai langsung. Pakai **Buat salinan** dari `AMS Template` untuk tiap pelanggan. Simpan template ini tetap bersih.

## B. Pasang untuk satu pelanggan
1. Salin `AMS Template` → ganti nama `AMS - <Nama Pelanggan>`. (Bisa dimiliki akun Google pelanggan: pindahkan kepemilikan atau buat salinan dari akun mereka.)
2. Buka salinan → **Apps Script → Project Settings → Script Properties**, isi:
   - `GEMINI_API_KEY` (kunci milik pelanggan agar biaya AI terpisah)
   - `REPORT_EMAIL` (penerima laporan bulanan)
   - `LEGACY_PASSWORD_SALT` hanya jika ada akun migrasi (biasanya dilewati)
3. Jalankan `cekKesiapanInstalasi` untuk melihat daftar yang masih harus diisi.
4. **Deploy → Deployment baru → Web app**: Jalankan sebagai *Saya*, akses *Siapa saja*. Salin URL `/exec`.
5. Jika memakai Netlify/situs sendiri: ganti URL `/exec` di `JavaScript.html`, `AbsensiOnline.html`, `FormIzin.html` (konstanta `GAS_WEBAPP_URL_`) lalu `node build.js`. Jika tidak, pelanggan bisa langsung memakai URL `/exec`.
6. Buka URL, login `admin` dengan password dari langkah A4, lalu **segera ganti password** (Setting → Account).
7. **Setting → Perusahaan**: isi nama aplikasi, perusahaan, departemen, logo (alamat https), batas lembur.
8. Isi data awal:
   - Setting → Jam Kerja: jam masuk/pulang tiap kode shift.
   - Setting → Hak Akses: atur role dan menu.
   - Setting → Account: buat akun Supervisor/HOD.
   - Daftar Karyawan → tambah karyawan; Setting → Assign ke jadwal.
   - Sheet tersembunyi `SettingLokasiAbsen` baris 2: latitude, longitude, radius kantor (untuk absen online).
   - Opsional: sheet `HariLibur` (kolom A tanggal libur nasional) agar header tanggal merah.
9. Uji: login, absen online dari ponsel, ajukan izin, unduh Excel jadwal.

## C. Berkas yang spesifik per pelanggan (edit manual bila perlu)
- `manifest.json` dan `assetlinks.json`: nama aplikasi PWA/Android (masih "AMS Housekeeping" / `ams_wisma46`).
- Logo bawaan di `index.html` (cukup isi logo di Setting → Perusahaan).
- Nama folder Drive foto absen: `FOTO_ABSEN_ROOT_FOLDER_NAME_` dan `FOTO_PROFIL_ROOT_FOLDER_NAME_` di `20_AbsenOnline.gs`.

## D. Memperbarui pelanggan saat ada versi baru
Saat ini manual: tempel file `.gs`/`.html` yang berubah ke proyek Apps Script pelanggan, lalu **Deploy → Versi baru**. Untuk banyak pelanggan, gunakan `clasp` (`clasp push` per proyek) agar tidak menempel satu per satu.

## E. Sebelum dijual (di luar kode)
- Izin tertulis dari PT Swadharma Primautama atas hak komersial.
- Kebijakan Privasi, Syarat Layanan, perjanjian pemrosesan data (UU PDP).
- Pastikan data Swadharma tidak ada di template dan repo.
