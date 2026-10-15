# Modul satelit batch 2: `19_PengajuanIzin.gs` dan `20_AbsenOnline.gs`

Tempel bagian ini ke MIGRASI.md.

## Cara pasang (di proyek SALINAN v2, jangan di produksi)

1. Ganti isi `19_PengajuanIzin.gs` dan `20_AbsenOnline.gs`. Tidak ada file yang ditambah atau dihapus.
2. `07_TextUtils.gs` TIDAK perlu diganti (sudah memuat `guardSheetText_` dari batch 1).
3. Deploy > Manage deployments > pilih deployment > Edit > New version > Deploy (URL tidak berubah).
4. Jalankan `auditRoutes()` di editor, hasilnya harus bersih.

## Cara menguji

Tes dilakukan dengan mock Apps Script di Node: kode lama (Fase 8) dibandingkan dengan kode baru pada skenario acak, lalu perbedaan yang disengaja diuji satu per satu. Hasil: `19_PengajuanIzin` 25 pengecekan, `20_AbsenOnline` 18 pengecekan.

## 19_PengajuanIzin.gs: bug yang diperbaiki

1. **Email notifikasi memanggil dirinya sendiri tanpa henti.** `kirimEmailNotifikasiPengajuan_` memanggil dirinya sendiri sehingga mengirim email berulang dan statusnya bisa salah ("Disetujui" padahal ditolak). Sekarang tepat 1 email dengan status yang benar.
2. **Approve tanpa menerapkan ke Realisasi.** Pengajuan bisa berstatus Approved tetapi sel jadwal tidak berubah. Sekarang cek status, terapkan, dan tandai dilakukan dalam satu kunci. Validasi dilakukan sebelum menulis apa pun.
3. **Penulisan sel satu per satu** diganti satu blok sel per rentang hari berurutan (jauh lebih cepat).
4. **Approve dan reject kini memakai kunci** dan menulis Audit Log.
5. **`autoTandaiAlfaHarian_`** membaca sel di dalam kunci dan melindungi pengajuan Pending yang tanggalnya bertipe Date.
6. **Batas:** alasan 1000 karakter, catatan 500, rentang maksimal 62 hari, lampiran 10 MB.

## 20_AbsenOnline.gs: bug yang diperbaiki

1. **`searchKaryawanUntukAbsen` membocorkan NIK** (endpoint publik). Sekarang `nik:''`.
2. **Retensi foto** salah hitung di batas bulan (31 Mei dan 31 Des). Sekarang memakai `new Date(y, m - RETENSI, 1)`.
3. **Penghapusan foto lama** hanya menandai kolom FotoUrl, tidak menyentuh kolom lain.
4. **`setFotoProfilKaryawan`:** ukuran dicek sebelum membuat file, gagal berbagi tidak membatalkan, memakai kunci.
5. **Foto gagal disimpan:** absen tetap tercatat dengan FotoUrl kosong, bukan error.
6. **Validasi GPS** pada `submitAbsenOnline`.
7. **Batas JamKerja** 16 jam.

## Keputusan yang perlu kamu tentukan

- Batas 16 jam JamKerja: apakah sesuai kebijakan?
- Endpoint publik: NIK tetap menjadi satu-satunya rahasia dan belum ada pembatasan percobaan (rate limit).
- Apakah Pengajuan Izin masuk ke matriks hak akses?
- Batas form-izin.html (10 MB, 62 hari, 1000 karakter): sesuaikan jika perlu.
- Hak akses chatbot untuk Viewer (dari batch 1).

## Sisa modul satelit

`18_PolaJadwalGenerator.gs` (11 fungsi) dan `24_BackupReliever.gs` (15 fungsi).
