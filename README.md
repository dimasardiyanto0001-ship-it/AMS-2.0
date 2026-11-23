# AMS — Attendance Monitoring System

Aplikasi web untuk memantau jadwal, kehadiran, dan absensi karyawan per departemen. Backend berjalan di **Google Apps Script** dengan **Google Sheets** sebagai basis data. Frontend dilayani langsung oleh Apps Script atau sebagai situs statis di **Netlify**.

Dikembangkan oleh Dimas Ardiyanto. Dipakai pertama kali oleh Departemen Housekeeping, PT Swadharma Primautama.

> Catatan riwayat penulisan ulang (v1 → v2), daftar bug yang sudah diperbaiki, dan hasil uji ada di [`MIGRASI.md`](MIGRASI.md). Langkah memasang AMS untuk pelanggan baru ada di [`PANDUAN-PEMASANGAN.md`](PANDUAN-PEMASANGAN.md).

---

## Fitur

| Area | Isi |
|---|---|
| **Dashboard** | Kartu KPI kehadiran, tren harian dan tren bulanan, ringkasan Reliever |
| **Jadwal Kerja & Input Realisasi** | Tabel per tanggal, pilih banyak sel lalu *Simpan Batch*, template dan upload Excel, generator jadwal dari Pola Rotasi Shift, draf jadwal bulan depan |
| **Persetujuan Jadwal** | Perubahan jadwal oleh Supervisor masuk antrean; HOD menyetujui atau menolak |
| **Rekap & Log Absensi** | Rekap kehadiran per karyawan, log scan, ekspor Excel/PDF, laporan email bulanan |
| **Absen Online** | Halaman publik berbasis GPS dan selfie untuk karyawan dan Reliever (bisa dipasang sebagai PWA) |
| **Pengajuan Izin** | Form publik untuk karyawan, persetujuan oleh admin, notifikasi email |
| **Kepegawaian** | Daftar Karyawan, Assign ke Jadwal, Daftar Reliever, upload data mesin finger |
| **Periode & Arsip** | Mulai periode baru, arsipkan periode, baca arsip lama |
| **Setting** | Account, Hak Akses (RBAC), Jam Kerja per shift, Hari Libur, Perusahaan (white-label), Audit Log |
| **Chatbot AI** | Tanya jawab data kehadiran dan kirim laporan lewat email (Google Gemini) |
| **Akses per departemen** | Akun non-Super Admin bisa dibatasi ke satu departemen |

Menu yang tampil bagi tiap role diatur lewat **Setting → Hak Akses**. Daftar kunci menu ada di `MENU_LIST` pada `00_Config.gs`.

## Peran pengguna

- **Super Admin (HOD):** melewati semua pembatasan dan satu-satunya yang boleh melakukan aksi sensitif (edit jadwal langsung, mulai periode baru, kirim laporan, dsb.).
- **Admin (Supervisor):** mengikuti matriks Hak Akses; perubahan jadwalnya masuk antrean persetujuan.
- **Role lain:** mengikuti matriks Hak Akses.
- Kolom *Departemen* di sheet `Users` membatasi akun non-Super Admin hanya ke karyawan departemennya.

## Arsitektur

```
Browser / PWA ──▶ Netlify (statis)         ┐
        │                                   ├─ POST text/plain ─▶ Apps Script Web App ─▶ Google Sheets
        └────────▶ URL /exec Apps Script    ┘     (01_Router.gs)                        + Drive + Gmail
```

- Frontend mengirim `POST` dengan `Content-Type: text/plain;charset=utf-8` (bukan `application/json`) agar tidak memicu preflight CORS yang tidak didukung Apps Script. Isi badan tetap JSON.
- `01_Router.gs` memegang tabel **`ROUTES`**: satu sumber kebenaran untuk nama aksi, fungsi backend, dan urutan parameter. Route non-publik wajib menerima `token`.
- Respons selalu berbentuk `{success, data}` atau `{success, error}`.
- Pembatasan departemen dipasang di satu titik (`verifySession_`), sehingga pembacaan data, arsip, upload, dan draf ikut terbatas.

## Struktur repo

```
.
├── 00_Config.gs … 99_Maintenance.gs   Backend Apps Script (satu file per modul)
├── index.html                         Halaman utama (memanggil Stylesheet, SettingJadwalKerja, JavaScript)
├── Stylesheet.html
├── JavaScript.html                    Seluruh logika frontend, satu file (ada DAFTAR ISI di bagian atas)
├── SettingJadwalKerja.html            Panel Setting Jadwal Kerja
├── AbsensiOnline.html                 Halaman publik Absen Online dan Reliever
├── FormIzin.html                      Halaman publik Pengajuan Izin
├── DashboardRange.html                Filter rentang tanggal dan kartu Reliever (modul tambahan, lihat catatan)
├── appsscript.json                    Manifest Apps Script (zona waktu, scope OAuth, web app)
├── build.js · build.config.json       Perakit halaman Netlify
├── netlify.toml                       Konfigurasi build, redirect, header keamanan
├── manifest.json · sw.js · icon-*.png · assetlinks.json   PWA / aplikasi Android (TWA)
├── MIGRASI.md                         Riwayat migrasi v1 → v2
└── PANDUAN-PEMASANGAN.md              Pemasangan untuk pelanggan baru
```

### Modul backend

| File | Fungsi |
|---|---|
| `00_Config` | `CONFIG`, `MENU_LIST`, konstanta auth dan role |
| `01_Router` | Tabel `ROUTES`, `doGet`/`doPost`, `auditRoutes()`, `getRouteParamMap()` |
| `02_SheetUtils` · `06_FileUtils` · `07_TextUtils` | Helper sheet, Drive/Excel, dan teks |
| `03_Auth` · `04_Access` · `39_AksesDepartemen` | Login, sesi, role, matriks hak akses, pembatasan departemen |
| `10_DataApi` · `11_InputData` | `getAllData` (dengan cache), simpan jadwal dan realisasi |
| `12_Period` · `13_Archive` · `14_MonthlyTrend` | Periode, arsip, snapshot tren bulanan |
| `16_Rekap` | Rekap periode aktif dan arsip |
| `17_UploadGrid` · `18_PolaJadwalGenerator` · `38_JadwalDraf` | Upload Excel, generator pola shift, draf bulan depan |
| `35_PersetujuanJadwal` | Antrean persetujuan perubahan jadwal |
| `19_PengajuanIzin` | Pengajuan izin dan notifikasi email |
| `20_AbsenOnline` · `32_JadwalHariIni` | Absen GPS/selfie dan jadwal hari ini |
| `24_BackupReliever` · `33_RelieverData` · `34_RelieverMaster` | Absen Reliever, tabel dan ekspor Data Reliever, Daftar Reliever |
| `25_FingerAbsen` · `26_JamKerja` · `23_SettingJamKerja` | Data mesin finger, jam kerja, pengaturan shift |
| `27_EmployeeData` · `28_EmployeeLifecycle` | Data karyawan, tambah, assign, hapus |
| `29_Export` · `30_AttendanceLog` · `31_MonthlyReport` | Ekspor Excel/PDF, log absensi, laporan email bulanan |
| `15_AIchat` | Chatbot Gemini |
| `21_AuditLog` · `23_Notifications` | Audit log dan notifikasi lonceng |
| `36_Brand` · `37_SetupInstalasi` | Pengaturan perusahaan dan penyiapan instalasi pelanggan baru |
| `98_SecurityCheck` · `99_Maintenance` | Alat diagnosis dan perawatan yang dijalankan manual dari editor |

## Menjalankan dan men-deploy

### 1. Backend (Apps Script)

1. Buat project Apps Script yang terikat ke spreadsheet. Tempel semua file `.gs`.
2. Tempel semua file `.html` tanpa ekstensi di nama editor (`index.html` → `Index`). Nama file utama harus sama persis dengan `HTML_ENTRY_FILE` di `01_Router.gs` (`Index`, huruf besar).
3. Ganti isi `appsscript.json` dengan versi di repo.
4. Isi **Script Properties** (tabel di bawah).
5. Jalankan `auditRoutes()` dari editor. Hasilnya harus bersih.
6. **Deploy → New deployment → Web app**: *Execute as* pemilik, *Who has access* siapa saja. Untuk pembaruan berikutnya pakai *Manage deployments → Edit → New version* agar URL tidak berubah.

### 2. Frontend (Netlify)

```bash
node build.js            # target produksi  (build.config.json → "prod")
node build.js --test     # target salinan   (build.config.json → "test"), ada banner merah
```

`build.js` menggabungkan `index.html` dengan file yang di-`include`, mengisi `GAS_WEBAPP_URL_` dari `build.config.json`, memeriksa sintaks setiap blok `<script>`, lalu menulis hasilnya ke `public/`:

| Sumber | Hasil | Alamat |
|---|---|---|
| `index.html` | `public/index.html` | `/` |
| `AbsensiOnline.html` | `public/absensi-online.html` | `/absensi` |
| `FormIzin.html` | `public/izin.html` | `/izin` |

File sumber tidak diubah, jadi tetap bisa ditempel apa adanya ke editor Apps Script. Di Netlify, branch selain produksi dan Deploy Preview otomatis memakai target `test` (sheet salinan), diatur di `netlify.toml`.

### 3. Script Properties

| Kunci | Fungsi |
|---|---|
| `LEGACY_PASSWORD_SALT` | Salt hash akun lama (hanya bila ada akun hasil migrasi) |
| `REPORT_EMAIL` | Penerima laporan bulanan |
| `GEMINI_API_KEY` | Kunci API chatbot |
| `AI_MODEL` | Opsional; mengganti model Gemini bawaan |
| `BRAND_CONFIG` | Diisi otomatis dari Setting → Perusahaan |
| `IZIN_RESET_INSTALASI` | Pengaman sementara untuk `siapkanInstalasiBaru()`; terhapus otomatis |

Rahasia tidak boleh ditulis di kode. Jalankan `periksaRahasiaScriptProperties()` untuk melihat mana yang sudah terisi (nilainya tidak ditampilkan).

### 4. Alat yang dijalankan manual dari editor

| Fungsi | Kegunaan |
|---|---|
| `auditRoutes()` | Memeriksa tabel `ROUTES` terhadap fungsi dan parameter |
| `getRouteParamMap()` | Menghasilkan peta parameter untuk `GAS_PARAM_MAP_` di frontend |
| `setupMonthlyReportTrigger()` | Memasang ulang trigger laporan bulanan (tanggal 5, jam 07:00) |
| `periksaRahasiaScriptProperties()` · `periksaAkunBelumTermigrasi()` | Diagnosis keamanan, hanya membaca |
| `cekKesiapanInstalasi()` · `siapkanInstalasiBaru()` | Penyiapan salinan untuk pelanggan baru. **`siapkanInstalasiBaru()` menghapus data; jangan pernah dijalankan di produksi** |
| `debugNameMismatch()` · `findBlankRowGaps()` · `auditJamNyangkutOFF()` · `backfillRekapBulananDariArsip()` dst. | Perawatan di `99_Maintenance.gs` |

Fungsi berakhiran `_` tidak muncul di dropdown *Jalankan*; alat manual sengaja diberi nama tanpa garis bawah.

## Konvensi kode

- Nama route, urutan parameter, dan bentuk respons tidak diubah tanpa menyesuaikan `GAS_PARAM_MAP_` di `JavaScript.html`.
- Setiap fungsi sensitif memakai `isSuperAdminRole_(session.role)`, bukan membandingkan string `'admin'`.
- Koordinat sel (baris header, kolom hari pertama, dst.) selalu dari `CONFIG`, tidak ditulis ulang di modul.
- Tulis ke sheet dalam blok (`setValues`) dan di dalam `withScriptLock_`; operasi batch bersifat semua-atau-tidak-sama-sekali.
- Teks dari pengguna yang ditulis ke sheet dilewatkan `guardSheetText_` agar tidak menjadi rumus.
- `JavaScript.html` sengaja tetap satu file. Cari bagian lewat DAFTAR ISI di atasnya (Ctrl+F pada penanda `============ NAMA ============`).

## Catatan

- `DashboardRange.html` belum dipanggil dari `index.html`. Cara memasangnya ada di komentar bagian atas file itu (tambahkan `<?!= include('DashboardRange'); ?>` tepat setelah `include('JavaScript')`).
- Repositori ini publik. Kode tidak boleh memuat rahasia (lihat tabel Script Properties). URL deployment di `build.config.json` adalah alamat web app yang memang bisa diakses publik, tetapi pertimbangkan menjadikan repositori *Private* (Netlify tetap bisa membangun lewat integrasi GitHub).
- Sebelum dipasang untuk pelanggan lain, baca bagian *Sebelum dijual* di `PANDUAN-PEMASANGAN.md` (hak komersial, kebijakan privasi, UU PDP).
