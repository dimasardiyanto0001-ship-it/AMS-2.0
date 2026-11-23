# AMS v2 — Catatan Migrasi

Riwayat penulisan ulang `Code.gs` (v1) menjadi modul-modul v2: apa yang berubah, bug yang ditemukan dan diperbaiki, cara memasang, hasil uji, dan keputusan yang masih menunggu. Untuk gambaran proyek dan cara menjalankannya, lihat [`README.md`](README.md).

## Status

- **Fase 1–8 selesai.** Seluruh isi `Code.gs` lama sudah ditulis ulang, `05_LegacyCode.gs` dihapus, dan frontend dibersihkan (Fase 8 versi ringan; `JavaScript.html` sengaja tetap satu file).
- **Modul satelit:** `15_AIchat`, `21_AuditLog`, `23_SettingJamKerja` (batch 1) serta `19_PengajuanIzin` dan `20_AbsenOnline` (batch 2) sudah ditulis ulang. Sisa: `18_PolaJadwalGenerator` dan `24_BackupReliever`. `23_Notifications` diperbaiki di Fase 1.
- **Belum diuji di spreadsheet asli.** Semua uji di bawah memakai tiruan Apps Script (Node) atau browser headless. Selalu uji di **salinan** spreadsheet dulu.

> **Tentang angka route.** Angka "65 route" dan "66 route" di catatan fase adalah jumlah pada saat fase itu. Tabel `ROUTES` di `01_Router.gs` sekarang berisi 97 entri (16 di antaranya publik). Jalankan `auditRoutes()` untuk angka dan status terbaru.

## Daftar isi

1. [Peta file lama → baru](#1-peta-file-lama--baru)
2. [Modul yang ditambahkan setelah Fase 8](#2-modul-yang-ditambahkan-setelah-fase-8)
3. [Riwayat perubahan dan bug per fase](#3-riwayat-perubahan-dan-bug-per-fase)
4. [Cara memasang](#4-cara-memasang)
5. [Menghapus rahasia dari kode](#5-menghapus-rahasia-dari-kode)
6. [Hasil uji](#6-hasil-uji)
7. [Temuan yang menunggu keputusan](#7-temuan-yang-menunggu-keputusan)
8. [Status fase](#8-status-fase)

---

## 1. Peta file lama → baru

Perilaku yang sengaja dipertahankan: nama route, urutan parameter, bentuk respons `{success,data}` / `{success,error}`, skema hash password (akun lama tetap bisa login), dan logika RBAC.

| File baru | Isi | Pengganti |
|---|---|---|
| `00_Config.gs` | `CONFIG`, daftar menu, konstanta auth/role | `Code.gs` bagian 00 |
| `01_Router.gs` | Tabel `ROUTES`, `dispatchRoute_`, `doGet/doPost`, `auditRoutes()`, `getRouteParamMap()` | `Code.gs` bagian 01 (switch ±200 baris) |
| `02_SheetUtils.gs` | `getSheet_`, `readGrid_`, `withScriptLock_`, `checkRowAlignment_` | bagian 03 + duplikat di 16 |
| `03_Auth.gs` | login, sesi, hash, rate limit, ganti password | bagian 02 |
| `04_Access.gs` | user, role, matriks hak akses | bagian 17 |
| `06_FileUtils.gs` | helper Drive/Excel bersama: unggah-konversi file sementara (selalu dibuang), ekspor xlsx | kode ganda di bagian 07, 11, 13 |
| `07_TextUtils.gs` | `escapeHtmlServer_` (email laporan dan email pengajuan izin), `guardSheetText_` | bagian 09 |
| `10_DataApi.gs` | `getAllData` (+cache), tren harian, ringkasan dashboard, sinkron OFF | bagian 04 + patch performa |
| `11_InputData.gs` | `saveRealisasi`, `saveRealisasiBatch`, `saveJadwal`, validasi sel | bagian 05 |
| `12_Period.gs` | parsing nama periode, `mulaiPeriodeBaru`, reset sheet bulan baru | bagian 08 + util tanggal di bagian 10 |
| `13_Archive.gs` | `archiveCurrentPeriod`, `ArsipIndex`, baca/hapus arsip, `getPeriodMeta` | bagian 08 + `readArchivedJadwalGrid_` |
| `14_MonthlyTrend.gs` | snapshot `Rekap_Bulanan`, `getMonthlyTrendYear` | `20_TrendComparison.gs` (dihapus) |
| `16_Rekap.gs` | `getRekap_` (periode aktif) dan `computeRekapDariGrid_` (arsip/tren), penghitung bersama `tallyAttendance_` | `16_RepairAlignment.gs` + bagian 08 |
| `17_UploadGrid.gs` | `uploadJadwal`, `uploadRealisasi`, `generateJadwalTemplateExcel` | bagian 07 + 11 |
| `25_FingerAbsen.gs` | `uploadFingerAbsen`, parsing file mesin finger, terapkan ke Realisasi/JamKerja/JamScan | bagian 13 (upload) |
| `26_JamKerja.gs` | sheet tersembunyi JamKerja/JamScan, `attachJamKerjaToRekap_`, `attachNikToRekap_`, ringkasan jam | bagian 13 (jam kerja) |
| `27_EmployeeData.gs` | sheet `DataKaryawan`: master data, daftar karyawan, `setEmployeeMasterData`, `setEmployeeNik` | bagian 15_Kepegawaian |
| `28_EmployeeLifecycle.gs` | `addEmployee`, `assignEmployeeToJadwal`, `deleteEmployee`, `deleteUnassignedEmployee`, `setKaryawanResign` | `addEmployee` + `21_Resign.gs`, `22_AssignJadwal.gs` (dihapus) |
| `29_Export.gs` | `exportSheetPdf`, `generateReportExcel`, `generateReportPdf` | bagian 06_ExportPdf + laporan generik |
| `30_AttendanceLog.gs` | `generateAttendanceLogExcel`, `readPeriodGrids_`, parser scan | bagian 10_AttendanceLog |
| `31_MonthlyReport.gs` | `sendMonthlyReportEmail` (laporan **bulan lalu**, tanggal 5 jam 07:00), pemilihan sumber periode, analisis tren, HTML email, `setupMonthlyReportTrigger` | bagian 09_MonthlyReport |
| `32_JadwalHariIni.gs` | `getJadwalHariIni` (route publik Absen Online) | sisa `Code.gs` |
| `99_Maintenance.gs` | alat perawatan/diagnosis yang dijalankan manual dari editor (`debugNameMismatch`, `findBlankRowGaps`, `auditJamNyangkutOFF`, `bersihkanJamNyangkutOFF`, `fixJamKerjaPeriodeAktifSekarang`, `backfillRekapBulananDariArsip`, dst.) | bagian 14_DebugNama + fungsi perbaikan sekali-jalan |

Modul satelit yang sudah ada sebelum migrasi dan ditulis ulang bertahap: `15_AIchat`, `18_PolaJadwalGenerator`, `19_PengajuanIzin`, `20_AbsenOnline`, `21_AuditLog`, `23_Notifications`, `23_SettingJamKerja`, `24_BackupReliever`.

## 2. Modul yang ditambahkan setelah Fase 8

Modul ini fitur baru, bukan hasil migrasi. Ringkasannya diambil dari komentar header tiap file.

| File | Isi |
|---|---|
| `33_RelieverData.gs` | Menu "Data Reliever": tabel absen reliever per tanggal (rentang bebas, boleh lintas bulan) dan unduh Excel/PDF. Izin: menu `relieverdata`. |
| `34_RelieverMaster.gs` | "Daftar Reliever" di Kepegawaian (tambah, ubah Perusahaan/No HP, Aktif/Nonaktif). Pengganti input manual di sheet `DataReliever`. Tidak ada hapus, karena riwayat absen memakai nama sebagai kunci. Izin: menu `daftarreliever`. |
| `35_PersetujuanJadwal.gs` | Perubahan Jadwal oleh Supervisor (role Admin) masuk antrean *Pending*; HOD (Super Admin) menyetujui atau menolak. Super Admin tetap langsung berlaku. Sel dikenali lewat nama karyawan, bukan nomor baris. |
| `36_Brand.gs` | Pengaturan Perusahaan (white-label): nama aplikasi, perusahaan, departemen, logo, batas lembur. Disimpan di Script Properties (`BRAND_CONFIG`). |
| `37_SetupInstalasi.gs` | `siapkanInstalasiBaru()` dan `cekKesiapanInstalasi()` untuk menyiapkan salinan bersih bagi pelanggan baru (lihat `PANDUAN-PEMASANGAN.md`). |
| `38_JadwalDraf.gs` | Draf jadwal bulan depan, bisa diunggah sebelum periode diganti, otomatis diterapkan saat HOD menjalankan "Mulai Periode Baru". |
| `39_AksesDepartemen.gs` | Pembatasan akses per departemen lewat `verifySession_()`. |
| `98_SecurityCheck.gs` | `periksaRahasiaScriptProperties()` dan `periksaAkunBelumTermigrasi()`. |
| `DashboardRange.html` | Filter rentang tanggal dan kartu Jumlah Reliever di Dashboard. Belum dipanggil dari `index.html`. |

---

## 3. Riwayat perubahan dan bug per fase

Urutan dari yang terbaru ke yang terlama.

### Tata letak halaman (`Stylesheet.html`)

- **Halaman melebar di layar sempit, kolom tanggal terpotong.** `main` memakai `margin: 0 auto` di dalam wadah flex kolom (`.main-area`), sehingga lebarnya mengikuti lebar isi tabel (31 kolom, teks tidak boleh turun baris), bukan lebar layar. Akibatnya seluruh halaman bisa digeser ke samping (di HP sampai ±1.000 px) dan tabel tidak punya batas untuk menggulir di dalam kotaknya.
- **Perbaikan:** `main { width: 100%; min-width: 0 }`; tabel Jadwal/Realisasi dirapatkan (kolom tanggal lebih sempit, Jabatan dipotong dengan "..." dan disembunyikan di HP); tab Jadwal/Input boleh selebar 1.760 px (memakai `:has()`, browser lama tetap memakai batas 1.400 px); tinggi kotak tabel dibatasi `100vh - 250px` agar tidak tertutup bar Simpan Batch.

### Input Realisasi & Jadwal: pilih sel lalu Simpan Batch

- **Perilaku berubah.** Tab Input Realisasi dan Jadwal Kerja tidak lagi menyimpan per klik. Klik sel (boleh banyak; Shift+klik untuk rentang di baris yang sama; Esc membatalkan pilihan), pilih kode/shift di panel, perubahan ditahan (sel bergaris putus-putus oranye, hitungan "N perubahan belum disimpan"), lalu **Simpan Batch** mengirim semuanya sekaligus. Popup "pilih status" lama beserta `openPicker`, `submitStatus`, dan markup `statusPicker` dihapus. Route `saveRealisasi` dan `saveJadwal` (satu sel) tetap ada di backend tetapi tidak lagi dipanggil UI.
- **Route baru `saveJadwalBatch(updates, token)`** (khusus Super Admin). Semua-atau-tidak-sama-sekali (satu kode shift salah = tidak ada sel yang ditulis), memakai script lock, maksimum 6.000 sel per batch, sel yang sama dua kali = yang terakhir menang, daftar kosong = tidak melakukan apa pun. Satu batch = **satu** baris Audit Log: aksi `Edit Jadwal (Batch)`, modul `Jadwal`, target `N sel`, detail `Update batch N sel jadwal pada M karyawan (P×5, S×3, ...)`.
- **`saveRealisasiBatch` jauh lebih cepat, hasilnya sama.** Dulu setiap sel membaca area data dua kali (validasi) lalu menulis satu per satu (±3 panggilan sheet per sel). Kini area data dibaca sekali dan hari berurutan di satu baris ditulis dengan satu `setValues`. Audit Log, nilai (`1`/`2` tetap angka), dan pesan error tidak berubah. Helper bersama: `validateCellUpdates_`, `writeGridCells_`; `validateCellTarget_` dipecah menjadi `assertIntegerCell_` + `assertCellInRange_` (perilaku sama).
- Perubahan tertunda dipertahankan saat polling menggambar ulang tabel, dan dibuang otomatis bila: logout, tampilan arsip/periode berubah, role kehilangan izin, atau karyawan di baris itu berubah. Realisasi dan Jadwal punya perubahan tertunda terpisah. Memilih kode yang sama dengan isi sheet tidak dihitung perubahan. Browser memperingatkan bila halaman ditutup dengan perubahan belum disimpan.
- Tombol kode memuat semua kode dari backend, dan `✕` (Kosongkan) selalu di paling kanan. Header kedua tab diseragamkan (`Unduh Template`, `Upload ...`, legenda di baris sendiri). `downloadJadwalTemplate()` kini menerima id tombol; tombol template di tab Input memakai template Jadwal (struktur kolom sama) dan butuh izin menu `templatejadwal`.
- Peta Netlify `GAS_PARAM_MAP_` ditambah `saveJadwalBatch`.

### Modul satelit batch 2: `19_PengajuanIzin.gs` dan `20_AbsenOnline.gs`

**`19_PengajuanIzin.gs`**

1. **Email notifikasi memanggil dirinya sendiri tanpa henti.** `kirimEmailNotifikasiPengajuan_` memanggil dirinya sendiri sehingga mengirim email berulang dan statusnya bisa salah ("Disetujui" padahal ditolak). Sekarang tepat satu email dengan status yang benar.
2. **Approve tanpa menerapkan ke Realisasi.** Pengajuan bisa berstatus Approved tetapi sel jadwal tidak berubah. Sekarang cek status, terapkan, dan tandai dilakukan dalam satu kunci. Validasi dilakukan sebelum menulis apa pun.
3. Penulisan sel satu per satu diganti satu blok sel per rentang hari berurutan (jauh lebih cepat).
4. Approve dan reject kini memakai kunci dan menulis Audit Log.
5. `autoTandaiAlfaHarian_` membaca sel di dalam kunci dan melindungi pengajuan Pending yang tanggalnya bertipe Date.
6. Batas baru: alasan 1.000 karakter, catatan 500, rentang maksimal 62 hari, lampiran 10 MB.

**`20_AbsenOnline.gs`**

1. **`searchKaryawanUntukAbsen` membocorkan NIK** (endpoint publik). Sekarang `nik:''`.
2. **Retensi foto salah hitung di batas bulan** (31 Mei dan 31 Des). Sekarang memakai `new Date(y, m - RETENSI, 1)`.
3. Penghapusan foto lama hanya menandai kolom `FotoUrl`, tidak menyentuh kolom lain.
4. `setFotoProfilKaryawan`: ukuran dicek sebelum membuat file, gagal berbagi tidak membatalkan, memakai kunci.
5. Foto gagal disimpan: absen tetap tercatat dengan `FotoUrl` kosong, bukan error.
6. Validasi GPS pada `submitAbsenOnline`.
7. Batas JamKerja 16 jam.

**Pemasangan batch 2:** ganti isi `19_PengajuanIzin.gs` dan `20_AbsenOnline.gs` (tidak ada file ditambah atau dihapus). `07_TextUtils.gs` tidak perlu diganti (sudah memuat `guardSheetText_` dari batch 1). Deploy versi baru lalu jalankan `auditRoutes()`.

### Modul satelit batch 1: `21_AuditLog.gs`, `15_AIchat.gs`, `23_SettingJamKerja.gs`

**Audit Log**
- **Filter tanggal kehilangan baris dini hari.** `new Date('2026-09-10')` dibaca sebagai 00:00 UTC = 07:00 WIB, jadi baris pukul 00:00–06:59 pada tanggal awal tidak pernah muncul. Pada 6 baris uji di sekitar batas tanggal, versi lama hanya menampilkan 2 dari 4 baris tanggal itu. Kini tanggal dibaca sebagai waktu lokal, dan format tanggal salah memberi pesan jelas (dulu filter diam-diam diabaikan).
- **Teks dari pengguna bisa menjadi rumus.** Nama/target/catatan berawalan `=` ditulis `appendRow` sebagai rumus (mis. `=IMPORTXML(...)`). Kini nilai berawalan `= + - @` disimpan sebagai teks (`guardSheetText_` di `07_TextUtils.gs`, dipakai juga untuk label shift). Detail dipotong 5.000 karakter (sel maksimal 50.000; JSON setting panjang dulu menggagalkan pencatatan diam-diam). Baris yang ditampilkan maksimal 5.000.

**Setting Jam Kerja**
- **Menulis format ±2.000 sel di setiap pembacaan,** padahal `getShiftCodesDynamic_` dipanggil di tiap `getAllData`. Kini hanya memeriksa format baris data dan baris cadangan paling bawah, dan menulis hanya bila ada yang salah. Uji 10 pembacaan berturut-turut: lama menulis >10.000 sel, baru 0. Baris di tengah yang formatnya rusak tetap diperbaiki.
- **Menerima jam sembarang** (`25:99`, `9am`) dan target `abc`/`99`. Kini divalidasi dulu (jam `H:mm`/`HH:mm`, dirapikan jadi `06:05`; target 0–24) dan **semua-atau-tidak-sama-sekali**: bila satu shift salah, tidak ada yang tertulis. Label shift baru dibatasi 40 karakter.
- Simpan kini memakai lock dan hanya menulis baris yang diubah (dulu 4 panggilan per baris, tanpa lock). Tambah Shift juga memakai lock. Target jam tak terbaca dibaca sebagai 8 (dulu `NaN`).

**Chatbot AI**
- **AI tidak tahu jumlah hari periode.** Instruksi "tglAkhir = hari terakhir periode" tanpa data jumlah hari membuat AI menebak (Februari/30 hari salah). Kini `dayCount` ikut dalam konteks, tanggal dari AI dijepit ke 1..dayCount dan dijadikan angka; rentang terbalik dijawab sebagai pesan, bukan konfirmasi kirim.
- **Balasan AI bukan JSON (atau berpagar ```json) menimbulkan galat teknis.** Kini pagar dibuang; JSON rusak memberi pesan jelas; `action` tak dikenal atau `teks` kosong menjadi jawaban biasa. Pesan kosong ditolak dan pesan >1.000 karakter ditolak sebelum memanggil AI. Pesan galat HTTP dipotong 300 karakter. Model bisa diganti lewat Script Property `AI_MODEL` (bawaan `gemini-3.5-flash`).
- **Kirim laporan lewat chatbot:** alamat tujuan harus satu alamat email yang sah (daftar `a@x.com,b@y.com` ditolak, dicek sebelum Excel dibuat), dan pengiriman tercatat di Audit Log (aksi "Kirim Laporan Email", target = alamat tujuan). Hanya Super Admin, seperti sebelumnya.

### Perbaikan pasca-Fase 7

- **`submitAbsenOnline` kini mengecek bulan periode.** Bila periode aktif bukan bulan berjalan, absen tetap tercatat di `LogAbsenOnline` tetapi **tidak** ditulis ke Realisasi/JamScan/JamKerja (dulu menimpa kolom tanggal yang sama di bulan lalu). Kembalian membawa `periodeBelumDiganti` dan `periode`, dan halaman Absen Online menampilkan peringatan. Perilaku identik bila periode sesuai.
- **`getJadwalHariIni` kini mengecek bulan periode.** Bila bulan/tahun periode aktif (A2 sheet Jadwal) ≠ hari ini, halaman Absen Online menampilkan "Periode belum diganti (NAMA PERIODE)" (dulu mis. 1 Oktober dengan periode masih September menampilkan jadwal 1 September). Bila nama periode tak terbaca bulan/tahunnya, pengecekan dilewati (perilaku lama).
- **`AbsensiOnline.html` rusak:** teks README Backup Reliever sempat tertempel di tengah blok `<script>` (±baris 640–1291) sehingga skrip error sintaks dan halaman tidak berfungsi. Bagian yang hilang (pengambilan GPS + pengiriman absen, serta fungsi `pilihMode`) dipulihkan dari tambalan Reliever.

### Fase 8: frontend (versi ringan)

- **`index.html` 327 KB → 82 KB.** Logo (PNG 460×512, 107 KB) tertanam dua kali sebagai base64 (login dan sidebar) = 285 KB. Kini diperkecil ke tinggi 136 px (2× ukuran tampil terbesar, 68 px), warna penuh dan transparansi dipertahankan, selisih piksel 0 terhadap logo asli yang diperkecil. Tetap tertanam (tanpa file luar) sehingga jalan di Apps Script maupun Netlify.
- **`JavaScript.html` tidak dipecah** (tetap satu file; pemecahan menambah belasan `include` dan risiko salah urut tanpa mempercepat apa pun). Ditambah **DAFTAR ISI** di atas blok skrip utama (32 judul bagian). Penanda bagian seragam: `============ NAMA ============`, sub-bagian `---- nama ----`.
- **Kode mati: tidak ada.** Dari 191 fungsi di `JavaScript.html`, tidak satu pun hanya muncul di baris definisinya (dicek terhadap seluruh file frontend); tidak ada nama fungsi ganda dan tidak ada pemanggilan dinamis (`window[...]`, `eval`).
- `GAS_WEBAPP_URL_` (adapter Netlify) diisi URL deployment terbaru. Bila memakai deployment lain, ubah di baris `var GAS_WEBAPP_URL_ = ...` (blok skrip pertama) atau lewat `build.config.json`.
- Peta `GAS_PARAM_MAP_` frontend diverifikasi ulang sama dengan `getRouteParamMap()` backend.

### Fase 7: sisa `Code.gs` dan alat debug

- **`backfillRekapBulananDariArsip_` dan `fixSnapshotAgustus2026_` tidak bisa dijalankan dari editor:** fungsi berakhiran `_` disembunyikan dari dropdown "Jalankan". Kini `backfillRekapBulananDariArsip()` (tanpa garis bawah, perbaikan argumen dari Fase 3 ikut). `fixSnapshotAgustus2026_` (khusus Agustus 2026, hardcode) dihapus karena sudah tercakup backfill.
- **`findBlankRowGaps` melaporkan berlebihan:** mengira sistem berhenti di baris kosong *pertama*, padahal `getLastEmployeeRow_` baru berhenti setelah 5 baris kosong berturut-turut. Jeda 2 atau 4 baris kosong dilaporkan "4 baris hilang" oleh versi lama, padahal tidak ada yang hilang. Kini mengikuti algoritma sebenarnya (konstanta `EMPLOYEE_SCAN_ROWS_` dan `MAX_CONSECUTIVE_BLANK_` dibagi bersama).
- `bersihkanJamNyangkutOFF` kini menghitung juga sel JamScan yang dikosongkan (dulu hanya JamKerja, sehingga pesan "N sel" kurang separuh). Isi sheet hasilnya tidak berubah.
- `debugShiftSheet` tidak lagi crash bila sheet `SettingJamKerja` tidak ada; `testGetJadwalHariIni` memakai karyawan pertama di Jadwal bila tanpa argumen (dulu nama orang tertentu di-hardcode); `detectSuspiciousChars_` kini juga mengenali karakter zero-width.
- `getRekap_` dan `computeRekapDariGrid_` memakai penghitung yang sama (kode ganda hilang), hasil identik.

### Fase 6: laporan dan ekspor

- **Laporan email = bulan lalu, terkirim tanggal 5 jam 07:00.** Ubah `MONTHLY_REPORT_DAY_` / `MONTHLY_REPORT_HOUR_` lalu jalankan ulang `setupMonthlyReportTrigger()`. Sumber data dipilih berlapis: (1) arsip periode bulan lalu bila ada, dengan PDF Dashboard/Rekap diambil dari sheet arsip; (2) periode aktif bila periode aktif itu bulan lalu (belum ditutup); (3) selain itu arsip terbaru (atau periode aktif bila belum ada arsip) dengan **peringatan merah di badan email**. Bila arsip bulan lalu tidak bisa dibaca, otomatis jatuh ke periode aktif dengan peringatan. Analisis tren dibandingkan dengan arsip tepat sebelum periode yang dilaporkan (bukan dengan dirinya sendiri).
- **Analisis tren di email bulanan tidak sebanding.** Rata-rata periode berjalan dihitung sebagai rata-rata persen per karyawan, sedangkan pembandingnya (periode arsip terakhir) dihitung total hadir / total jadwal lalu dibulatkan ke persen bulat. Pada data yang sama: 0,3200 (lama) vs 0,3296 (baru), selisih ±1 poin, lebih besar daripada ambang "STABIL" (0,5 poin). Kini keduanya memakai cara hitung yang sama (juga sama dengan grafik tren `Rekap_Bulanan`). Bila arsip tidak punya baris valid, hasilnya "belum ada data pembanding" (dulu dianggap 0%).
- **Export Excel/PDF generik gagal diam-diam:** bila Google mengembalikan error, versi lama tetap mengirim file rusak (halaman error) sebagai "Excel". Kini error jelas ("Gagal membuat laporan Excel (kode N)"). Spreadsheet sementara dibuang walau export gagal (dulu tertinggal di Drive).
- Input laporan generik divalidasi: kolom kosong/bukan daftar memberi pesan jelas, baris kosong/undefined dianggap tanpa baris.
- `escapeHtmlServer_` dipindah ke `07_TextUtils.gs` karena juga dipakai `19_PengajuanIzin.gs`.
- Pembuatan salinan sheet sementara (template jadwal, PDF per department) kini satu helper `withTempSheetCopy_` di `06_FileUtils.gs`; filter baris per department memakai `keepOnlyDepartmentRows_` yang sama.

### Fase 5: kepegawaian

- **`setKaryawanResign` hanya kerangka:** tidak menulis apa pun ke sheet, tetapi mengembalikan `success:true` dan mencatat "Tandai Resign" di Audit Log (jejak palsu). Tidak ada tombol frontend yang memakainya. Kini menolak dengan pesan jelas dan tidak mencatat apa-apa. Perlu keputusan rancangan fitur (lihat bagian 7).
- **Status Assignment kosong** tampil "Belum Terjadwal" di Daftar Karyawan, tetapi assign menolaknya dengan alasan kosong. Kini diperlakukan sama seperti tampilan (bisa di-assign).
- **Assign nama yang sudah ada di Jadwal/Realisasi** dulu menyisipkan baris duplikat; kini dilewati dengan alasan "sudah ada di sheet Jadwal/Realisasi" (mencegah baris Jadwal dan Realisasi tidak sejajar).
- Assign Karyawan ke Jadwal dan Hapus Karyawan dari Jadwal kini tercatat di Audit Log (sebelumnya tidak).
- `setEmployeeNik` (tombol "Isi NIK") mengembalikan state terbaru tanpa cache 8 detik, supaya NIK langsung terlihat.
- **Performa:** `getOrCreateNikSheet_` dipanggil di setiap pembacaan data karyawan (termasuk tiap `getAllData`) dan dulu selalu menulis format teks ke seluruh kolom No HP (±1.000 sel). Kini hanya menulis bila header berubah atau baris paling bawah belum berformat teks.
- Kode mati dihapus: `ensureDataKaryawanColumns_` (skema 6 kolom usang) dan `getUnassignedEmployees_`.

### Fase 4: upload dan finger

- Upload Jadwal, Upload Realisasi, dan Upload Finger sebelumnya tidak tercatat di Audit Log. Upload Jadwal/Realisasi kini mencatat jumlah karyawan cocok, nama tidak ditemukan, dan kode tidak valid. (Upload Finger belum dicatat; menunggu keputusan.)
- Sheet kosong (tanpa karyawan) memberi pesan jelas ("Sheet X belum berisi karyawan"), bukan error teknis "number of rows must be at least 1".
- Hak akses upload dicek **sebelum** daftar kode shift dibaca, jadi akun non-Super Admin ditolak tanpa menyentuh sheet.
- Sinkron JamKerja/JamScan tidak lagi gagal bila Realisasi belum punya karyawan.

### Fase 3: periode dan arsip

- **`backfillRekapBulananDariArsip_` salah urutan argumen:** memanggil snapshot dengan (sheet Rekap arsip, sheet Jadwal arsip) padahal yang diminta (Jadwal, Realisasi), sehingga menulis baris bulan berisi nol semua. Lama menghasilkan `[2026,10,0,0,0,0,0,0,0,0]`, baru `[2026,10,36,19,...]` (sama dengan snapshot normal).
- `getMonthlyTrendYear` mengabaikan baris dengan nilai bulan tidak valid (sebelumnya menulis properti aneh di array). Duplikasi kode pembacaan `ArsipIndex` (5 tempat) jadi satu helper `findArchiveRow_`.

### Fase 1–2: fondasi, data API, input

- **`dayIndex` berupa string (Fase 2):** `saveRealisasi`/`saveJadwal` lama menghitung kolom `4 + '5'` = `'45'` (kolom 45, bukan 9) dan menulis log "Hari ke-51". Kini dipaksa jadi angka. `saveRealisasiBatch` juga menolak input non-daftar dan tetap all-or-nothing.
- **`setEmployeeNik` tidak ada:** tombol "Isi NIK" di frontend memanggilnya, tetapi fungsinya tidak pernah didefinisikan. Sekarang pembungkus `setEmployeeMasterData`.
- **Route `getPengajuanPendingCount` tidak ada** di `doPost`, padahal frontend memanggilnya (badge pengajuan). Ditambahkan.
- **Peta `GAS_PARAM_MAP_` frontend tidak lengkap:** `deleteUnassignedEmployee` dan `addShiftJamKerja` dipanggil frontend tetapi tidak terdaftar (error di mode Netlify). Peta kini dibuat dari `ROUTES`.
- **`16_RepairAlignment.gs` rusak** (komentar dan fungsi menempel dalam satu baris) dan menduplikasi `checkRowAlignment_`. Menjadi `16_Rekap.gs`.
- **`24_BackupReliever.gs`** mendefinisikan dua fungsi test dua kali. Salinan kedua dihapus.
- **Notifikasi lonceng salah baca kolom:** mengasumsikan `[Waktu,User,Aksi,Modul,Detail]` padahal AuditLog `[Timestamp,Username,Nama,Role,Aksi,Modul,Target,Detail]`, dan memakai kolom `TanggalPengajuan` yang tidak ada (sebenarnya `DiajukanPada`). Kini membaca lewat nama header. Fungsi mati `getNotifications_` (memanggil `validateToken_` yang tidak ada) dihapus.
- **Rahasia di kode:** `PASSWORD_SALT` dan `REPORT_EMAIL` kini dibaca dari Script Properties (`LEGACY_PASSWORD_SALT`, `REPORT_EMAIL`). Fungsi `resetPasswordSaya()` (password tertulis di kode) dan fungsi dedupe sekali-jalan dihapus. Lanjutannya ada di bagian 5.

---

## 4. Cara memasang

### Pemasangan awal (dari nol)

1. Di Apps Script: **hapus `Code.gs`** dan semua file `.gs` lama, lalu tambahkan semua file backend. Jangan biarkan dua versi hidup bersamaan (deklarasi ganda akan error).
2. Ganti isi `appsscript.json` dengan versi di repo.
3. Frontend: salin file `.html`. Isi ulang `GAS_WEBAPP_URL_` bila memakai Netlify (atau gunakan `build.js`).
4. Jalankan `auditRoutes()`; hasilnya harus bersih.
5. Uji: login, buka Dashboard, input realisasi, Setting > Account, Setting > Hak Akses, lonceng notifikasi.
6. (Disarankan) Isi Script Properties: `LEGACY_PASSWORD_SALT`, `REPORT_EMAIL`, `GEMINI_API_KEY`.

> Nama file template di `HTML_ENTRY_FILE` adalah `Index` (huruf besar). Di repo berkas bernama `index.html`; pastikan nama di editor Apps Script sama persis dengan konstanta tersebut.

### Pembaruan antar fase (arsip)

Langkah-langkah ini mencatat cara pembaruan di masing-masing fase, dari yang terbaru. Berguna bila Anda harus mengurai proyek yang masih berada di fase lama.

<details>
<summary><b>Input Realisasi & Jadwal Batch (backend dan frontend)</b></summary>

1. Backend: ganti isi `01_Router.gs` dan `11_InputData.gs`.
2. Frontend: ganti isi `index.html`, `JavaScript.html`, `Stylesheet.html`. Pastikan `GAS_WEBAPP_URL_` di `JavaScript.html` masih berisi URL deployment Anda.
3. Jalankan `auditRoutes()`. Terapkan ulang Web App (Deploy > Manage deployments > Edit > New version) dan deploy ulang Netlify bila dipakai. Bila sebelumnya sudah memasang paket "jadwal batch" tanpa perbaikan tata letak, cukup ganti `Stylesheet.html`.
4. Uji di salinan: (a) tab Input Realisasi: pilih beberapa sel, beri kode, Simpan Batch, cek di sheet Realisasi; (b) tab Jadwal Kerja sebagai Super Admin: pilih rentang sel, beri shift, Simpan Batch, cek di sheet Jadwal; (c) Setting > Audit Log: satu batch tampil sebagai satu baris `Edit Jadwal (Batch)` / `Edit Realisasi (Batch)`; (d) login sebagai role non-admin: tab Jadwal tidak boleh bisa diubah; (e) logout dengan perubahan belum disimpan: perubahan dibuang.
</details>

<details>
<summary><b>Fase 8 → modul satelit batch 1 (backend saja)</b></summary>

1. Ganti isi `07_TextUtils.gs`, `15_AIchat.gs`, `21_AuditLog.gs`, `23_SettingJamKerja.gs`.
2. Jalankan `auditRoutes()`, lalu uji: Setting > Jam Kerja (ubah jam, tambah shift, coba jam `25:00`), Setting > Audit Log (filter tanggal hari ini dan Modul), dan chatbot (tanya satu pertanyaan; minta kirim email ke diri sendiri lalu cek Audit Log).
3. Opsional: isi Script Property `AI_MODEL`.
</details>

<details>
<summary><b>Fase 7 → Fase 8 (frontend saja)</b></summary>

1. Ganti `index.html` dan `JavaScript.html`. `Stylesheet.html` dan `SettingJadwalKerja.html` tidak berubah.
2. Terapkan ulang Web App / deploy ulang Netlify.
3. Uji: login gate tampil dengan logo, login, Dashboard, buka satu tab tiap menu. Bila logo di layar sangat besar terlihat kurang tajam, tinggi 136 px bisa dinaikkan (perlu mengganti gambar tertanam).
</details>

<details>
<summary><b>Fase 6 → Fase 7</b></summary>

1. **Hapus `05_LegacyCode.gs`** (isinya sudah habis dipindah; bila dibiarkan, fungsi di dalamnya bentrok).
2. Tambahkan `32_JadwalHariIni.gs` dan `99_Maintenance.gs`.
3. Ganti isi `02_SheetUtils.gs` dan `16_Rekap.gs`.
4. Jalankan `auditRoutes()`, lalu uji: halaman Absen Online (jadwal shift hari ini muncul), Dashboard dan Rekap Karyawan, serta pilih satu periode arsip (kolom Rank dan angka sama seperti sebelumnya).
</details>

<details>
<summary><b>Fase 5 → Fase 6</b></summary>

1. Tambahkan `07_TextUtils.gs`, `29_Export.gs`, `30_AttendanceLog.gs`, `31_MonthlyReport.gs`.
2. Ganti isi `05_LegacyCode.gs`, `06_FileUtils.gs`, `17_UploadGrid.gs`.
3. `appsscript.json` tidak berubah. **Laporan bulanan kini untuk bulan lalu dan terkirim tanggal 5 jam 07:00.** Trigger lama (tanggal 1) tetap aktif sampai Anda menjalankan `setupMonthlyReportTrigger()` sekali dari editor; fungsi itu menghapus trigger lama lalu memasang yang baru.
4. Jalankan `auditRoutes()`, lalu uji: Export PDF (Rekap/Dashboard/Jadwal/Realisasi, dengan dan tanpa department), Report > Export Excel/PDF, Log Absensi (periode aktif dan arsip). Jalankan `sendMonthlyReportEmail` dari editor hanya bila `REPORT_EMAIL` di salinan sudah Anda atur sendiri (email sungguhan terkirim).
</details>

<details>
<summary><b>Fase 4 → Fase 5</b></summary>

1. **Hapus dulu `21_Resign.gs` dan `22_AssignJadwal.gs`** (digantikan `28_EmployeeLifecycle.gs`; bila tidak dihapus akan error deklarasi ganda).
2. Tambahkan `27_EmployeeData.gs` dan `28_EmployeeLifecycle.gs`.
3. Ganti isi `01_Router.gs` dan `05_LegacyCode.gs`.
4. Jalankan `auditRoutes()`, lalu uji: Daftar Karyawan (tambah, edit NIK/department), Assign Karyawan ke Jadwal, Hapus Karyawan (tanpa dan dengan "Hapus Paksa"), dan hapus karyawan yang belum terjadwal.
</details>

<details>
<summary><b>Fase 3 → Fase 4</b></summary>

1. Tambahkan `06_FileUtils.gs`, `17_UploadGrid.gs`, `25_FingerAbsen.gs`, `26_JamKerja.gs`.
2. Ganti isi `05_LegacyCode.gs` (kode upload/finger/template/jam kerja lama sudah dihapus dari situ; bila tidak diganti akan error deklarasi ganda).
3. `appsscript.json` tidak berubah.
4. Jalankan `auditRoutes()`, lalu uji: Download Template Excel (Jadwal Kerja), Upload Jadwal, Upload Realisasi, Upload Finger Absen dengan file contoh kecil.
</details>

<details>
<summary><b>Fase 1–2 → Fase 3</b></summary>

1. **Hapus `20_TrendComparison.gs`** (digantikan `14_MonthlyTrend.gs`; bila tidak dihapus akan error deklarasi ganda).
2. Tambah `12_Period.gs`, `13_Archive.gs`, `14_MonthlyTrend.gs`.
3. Ganti isi `01_Router.gs` dan `05_LegacyCode.gs`.
4. Jalankan `auditRoutes()`, lalu uji: Periode Aktif ▾ pilih arsip, Arsipkan Periode Ini, Mulai Periode Baru.
</details>

## 5. Menghapus rahasia dari kode

Repositori ini publik, jadi rahasia tidak boleh ada di kode. Yang berubah: `00_Config.gs` (salt hash lama dan email penerima laporan tidak lagi tertulis di kode, hanya dibaca dari Script Properties) dan file baru `98_SecurityCheck.gs` (alat diagnosis yang tidak menampilkan nilai rahasia).

**Urutan pemasangan (urutan penting, jangan dibalik):**

1. Apps Script produksi > Project Settings > **Script Properties**, tambahkan:
   - `LEGACY_PASSWORD_SALT` = nilai salt lama (yang sempat terlihat di `00_Config.gs` publik)
   - `REPORT_EMAIL` = alamat penerima laporan bulanan
   - `GEMINI_API_KEY` bila chatbot dipakai (sudah tidak ada di kode)
2. Tambahkan `98_SecurityCheck.gs`, lalu jalankan `periksaRahasiaScriptProperties()`. Log harus menunjukkan semua yang Anda pakai "terisi".
3. Ganti isi `00_Config.gs` dengan versi baru. Jalankan `auditRoutes()` (harus bersih). Uji login satu akun lama dan satu akun baru.
4. Jalankan `periksaAkunBelumTermigrasi()`. Daftar itu = akun yang masih memakai hash lama; mereka otomatis pindah ke hash per-user saat login berikutnya. Akun yang tidak akan login lagi: reset password lewat Setting > Account.
5. Setelah daftar kosong, ganti nilai `LEGACY_PASSWORD_SALT` dengan teks acak panjang. Salt lama yang sempat publik jadi tidak berguna.

**Satu langkah lagi di GitHub:** ubah repositori menjadi **Private** (Settings > General > Danger Zone > Change visibility). Riwayat commit lama tetap memuat nilai lama selama repositori publik; langkah 5 menetralkannya. Netlify tetap bisa membangun dari repositori private lewat integrasi GitHub-nya.

## 6. Hasil uji

Semua uji memakai tiruan Apps Script di Node atau browser headless/jsdom. **Belum ada yang diuji di Apps Script dan spreadsheet sungguhan.** Uji kode lama vs baru hanya membuktikan perilaku identik pada skenario yang dicoba; selisih yang disengaja diuji terpisah.

| Cakupan | Jenis uji | Hasil |
|---|---|---|
| Input Realisasi & Jadwal Batch, backend | Tiruan Node, 1.025 pemeriksaan | `validateCellTarget_` lama = baru pada 324 kombinasi masukan (string, desimal, NaN, null); `saveRealisasiBatch` lama = baru pada 60 batch acak (isi sheet, nilai angka 1/2, kembalian, Audit Log identik; panggilan sheet lebih sedikit) dan 10 kasus error (pesan sama, tidak ada yang ditulis, lock dilepas); `saveJadwalBatch` = `saveJadwal` berurutan pada 60 batch acak (isi identik, tepat satu catatan Audit Log, lock dipakai sekali dan dilepas); 9 kasus penolakan; sel ganda, kode kosong, daftar kosong; 930 sel berurutan = 30 panggilan `setValues` dan 2 pembacaan sheet (lama ±3 panggilan per sel) |
| Input Realisasi & Jadwal Batch, tampilan | Chrome headless, 57 pemeriksaan | Pilih sel, Shift+klik, kode sama tidak dihitung, polling tidak menghapus perubahan, Simpan Batch satu panggilan dengan token benar, gagal simpan mempertahankan perubahan, baris bergeser membuang perubahan, Batalkan, Esc, arsip, role tanpa izin, Realisasi dan Jadwal terpisah, popup lama hilang, logout; tanpa error JavaScript |
| Tata letak tabel | Chrome headless, 31 hari × 60 karyawan, tiga lebar layar | Sebelum perbaikan halaman melebar 274 px (laptop 1366) dan 1.010 px (HP 390); sesudahnya 0 px di semua ukuran. Sapuan 19 tab × 3 lebar (57 kombinasi): 8 membaik, 49 sama, 0 memburuk |
| Modul satelit batch 1 | Tiruan Node, 26 pemeriksaan | Audit Log, Setting Jam Kerja, dan Chatbot identik dengan versi lama pada skenario umum; selisih disengaja diuji terpisah. Belum diuji di Google Sheets sungguhan: perilaku apostrof di depan teks (`'=...`). Cek sekali dengan menambah shift berlabel `=1+1` di salinan (label harus tampil `=1+1`, bukan `2`) |
| Modul satelit batch 2 | Tiruan Node | `19_PengajuanIzin` 25 pengecekan, `20_AbsenOnline` 18 pengecekan |
| Fase 8, `index.html` rakitan | jsdom | Halaman termuat tanpa galat skrip; logo login dan sidebar data PNG < 30 KB; adapter Netlify aktif dan menerjemahkan `login(username, password)` ke permintaan aksi `login`; seluruh blok `<script>` lolos cek sintaks; `GAS_PARAM_MAP_` = route backend |
| `submitAbsenOnline` (pasca-Fase 7) | 800 skenario acak | Identik dengan versi sebelumnya bila periode sesuai; 4 skenario periode tidak sesuai × Masuk/Pulang: Realisasi/JamScan/JamKerja tidak berubah, log tetap tercatat, kembalian membawa penanda |
| `AbsensiOnline.html` | jsdom, 19 pemeriksaan | Login karyawan, tab Reliever (pencarian, pilih nama, dropdown "menggantikan siapa"), penolakan tanpa pilihan, isi kiriman `submitAbsenOnlineReliever`, logout, pesan periode belum diganti |
| `getJadwalHariIni` (pasca-Fase 7) | 3.000 kombinasi acak | Identik bila periode = bulan berjalan; 5 skenario periode tidak sesuai menghasilkan "Periode belum diganti"; nama periode tak terbaca, sheet hilang, dan header tak dikenal tetap seperti sebelumnya |
| Fase 7 | 298 pemeriksaan | Rekap aktif (40 kasus) dan rekap arsip (150 kasus acak) identik; `getJadwalHariIni` 48 kombinasi identik; tiga alat perbaikan dan `debugNameMismatch` identik pada 5 variasi sheet; `findBlankRowGaps` cocok dengan perhitungan independen; backfill sama dengan snapshot normal |
| Fase 6 | 165 pemeriksaan | Termasuk 8 skenario "laporan bulan lalu" (arsip ada, belum diarsipkan tapi aktif, tidak ada keduanya, arsip rusak, pergantian tahun, lampiran gagal); export PDF 14 kombinasi; laporan Excel/PDF generik; log absensi 11 kombinasi; email bulanan 14 skenario berikut HTML-nya; pemasangan trigger |
| Fase 5 | 198 perbandingan | Lapisan DataKaryawan dari nol/header salah, edit master data, tambah/assign/hapus, hak akses, hasil di semua sheet termasuk format. Selisih yang disengaja (status assignment kosong, assign duplikat, resign) diperiksa terpisah |
| Fase 4 | 503 perbandingan | Impor grid acak ×120, terapkan finger acak ×160, parsing file finger, sinkron JamKerja/JamScan, pengayaan rekap, template per department, upload end-to-end (isi request ke Drive, pembuangan file sementara, penolakan non-Super Admin). Satu-satunya selisih = catatan Audit Log baru |
| Fase 2 | 603 perbandingan acak | Tren harian, ringkasan dashboard, peringatan alignment identik; `syncOffFromJadwal_`, `saveRealisasi`, `saveRealisasiBatch` identik (selisih tunggal = perbaikan `dayIndex`) |
| Umum | Sintaks dan dispatcher | Semua file `.gs` lolos cek sintaks dan dimuat bersama tanpa deklarasi ganda; logika role (termasuk "admin" lama vs "Admin" baru), hash password skema baru dan lama, dan dispatcher route diuji dengan mock |

## 7. Temuan yang menunggu keputusan

1. **Tanggal tanpa cek bulan — sudah diperbaiki.** `getJadwalHariIni` dan `submitAbsenOnline` kini mengecek bulan periode (lihat Perbaikan pasca-Fase 7). Yang masih terbuka: absen yang tercatat di log selama periode belum diganti perlu diproses admin; belum ada alat otomatis untuk memasukkannya ke Realisasi setelah periode diganti. Modul Backup Reliever (`24`) sudah punya pengecekan serupa (`getKaryawanAbsenHariIni` mengembalikan daftar kosong bila bulan periode ≠ bulan sekarang); pengamannya tinggal diseragamkan.
2. **Arti kolom Rank berbeda** antara periode aktif (`getRekap_`: Rank 1 = % **terendah**, sesuai label "Rank (Terendah)") dan periode arsip (`computeRekapDariGrid_`: Rank 1 = % **tertinggi**). Frontend menampilkan angkanya apa adanya. Perilaku warisan dipertahankan; perlu diputuskan satu aturan.
3. **Tandai Resign** (lihat Fase 5): perlu rancangan fitur.
4. **Batas 16 jam JamKerja** (batch 2): apakah sesuai kebijakan?
5. **Endpoint publik Absen Online:** NIK tetap menjadi satu-satunya rahasia dan belum ada pembatasan percobaan (rate limit).
6. **Pengajuan Izin masuk matriks hak akses?** Sekarang `getPengajuanList`/`approvePengajuan`/`rejectPengajuan`/`getPengajuanPendingCount` dibatasi Super Admin saja, padahal menu `pengajuanizin` ada di `MENU_LIST`.
7. **Batas form izin** (10 MB, 62 hari, 1.000 karakter): sesuaikan jika perlu.
8. **Hak akses chatbot untuk Viewer** (dari batch 1).
9. **Upload Finger** belum dicatat di Audit Log (Fase 4).

## 8. Status fase

1. ~~Data API~~ selesai (Fase 2)
2. ~~Input~~ selesai (Fase 2)
3. ~~Periode & arsip~~ selesai (Fase 3)
4. ~~Upload & finger~~ selesai (Fase 4)
5. ~~Kepegawaian~~ selesai (Fase 5)
6. ~~Laporan~~ selesai (Fase 6)
7. ~~Sisa `Code.gs` + alat debug~~ selesai (Fase 7)
8. ~~Frontend~~ selesai (Fase 8 versi ringan: logo diperkecil, daftar isi, tanpa memecah file)
9. **Modul satelit** — batch 1 dan 2 selesai (Audit Log, Chatbot AI, Setting Jam Kerja, Pengajuan Izin, Absen Online); sisa: `18_PolaJadwalGenerator`, `24_BackupReliever`
