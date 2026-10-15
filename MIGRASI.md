# AMS v2 — Catatan Migrasi

Status: **Fase 1-8 selesai.** Seluruh isi `Code.gs` lama sudah ditulis ulang, `05_LegacyCode.gs` dihapus, dan frontend sudah dibersihkan (Fase 8 versi ringan; `JavaScript.html` sengaja TETAP SATU FILE). Modul satelit: 3 dari 7 sudah ditulis ulang (Audit Log, Chatbot AI, Setting Jam Kerja); 4 lagi belum (lihat bagian "Belum ditulis ulang").

## Yang sudah ditulis ulang (Fase 1)

| File baru | Isi | Pengganti |
|---|---|---|
| `00_Config.gs` | CONFIG, daftar menu, konstanta auth/role | Code.gs bagian 00 |
| `01_Router.gs` | Tabel `ROUTES` (66 aksi), `dispatchRoute_`, `doGet/doPost`, `auditRoutes()`, `getRouteParamMap()` | Code.gs bagian 01 (switch ±200 baris) |
| `02_SheetUtils.gs` | `getSheet_`, `readGrid_`, `withScriptLock_`, `checkRowAlignment_` | bagian 03 + duplikat di 16 |
| `03_Auth.gs` | login, session, hash, rate limit, ganti password | bagian 02 |
| `04_Access.gs` | user, role, matrix hak akses | bagian 17 |
| `10_DataApi.gs` | `getAllData` (+cache), tren harian, ringkasan dashboard, sinkron OFF | bagian 04 + patch performa |
| `11_InputData.gs` | `saveRealisasi`, `saveRealisasiBatch`, `saveJadwal`, validasi sel | bagian 05 |
| `12_Period.gs` | parsing nama periode, `mulaiPeriodeBaru`, reset sheet bulan baru | bagian 08 + util tanggal di bagian 10 |
| `13_Archive.gs` | `archiveCurrentPeriod`, `ArsipIndex`, baca/hapus arsip, `getPeriodMeta` | bagian 08 + `readArchivedJadwalGrid_` |
| `14_MonthlyTrend.gs` | snapshot `Rekap_Bulanan`, `getMonthlyTrendYear` | **`20_TrendComparison.gs` (dihapus)** |
| `06_FileUtils.gs` | helper Drive/Excel bersama: unggah-konversi file sementara (selalu dibuang), ekspor xlsx | kode ganda di bagian 07, 11, 13 |
| `17_UploadGrid.gs` | `uploadJadwal`, `uploadRealisasi`, `generateJadwalTemplateExcel` | bagian 07 + 11 |
| `25_FingerAbsen.gs` | `uploadFingerAbsen`, parsing file mesin finger, terapkan ke Realisasi/JamKerja/JamScan | bagian 13 (bagian upload) |
| `26_JamKerja.gs` | sheet tersembunyi JamKerja/JamScan, `attachJamKerjaToRekap_`, `attachNikToRekap_`, ringkasan jam | bagian 13 (bagian jam kerja) |
| `27_EmployeeData.gs` | sheet `DataKaryawan`: master data, daftar karyawan, `setEmployeeMasterData`, `setEmployeeNik` | bagian 15_Kepegawaian |
| `28_EmployeeLifecycle.gs` | `addEmployee`, `assignEmployeeToJadwal`, `deleteEmployee`, `deleteUnassignedEmployee`, `setKaryawanResign` | `addEmployee` + **`21_Resign.gs`, `22_AssignJadwal.gs` (dihapus)** |
| `07_TextUtils.gs` | `escapeHtmlServer_` (dipakai email laporan dan email pengajuan izin) | bagian 09 |
| `29_Export.gs` | `exportSheetPdf`, `generateReportExcel`, `generateReportPdf` | bagian 06_ExportPdf + laporan generik |
| `30_AttendanceLog.gs` | `generateAttendanceLogExcel`, `readPeriodGrids_`, parser scan | bagian 10_AttendanceLog |
| `31_MonthlyReport.gs` | `sendMonthlyReportEmail` (laporan **bulan lalu**, tanggal 5 jam 07:00), pemilihan sumber periode, analisis tren, HTML email, `setupMonthlyReportTrigger` | bagian 09_MonthlyReport |
| `16_Rekap.gs` | `getRekap_` (periode aktif) dan `computeRekapDariGrid_` (arsip/tren) dengan penghitung bersama `tallyAttendance_` | `16_RepairAlignment.gs` + bagian 08 |
| `32_JadwalHariIni.gs` | `getJadwalHariIni` (route publik Absen Online) | sisa `Code.gs` |
| `99_Maintenance.gs` | 9 alat perawatan/diagnosis yang dijalankan manual dari editor (`debugNameMismatch`, `findBlankRowGaps`, `auditJamNyangkutOFF`, `bersihkanJamNyangkutOFF`, `fixJamKerjaPeriodeAktifSekarang`, `backfillRekapBulananDariArsip`, dst.) | bagian 14_DebugNama + fungsi perbaikan sekali-jalan |

Perilaku sengaja dipertahankan: nama route, urutan parameter, bentuk respons `{success,data}/{success,error}`, skema hash password (akun lama tetap bisa login), logika RBAC.

## Bug yang ditemukan dan sudah diperbaiki

0000000. **(Tata letak) Halaman melebar di layar sempit, kolom tanggal terpotong:** `main` memakai `margin: 0 auto` di dalam wadah flex kolom (`.main-area`), sehingga lebarnya mengikuti lebar isi tabel (31 kolom, teks tidak boleh turun baris), bukan lebar layar; akibatnya seluruh halaman bisa digeser ke samping (di HP sampai ±1.000 px) dan tabel tidak punya batas untuk menggulir di dalam kotaknya. Perbaikan di `Stylesheet.html`: `main { width: 100%; min-width: 0 }`; tabel Jadwal/Realisasi dirapatkan (kolom tanggal lebih sempit, Jabatan dipotong dengan "..." dan disembunyikan di HP); tab Jadwal/Input boleh selebar 1.760 px (memakai `:has()`; browser lama tetap memakai batas 1.400 px); tinggi kotak tabel dibatasi `100vh - 250px` agar tidak tertutup bar Simpan Batch.

0000000. **(Input Realisasi & Jadwal, pilih sel lalu Simpan Batch) Perubahan perilaku dan perbaikan:**
   - **Perilaku BERUBAH:** tab **Input Realisasi** dan **Jadwal Kerja** tidak lagi menyimpan per klik. Klik sel (boleh banyak, Shift+klik untuk rentang di baris yang sama, Esc membatalkan pilihan), pilih kode/shift di panel, perubahan ditahan (sel bergaris putus-putus oranye, hitungan "N perubahan belum disimpan"), lalu **Simpan Batch** mengirim semuanya sekaligus. Popup "pilih status" lama beserta `openPicker`, `submitStatus`, dan markup `statusPicker` **dihapus**. Route `saveRealisasi` dan `saveJadwal` (satu sel) tetap ada di backend tetapi tidak lagi dipanggil UI.
   - **Route baru `saveJadwalBatch(updates, token)`** (khusus Super Admin; total 66 route). Semua-atau-tidak-sama-sekali (satu kode shift salah = tidak ada sel yang ditulis), memakai script lock, maksimum 6.000 sel per batch, sel yang sama dua kali = yang terakhir menang, daftar kosong = tidak melakukan apa pun. **Satu batch = SATU baris Audit Log**: aksi `Edit Jadwal (Batch)`, modul `Jadwal`, target `N sel`, detail `Update batch N sel jadwal pada M karyawan (P×5, S×3, ...)`.
   - **`saveRealisasiBatch` jauh lebih cepat, hasilnya sama.** Sebelumnya setiap sel membaca area data dua kali (validasi) lalu menulis satu per satu (±3 panggilan sheet per sel). Kini area data dibaca sekali dan hari yang berurutan di satu baris ditulis dengan satu `setValues`. Audit Log, nilai (`1`/`2` tetap angka), dan pesan error tidak berubah. Helper bersama: `validateCellUpdates_`, `writeGridCells_`; `validateCellTarget_` dipecah menjadi `assertIntegerCell_` + `assertCellInRange_` (perilaku sama).
   - Perubahan tertunda dipertahankan saat polling menggambar ulang tabel, dan dibuang otomatis bila: logout, tampilan arsip/periode berubah, role kehilangan izin, atau karyawan di baris itu berubah. Realisasi dan Jadwal punya perubahan tertunda yang terpisah. Memilih kode yang sama dengan isi sheet tidak dihitung perubahan. Browser memperingatkan bila halaman ditutup dengan perubahan belum disimpan.
   - Tombol kode memuat semua kode dari backend, dan `✕` (Kosongkan) selalu di paling kanan. Header kedua tab diseragamkan (`Unduh Template`, `Upload ...`, legenda di baris sendiri). `downloadJadwalTemplate()` kini menerima id tombol; tombol template di tab Input memakai template Jadwal (struktur kolom sama) dan butuh izin menu `templatejadwal`.
   - Peta Netlify `GAS_PARAM_MAP_` ditambah `saveJadwalBatch`.

0000000. **(Modul satelit batch 1: `21_AuditLog.gs`, `15_AIchat.gs`, `23_SettingJamKerja.gs`) Temuan dan perubahan perilaku:**
   - **Filter tanggal Audit Log kehilangan baris dini hari.** `new Date('2026-09-10')` dibaca sebagai 00:00 UTC = 07:00 WIB, jadi baris tanggal awal pukul 00:00-06:59 tidak pernah muncul. Di uji: 6 baris sekitar batas tanggal, versi lama hanya menampilkan 2 dari 4 baris tanggal itu. Kini tanggal dibaca sebagai waktu lokal. Format tanggal salah kini memberi pesan jelas (dulu filter diam-diam diabaikan).
   - **Audit Log: teks dari pengguna bisa menjadi rumus.** Nama/target/catatan berawalan `=` ditulis `appendRow` sebagai rumus (mis. `=IMPORTXML(...)`). Kini nilai berawalan `= + - @` disimpan sebagai teks (helper `guardSheetText_` di `07_TextUtils.gs`, dipakai juga untuk label shift). Detail dipotong 5.000 karakter (sel maksimal 50.000; JSON setting panjang dulu menggagalkan pencatatan diam-diam). Jumlah baris yang ditampilkan maksimal 5.000.
   - **Setting Jam Kerja menulis format ±2.000 sel di setiap pembacaan**, padahal `getShiftCodesDynamic_` dipanggil di tiap `getAllData`. Kini hanya memeriksa format baris data dan baris cadangan paling bawah, menulis hanya bila ada yang salah. Di uji: 10 pembacaan berturut-turut, lama menulis >10.000 sel, baru 0. Baris di tengah yang formatnya rusak tetap diperbaiki.
   - **Setting Jam Kerja menerima jam sembarang** (`25:99`, `9am`) dan target `abc`/`99`, lalu menyimpannya. Kini divalidasi dulu (jam `H:mm`/`HH:mm`, dirapikan jadi `06:05`; target 0-24) dan **semua-atau-tidak-sama-sekali**: bila satu shift salah, tidak ada yang tertulis. Label shift baru dibatasi 40 karakter.
   - Simpan Setting Jam Kerja kini memakai lock dan hanya menulis baris yang diubah (dulu 4 panggilan per baris, tanpa lock); baris yang tidak diubah tidak disentuh sama sekali. Tambah Shift juga memakai lock. Target jam tak terbaca di sheet dibaca sebagai 8 (dulu `NaN`).
   - **Chatbot: AI tidak tahu jumlah hari periode.** Instruksi "tglAkhir = hari terakhir periode" tanpa data jumlah hari membuat AI menebak (Februari/30 hari salah). Kini `dayCount` ikut dalam konteks, dan tanggal dari AI dijepit ke 1..dayCount serta dijadikan angka; rentang terbalik dijawab sebagai pesan, bukan konfirmasi kirim.
   - **Chatbot: balasan AI yang bukan JSON (atau berpagar ```json) membuat galat teknis.** Kini pagar dibuang; JSON rusak memberi pesan jelas; `action` tak dikenal atau `teks` kosong menjadi jawaban biasa. Pesan kosong ditolak dan pesan >1.000 karakter ditolak sebelum memanggil AI. Pesan galat HTTP dipotong 300 karakter. Model bisa diganti lewat Script Property `AI_MODEL` (bawaan tetap `gemini-3.5-flash`).
   - **Kirim laporan lewat chatbot**: alamat tujuan kini harus satu alamat email yang sah (daftar `a@x.com,b@y.com` ditolak, dicek sebelum Excel dibuat), dan pengiriman **tercatat di Audit Log** (aksi "Kirim Laporan Email", target = alamat tujuan). Hanya Super Admin, seperti sebelumnya.

000000. **(Fase 7) Temuan dan perubahan perilaku:**
   - **`backfillRekapBulananDariArsip_` dan `fixSnapshotAgustus2026_` tidak bisa dijalankan dari editor**: fungsi berakhiran `_` disembunyikan dari dropdown "Jalankan". Kini `backfillRekapBulananDariArsip()` (tanpa garis bawah, perbaikan argumen dari Fase 3 ikut). `fixSnapshotAgustus2026_` (khusus Agustus 2026, hardcode) dihapus karena sudah tercakup backfill.
   - **`findBlankRowGaps` melaporkan berlebihan**: mengira sistem berhenti di baris kosong *pertama*, padahal `getLastEmployeeRow_` baru berhenti setelah 5 baris kosong berturut-turut. Di uji: jeda 2 atau 4 baris kosong dilaporkan "4 baris hilang" oleh versi lama, padahal tidak ada yang hilang. Kini mengikuti algoritma sebenarnya (konstanta `EMPLOYEE_SCAN_ROWS_` dan `MAX_CONSECUTIVE_BLANK_` dibagi bersama).
   - `bersihkanJamNyangkutOFF` kini menghitung juga sel JamScan yang dikosongkan (dulu hanya JamKerja, sehingga pesan "N sel" kurang separuh). Isi sheet hasilnya tidak berubah.
   - `debugShiftSheet` tidak lagi crash bila sheet `SettingJamKerja` tidak ada; `testGetJadwalHariIni` memakai karyawan pertama di Jadwal bila tanpa argumen (dulu nama orang tertentu yang di-hardcode); `detectSuspiciousChars_` kini juga mengenali karakter zero-width.
   - `getRekap_` dan `computeRekapDariGrid_` memakai penghitung yang sama (kode ganda hilang), hasil identik.
   - **(Fase 8, frontend versi ringan)**
     - **`index.html` 327 KB → 82 KB**: logo (PNG 460×512, 107 KB) tertanam dua kali sebagai base64 (login & sidebar) = 285 KB. Kini diperkecil ke 136 px tinggi (2× ukuran tampil terbesar, 68 px), warna penuh dan transparansi dipertahankan, selisih piksel 0 terhadap logo asli yang diperkecil; tetap tertanam (tanpa file luar) sehingga jalan di Apps Script maupun Netlify. Tidak ada baris lain yang berubah.
     - **`JavaScript.html` tidak dipecah** (keputusan: tetap satu file; pemecahan menambah belasan `include` dan risiko salah urut tanpa mempercepat apa pun). Ditambah **DAFTAR ISI** di bagian atas blok skrip utama (32 judul bagian, cari dengan Ctrl+F). Penanda bagian sudah seragam (`============ NAMA ============`; sub-bagian `---- nama ----`).
     - **Kode mati: tidak ada.** Dari 191 fungsi di `JavaScript.html`, tidak satu pun hanya muncul di baris definisinya (dicek terhadap seluruh file frontend); tidak ada nama fungsi ganda dan tidak ada pemanggilan dinamis (`window[...]`, `eval`) yang bisa menyembunyikan pemakaian.
     - **`GAS_WEBAPP_URL_`** (adapter Netlify) sudah diisi URL deployment terbaru. Mengganti file ini di repo tidak lagi mengosongkan URL. Bila memakai deployment lain, ubah di baris `var GAS_WEBAPP_URL_ = ...` (blok skrip pertama).
     - Peta `GAS_PARAM_MAP_` frontend diverifikasi ulang = `getRouteParamMap()` backend (65 route, urutan parameter sama).
   - **Perbaikan pasca-Fase 7: `submitAbsenOnline` kini mengecek bulan periode** (lihat temuan no. 1). Tanpa periode yang sama dengan bulan berjalan, Realisasi/JamScan/JamKerja tidak disentuh. Perilaku identik bila periode sesuai (800 skenario acak).
   - **Perbaikan berkas `AbsensiOnline.html`**: teks README Backup Reliever sempat tertempel di tengah blok `<script>` (± baris 640-1291) sehingga skrip error sintaks dan halaman tidak berfungsi. Bagian yang hilang (pengambilan GPS + pengiriman absen, serta fungsi `pilihMode`) dipulihkan dari tambalan Reliever; alur karyawan tetap dan reliever diuji di simulasi browser (19 pemeriksaan).
   - **Perbaikan pasca-Fase 7: `getJadwalHariIni` kini mengecek bulan periode.** Bila bulan/tahun periode aktif (A2 sheet Jadwal) ≠ hari ini, halaman Absen Online menampilkan "Periode belum diganti (NAMA PERIODE)", bukan shift pada tanggal yang sama di bulan lalu (dulu mis. 1 Oktober dengan periode masih September menampilkan jadwal 1 September). Bila nama periode tak terbaca bulan/tahunnya, pengecekan dilewati (perilaku lama). Selain itu perilaku identik (3.000 kombinasi acak).

00000. **(Fase 6) Temuan dan perubahan perilaku:**
   - **Laporan email = bulan lalu, terkirim tanggal 5 jam 07:00** (permintaan Anda; ubah `MONTHLY_REPORT_DAY_`/`MONTHLY_REPORT_HOUR_` lalu jalankan ulang `setupMonthlyReportTrigger()`). Sumber data dipilih berlapis: (1) arsip periode bulan lalu bila ada, dengan PDF Dashboard/Rekap diambil dari sheet arsip; (2) periode aktif bila periode aktif itu bulan lalu (belum ditutup); (3) selain itu arsip terbaru (atau periode aktif bila belum ada arsip) dengan **peringatan merah di badan email**. Bila arsip bulan lalu tidak bisa dibaca, otomatis jatuh ke periode aktif dengan peringatan. Analisis tren dibandingkan dengan arsip tepat sebelum periode yang dilaporkan (bukan dengan dirinya sendiri).
   - **Analisis tren di email bulanan tidak sebanding.** Rata-rata periode berjalan dihitung sebagai rata-rata persen per karyawan, tetapi pembandingnya (periode arsip terakhir) dihitung total hadir / total jadwal lalu **dibulatkan ke persen bulat**. Di uji: 0,3200 (lama) vs 0,3296 (baru) untuk data yang sama, selisih ±1 poin, lebih besar daripada ambang "STABIL" (0,5 poin). Kini keduanya memakai cara hitung yang sama (juga sama dengan grafik tren `Rekap_Bulanan`). Bila arsip tidak punya baris valid, hasilnya "belum ada data pembanding" (dulu dianggap 0%).
   - **Export Excel/PDF generik gagal diam-diam**: bila Google mengembalikan error, versi lama tetap mengirim file rusak (halaman error) ke pengguna sebagai "Excel". Kini error jelas ("Gagal membuat laporan Excel (kode N)"). Spreadsheet sementara juga dibuang walau export gagal (dulu tertinggal di Drive).
   - Input laporan generik divalidasi: kolom kosong/bukan daftar memberi pesan jelas, baris kosong/undefined dianggap tanpa baris.
   - `escapeHtmlServer_` dipindah ke `07_TextUtils.gs` karena juga dipakai `19_PengajuanIzin.gs`.
   - Pembuatan salinan sheet sementara (template jadwal, PDF per department) kini satu helper `withTempSheetCopy_` di `06_FileUtils.gs`; filter baris per department memakai `keepOnlyDepartmentRows_` yang sama.

0000. **(Fase 5) Temuan dan perubahan perilaku:**
   - **`setKaryawanResign` hanya kerangka**: tidak menulis apa pun ke sheet, tetapi mengembalikan `success:true` dan mencatat "Tandai Resign" di Audit Log (jejak palsu). Tidak ada tombol frontend yang memakainya. Kini menolak dengan pesan jelas dan tidak mencatat apa-apa. **Perlu keputusan rancangan fitur** (lihat pertanyaan di akhir).
   - **Status Assignment kosong** tampil "Belum Terjadwal" di Daftar Karyawan, tetapi assign menolaknya dengan alasan kosong ("sudah "). Kini diperlakukan sama seperti tampilan (bisa di-assign).
   - **Assign nama yang sudah ada di Jadwal/Realisasi** dulu menyisipkan baris **duplikat**; kini dilewati dengan alasan "sudah ada di sheet Jadwal/Realisasi" (mencegah baris Jadwal dan Realisasi tidak sejajar).
   - **Audit Log**: Assign Karyawan ke Jadwal dan Hapus Karyawan dari Jadwal kini tercatat (sebelumnya tidak).
   - **`setEmployeeNik`** (tombol "Isi NIK") sekarang mengembalikan state terbaru tanpa cache 8 detik, supaya NIK langsung terlihat.
   - **Performa**: `getOrCreateNikSheet_` dipanggil di setiap pembacaan data karyawan (termasuk tiap `getAllData`) dan dulu selalu menulis format teks ke seluruh kolom No HP (±1000 sel). Kini hanya menulis bila header berubah atau baris paling bawah belum berformat teks.
   - Kode mati dihapus: `ensureDataKaryawanColumns_` (skema 6 kolom usang) dan `getUnassignedEmployees_` (tak dipanggil siapa pun).

000. **(Fase 4) Perubahan perilaku yang disengaja:**
   - Upload Jadwal, Upload Realisasi, dan Upload Finger sebelumnya **tidak tercatat di Audit Log**. Upload Jadwal/Realisasi kini mencatat jumlah karyawan cocok, nama tidak ditemukan, dan kode tidak valid. (Upload Finger belum dicatat; menunggu keputusan Anda.)
   - Sheet kosong (tanpa karyawan) memberi pesan jelas ("Sheet X belum berisi karyawan"), bukan error teknis "number of rows must be at least 1".
   - Hak akses upload dicek **sebelum** daftar kode shift dibaca, jadi akun non-Super Admin ditolak tanpa menyentuh sheet.
   - Sinkron JamKerja/JamScan tidak lagi gagal bila Realisasi belum punya karyawan.

00. **(Fase 3) `backfillRekapBulananDariArsip_` salah urutan argumen** — memanggil snapshot dengan (sheet Rekap arsip, sheet Jadwal arsip) padahal yang diminta (Jadwal, Realisasi), sehingga menulis baris bulan berisi nol semua. Terbukti di uji: lama menghasilkan `[2026,10,0,0,0,0,0,0,0,0]`, baru `[2026,10,36,19,...]` (sama dengan snapshot normal). Fungsi ini dibiarkan di `05_LegacyCode.gs` (akan pindah ke `99_Debug.gs` di Fase 7), tetapi argumennya sudah dibetulkan.
01. **(Fase 3)** `getMonthlyTrendYear` mengabaikan baris dengan nilai bulan tidak valid (sebelumnya menulis properti aneh di array). Duplikasi kode pembacaan `ArsipIndex` (5 tempat) jadi satu helper `findArchiveRow_`.

0. **(Fase 2) `dayIndex` berupa string** — `saveRealisasi`/`saveJadwal` lama menghitung kolom `4 + '5'` = `'45'` (kolom 45, bukan 9) dan menulis log "Hari ke-51". Kini dipaksa jadi angka. `saveRealisasiBatch` juga menolak input non-daftar, dan tetap all-or-nothing (satu sel salah = tidak ada yang ditulis).

1. **`setEmployeeNik` tidak ada** — tombol "Isi NIK" di frontend memanggilnya, tetapi fungsinya tidak pernah didefinisikan. Sekarang pembungkus `setEmployeeMasterData` (mengembalikan state penuh, sesuai handler frontend).
2. **Route `getPengajuanPendingCount` tidak ada** di `doPost`, padahal frontend memanggilnya (badge pengajuan). Ditambahkan.
3. **Peta `GAS_PARAM_MAP_` frontend tidak lengkap** — `deleteUnassignedEmployee` dan `addShiftJamKerja` dipanggil frontend tetapi tidak terdaftar (error di mode Netlify). Peta kini dibuat dari `ROUTES` (65 entri, sinkron 100%).
4. **`16_RepairAlignment.gs` rusak** (komentar & fungsi menempel dalam satu baris) dan menduplikasi `checkRowAlignment_`. Menjadi `16_Rekap.gs`.
5. **`24_BackupReliever.gs`** mendefinisikan dua fungsi test dua kali. Salinan kedua dihapus.
6. **Notifikasi lonceng salah baca kolom** — mengasumsikan `[Waktu,User,Aksi,Modul,Detail]` padahal AuditLog `[Timestamp,Username,Nama,Role,Aksi,Modul,Target,Detail]`, dan memakai kolom `TanggalPengajuan` yang tidak ada (sebenarnya `DiajukanPada`). Kini membaca lewat nama header. Fungsi mati `getNotifications_` (memanggil `validateToken_` yang tidak ada) dihapus.
7. **Rahasia di kode**: `PASSWORD_SALT` dan `REPORT_EMAIL` kini dibaca dari Script Properties (`LEGACY_PASSWORD_SALT`, `REPORT_EMAIL`) dengan nilai lama sebagai cadangan. Fungsi `resetPasswordSaya()` (password tertulis di kode) dan fungsi dedupe sekali-jalan dihapus.

## Cara memasang

**Update dari satelit batch 1 ke Input Realisasi & Jadwal Batch (backend DAN frontend):**
1. Backend: ganti isi `01_Router.gs` dan `11_InputData.gs`. File backend lain tidak berubah.
2. Frontend: ganti isi `index` (index.html), `JavaScript` (JavaScript.html), `Stylesheet` (Stylesheet.html). `AbsensiOnline` dan `SettingJadwalKerja` tidak berubah. Pastikan `GAS_WEBAPP_URL_` di `JavaScript.html` masih berisi URL deployment Anda (paket ini mempertahankannya).
3. Jalankan `auditRoutes()`: harus melaporkan **66 route** tanpa masalah. (Bila Anda sebelumnya sudah memasang paket "jadwal batch" tanpa perbaikan tata letak, cukup ganti `Stylesheet` saja.) Terapkan ulang Web App (Deploy > Manage deployments > Edit > New version) dan deploy ulang Netlify bila dipakai.
4. Uji **di salinan**: (a) tab Input Realisasi: pilih beberapa sel, beri kode, `Simpan Batch`, cek di sheet Realisasi; (b) tab Jadwal Kerja sebagai Super Admin: pilih rentang sel, beri shift, `Simpan Batch`, cek di sheet Jadwal; (c) buka **Setting > Audit Log**: satu batch harus tampil sebagai SATU baris `Edit Jadwal (Batch)` / `Edit Realisasi (Batch)`; (d) login sebagai role non-admin: tab Jadwal tidak boleh bisa diubah; (e) logout dengan perubahan belum disimpan: perubahan dibuang.

**Update dari Fase 8 ke modul satelit batch 1 (backend saja, frontend tidak berubah):**
1. Ganti isi `07_TextUtils.gs`, `15_AIchat.gs`, `21_AuditLog.gs`, `23_SettingJamKerja.gs` (tidak ada file yang ditambah atau dihapus).
2. Jalankan `auditRoutes()` (harus "Semua 65 route OK"), lalu uji **di salinan**: Setting > Jam Kerja (ubah jam, tambah shift, coba jam `25:00`), Setting > Audit Log (filter tanggal hari ini dan Modul), dan chatbot (tanya satu pertanyaan; minta kirim email ke diri sendiri lalu cek Audit Log).
3. Opsional: isi Script Property `AI_MODEL` bila ingin model lain.

**Update dari Fase 7 ke Fase 8 (frontend saja, backend tidak berubah):**
1. Ganti `index.html` dan `JavaScript.html` (di Apps Script dan/atau repo Netlify). `Stylesheet.html`, `SettingJadwalKerja.html` tidak berubah.
2. Terapkan ulang Web App (Apps Script) / deploy ulang (Netlify).
3. Uji: login gate tampil dengan logo, login, Dashboard, buka satu tab tiap menu. Logo pada layar sangat besar boleh dibandingkan dengan sebelumnya; bila terlihat kurang tajam, tinggi 136 px bisa dinaikkan (perlu mengganti gambar tertanam).


**Update dari Fase 6 ke Fase 7 (di project Apps Script uji):**
1. **Hapus `05_LegacyCode.gs`** (isinya sudah habis dipindah; kalau dibiarkan, fungsi di dalamnya bentrok dengan file baru).
2. Tambahkan `32_JadwalHariIni.gs` dan `99_Maintenance.gs`.
3. Ganti isi `02_SheetUtils.gs` dan `16_Rekap.gs` dengan versi baru.
4. Jalankan `auditRoutes()` (harus "Semua 65 route OK"), lalu uji **di salinan**: buka halaman Absen Online (jadwal shift hari ini muncul), Dashboard dan Rekap Karyawan, serta pilih satu periode arsip (kolom Rank dan angka sama seperti sebelumnya).

**Update dari Fase 5 ke Fase 6 (di project Apps Script uji):**
1. Tambahkan `07_TextUtils.gs`, `29_Export.gs`, `30_AttendanceLog.gs`, `31_MonthlyReport.gs`.
2. Ganti isi `05_LegacyCode.gs`, `06_FileUtils.gs`, dan `17_UploadGrid.gs` dengan versi baru (`29_Export.gs` dan `31_MonthlyReport.gs` juga ikut yang terbaru).
3. `appsscript.json` tidak berubah. **Laporan bulanan sekarang untuk BULAN LALU dan terkirim tanggal 5 jam 07:00** (konstanta `MONTHLY_REPORT_DAY_` di `31_MonthlyReport.gs`). Trigger lama (tanggal 1) tetap aktif sampai Anda menjalankan `setupMonthlyReportTrigger()` sekali dari editor; fungsi itu menghapus trigger lama lalu memasang yang baru.
4. Jalankan `auditRoutes()`, lalu uji **di salinan**: Export PDF (Rekap/Dashboard/Jadwal/Realisasi, dengan dan tanpa department), Report > Export Excel/PDF, Log Absensi (periode aktif dan arsip). Jalankan `sendMonthlyReportEmail` dari editor hanya bila alamat `REPORT_EMAIL` di salinan sudah Anda atur sendiri (email sungguhan terkirim).

**Update dari Fase 4 ke Fase 5 (di project Apps Script uji):**
1. **Hapus dulu `21_Resign.gs` dan `22_AssignJadwal.gs`** (digantikan `28_EmployeeLifecycle.gs`; kalau tidak dihapus akan error deklarasi ganda).
2. Tambahkan `27_EmployeeData.gs` dan `28_EmployeeLifecycle.gs`.
3. Ganti isi `01_Router.gs` dan `05_LegacyCode.gs` dengan versi baru.
4. Jalankan `auditRoutes()`, lalu uji **di salinan**: Daftar Karyawan (tambah, edit NIK/department), Assign Karyawan ke Jadwal, Hapus Karyawan (tanpa dan dengan "Hapus Paksa"), dan hapus karyawan yang belum terjadwal.

**Update dari Fase 3 ke Fase 4 (di project Apps Script uji):**
1. Tambahkan `06_FileUtils.gs`, `17_UploadGrid.gs`, `25_FingerAbsen.gs`, `26_JamKerja.gs`.
2. Ganti isi `05_LegacyCode.gs` dengan versi baru (kode upload/finger/template/jam kerja lama sudah dihapus dari situ; kalau tidak diganti akan error deklarasi ganda).
3. `appsscript.json` tidak berubah. Scope Drive/Spreadsheet/UrlFetch yang sudah Anda setujui sebelumnya cukup.
4. Jalankan `auditRoutes()`, lalu uji **di salinan**: Download Template Excel (Jadwal Kerja), Upload Jadwal, Upload Realisasi, Upload Finger Absen dengan file contoh kecil.

**Update dari Fase 1-2 ke Fase 3 (di project Apps Script uji):**
1. **Hapus file `20_TrendComparison.gs`** (digantikan `14_MonthlyTrend.gs`; kalau tidak dihapus akan error deklarasi ganda).
2. Tambah `12_Period.gs`, `13_Archive.gs`, `14_MonthlyTrend.gs`.
3. Ganti isi `01_Router.gs` dan `05_LegacyCode.gs` dengan versi baru.
4. Jalankan `auditRoutes()` (harus "Semua 65 route OK"), lalu uji di salinan: Periode Aktif ▾ pilih arsip, **Arsipkan Periode Ini**, **Mulai Periode Baru**.

**Pemasangan awal (dari nol):**

1. Di Apps Script: **hapus `Code.gs`** dan semua file `.gs` lama, lalu tambahkan semua file di `Backend/` (jangan biarkan dua versi hidup bersamaan; deklarasi ganda akan error).
2. Ganti isi `appsscript.json` dengan versi di folder ini (sama dengan sebelumnya).
3. Frontend: salin file di `Frontend/`. Hanya `JavaScript.html` yang berubah (blok `GAS_PARAM_MAP_`). Isi ulang `GAS_WEBAPP_URL_` di blok adapter jika memakai Netlify.
4. Jalankan `auditRoutes()` → harus log "Semua 65 route OK".
5. Uji: login, buka Dashboard, input realisasi, Setting > Account, Setting > Hak Akses, lonceng notifikasi.
6. (Disarankan) Isi Script Properties: `LEGACY_PASSWORD_SALT`, `REPORT_EMAIL`, `GEMINI_API_KEY`.

> Catatan: nama file template di `HTML_ENTRY_FILE` adalah `Index` (huruf besar). Di ZIP berkas bernama `index.html`; pastikan nama di editor Apps Script sama persis dengan konstanta tersebut.

## Yang sudah diuji

- **Input Realisasi & Jadwal Batch, backend (Apps Script tiruan di Node, 1.025 pemeriksaan)**: `validateCellTarget_` lama = baru pada 324 kombinasi masukan (termasuk string, desimal, NaN, null); `saveRealisasiBatch` lama = baru pada 60 batch acak (isi sheet, nilai angka 1/2, kembalian, Audit Log identik; panggilan sheet lebih sedikit) dan 10 kasus error (pesan sama, tidak ada yang ditulis atau dicatat, lock dilepas); `saveJadwalBatch` = `saveJadwal` dijalankan berurutan pada 60 batch acak (isi sheet identik, tepat SATU catatan Audit Log berformat benar, lock dipakai sekali dan dilepas); 9 kasus penolakan (kode salah di tengah batch, baris/hari di luar data, bukan bilangan bulat, bukan array, elemen null, bukan Super Admin, lebih dari 6.000 sel); sel ganda, kode kosong, dan daftar kosong; 930 sel berurutan = 30 panggilan `setValues` dan hanya 2 pembacaan sheet (lama: ±3 panggilan per sel); route terdaftar (66).
- **Input Realisasi & Jadwal Batch, tampilan (Chrome headless, server Apps Script tiruan, 57 pemeriksaan)**: pilih sel, Shift+klik rentang, terapkan kode, kode sama tidak dihitung, polling tidak menghapus perubahan, Simpan Batch mengirim satu panggilan dengan sel dan token benar (Realisasi dan Jadwal), gagal simpan mempertahankan perubahan, baris bergeser membuang perubahan, Batalkan, Esc, tampilan arsip, role tanpa izin (termasuk non Super Admin untuk Jadwal), perubahan Realisasi dan Jadwal terpisah, tidak ada lagi pemanggilan `saveJadwal`/`saveRealisasi`, popup lama hilang, Unduh Template, logout; tanpa error JavaScript. **Belum diuji pada Apps Script/spreadsheet sungguhan.**
- **Tata letak tabel (Chrome headless, data 31 hari x 60 karyawan, tiga lebar layar)**: sebelum perbaikan halaman melebar ke samping 274 px (laptop 1366) dan 1.010 px (HP 390); sesudahnya 0 px di semua ukuran. Pada layar 1920 seluruh 31 hari muat tanpa digeser; pada 1366 dan HP tabel digeser di dalam kotaknya dengan kolom nama tetap terlihat. Sapuan 19 tab pada 3 lebar layar (57 kombinasi) dibandingkan dengan versi sebelum perbaikan: 8 membaik (Jadwal, Input, Rekap, Daftar Karyawan, Setting Jam Kerja), 49 sama, 0 memburuk. Tabel + bar Simpan Batch tidak saling menutupi (ujung bawah tabel dan scrollbar horizontalnya tetap terlihat di atas bar).

- **Modul satelit batch 1, uji banding lama vs baru (mock Apps Script di Node, 26 pemeriksaan)**: Audit Log (pencatatan 100 baris dan header salah identik, 60 kombinasi filter modul/limit identik, penolakan non-Super Admin identik, filter tanggal dicocokkan dengan perhitungan waktu lokal independen), Setting Jam Kerja (pembuatan sheet dari nol dan migrasi sheet lama termasuk nilai Date identik, 60 urutan acak ubah jam identik pada isi, format sel, kembalian, dan Audit Log; 10 kasus tambah shift dan semua pesan galat validasi lama identik), Chatbot (isi permintaan ke Gemini, balasan jawab/kirim_email, galat tanpa key/HTTP 500/respons kosong identik; email terkirim identik). Selisih disengaja diuji terpisah: filter tanggal, rumus di Audit Log dan label shift, jam/target tak valid, JSON AI rusak, tanggal di luar rentang, alamat email tak valid, audit pengiriman email.
  **Belum diuji di Google Sheets sungguhan**: perilaku apostrof di depan teks (`'=...`) lewat `appendRow`/`setValues` mengikuti perilaku Sheets yang diketahui, tetapi mock tidak bisa membuktikannya; cek sekali dengan menambah shift berlabel `=1+1` di salinan (label harus tampil `=1+1`, bukan `2`).
- **Fase 8, uji asap `index.html` rakitan (jsdom)**: halaman (index + Stylesheet + SettingJadwalKerja + JavaScript) termuat tanpa galat skrip; logo login & sidebar berupa data PNG < 30 KB; adapter Netlify aktif dan menerjemahkan `login(username, password)` ke permintaan aksi `login` ke URL deployment; elemen yang dipasangi listener tersedia; seluruh blok `<script>` lolos cek sintaks; `GAS_PARAM_MAP_` = 65 route backend.
- **Perbaikan `submitAbsenOnline` (pasca-Fase 7)**: 800 skenario acak (periode sesuai; nama dengan spasi/huruf besar-kecil, NIK salah, jenis tidak valid, kode Realisasi/JamScan awal acak) identik dengan versi sebelumnya pada isi seluruh sheet, kembalian, dan galat; 4 skenario periode tidak sesuai (bulan lalu, bulan depan dari periode, pergantian tahun, tahun berbeda) × Masuk/Pulang: Realisasi/JamScan/JamKerja tidak berubah, log tetap tercatat identik, kembalian membawa penanda.
- **`AbsensiOnline.html` (simulasi browser jsdom)**: login karyawan, tab Reliever (pencarian dari DataReliever, pilih nama, dropdown "menggantikan siapa"), penolakan tanpa pilihan, isi kiriman `submitAbsenOnlineReliever`, logout, dan pesan periode belum diganti.
- **Perbaikan `getJadwalHariIni` (pasca-Fase 7)**: 3.000 kombinasi acak periode = bulan berjalan identik dengan versi sebelumnya; 5 skenario periode tidak sesuai (bulan lalu, 3 November, pergantian tahun, tahun berbeda, periode di depan) menghasilkan "Periode belum diganti"; 4 bentuk nama periode tak terbaca, sheet hilang, dan header tak dikenali tetap sama seperti sebelumnya; `auditRoutes()` 65 route OK.
- **Fase 7, uji banding lama vs baru**: 298 pemeriksaan (rekap aktif 40 kasus dan rekap arsip 150 kasus acak identik, `getJadwalHariIni` 48 kombinasi tanggal/nama/layout identik, tiga alat perbaikan dan `debugNameMismatch` identik pada 5 variasi sheet, penghitungan `findBlankRowGaps` terhadap perhitungan independen, backfill sama dengan snapshot normal, dan semua alat bernama tanpa garis bawah).\n- **Fase 6, uji banding lama vs baru**: 165 pemeriksaan (termasuk 8 skenario baru "laporan bulan lalu" dengan tanggal terkendali: arsip ada, belum diarsipkan tapi aktif, tidak ada keduanya, arsip rusak, pergantian tahun, lampiran gagal) dan perbandingan identik (export PDF 14 kombinasi sheet/department termasuk filter dan error, laporan Excel/PDF generik, log absensi 11 kombinasi argumen plus arsip, parser kode/scan, dan email bulanan 14 skenario berikut HTML-nya, pemasangan trigger). Selisih disengaja diperiksa terpisah (rata-rata periode sebelumnya, error export).
- **Fase 5, uji banding lama vs baru**: 198 perbandingan identik (lapisan DataKaryawan dari nol/header salah, edit master data, tambah/assign/hapus karyawan dengan berbagai kombinasi, hak akses, dan hasil di semua sheet termasuk format). Selisih yang disengaja diperiksa terpisah dan terbukti: status assignment kosong, assign duplikat, dan resign.
- **Fase 4, uji banding lama vs baru**: 503 perbandingan identik (impor grid acak ×120, terapkan finger acak ×160, parsing file finger, sinkron JamKerja/JamScan saat urutan karyawan berubah, pengayaan rekap, template per department, dan upload end-to-end termasuk isi request ke Drive, pembuangan file sementara, serta penolakan non-Super Admin). Satu-satunya selisih = catatan Audit Log baru.
- **Fase 2, uji banding lama vs baru**: 603 perbandingan acak (tren harian dengan/tanpa filter, ringkasan dashboard, peringatan alignment) hasilnya identik; `syncOffFromJadwal_`, `saveRealisasi`, `saveRealisasiBatch` pada sheet tiruan menghasilkan isi sel identik (selisih tunggal = perbaikan bug no. 0).

- Semua 18 file `.gs` lolos cek sintaks dan dimuat bersama tanpa deklarasi ganda.
- Uji logika dengan mock: normalisasi role (termasuk "admin" lama vs "Admin" baru), hash password skema baru & lama, dispatcher route.
- `auditRoutes()`: 65 route, 0 masalah. Peta frontend = tabel route (nama dan urutan parameter).

**Belum diuji**: berjalan di spreadsheet asli (tidak ada akses ke Apps Script dari sini). Lakukan uji langkah 5 di salinan spreadsheet dulu.

## Belum ditulis ulang

Empat modul satelit (±1.270 baris, 65 fungsi) masih berisi kode asli, hanya dirapikan seperlunya:
`18_PolaJadwalGenerator.gs` (11), `19_PengajuanIzin.gs` (18), `20_AbsenOnline.gs` (21), `24_BackupReliever.gs` (15). Sudah ditulis ulang: `15_AIchat.gs`, `21_AuditLog.gs`, `23_SettingJamKerja.gs` (batch 1). `23_Notifications.gs` sudah diperbaiki di Fase 1.

## Temuan yang menunggu keputusan

1. **Tanggal tanpa cek bulan (data bisa salah bulan).** ~~`getJadwalHariIni` (32)~~ **sudah diperbaiki** (lihat catatan Fase 7). ~~`submitAbsenOnline` (20)~~ **juga sudah diperbaiki**: bila periode aktif bukan bulan berjalan, absen tetap tercatat di `LogAbsenOnline` tetapi TIDAK ditulis ke Realisasi/JamScan/JamKerja (dulu menimpa kolom tanggal yang sama di bulan lalu); kembalian membawa `periodeBelumDiganti` dan `periode`, dan halaman Absen Online menampilkan peringatan. Absen yang tercatat di log selama periode belum diganti perlu diproses admin (belum ada alat otomatis untuk memasukkannya ke Realisasi setelah periode diganti). Bila periode belum diganti (mis. hari ini 3 November, periode aktif masih Oktober), jadwal yang tampil dan sel Realisasi/JamScan/JamKerja yang terisi adalah milik **3 Oktober**. Modul Backup Reliever (24) sudah punya pengecekan semacam ini (`getKaryawanAbsenHariIni` mengembalikan daftar kosong bila bulan periode ≠ bulan sekarang), jadi pengamannya tinggal diseragamkan. Usulan: bila bulan periode aktif ≠ bulan hari ini, jadwal hari ini menampilkan "Periode belum diganti" dan absen tetap tercatat di log tetapi tidak menulis ke Realisasi.
2. **Arti kolom Rank berbeda** antara periode aktif (`getRekap_`: Rank 1 = % **terendah**, sesuai label "Rank (Terendah)") dan periode arsip (`computeRekapDariGrid_`: Rank 1 = % **tertinggi**). Frontend menampilkan angkanya apa adanya. Perilaku warisan dipertahankan; perlu diputuskan satu aturan.
3. **Tandai Resign** (lihat catatan Fase 5): perlu rancangan fitur.

## Fase berikutnya

1. ~~Data API~~ selesai (Fase 2)
2. ~~Input~~ selesai (Fase 2)
3. ~~Periode & arsip~~ selesai (Fase 3)
4. ~~Upload & finger~~ selesai (Fase 4)
5. ~~Kepegawaian~~ selesai (Fase 5)
6. ~~Laporan~~ selesai (Fase 6)
7. ~~Sisa Code.gs + alat debug~~ selesai (Fase 7)
8. ~~Frontend~~ selesai (Fase 8 versi ringan: logo diperkecil, daftar isi, tanpa memecah file)
9. **Modul satelit** — batch 1 selesai (Audit Log, Chatbot AI, Setting Jam Kerja); sisa: PolaJadwalGenerator, PengajuanIzin, AbsenOnline, BackupReliever
