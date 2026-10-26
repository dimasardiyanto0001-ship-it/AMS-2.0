/*************************************************************
 * 20_AbsenOnline.gs
 * Backend halaman publik "Absen Online": absen GPS + selfie
 * (submitAbsenOnline), riwayat (getRiwayatAbsen), pencarian nama,
 * foto profil, dan penghapusan foto lama (retensi).
 * Tanpa token: identitas = Nama + NIK yang cocok dengan DataKaryawan.
 *
 * Sheet: LogAbsenOnline (log mentah, tersembunyi),
 *        SettingLokasiAbsen (lat, lng, radius; tersembunyi).
 * Dipakai juga oleh 24_BackupReliever.gs: getSettingLokasiAbsen_,
 * haversineDistanceMeter_, getOrCreateFotoAbsenBulanFolder_, dan
 * oleh 32_JadwalHariIni.gs: formatJam_.
 *
 * Perubahan terbaru:
 *  - Absen di luar radius DITOLAK (tidak dicatat, foto tidak disimpan).
 *  - Lokasi kantor yang belum diatur (0,0) ditolak dengan pesan jelas.
 *  - Masuk ganda: jam Masuk PALING AWAL yang dipertahankan di JamScan.
 *  - Pulang lintas tengah malam (mis. Masuk 22:00 -> Pulang 06:00):
 *    JamScan & JamKerja ditulis ke kolom hari Masuk.
 *************************************************************/

const SHEET_LOG_ABSEN_ONLINE = 'LogAbsenOnline';
const SHEET_SETTING_LOKASI_ABSEN = 'SettingLokasiAbsen';
const LOG_ABSEN_ONLINE_HEADERS_ = ['Timestamp', 'Tanggal', 'Nama', 'NIK', 'Jenis', 'Jam', 'Latitude', 'Longitude', 'JarakMeter', 'Status', 'FotoUrl'];
const LOG_ABSEN_FOTO_COL_ = 11;
const FOTO_ABSEN_ROOT_FOLDER_NAME_ = 'Foto Absen Online - AMS Housekeeping';
const FOTO_PROFIL_ROOT_FOLDER_NAME_ = 'Foto Profil Karyawan - AMS Housekeeping';
const RETENSI_FOTO_ABSEN_BULAN_ = 3;
const FOTO_MAX_BYTES_ = 6 * 1024 * 1024;   // foto selfie/profil, setelah didekode
const JAM_KERJA_MAKS_ = 16;                // durasi absen Masuk-Pulang di atas ini dianggap salah ketik, tidak ditulis ke JamKerja

/* ============================ HELPER SPREADSHEET ============================ */

/*
 * CATATAN: getSpreadsheet_() SENGAJA TIDAK didefinisikan lagi di file ini. Fungsi itu sudah ada di
 * 02_SheetUtils.gs. Versi lama di sini (getSheet_(...).getParent()) bentrok dengan versi 02_SheetUtils:
 * file ini dimuat belakangan sehingga menimpanya, lalu getSheet_ <-> getSpreadsheet_ saling memanggil
 * tanpa henti ("RangeError: Maximum call stack size exceeded") dan login/semua fitur gagal.
 */

/**
 * Jalankan SEKALI dari editor (pilih fungsi ini lalu Run) untuk melihat helper/konstanta
 * yang dibutuhkan file ini tetapi belum ada di project. Hasil ada di Execution log.
 */
function cekDependensiAbsenOnline() {
  requireScriptOwner_();
  const fungsi = ['getSheet_', 'getKaryawanMasterMap_', 'getOrCreateNikSheet_', 'withScriptLock_',
    'parsePeriodMonthYear_', 'getLastEmployeeRow_', 'getDayColCount_', 'getOrSyncJamScanSheet_',
    'getOrSyncJamKerjaSheet_', 'splitScanPair_', 'formatScanPair_'];
  const objek = ['CONFIG', 'EMP_COL'];
  const hilang = [];
  fungsi.forEach(n => { try { if (typeof eval(n) !== 'function') hilang.push(n); } catch (e) { hilang.push(n); } });
  objek.forEach(n => { try { if (typeof eval(n) !== 'object') hilang.push(n); } catch (e) { hilang.push(n); } });
  Logger.log(hilang.length ? 'BELUM ADA di project: ' + hilang.join(', ') : 'Semua dependensi ada.');
}

/* ============================ LOKASI ============================ */

function getOrCreateSettingLokasiAbsenSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(SHEET_SETTING_LOKASI_ABSEN);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SETTING_LOKASI_ABSEN);
    sheet.getRange(1, 1, 1, 3).setValues([['Latitude', 'Longitude', 'RadiusMeter']]).setFontWeight('bold');
    // NILAI PLACEHOLDER -- WAJIB diganti manual di baris 2 ke koordinat kantor sebenarnya.
    sheet.getRange(2, 1, 1, 3).setValues([[0, 0, 100]]);
    sheet.getRange('E1').setValue('Isi Latitude/Longitude kantor & radius toleransi (meter) di baris 2.');
    sheet.hideSheet();
  }
  return sheet;
}

function getSettingLokasiAbsen_() {
  const sheet = getOrCreateSettingLokasiAbsenSheet_();
  const row = sheet.getRange(2, 1, 1, 3).getValues()[0];
  return { lat: Number(row[0]) || 0, lng: Number(row[1]) || 0, radius: Number(row[2]) || 100 };
}

function haversineDistanceMeter_(lat1, lng1, lat2, lng2) {
  const R = 6371000;
  const toRad = d => d * Math.PI / 180;
  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLng / 2) ** 2;
  return Math.round(R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a)));
}

/* ============================ LOG ============================ */

function getOrCreateLogAbsenOnlineSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(SHEET_LOG_ABSEN_ONLINE);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_LOG_ABSEN_ONLINE);
    sheet.getRange(1, 1, 1, LOG_ABSEN_ONLINE_HEADERS_.length).setValues([LOG_ABSEN_ONLINE_HEADERS_]).setFontWeight('bold');
    sheet.hideSheet();
  } else {
    ensureLogAbsenOnlineHasFotoColumn_(sheet);
  }
  return sheet;
}

function ensureLogAbsenOnlineHasFotoColumn_(sheet) {
  const header = sheet.getRange(1, LOG_ABSEN_FOTO_COL_).getValue();
  if (header !== 'FotoUrl') {
    sheet.getRange(1, LOG_ABSEN_FOTO_COL_).setValue('FotoUrl').setFontWeight('bold');
  }
}

/* ============================ VERIFIKASI ============================ */

const PIN_MAX_GAGAL_ = 5;        // salah PIN/NIK berturut-turut per nama
const PIN_KUNCI_DETIK_ = 900;    // 15 menit (CacheService maks 6 jam)
const PIN_BELUM_DIATUR_ = 'PIN_BELUM_DIATUR';

function pinFailKey_(namaKey) {
  return 'pinfail_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, namaKey)).substring(0, 40);
}

/** Lempar error bila nama ini sudah terlalu sering salah; kembalikan hitungan saat ini. */
function cekKunciPin_(namaKey) {
  const gagal = Number(CacheService.getScriptCache().get(pinFailKey_(namaKey)) || 0);
  if (gagal >= PIN_MAX_GAGAL_) {
    throw new Error('Terlalu banyak percobaan salah. Coba lagi dalam ' + Math.round(PIN_KUNCI_DETIK_ / 60) + ' menit atau hubungi admin.');
  }
  return gagal;
}

function catatGagalPin_(namaKey, gagal) {
  CacheService.getScriptCache().put(pinFailKey_(namaKey), String(gagal + 1), PIN_KUNCI_DETIK_);
}

function buatPinHash_(pin) {
  const salt = generateSalt_();
  return salt + '$' + hashPasswordWithSalt_(String(pin), salt);
}

function pinCocok_(pin, stored) {
  const parts = String(stored || '').split('$');
  return parts.length === 2 && parts[0] !== '' && hashPasswordWithSalt_(String(pin), parts[0]) === parts[1];
}

function validasiFormatPin_(pin) {
  pin = String(pin || '').trim();
  if (!/^\d{6}$/.test(pin)) throw new Error('PIN harus 6 digit angka.');
  if (/^(\d)\1{5}$/.test(pin) || '0123456789012345'.indexOf(pin) !== -1 || '9876543210987654'.indexOf(pin) !== -1) {
    throw new Error('PIN terlalu mudah ditebak (mis. 111111 atau 123456). Pilih PIN lain.');
  }
  return pin;
}

/**
 * Verifikasi nama + PIN untuk endpoint publik. 5 kali salah untuk satu nama mengunci
 * nama itu 15 menit (termasuk untuk PIN yang benar). Karyawan yang PIN-nya belum
 * diaktifkan ditolak dengan pesan 'PIN_BELUM_DIATUR' (frontend lalu menawarkan aktivasi).
 */
function verifyKaryawanByNamaPin_(nama, pin) {
  const namaKey = String(nama || '').trim().toLowerCase();
  const gagal = cekKunciPin_(namaKey);
  const master = getKaryawanMasterMap_()[namaKey];
  if (!master) throw new Error('Nama karyawan tidak ditemukan.');
  if (!master.pinHash) throw new Error(PIN_BELUM_DIATUR_);
  if (!pinCocok_(pin, master.pinHash)) {
    catatGagalPin_(namaKey, gagal);
    throw new Error('PIN salah.');
  }
  if (gagal) CacheService.getScriptCache().remove(pinFailKey_(namaKey));
  return master;
}

/**
 * Aktivasi PIN pertama kali: nama + NIK (sekali pakai, hanya bila PIN belum ada) + PIN baru.
 * Setelah aktif, NIK tidak berlaku lagi untuk absen; PIN hanya bisa diganti lewat reset admin.
 */
function aktivasiPinKaryawan(nama, nik, pinBaru) {
  nama = String(nama || '').trim();
  const namaKey = nama.toLowerCase();
  nik = String(nik || '').trim();
  if (!nama || !nik) throw new Error('Nama dan NIK wajib diisi.');
  pinBaru = validasiFormatPin_(pinBaru);

  const gagal = cekKunciPin_(namaKey);
  withScriptLock_(() => {
    const sheet = getOrCreateNikSheet_();
    const rows = readEmployeeRows_(sheet, EMP_NUM_COLS_);
    const idx = indexOfEmployee_(rows, nama);
    if (idx === -1) throw new Error('Nama karyawan tidak ditemukan.');
    const row = rows[idx];
    if (cellText_(row[EMP_COL.PIN_HASH - 1])) {
      throw new Error('PIN sudah pernah diatur. Hubungi admin bila lupa PIN.');
    }
    const nikTerdaftar = cellText_(row[EMP_COL.NIK - 1]);
    if (!nikTerdaftar || nikTerdaftar !== nik) {
      catatGagalPin_(namaKey, gagal);
      throw new Error('NIK tidak sesuai dengan data karyawan.');
    }
    const cell = sheet.getRange(idx + 2, EMP_COL.PIN_HASH);
    cell.setNumberFormat('@');
    cell.setValue(buatPinHash_(pinBaru));
    SpreadsheetApp.flush();
  }, 15000);
  CacheService.getScriptCache().remove(pinFailKey_(namaKey));
  return { success: true };
}


/** Admin / Super Admin: kosongkan PIN karyawan supaya bisa diaktifkan ulang dengan NIK. */
function resetPinKaryawan(nama, token) {
  const session = requireUploadAccess_(token, 'Hanya akun dengan role "Super Admin" atau "Admin" yang boleh mereset PIN karyawan.');
  nama = String(nama || '').trim();
  if (!nama) throw new Error('Nama karyawan wajib diisi.');
  withScriptLock_(() => {
    const sheet = getOrCreateNikSheet_();
    const idx = indexOfEmployee_(readEmployeeRows_(sheet, 1), nama);
    if (idx === -1) throw new Error('Karyawan "' + nama + '" tidak ditemukan.');
    sheet.getRange(idx + 2, EMP_COL.PIN_HASH).setValue('');
    SpreadsheetApp.flush();
  }, 15000);
  CacheService.getScriptCache().remove(pinFailKey_(nama.toLowerCase()));
  catatAuditLog_(session, 'Reset PIN Karyawan', 'Karyawan', nama, 'PIN absen dikosongkan; karyawan harus mengaktifkan ulang');
  return getDaftarKaryawanInternal_();
}

function verifyLoginAbsenOnline(nama, pin) {
  const master = verifyKaryawanByNamaPin_(nama, pin);
  return { department: master.department || '-', fotoProfilUrl: master.fotoProfilUrl || '', nik: master.nik || '' };
}

/**
 * Pencarian nama untuk autocomplete (maks 10 hasil paling relevan).
 * Hasil: { nama, nik:'', label }. NIK TIDAK dikembalikan: fungsi ini publik dan
 * NIK adalah satu-satunya "kata sandi" absen, jadi siapa pun bisa absen atas
 * nama orang lain bila NIK ikut terkirim. Frontend meminta NIK diketik sendiri.
 */
function searchKaryawanUntukAbsen(query) {
  query = String(query || '').trim().toLowerCase();
  if (!query) return [];

  const sheet = getOrCreateNikSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); // kolom A = Nama
  const results = [];

  data.forEach(row => {
    const nama = String(row[0] || '').trim();
    if (!nama) return;
    const namaLower = nama.toLowerCase();

    let score = -1;
    if (namaLower === query) score = 100;
    else if (namaLower.startsWith(query)) score = 80;
    else if (namaLower.includes(query)) score = 50;

    if (score > -1) results.push({ nama: nama, score: score });
  });

  results.sort((a, b) => b.score - a.score || a.nama.localeCompare(b.nama));
  return results.slice(0, 10).map(r => ({ nama: r.nama, nik: '', label: r.nama }));
}

/* ============================ FOTO ============================ */

function getOrCreateFotoAbsenRootFolder_() {
  const it = DriveApp.getFoldersByName(FOTO_ABSEN_ROOT_FOLDER_NAME_);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(FOTO_ABSEN_ROOT_FOLDER_NAME_);
}

function getOrCreateFotoAbsenBulanFolder_(tahun, bulan) {
  const root = getOrCreateFotoAbsenRootFolder_();
  const namaFolder = tahun + '-' + String(bulan).padStart(2, '0');
  const it = root.getFoldersByName(namaFolder);
  if (it.hasNext()) return it.next();
  return root.createFolder(namaFolder);
}

/** Base64 gambar (boleh berawalan data:image/...;base64,) -> Blob JPEG bernama fileName. Error bila kosong atau > FOTO_MAX_BYTES_. */
function decodeFotoBase64_(fotoBase64, fileName) {
  const raw = String(fotoBase64).replace(/^data:image\/\w+;base64,/, '');
  if (Math.floor(raw.length * 3 / 4) > FOTO_MAX_BYTES_) {
    throw new Error('Ukuran foto maksimal ' + (FOTO_MAX_BYTES_ / 1024 / 1024) + ' MB.');
  }
  return Utilities.newBlob(Utilities.base64Decode(raw), 'image/jpeg', fileName);
}

/** Simpan foto selfie ke folder bulan berjalan; bila gagal, kembalikan '' (absen tetap dicatat). */
function simpanFotoAbsen_(nama, jenis, fotoBase64, now, tz) {
  try {
    const folder = getOrCreateFotoAbsenBulanFolder_(now.getFullYear(), now.getMonth() + 1);
    const namaFile = nama.replace(/[^\w]+/g, '_') + '_' + jenis + '_' +
      Utilities.formatDate(now, tz, 'yyyyMMdd_HHmmss') + '.jpg';
    return folder.createFile(decodeFotoBase64_(fotoBase64, namaFile)).getUrl();
  } catch (e) {
    Logger.log('Gagal simpan foto absen untuk ' + nama + ': ' + e.message);
    return '';
  }
}

/* ============================ ABSEN ============================ */

/** Durasi (jam) antara dua jam "HH:mm"; keluar lebih awal dari masuk dianggap lewat tengah malam. */
function hitungDurasiJamKerja_(masukStr, keluarStr) {
  let diff = menitDariJam_(keluarStr) - menitDariJam_(masukStr);
  if (diff < 0) diff += 24 * 60;
  return diff / 60;
}

/** "HH:mm" -> menit sejak 00:00. */
function menitDariJam_(s) {
  const parts = String(s).split(':');
  const h = parseInt(parts[0], 10) || 0;
  const m = parseInt(parts[1], 10) || 0;
  return h * 60 + m;
}

/**
 * Catat absen Masuk/Pulang. Absen di luar radius DITOLAK (error, tidak ada yang dicatat).
 * Absen valid selalu masuk ke LogAbsenOnline (beserta foto).
 * Ke Realisasi (+ JamScan, JamKerja) hanya bila periode aktif = bulan berjalan
 * dan karyawan ada di Realisasi: sel yang sudah S/I/C/A/1/2 tidak diubah.
 * Hasil: { status, jam, jarak [, periodeBelumDiganti, periode] }.
 */
function submitAbsenOnline(nama, pin, jenis, lat, lng, akurasi, fotoBase64) {
  nama = String(nama || '').trim();
  pin = String(pin || '').trim();
  jenis = String(jenis || '').trim();
  const latRaw = lat, lngRaw = lng;
  lat = Number(lat); lng = Number(lng);

  if (!nama || !pin) throw new Error('Nama dan PIN wajib diisi.');
  if (jenis !== 'Masuk' && jenis !== 'Pulang') throw new Error('Jenis absen tidak valid.');
  const gpsKosong = (latRaw === null || latRaw === '' || lngRaw === null || lngRaw === '');
  if (gpsKosong || !isFinite(lat) || !isFinite(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180) {
    throw new Error('Data lokasi GPS tidak valid.');
  }
  if (!fotoBase64) throw new Error('Foto selfie wajib diambil sebelum absen.');

  const nik = verifyKaryawanByNamaPin_(nama, pin).nik;   // NIK tetap dicatat di log dari data master

  // Cek lokasi SEBELUM foto disimpan / log ditulis: absen di luar radius tidak meninggalkan jejak.
  const lokasi = getSettingLokasiAbsen_();
  if (!lokasi.lat && !lokasi.lng) {
    throw new Error('Lokasi kantor belum diatur oleh admin.');
  }
  const jarak = haversineDistanceMeter_(lat, lng, lokasi.lat, lokasi.lng);
  if (jarak > lokasi.radius) {
    throw new Error('Absen ditolak: Anda berada ' + jarak + ' m dari lokasi absen (maksimal ' + lokasi.radius + ' m).');
  }
  const status = 'Valid';

  const now = new Date();
  const tz = Session.getScriptTimeZone() || 'GMT+7';
  const jamStr = Utilities.formatDate(now, tz, 'HH:mm');
  const tanggalStr = Utilities.formatDate(now, tz, 'dd/MM/yyyy');

  const fotoUrl = simpanFotoAbsen_(nama, jenis, fotoBase64, now, tz);

  let periodeBelumDiganti = false;
  let namaPeriode = '';

  withScriptLock_(() => {
    getOrCreateLogAbsenOnlineSheet_().appendRow([
      now, tanggalStr, nama, nik, jenis, jamStr, lat, lng, jarak, status, fotoUrl
    ]);

    // Periode aktif bukan bulan berjalan (mis. tanggal 1 sebelum "Mulai Periode Baru"):
    // kolom tanggalnya milik bulan lalu, jadi Realisasi/JamScan/JamKerja tidak disentuh.
    const realisasiSheet = getSheet_(CONFIG.SHEET_REALISASI);
    const periodText = String(realisasiSheet.getRange(2, 1).getValue() || '').trim();
    const my = parsePeriodMonthYear_(periodText);
    if (my && (now.getFullYear() !== my.year || now.getMonth() !== my.monthIndex)) {
      periodeBelumDiganti = true;
      namaPeriode = periodText.replace(/^periode\s*:?\s*/i, '');
    }

    const lastRow = getLastEmployeeRow_(realisasiSheet);
    const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
    const names = numRows > 0
      ? realisasiSheet.getRange(CONFIG.DATA_START_ROW, 2, numRows, 1).getValues().map(r => String(r[0] || '').trim().toLowerCase())
      : [];
    const rowIdx = periodeBelumDiganti ? -1 : names.indexOf(nama.toLowerCase());

    if (rowIdx !== -1) {
      tulisAbsenKeRekap_(realisasiSheet, CONFIG.DATA_START_ROW + rowIdx, now, jenis, jamStr);
    }

    SpreadsheetApp.flush();
  }, 15000);

  const hasil = { status: status, jam: jamStr, jarak: jarak };
  if (periodeBelumDiganti) { hasil.periodeBelumDiganti = true; hasil.periode = namaPeriode; }
  return hasil;
}

/**
 * Tulis Realisasi (=1 bila masih kosong), JamScan, dan JamKerja (saat Pulang) untuk hari ini.
 * Dipanggil di dalam lock.
 *
 * Aturan:
 *  - Masuk ganda: jam Masuk PALING AWAL yang dipertahankan.
 *  - Pulang berulang: jam Pulang TERAKHIR yang dipakai.
 *  - Pulang lintas tengah malam: bila hari ini belum ada Masuk, tetapi kemarin ada Masuk
 *    tanpa Pulang dan jam Pulang lebih kecil dari jam Masuk kemarin, maka Pulang & JamKerja
 *    ditulis ke kolom hari kemarin, dan Realisasi hari ini tidak diubah.
 */
function tulisAbsenKeRekap_(realisasiSheet, rowNum, now, jenis, jamStr) {
  const dayCount = getDayColCount_(realisasiSheet);
  const dayIdx = now.getDate() - 1;
  if (dayIdx < 0 || dayIdx >= dayCount) return;

  const ss = getSpreadsheet_();
  let jamScanSheet = ss.getSheetByName(CONFIG.SHEET_JAMSCAN);
  if (!jamScanSheet) jamScanSheet = getOrSyncJamScanSheet_();

  const col = CONFIG.FIRST_DAY_COL + dayIdx;
  const scanCell = jamScanSheet.getRange(rowNum, col);
  const existing = splitScanPair_(scanCell.getValue());

  // --- Pulang lintas tengah malam ---
  if (jenis === 'Pulang' && !existing.masuk && dayIdx > 0) {
    const prevCol = col - 1;
    const prevCell = jamScanSheet.getRange(rowNum, prevCol);
    const prev = splitScanPair_(prevCell.getValue());
    if (prev.masuk && !prev.keluar && menitDariJam_(jamStr) < menitDariJam_(prev.masuk)) {
      const jam = hitungDurasiJamKerja_(prev.masuk, jamStr);
      if (jam > 0 && jam <= JAM_KERJA_MAKS_) {
        prevCell.setValue(formatScanPair_(prev.masuk, jamStr));
        tulisJamKerja_(ss, rowNum, prevCol, jam);
        return;
      }
    }
  }

  // --- Alur biasa (shift dalam satu hari) ---
  const currentVal = realisasiSheet.getRange(rowNum, col).getValue();
  const currentCode = (currentVal === null || currentVal === undefined) ? '' : String(currentVal).trim().toUpperCase();
  const isProtected = ['S', 'I', 'C', 'A'].includes(currentCode);
  if (!isProtected && currentCode !== '1' && currentCode !== '2') {
    realisasiSheet.getRange(rowNum, col).setValue(1);
  }

  const masuk = jenis === 'Masuk' ? (existing.masuk || jamStr) : existing.masuk;
  const keluar = jenis === 'Pulang' ? jamStr : existing.keluar;
  scanCell.setValue(formatScanPair_(masuk, keluar));

  // Jam kerja terhitung ditulis ke JamKerja saat absen Pulang (bila Masuk-nya ada).
  if (jenis === 'Pulang' && masuk && keluar) {
    const jam = hitungDurasiJamKerja_(masuk, keluar);
    if (jam > 0 && jam <= JAM_KERJA_MAKS_) tulisJamKerja_(ss, rowNum, col, jam);
  }
}

/** Tulis jam kerja (jam, dibulatkan 2 desimal) ke sheet JamKerja pada baris/kolom tertentu. */
function tulisJamKerja_(ss, rowNum, col, jam) {
  let jamKerjaSheet = ss.getSheetByName(CONFIG.SHEET_JAMKERJA);
  if (!jamKerjaSheet) jamKerjaSheet = getOrSyncJamKerjaSheet_();
  jamKerjaSheet.getRange(rowNum, col).setValue(Math.round(jam * 100) / 100);
}

/* ============================ RIWAYAT ============================ */

function getRiwayatAbsen(nama, pin) {
  nama = String(nama || '').trim();
  if (!nama || !String(pin || '').trim()) throw new Error('Nama dan PIN wajib diisi.');
  const nik = verifyKaryawanByNamaPin_(nama, pin).nik;
  if (!nik) return [];

  const sheet = getOrCreateLogAbsenOnlineSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  return sheet.getRange(2, 1, lastRow - 1, 10).getValues()
    .filter(r => String(r[3] || '').trim() === nik)
    .map(r => ({
      tanggal: formatTanggalAbsen_(r[1]),
      jenis: r[4],
      jam: formatJam_(r[5]),
      status: r[9]
    }))
    .reverse();
}

/**
 * Periode aktif (bulan/tahun dari label Realisasi!A2) untuk statistik di halaman Absensi Online.
 * Hasil: { tahun, bulan (1-12), label }. Label tak terbaca -> bulan berjalan.
 */
function getPeriodeAktif() {
  const now = new Date();
  let tahun = now.getFullYear(), bulan = now.getMonth() + 1, label = '';
  const sheet = getSpreadsheet_().getSheetByName(CONFIG.SHEET_REALISASI);
  if (sheet) {
    label = String(sheet.getRange(2, 1).getValue() || '').trim();
    const my = parsePeriodMonthYear_(label);
    if (my) { tahun = my.year; bulan = my.monthIndex + 1; }
  }
  return { tahun: tahun, bulan: bulan, label: label.replace(/^periode\s*:?\s*/i, '') };
}

function formatTanggalAbsen_(value) {
  if (value === null || value === undefined || value === '') return '-';
  if (value instanceof Date) {
    return Utilities.formatDate(value, getSpreadsheet_().getSpreadsheetTimeZone(), 'dd/MM/yyyy');
  }
  return String(value).trim();
}

function formatJam_(value) {
  if (value === null || value === undefined || value === '') return '-';
  if (value instanceof Date) {
    return Utilities.formatDate(value, getSpreadsheet_().getSpreadsheetTimeZone(), 'HH:mm');
  }
  return String(value).trim();
}

/* ============================ RETENSI FOTO ============================ */

/** Penanda di LogAbsenOnline untuk foto yang sudah dihapus. */
function labelFotoDihapus_() {
  return '[Foto dihapus - retensi ' + RETENSI_FOTO_ABSEN_BULAN_ + ' bulan]';
}

/** Awal bulan batas retensi: folder foto SEBELUM tanggal ini dibuang. Dihitung per bulan (tidak meluber di tanggal 29-31). */
function batasAwalRetensiFoto_(now) {
  return new Date(now.getFullYear(), now.getMonth() - RETENSI_FOTO_ABSEN_BULAN_, 1);
}

function hapusFotoAbsenLama_() {
  const root = getOrCreateFotoAbsenRootFolder_();
  const batasAwalBulan = batasAwalRetensiFoto_(new Date());

  const folders = root.getFolders();
  let totalDihapus = 0;

  while (folders.hasNext()) {
    const folder = folders.next();
    const match = folder.getName().match(/^(\d{4})-(\d{2})$/);
    if (!match) continue;

    const folderDate = new Date(parseInt(match[1], 10), parseInt(match[2], 10) - 1, 1);
    if (folderDate < batasAwalBulan) {
      const files = folder.getFiles();
      while (files.hasNext()) {
        files.next().setTrashed(true);
        totalDihapus++;
      }
      folder.setTrashed(true);
    }
  }

  Logger.log('hapusFotoAbsenLama_: ' + totalDihapus + ' foto (folder > ' + RETENSI_FOTO_ABSEN_BULAN_ + ' bulan) dipindah ke trash.');
  tandaiFotoTerhapusDiSheet_(batasAwalBulan);
}

/** Ganti FotoUrl baris log yang lebih lama dari batas dengan penanda. Hanya kolom FotoUrl yang ditulis. */
function tandaiFotoTerhapusDiSheet_(batasWaktu) {
  const sheet = getOrCreateLogAbsenOnlineSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, LOG_ABSEN_FOTO_COL_).getValues();
  const label = labelFotoDihapus_();
  const fotoCol = data.map(row => [row[LOG_ABSEN_FOTO_COL_ - 1]]);
  let changed = false;

  data.forEach((row, i) => {
    const timestamp = row[0];
    const fotoUrl = row[LOG_ABSEN_FOTO_COL_ - 1];
    if (timestamp instanceof Date && timestamp < batasWaktu && fotoUrl && fotoUrl !== label) {
      fotoCol[i][0] = label;
      changed = true;
    }
  });

  if (changed) {
    sheet.getRange(2, LOG_ABSEN_FOTO_COL_, fotoCol.length, 1).setValues(fotoCol);
    SpreadsheetApp.flush();
  }
}

/** Jalankan SEKALI dari editor untuk memasang trigger bulanan. */
function setupTriggerHapusFotoAbsenLama() {
  requireScriptOwner_();
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'hapusFotoAbsenLama_') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('hapusFotoAbsenLama_')
    .timeBased()
    .onMonthDay(1)
    .atHour(3)
    .create();

  Logger.log('Trigger hapus foto absen lama terpasang: tiap tanggal 1 jam 03:00.');
}

/** Alat uji manual: tampilkan riwayat absen karyawan pertama di DataKaryawan (atau yang namanya mengandung `namaCari`). */
function testRiwayat(namaCari) {
  requireScriptOwner_();
  const sheet = getOrCreateNikSheet_();
  if (sheet.getLastRow() < 2) { Logger.log('DataKaryawan kosong.'); return; }
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const cari = String(namaCari || '').toLowerCase();
  const baris = cari ? data.find(r => String(r[0]).toLowerCase().includes(cari)) : data[0];
  Logger.log('NIK ditemukan: ' + JSON.stringify(baris));

  if (baris) {
    Logger.log('Riwayat kini butuh PIN karyawan; uji lewat halaman Absen Online.');
  }
}

/* ============================ FOTO PROFIL ============================ */

function getOrCreateFotoProfilFolder_() {
  const it = DriveApp.getFoldersByName(FOTO_PROFIL_ROOT_FOLDER_NAME_);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(FOTO_PROFIL_ROOT_FOLDER_NAME_);
}

function setFotoProfilKaryawan(nama, pin, fotoBase64) {
  nama = String(nama || '').trim();
  pin = String(pin || '').trim();
  if (!nama || !pin) throw new Error('Nama dan PIN wajib diisi.');
  if (!fotoBase64) throw new Error('Foto wajib dipilih.');

  verifyKaryawanByNamaPin_(nama, pin);

  const namaFile = nama.replace(/[^\w]+/g, '_') + '_profil_' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+7', 'yyyyMMdd_HHmmss') + '.jpg';
  const blob = decodeFotoBase64_(fotoBase64, namaFile);   // ukuran diperiksa sebelum ada file dibuat

  const folder = getOrCreateFotoProfilFolder_();
  const file = folder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    // Domain Workspace yang melarang sharing publik: foto tetap tersimpan, hanya mungkin tidak tampil bagi yang tak punya akses.
    Logger.log('Gagal set sharing foto profil: ' + e.message);
  }
  const fotoUrl = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w400';

  withScriptLock_(() => {
    const sheet = getOrCreateNikSheet_();
    const lastRow = sheet.getLastRow();
    const names = lastRow < 2 ? [] : sheet.getRange(2, 1, lastRow - 1, 1).getValues()
      .map(r => String(r[0] || '').trim().toLowerCase());
    const idx = names.indexOf(nama.toLowerCase());
    if (idx === -1) throw new Error('Karyawan tidak ditemukan di DataKaryawan.');
    sheet.getRange(idx + 2, EMP_COL.FOTO_PROFIL).setValue(fotoUrl);
    SpreadsheetApp.flush();
  }, 15000);

  return { fotoProfilUrl: fotoUrl };
}
