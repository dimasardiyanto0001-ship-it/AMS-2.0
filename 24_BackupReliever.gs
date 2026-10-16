/*************************************************************
 * 24_BackupReliever.gs
 * Backup Reliever: reliever (pengganti dari luar) absen lewat halaman
 * Absen Online, dicatat terpisah dari karyawan tetap. Endpoint PUBLIK
 * (tanpa token): identitas reliever = namanya di DataReliever.
 *
 * Sheet:
 *  - DataReliever           : Nama | Perusahaan | NoHP | Status (diisi manual oleh admin)
 *  - LogAbsenOnlineReliever : log mentah tiap absen reliever (tersembunyi)
 *  - BackupReliever         : grid periode (struktur sama dengan Jadwal/Realisasi):
 *                             baris = reliever, kolom = tanggal, isi sel = nama karyawan
 *                             yang digantikan. Grid ini ikut periode aktif: saat label
 *                             periode berganti, grid lama disimpan dengan nama
 *                             "BackupReliever <periode lama>" dan dibuat grid baru.
 *
 * Memakai dari 20_AbsenOnline.gs: getSettingLokasiAbsen_, haversineDistanceMeter_,
 * simpanFotoAbsen_.
 *************************************************************/

const RELIEVER_FIELD_MAKS_ = 80;
const RELIEVER_QUERY_MAKS_ = 60;
const BACKUP_GRID_NAMA_ARSIP_MAKS_ = 90;

/* ============================ SETUP SHEET ============================ */

function getOrCreateDataRelieverSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(CONFIG.SHEET_DATARELIEVER);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_DATARELIEVER);
    sheet.getRange(1, 1, 1, 4).setValues([['Nama', 'Perusahaan', 'NoHP', 'Status']]).setFontWeight('bold');
    sheet.getRange(1, 3, sheet.getMaxRows(), 1).setNumberFormat('@'); // 0 di depan nomor HP jangan hilang
    sheet.setColumnWidths(1, 4, 160);
    sheet.getRange('F1').setValue('Daftar reliever yang boleh absen lewat halaman Absen Online. Tambah baris manual di sini dulu (Status="Aktif") sebelum reliever bisa dicari di form.');
  }
  return sheet;
}

const LOG_ABSEN_RELIEVER_HEADERS_ = [
  'Timestamp', 'Tanggal', 'Nama', 'AsalPerusahaan', 'NamaDigantikan', 'JenisDigantikan',
  'Jenis', 'Jam', 'Latitude', 'Longitude', 'JarakMeter', 'Status', 'FotoUrl'
];

function getOrCreateLogAbsenRelieverSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(CONFIG.SHEET_LOGABSENRELIEVER);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_LOGABSENRELIEVER);
    sheet.getRange(1, 1, 1, LOG_ABSEN_RELIEVER_HEADERS_.length).setValues([LOG_ABSEN_RELIEVER_HEADERS_]).setFontWeight('bold');
    sheet.hideSheet();
  }
  return sheet;
}

/** Salin label periode + header tanggal + nama hari dari Realisasi ke sheet grid. */
function isiHeaderGridDariRealisasi_(sheet, realisasiSheet) {
  const dayCount = getDayColCount_(realisasiSheet);
  const kolomButuh = CONFIG.FIRST_DAY_COL + 31 - 1;
  if (sheet.getMaxColumns() < kolomButuh) sheet.insertColumnsAfter(sheet.getMaxColumns(), kolomButuh - sheet.getMaxColumns());
  sheet.getRange(1, 1, 1, 3).setValues([['No', 'Nama Reliever', 'Perusahaan']]).setFontWeight('bold');
  sheet.getRange(2, 1).setValue(realisasiSheet.getRange(2, 1).getValue());
  if (dayCount > 0) {
    [CONFIG.HEADER_ROW, CONFIG.DAY_NAME_ROW].forEach(row => {
      sheet.getRange(row, CONFIG.FIRST_DAY_COL, 1, dayCount)
        .setValues(realisasiSheet.getRange(row, CONFIG.FIRST_DAY_COL, 1, dayCount).getValues());
    });
  }
}

function buatGridBackupReliever_(realisasiSheet) {
  const sheet = getSpreadsheet_().insertSheet(CONFIG.SHEET_BACKUPRELIEVER);
  isiHeaderGridDariRealisasi_(sheet, realisasiSheet);
  sheet.hideSheet();
  return sheet;
}

/**
 * Grid BackupReliever untuk periode aktif (label di Realisasi!A2).
 * Dibuat bila belum ada. Bila labelnya sudah beda dari periode aktif: grid yang ada isinya
 * disimpan (nama sheet diberi label lama) lalu dibuat grid baru; grid kosong cukup diperbarui headernya.
 * Struktur kolomnya SENGAJA sama persis dengan Jadwal/Realisasi supaya bisa dibaca readGridFromSheetObj_().
 */
function getOrCreateBackupRelieverSheet_() {
  const realisasiSheet = getSheet_(CONFIG.SHEET_REALISASI);
  const sheet = getSpreadsheet_().getSheetByName(CONFIG.SHEET_BACKUPRELIEVER);
  if (!sheet) return buatGridBackupReliever_(realisasiSheet);

  const labelAktif = String(realisasiSheet.getRange(2, 1).getValue() || '').trim();
  const labelGrid = String(sheet.getRange(2, 1).getValue() || '').trim();
  if (labelAktif === labelGrid) return sheet;

  if (getLastEmployeeRow_(sheet) < CONFIG.DATA_START_ROW) {
    isiHeaderGridDariRealisasi_(sheet, realisasiSheet);
    [CONFIG.HEADER_ROW, CONFIG.DAY_NAME_ROW].forEach(row => {
      const sisa = 31 - getDayColCount_(realisasiSheet);
      if (sisa > 0) sheet.getRange(row, CONFIG.FIRST_DAY_COL + 31 - sisa, 1, sisa).clearContent();
    });
    return sheet;
  }

  const ss = getSpreadsheet_();
  const dasar = ('BackupReliever ' + (labelGrid || 'lama')).replace(/[\[\]*?:\/\\]/g, '-').substring(0, BACKUP_GRID_NAMA_ARSIP_MAKS_);
  let nama = dasar;
  for (let n = 2; ss.getSheetByName(nama); n++) nama = dasar + ' (' + n + ')';
  sheet.setName(nama);
  return buatGridBackupReliever_(realisasiSheet);
}

/** Baris reliever di grid berdasarkan nama; belum ada -> ditambah di akhir. */
function getOrCreateReliverRow_(sheet, nama, perusahaan) {
  const lastRow = getLastEmployeeRow_(sheet);
  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;

  if (numRows > 0) {
    const target = nameKey_(nama);
    const names = sheet.getRange(CONFIG.DATA_START_ROW, 2, numRows, 1).getValues().map(r => nameKey_(r[0] === null || r[0] === undefined ? '' : r[0]));
    const idx = names.indexOf(target);
    if (idx !== -1) return CONFIG.DATA_START_ROW + idx;
  }

  const newRow = lastRow + 1;
  if (newRow > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 50);
  sheet.getRange(newRow, 1, 1, 3).setValues([[Math.max(numRows, 0) + 1, guardSheetText_(nama, RELIEVER_FIELD_MAKS_), guardSheetText_(perusahaan || '', RELIEVER_FIELD_MAKS_)]]);
  return newRow;
}

/* ============================ VERIFIKASI & PENCARIAN (PUBLIK) ============================ */

function isRelieverNonaktif_(status) {
  const s = String(status === null || status === undefined ? '' : status).trim().toLowerCase();
  return s === 'nonaktif' || s === 'tidak aktif';
}

function verifyReliverByNama_(nama) {
  const sheet = getOrCreateDataRelieverSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('Belum ada reliever terdaftar. Hubungi admin.');

  const target = nameKey_(nama || '');
  const row = sheet.getRange(2, 1, lastRow - 1, 4).getValues()
    .find(r => nameKey_(r[0] === null || r[0] === undefined ? '' : r[0]) === target && target !== '');
  if (!row) throw new Error('Nama reliever "' + nama + '" belum terdaftar. Hubungi admin untuk didaftarkan dulu.');
  if (isRelieverNonaktif_(row[3])) throw new Error('Reliever "' + nama + '" berstatus nonaktif.');

  return { nama: row[0], perusahaan: row[1], noHp: row[2] };
}

function verifyLoginReliever(nama) {
  verifyReliverByNama_(nama);
  return { success: true };
}

/** Autocomplete nama reliever di form Absen Online (maks 10 hasil paling relevan). */
function searchReliverUntukAbsen(query) {
  query = String(query || '').trim().toLowerCase().substring(0, RELIEVER_QUERY_MAKS_);
  if (!query) return [];

  const sheet = getOrCreateDataRelieverSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const results = [];
  sheet.getRange(2, 1, lastRow - 1, 2).getValues().forEach(row => {
    const nama = String(row[0] === null || row[0] === undefined ? '' : row[0]).trim();
    const perusahaan = String(row[1] === null || row[1] === undefined ? '' : row[1]).trim();
    if (!nama) return;
    const namaLower = nama.toLowerCase();

    let score = -1;
    if (namaLower === query) score = 100;
    else if (namaLower.startsWith(query)) score = 80;
    else if (namaLower.includes(query)) score = 50;
    if (score > -1) results.push({ nama: nama, perusahaan: perusahaan, score: score });
  });

  results.sort((a, b) => b.score - a.score || a.nama.localeCompare(b.nama));
  return results.slice(0, 10).map(r => ({ nama: r.nama, perusahaan: r.perusahaan, label: r.nama + ' - ' + r.perusahaan }));
}

/**
 * Karyawan pada kolom Realisasi hari dayNum yang berkode A/I/S: [{nama, jenis}].
 * Hanya bila periode Realisasi = bulan berjalan; selain itu [].
 * (dayNum kosong = tanggal hari ini.)
 */
function bacaKaryawanAbsenRealisasi_(dayNum) {
  const realisasiSheet = getSheet_(CONFIG.SHEET_REALISASI);
  const my = parsePeriodMonthYear_(String(realisasiSheet.getRange(2, 1).getValue() || '').trim());
  if (!my) return [];

  const today = new Date();
  if (today.getFullYear() !== my.year || today.getMonth() !== my.monthIndex) return [];

  const hari = dayNum || today.getDate();
  if (hari < 1 || hari > getDayColCount_(realisasiSheet)) return [];

  const numRows = getLastEmployeeRow_(realisasiSheet) - CONFIG.DATA_START_ROW + 1;
  if (numRows <= 0) return [];

  const names = realisasiSheet.getRange(CONFIG.DATA_START_ROW, 2, numRows, 1).getValues();
  const codes = realisasiSheet.getRange(CONFIG.DATA_START_ROW, CONFIG.FIRST_DAY_COL + hari - 1, numRows, 1).getValues();

  const result = [];
  for (let i = 0; i < numRows; i++) {
    const nama = String(names[i][0] === null || names[i][0] === undefined ? '' : names[i][0]).trim();
    const kode = String(codes[i][0] === null || codes[i][0] === undefined ? '' : codes[i][0]).trim().toUpperCase();
    if (nama && (kode === 'A' || kode === 'I' || kode === 'S')) {
      const info = CONFIG.STATUS_CODES[kode];
      result.push({ nama: nama, jenis: info ? info.label : kode });
    }
  }
  return result;
}

/** Dropdown "menggantikan siapa hari ini?" di form reliever. */
function getKaryawanAbsenHariIni() {
  return bacaKaryawanAbsenRealisasi_();
}

/* ============================ SUBMIT ABSEN RELIEVER (PUBLIK) ============================ */

/**
 * Catat absen Masuk/Pulang reliever. Selalu masuk ke LogAbsenOnlineReliever (beserta foto).
 * Grid BackupReliever hanya ditandai saat Masuk dan hanya bila periode aktif = bulan berjalan.
 * Hasil: { status, jam, jarak [, periodeBelumDiganti, periode] }.
 */
function submitAbsenOnlineReliever(namaReliever, asalPerusahaan, namaDigantikan, jenisDigantikan, jenis, lat, lng, akurasi, fotoBase64) {
  namaReliever = String(namaReliever || '').trim();
  asalPerusahaan = String(asalPerusahaan || '').trim();
  namaDigantikan = String(namaDigantikan || '').trim();
  jenisDigantikan = String(jenisDigantikan || '').trim();
  jenis = String(jenis || '').trim();
  const latRaw = lat, lngRaw = lng;
  lat = Number(lat); lng = Number(lng);

  if (!namaReliever) throw new Error('Nama reliever wajib diisi.');
  if (!asalPerusahaan) throw new Error('Asal perusahaan wajib diisi.');
  if (namaReliever.length > RELIEVER_FIELD_MAKS_ || asalPerusahaan.length > RELIEVER_FIELD_MAKS_) throw new Error('Nama atau perusahaan terlalu panjang.');
  if (!namaDigantikan) throw new Error('Wajib pilih karyawan yang digantikan.');
  if (jenis !== 'Masuk' && jenis !== 'Pulang') throw new Error('Jenis absen tidak valid.');
  const gpsKosong = (latRaw === null || latRaw === '' || lngRaw === null || lngRaw === '');
  if (gpsKosong || !isFinite(lat) || !isFinite(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180) {
    throw new Error('Data lokasi GPS tidak valid.');
  }
  if (!fotoBase64) throw new Error('Foto selfie wajib diambil sebelum absen.');

  // Reliever mengetik nama & perusahaan sendiri (tidak dicek ke DataReliever, sesuai perubahan terbaru).
  // Pengaman: karyawan yang digantikan harus benar-benar A/I/S hari ini; nama dicatat sesuai ejaan di Realisasi.
  const keyDigantikan = nameKey_(namaDigantikan);
  const cocok = bacaKaryawanAbsenRealisasi_().find(k => nameKey_(k.nama) === keyDigantikan);
  if (!cocok) throw new Error('Karyawan yang digantikan tidak berstatus Alfa/Izin/Sakit hari ini.');
  const dikenal = cocok.nama;
  const perusahaan = asalPerusahaan;

  const lokasi = getSettingLokasiAbsen_();
  const jarak = haversineDistanceMeter_(lat, lng, lokasi.lat, lokasi.lng);
  const status = jarak <= lokasi.radius ? 'Valid' : 'Di Luar Radius';

  const now = new Date();
  const tz = Session.getScriptTimeZone() || 'GMT+7';
  const jamStr = Utilities.formatDate(now, tz, 'HH:mm');
  const tanggalStr = Utilities.formatDate(now, tz, 'dd/MM/yyyy');

  const fotoUrl = simpanFotoAbsen_('RELIEVER_' + namaReliever, jenis, fotoBase64, now, tz);

  let periodeBelumDiganti = false;
  let namaPeriode = '';

  withScriptLock_(() => {
    const realisasiSheet = getSheet_(CONFIG.SHEET_REALISASI);
    const periodText = String(realisasiSheet.getRange(2, 1).getValue() || '').trim();
    const my = parsePeriodMonthYear_(periodText);
    if (my && (now.getFullYear() !== my.year || now.getMonth() !== my.monthIndex)) {
      periodeBelumDiganti = true;
      namaPeriode = periodText.replace(/^periode\s*:?\s*/i, '');
    }

    getOrCreateLogAbsenRelieverSheet_().appendRow([
      now, tanggalStr, namaReliever, guardSheetText_(perusahaan, RELIEVER_FIELD_MAKS_), dikenal,
      guardSheetText_(jenisDigantikan, RELIEVER_FIELD_MAKS_), jenis, jamStr, lat, lng, jarak, status, fotoUrl
    ]);

    if (jenis === 'Masuk' && !periodeBelumDiganti) {
      const grid = getOrCreateBackupRelieverSheet_();
      const dayIdx = now.getDate() - 1;
      if (dayIdx >= 0 && dayIdx < getDayColCount_(grid)) {
        const row = getOrCreateReliverRow_(grid, namaReliever, perusahaan);
        grid.getRange(row, CONFIG.FIRST_DAY_COL + dayIdx).setValue(dikenal);
      }
    }

    SpreadsheetApp.flush();
  }, 15000);

  const hasil = { status: status, jam: jamStr, jarak: jarak };
  if (periodeBelumDiganti) { hasil.periodeBelumDiganti = true; hasil.periode = namaPeriode; }
  return hasil;
}

/* ============================ ALAT BANTU (jalankan dari editor) ============================ */

function buatSheetDataReliever() {
  getOrCreateDataRelieverSheet_();
}

function testBuatSheetBackupReliever() {
  const sheet = getOrCreateBackupRelieverSheet_();
  Logger.log('Sheet berhasil dibuat/ditemukan: ' + sheet.getName());
  Logger.log('Isi baris ' + CONFIG.HEADER_ROW + ' (header tanggal): ' + JSON.stringify(sheet.getRange(CONFIG.HEADER_ROW, 1, 1, 10).getValues()));
  Logger.log('Isi baris ' + CONFIG.DAY_NAME_ROW + ' (nama hari): ' + JSON.stringify(sheet.getRange(CONFIG.DAY_NAME_ROW, 1, 1, 10).getValues()));
}

function testGetKaryawanAbsenHariIni() {
  const hasil = getKaryawanAbsenHariIni();
  Logger.log('Jumlah karyawan A/I/S hari ini: ' + hasil.length);
  Logger.log(JSON.stringify(hasil));
}

/** Diagnosa dropdown "digantikan": karyawan A/I/S pada tanggal tertentu (kosong = hari ini) di periode aktif. */
function debugKaryawanAbsenTanggalManual(tanggal) {
  const realisasiSheet = getSheet_(CONFIG.SHEET_REALISASI);
  const periodText = String(realisasiSheet.getRange(2, 1).getValue() || '').trim();
  Logger.log('Periode (Realisasi!A2): "' + periodText + '" -> ' + JSON.stringify(parsePeriodMonthYear_(periodText)));
  Logger.log('Tanggal server: ' + new Date() + ', jumlah kolom hari: ' + getDayColCount_(realisasiSheet));
  const hasil = bacaKaryawanAbsenRealisasi_(tanggal);
  Logger.log('Karyawan A/I/S (tanggal ' + (tanggal || 'hari ini') + '): ' + JSON.stringify(hasil));
  return hasil;
}
