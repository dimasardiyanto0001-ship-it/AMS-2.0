/*************************************************************
 * 24_BackupReliever.gs
 * Fitur Backup Reliever: reliever (pengganti dari luar) absen lewat
 * halaman Absen Online, dicatat terpisah dari absen karyawan tetap.
 *
 * Sheet terkait:
 *  - DataReliever        : master nama reliever (Nama|Perusahaan|NoHP|Status)
 *  - LogAbsenOnlineReliever : log mentah tiap absen reliever
 *  - BackupReliever       : grid periode (pola SAMA seperti Jadwal/Realisasi,
 *                           supaya bisa dibaca readGridFromSheetObj_ yang
 *                           sudah ada). Baris = reliever, kolom = tanggal,
 *                           isi sel = nama karyawan yang digantikan.
 *************************************************************/

/* ============================ SETUP SHEET ============================ */

function getOrCreateDataRelieverSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_DATARELIEVER);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_DATARELIEVER);
    sheet.getRange(1, 1, 1, 4).setValues([['Nama', 'Perusahaan', 'NoHP', 'Status']]).setFontWeight('bold');
    sheet.setColumnWidths(1, 4, 160);
    sheet.getRange('F1').setValue('Daftar reliever yang boleh absen lewat halaman Absen Online. Tambah baris manual di sini dulu (Status="Aktif") sebelum reliever bisa dicari di form.');
  }
  return sheet;
}

function getOrCreateLogAbsenRelieverSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_LOGABSENRELIEVER);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_LOGABSENRELIEVER);
    sheet.getRange(1, 1, 1, 13).setValues([[
      'Timestamp', 'Tanggal', 'Nama', 'AsalPerusahaan', 'NamaDigantikan', 'JenisDigantikan',
      'Jenis', 'Jam', 'Latitude', 'Longitude', 'JarakMeter', 'Status', 'FotoUrl'
    ]]).setFontWeight('bold');
    sheet.hideSheet();
  }
  return sheet;
}

/** Grid periode untuk reliever -- strukturnya SENGAJA disamakan persis dengan
 *  Jadwal/Realisasi (No|Nama|Perusahaan di kolom A-C, HEADER_ROW/DAY_NAME_ROW/
 *  DATA_START_ROW/FIRST_DAY_COL sama persis) supaya bisa langsung dibaca pakai
 *  readGridFromSheetObj_() yang sudah ada tanpa modifikasi apa pun. */
function getOrCreateBackupRelieverSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_BACKUPRELIEVER);
  if (sheet) return sheet;

  sheet = ss.insertSheet(CONFIG.SHEET_BACKUPRELIEVER);
  const realisasiSheet = getSheet_(CONFIG.SHEET_REALISASI);
  const dayCount = getDayColCount_(realisasiSheet);

  sheet.getRange(1, 1, 1, 3).setValues([['No', 'Nama Reliever', 'Perusahaan']]).setFontWeight('bold');
  sheet.getRange(2, 1).setValue(realisasiSheet.getRange(2, 1).getValue()); // samakan label periode

  if (dayCount > 0) {
    sheet.getRange(CONFIG.HEADER_ROW, CONFIG.FIRST_DAY_COL, 1, dayCount)
      .setValues(realisasiSheet.getRange(CONFIG.HEADER_ROW, CONFIG.FIRST_DAY_COL, 1, dayCount).getValues());
    sheet.getRange(CONFIG.DAY_NAME_ROW, CONFIG.FIRST_DAY_COL, 1, dayCount)
      .setValues(realisasiSheet.getRange(CONFIG.DAY_NAME_ROW, CONFIG.FIRST_DAY_COL, 1, dayCount).getValues());
  }

  sheet.hideSheet();
  return sheet;
}

/** Cari baris reliever berdasarkan nama; kalau belum ada, tambah baris baru di akhir. */
function getOrCreateReliverRow_(sheet, nama, perusahaan) {
  const lastRow = getLastEmployeeRow_(sheet);
  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;

  if (numRows > 0) {
    const names = sheet.getRange(CONFIG.DATA_START_ROW, 2, numRows, 1).getValues()
      .map(r => String(r[0] || '').trim().toLowerCase());
    const idx = names.indexOf(nama.trim().toLowerCase());
    if (idx !== -1) return CONFIG.DATA_START_ROW + idx;
  }

  const newRow = lastRow + 1;
  const newNo = numRows + 1;
  sheet.getRange(newRow, 1, 1, 3).setValues([[newNo, nama, perusahaan || '']]);
  return newRow;
}

/* ============================ VERIFIKASI & PENCARIAN (PUBLIK) ============================ */

function verifyReliverByNama_(nama) {
  const sheet = getOrCreateDataRelieverSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('Belum ada reliever terdaftar. Hubungi admin.');

  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const row = data.find(r => String(r[0] || '').trim().toLowerCase() === String(nama || '').trim().toLowerCase());
  if (!row) throw new Error('Nama reliever "' + nama + '" belum terdaftar. Hubungi admin untuk didaftarkan dulu.');
  if (String(row[3] || '').trim() === 'Nonaktif') throw new Error('Reliever "' + nama + '" berstatus nonaktif.');

  return { nama: row[0], perusahaan: row[1], noHp: row[2] };
}

function verifyLoginReliever(nama) {
  verifyReliverByNama_(nama);
  return { success: true };
}

/** Autocomplete nama reliever di form Absen Online, mirip searchKaryawanUntukAbsen. */
function searchReliverUntukAbsen(query) {
  query = String(query || '').trim().toLowerCase();
  if (!query) return [];

  const sheet = getOrCreateDataRelieverSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const results = [];
  data.forEach(row => {
    const nama = String(row[0] || '').trim();
    const perusahaan = String(row[1] || '').trim();
    if (!nama) return;
    const namaLower = nama.toLowerCase();
    let score = -1;
    if (namaLower === query) score = 100;
    else if (namaLower.startsWith(query)) score = 80;
    else if (namaLower.includes(query)) score = 50;
    if (score > -1) results.push({ nama, perusahaan, score });
  });

  results.sort((a, b) => b.score - a.score || a.nama.localeCompare(b.nama));
  return results.slice(0, 10).map(r => ({ nama: r.nama, perusahaan: r.perusahaan, label: r.nama + ' - ' + r.perusahaan }));
}

/** Dropdown "menggantikan siapa hari ini?" -- scan kolom Realisasi HARI INI, ambil yg A/I/S. */
function getKaryawanAbsenHariIni() {
  const realisasiSheet = getSheet_(CONFIG.SHEET_REALISASI);
  const periodText = String(realisasiSheet.getRange(2, 1).getValue() || '').trim();
  const my = parsePeriodMonthYear_(periodText);
  if (!my) return [];

  const today = new Date();
  if (today.getFullYear() !== my.year || today.getMonth() !== my.monthIndex) return [];

  const dayCount = getDayColCount_(realisasiSheet);
  const dayNum = today.getDate();
  if (dayNum > dayCount) return [];

  const lastRow = getLastEmployeeRow_(realisasiSheet);
  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  if (numRows <= 0) return [];

  const col = CONFIG.FIRST_DAY_COL + (dayNum - 1);
  const names = realisasiSheet.getRange(CONFIG.DATA_START_ROW, 2, numRows, 1).getValues();
  const codes = realisasiSheet.getRange(CONFIG.DATA_START_ROW, col, numRows, 1).getValues();

  const result = [];
  for (let i = 0; i < numRows; i++) {
    const nama = String(names[i][0] || '').trim();
    const kode = String(codes[i][0] || '').trim().toUpperCase();
    if (nama && (kode === 'A' || kode === 'I' || kode === 'S')) {
      const info = CONFIG.STATUS_CODES[kode];
      result.push({ nama: nama, jenis: info ? info.label : kode });
    }
  }
  return result;
}

/* ============================ SUBMIT ABSEN RELIEVER (PUBLIK) ============================ */

function submitAbsenOnlineReliever(namaReliever, asalPerusahaan, namaDigantikan, jenisDigantikan, jenis, lat, lng, akurasi, fotoBase64) {
  namaReliever = String(namaReliever || '').trim();
  asalPerusahaan = String(asalPerusahaan || '').trim();
  namaDigantikan = String(namaDigantikan || '').trim();
  jenisDigantikan = String(jenisDigantikan || '').trim();
  jenis = String(jenis || '').trim();
  lat = Number(lat); lng = Number(lng);

  if (!namaReliever) throw new Error('Nama reliever wajib diisi.');
  if (!asalPerusahaan) throw new Error('Asal perusahaan wajib diisi.');
  if (namaReliever.length > 80 || asalPerusahaan.length > 80) throw new Error('Nama atau perusahaan terlalu panjang.');
  if (!namaDigantikan) throw new Error('Wajib pilih karyawan yang digantikan.');
  if (jenis !== 'Masuk' && jenis !== 'Pulang') throw new Error('Jenis absen tidak valid.');
  if (isNaN(lat) || isNaN(lng)) throw new Error('Data lokasi GPS tidak valid.');
  if (!fotoBase64) throw new Error('Foto selfie wajib diambil sebelum absen.');

  // Tidak lagi dicek ke sheet DataReliever: reliever mengetik nama & perusahaan sendiri.
  // Pengaman: karyawan yang digantikan harus benar-benar A/I/S hari ini.
  const daftarHariIni = getKaryawanAbsenHariIni().map(k => k.nama.toLowerCase());
  if (daftarHariIni.indexOf(namaDigantikan.toLowerCase()) === -1) {
    throw new Error('Karyawan yang digantikan tidak berstatus Alfa/Izin/Sakit hari ini.');
  }

  const lokasi = getSettingLokasiAbsen_();
  const jarak = haversineDistanceMeter_(lat, lng, lokasi.lat, lokasi.lng);
  const status = jarak <= lokasi.radius ? 'Valid' : 'Di Luar Radius';

  const now = new Date();
  const tz = Session.getScriptTimeZone() || 'GMT+7';
  const jamStr = Utilities.formatDate(now, tz, 'HH:mm');
  const tanggalStr = Utilities.formatDate(now, tz, 'dd/MM/yyyy');

  let fotoUrl = '';
  try {
    const folder = getOrCreateFotoAbsenBulanFolder_(now.getFullYear(), now.getMonth() + 1);
    const namaFile = 'RELIEVER_' + namaReliever.replace(/[^\w]+/g, '_') + '_' + jenis + '_' +
      Utilities.formatDate(now, tz, 'yyyyMMdd_HHmmss') + '.jpg';
    const rawBase64 = fotoBase64.replace(/^data:image\/\w+;base64,/, '');
    const bytes = Utilities.base64Decode(rawBase64);
    const blob = Utilities.newBlob(bytes, 'image/jpeg', namaFile);
    fotoUrl = folder.createFile(blob).getUrl();
  } catch (e) {
    Logger.log('Gagal simpan foto absen reliever untuk ' + namaReliever + ': ' + e.message);
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    getOrCreateLogAbsenRelieverSheet_().appendRow([
      now, tanggalStr, namaReliever, asalPerusahaan, namaDigantikan, jenisDigantikan,
      jenis, jamStr, lat, lng, jarak, status, fotoUrl
    ]);

    // Tandai grid BackupReliever hanya saat absen Masuk.
    if (jenis === 'Masuk') {
      const grid = getOrCreateBackupRelieverSheet_();
      const row = getOrCreateReliverRow_(grid, namaReliever, asalPerusahaan);
      const dayCount = getDayColCount_(grid);
      const dayIdx = now.getDate() - 1;
      if (dayIdx >= 0 && dayIdx < dayCount) {
        grid.getRange(row, CONFIG.FIRST_DAY_COL + dayIdx).setValue(namaDigantikan);
      }
    }

    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  return { status: status, jam: jamStr, jarak: jarak };
}
function testBuatSheetBackupReliever() {
  const sheet = getOrCreateBackupRelieverSheet_();
  Logger.log('Sheet berhasil dibuat/ditemukan: ' + sheet.getName());
  Logger.log('Isi baris 4 (header tanggal): ' + JSON.stringify(sheet.getRange(4, 1, 1, 10).getValues()));
  Logger.log('Isi baris 5 (nama hari): ' + JSON.stringify(sheet.getRange(5, 1, 1, 10).getValues()));
}
function testGetKaryawanAbsenHariIni() {
  const hasil = getKaryawanAbsenHariIni();
  Logger.log('Jumlah karyawan A/I/S hari ini: ' + hasil.length);
  Logger.log(JSON.stringify(hasil));
}
function debugKaryawanAbsenHariIni() {
  const realisasiSheet = getSheet_(CONFIG.SHEET_REALISASI);
  const periodText = String(realisasiSheet.getRange(2, 1).getValue() || '').trim();
  Logger.log('Periode text (Realisasi!A2): "' + periodText + '"');

  const my = parsePeriodMonthYear_(periodText);
  Logger.log('Hasil parsePeriodMonthYear_: ' + JSON.stringify(my));

  const today = new Date();
  Logger.log('Tanggal server sekarang: ' + today + ' | getFullYear=' + today.getFullYear() + ' getMonth=' + today.getMonth());

  if (!my) { Logger.log('BERHENTI: periode tidak terbaca.'); return; }
  if (today.getFullYear() !== my.year || today.getMonth() !== my.monthIndex) {
    Logger.log('BERHENTI: tahun/bulan server tidak cocok dengan periode aktif.');
    return;
  }

  const dayCount = getDayColCount_(realisasiSheet);
  const dayNum = today.getDate();
  Logger.log('dayCount=' + dayCount + ', dayNum (tanggal hari ini)=' + dayNum);
  if (dayNum > dayCount) { Logger.log('BERHENTI: dayNum melebihi dayCount.'); return; }

  const col = CONFIG.FIRST_DAY_COL + (dayNum - 1);
  const lastRow = getLastEmployeeRow_(realisasiSheet);
  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const names = realisasiSheet.getRange(CONFIG.DATA_START_ROW, 2, numRows, 1).getValues();
  const codes = realisasiSheet.getRange(CONFIG.DATA_START_ROW, col, numRows, 1).getValues();

  Logger.log('Kolom yang dibaca untuk hari ini: kolom ke-' + col);
  for (let i = 0; i < numRows; i++) {
    Logger.log(names[i][0] + ' -> kode: "' + codes[i][0] + '"');
  }
}
function debugKaryawanAbsenTanggalManual(tanggalUjiDbl) {
  // tanggalUjiDbl = angka tanggal yang mau diuji, misal 15 untuk tanggal 15 di periode aktif
  const realisasiSheet = getSheet_(CONFIG.SHEET_REALISASI);
  const dayCount = getDayColCount_(realisasiSheet);
  const dayNum = tanggalUjiDbl;
  if (dayNum > dayCount) { Logger.log('dayNum melebihi dayCount.'); return; }

  const col = CONFIG.FIRST_DAY_COL + (dayNum - 1);
  const lastRow = getLastEmployeeRow_(realisasiSheet);
  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const names = realisasiSheet.getRange(CONFIG.DATA_START_ROW, 2, numRows, 1).getValues();
  const codes = realisasiSheet.getRange(CONFIG.DATA_START_ROW, col, numRows, 1).getValues();

  Logger.log('Menguji tanggal ke-' + dayNum + ' (kolom ke-' + col + ')');
  for (let i = 0; i < numRows; i++) {
    Logger.log(names[i][0] + ' -> kode: "' + codes[i][0] + '"');
  }
}

function testDebugTanggal15() {
  debugKaryawanAbsenTanggalManual(15); // ganti 15 dengan tanggal yang Anda tahu ada karyawan A/I/S-nya
}
function buatSheetDataReliever() {
  getOrCreateDataRelieverSheet_();
}
