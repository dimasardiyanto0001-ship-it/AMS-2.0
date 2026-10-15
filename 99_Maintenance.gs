/*************************************************************
 * 99_Maintenance.gs
 * Alat perawatan dan diagnosis. Dijalankan MANUAL dari Apps Script Editor
 * (pilih fungsi di dropdown lalu Jalankan); TIDAK ada route web ke sini.
 * Fungsi berakhiran "_" disembunyikan dari dropdown editor, jadi semua alat di sini
 * sengaja bernama tanpa garis bawah di belakang.
 *
 *   DIAGNOSIS (aman, hanya membuat sheet laporan "_Debug..." / "_Audit..." atau log)
 *     debugNameMismatch()      nama di Rekap yang tidak cocok dengan Jadwal/Realisasi/JamKerja
 *     findBlankRowGaps()       baris karyawan terisi yang tidak ikut terhitung sistem
 *     auditJamNyangkutOFF()    jam kerja/scan terisi padahal Realisasi = OFF
 *     cekSpreadsheetAktif()    nama, ID, URL spreadsheet yang sedang dipakai
 *     testGetJadwalHariIni(n)  uji jadwal shift hari ini untuk satu nama
 *     debugShiftSheet()        isi sheet SettingJamKerja
 *
 *   PERBAIKAN (MENGUBAH DATA periode aktif; sheet arsip tidak disentuh)
 *     bersihkanJamNyangkutOFF()           kosongkan jam yang nyangkut di sel OFF
 *     fixJamKerjaPeriodeAktifSekarang()   kosongkan semua JamKerja/JamScan periode aktif
 *     backfillRekapBulananDariArsip()     isi Rekap_Bulanan dari semua periode arsip
 *************************************************************/

/* ============================ HELPER ============================ */

/** Log + tampilkan alert bila dijalankan dari menu Sheets (dari editor, alert dilewati). */
function showAlertQuietly_(message) {
  Logger.log(message);
  try { SpreadsheetApp.getUi().alert(message); } catch (e) { /* dijalankan dari editor: abaikan */ }
}

/** Buat ulang sheet laporan diagnosis (yang lama dihapus). */
function recreateReportSheet_(name) {
  const ss = getSpreadsheet_();
  const old = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);
  return ss.insertSheet(name);
}

/** Tulis tabel (baris pertama = header) ke sheet laporan. */
function writeReportTable_(sheet, rows, headerBackground) {
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(1, 1, 1, rows[0].length).setFontWeight('bold').setBackground(headerBackground || '#fef2f2');
  sheet.autoResizeColumns(1, rows[0].length);
}

/** Ukuran area data Realisasi aktif, atau null bila kosong. */
function getActiveGridSize_() {
  const sheet = getSheet_(CONFIG.SHEET_REALISASI);
  const numRows = getLastEmployeeRow_(sheet) - CONFIG.DATA_START_ROW + 1;
  const dayCount = getDayColCount_(sheet);
  return (numRows <= 0 || dayCount <= 0) ? null : { realisasiSheet: sheet, numRows: numRows, dayCount: dayCount };
}

function gridRangeOf_(sheet, size) {
  return sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.FIRST_DAY_COL, size.numRows, size.dayCount);
}

function isFilledCell_(v) {
  return v !== '' && v !== null && v !== undefined;
}

/* ============================ DIAGNOSIS ============================ */

function cekSpreadsheetAktif() {
  const ss = getSpreadsheet_();
  Logger.log('Nama Spreadsheet: ' + ss.getName());
  Logger.log('ID Spreadsheet : ' + ss.getId());
  Logger.log('URL            : ' + ss.getUrl());
}

/** Catatan karakter mencurigakan pada nama (penyebab nama tampak sama tetapi tidak cocok). */
function detectSuspiciousChars_(name) {
  const s = String(name || '');
  const notes = [];
  if (/\u00A0/.test(s)) notes.push('mengandung spasi non-breaking (biasa dari copy-paste Excel/Word)');
  if (/[\u200B\u200C\u200D\uFEFF]/.test(s)) notes.push('mengandung karakter tak terlihat (zero-width)');
  if (/\t/.test(s)) notes.push('mengandung karakter tab');
  if (/\s{2,}/.test(s)) notes.push('ada spasi ganda/lebih');
  if (/^\s|\s$/.test(s)) notes.push('ada spasi di awal/akhir nama');
  return notes.join('; ');
}

/** Kolom Nama (apa adanya, tanpa dirapikan) dari sheet bergaya Jadwal/Realisasi. */
function getRawNameListFromSheetObj_(sheet) {
  const numRows = getLastEmployeeRow_(sheet) - CONFIG.DATA_START_ROW + 1;
  if (numRows <= 0) return [];
  return sheet.getRange(CONFIG.DATA_START_ROW, 2, numRows, 1).getValues().map(r => r[0]);
}

function getRawNameList_(sheetName) {
  return getRawNameListFromSheetObj_(getSheet_(sheetName));
}

function getRawNameListFromNikSheet_() {
  return readEmployeeRows_(getOrCreateNikSheet_(), 1).map(r => r[0]);
}

/**
 * Bandingkan nama di sheet Rekap dengan Jadwal, Realisasi, JamKerja, DataKaryawan.
 * Hasil di sheet "_DebugNamaMismatch" (hanya baris bermasalah).
 */
function debugNameMismatch() {
  const jamKerjaSheet = getSpreadsheet_().getSheetByName(CONFIG.SHEET_JAMKERJA);

  const rekapNames = getRawNameList_(CONFIG.SHEET_REKAP);
  const jamKerjaNames = jamKerjaSheet ? getRawNameListFromSheetObj_(jamKerjaSheet) : [];
  const nameSet = names => new Set(names.map(normalizeName_));
  const jadwalSet = nameSet(getRawNameList_(CONFIG.SHEET_JADWAL));
  const realisasiSet = nameSet(getRawNameList_(CONFIG.SHEET_REALISASI));
  const jamKerjaSet = nameSet(jamKerjaNames);
  const nikSet = nameSet(getRawNameListFromNikSheet_());
  const noJamKerjaSheet = jamKerjaNames.length === 0;

  const rows = [['Nama di Rekap', 'Ada di Jadwal?', 'Ada di Realisasi?', 'Ada di JamKerja?', 'Ada di DataKaryawan (NIK)?', 'Catatan (karakter mencurigakan)']];

  rekapNames.forEach(rawName => {
    if (!rawName) return;
    const norm = normalizeName_(rawName);
    const inJadwal = jadwalSet.has(norm);
    const inRealisasi = realisasiSet.has(norm);
    const inJamKerja = jamKerjaSet.has(norm);

    if (inJadwal && inRealisasi && (inJamKerja || noJamKerjaSheet)) return;

    rows.push([
      rawName,
      inJadwal ? 'Ya' : '❌ TIDAK KETEMU',
      inRealisasi ? 'Ya' : '❌ TIDAK KETEMU',
      noJamKerjaSheet ? '(sheet JamKerja belum ada)' : (inJamKerja ? 'Ya' : '❌ TIDAK KETEMU'),
      nikSet.has(norm) ? 'Ya' : '(belum diisi NIK)',
      detectSuspiciousChars_(rawName) || '-'
    ]);
  });

  const debugSheet = recreateReportSheet_('_DebugNamaMismatch');
  if (rows.length === 1) {
    debugSheet.getRange(1, 1).setValue('✅ Semua nama di Rekap cocok persis dengan Jadwal, Realisasi' +
      (noJamKerjaSheet ? ' (sheet JamKerja belum ada, jadi belum dicek).' : ', dan JamKerja.'));
  } else {
    writeReportTable_(debugSheet, rows);
    for (let r = 1; r < rows.length; r++) {
      debugSheet.getRange(r + 1, 1, 1, rows[0].length).setBackground('#fff7ed');
    }
  }

  const totalMasalah = rows.length - 1;
  showAlertQuietly_(totalMasalah > 0
    ? 'Ditemukan ' + totalMasalah + ' nama di Rekap yang TIDAK cocok persis dengan sheet lain.\n\nLihat detailnya di sheet "_DebugNamaMismatch".'
    : 'Semua nama cocok! Tidak ada masalah pencocokan nama ditemukan.');
}

/**
 * Cari baris karyawan yang terisi tetapi TIDAK ikut terhitung sistem. Sistem berhenti membaca
 * daftar karyawan setelah MAX_CONSECUTIVE_BLANK baris kosong berturut-turut di kolom Nama
 * (lihat getLastEmployeeRow_), jadi baris di luar batas itu "hilang". Hasil di sheet "_DebugBlankGaps".
 * (Versi lama mengira sistem berhenti di baris kosong PERTAMA sehingga laporannya berlebihan.)
 */
function findBlankRowGaps() {
  const ss = getSpreadsheet_();
  const rows = [['Sheet', 'Baris Kosong Pertama', 'Baris Terakhir yg Dihitung Sistem', 'Jumlah Baris Terisi yg TIDAK Dihitung Sistem', 'Contoh Nama yg Tidak Dihitung']];

  [CONFIG.SHEET_JADWAL, CONFIG.SHEET_REALISASI].forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      rows.push([sheetName, 'Sheet tidak ditemukan', '-', '-', '-']);
      return;
    }

    const names = sheet.getRange(CONFIG.DATA_START_ROW, 2, EMPLOYEE_SCAN_ROWS_, 1).getValues().map(r => r[0]);
    const firstBlankIdx = names.findIndex(v => !isFilledCell_(v));
    const systemLastRow = getLastEmployeeRow_(sheet);

    const missed = [];
    names.forEach((v, i) => {
      if (isFilledCell_(v) && CONFIG.DATA_START_ROW + i > systemLastRow) missed.push(v);
    });

    rows.push([
      sheetName,
      firstBlankIdx === -1 ? '(tidak ada baris kosong)' : 'Baris ' + (CONFIG.DATA_START_ROW + firstBlankIdx),
      systemLastRow,
      missed.length,
      missed.slice(0, 3).join(', ') || '-'
    ]);
  });

  writeReportTable_(recreateReportSheet_('_DebugBlankGaps'), rows);

  showAlertQuietly_('Selesai. Lihat detail di sheet "_DebugBlankGaps".\n\n' +
    'Kalau kolom "Jumlah Baris Terisi yg TIDAK Dihitung Sistem" > 0, ada baris karyawan terisi yang tidak ikut ' +
    'terhitung karena didahului ' + MAX_CONSECUTIVE_BLANK_ + ' baris kosong atau lebih di kolom Nama.');
}

/**
 * AUDIT (aman): sel JamScan/JamKerja yang terisi jam padahal Realisasi pada tanggal yang sama = OFF
 * (indikasi data "nyangkut" dari upload lama). Hasil di sheet "_AuditJamNyangkutOFF". Jalankan ini
 * DULU sebelum bersihkanJamNyangkutOFF().
 */
function auditJamNyangkutOFF() {
  const ss = getSpreadsheet_();
  const size = getActiveGridSize_();
  if (!size) {
    Logger.log('Tidak ada data untuk diaudit.');
    return;
  }

  const sheet = size.realisasiSheet;
  const info = sheet.getRange(CONFIG.DATA_START_ROW, 1, size.numRows, 2).getValues(); // No, Nama
  const realGrid = gridRangeOf_(sheet, size).getValues();
  const readAux = name => { const s = ss.getSheetByName(name); return s ? gridRangeOf_(s, size).getValues() : null; };
  const jamKerjaGrid = readAux(CONFIG.SHEET_JAMKERJA);
  const jamScanGrid = readAux(CONFIG.SHEET_JAMSCAN);

  const rows = [['Nama', 'Tanggal (ke-)', 'Kode Realisasi', 'Isi JamKerja', 'Isi JamScan']];
  for (let i = 0; i < size.numRows; i++) {
    const nama = info[i][1];
    if (!nama) continue;
    for (let d = 0; d < size.dayCount; d++) {
      const realCode = cellCode_(realGrid[i][d]);
      const jkVal = jamKerjaGrid ? jamKerjaGrid[i][d] : '';
      const jsVal = jamScanGrid ? jamScanGrid[i][d] : '';
      if (realCode === 'OFF' && (isFilledCell_(jkVal) || isFilledCell_(jsVal))) {
        rows.push([nama, d + 1, realCode || '(kosong)', jkVal, jsVal]);
      }
    }
  }

  const debugSheet = recreateReportSheet_('_AuditJamNyangkutOFF');
  if (rows.length === 1) {
    debugSheet.getRange(1, 1).setValue('✅ Tidak ditemukan sel JamKerja/JamScan yang terisi padahal Realisasi = OFF.');
  } else {
    writeReportTable_(debugSheet, rows);
  }

  showAlertQuietly_((rows.length - 1) + ' sel ditemukan nyangkut (JamKerja/JamScan terisi tapi Realisasi = OFF).\n\n' +
    'Lihat detail di sheet "_AuditJamNyangkutOFF".\n\n' +
    'Kalau hasilnya sudah sesuai dugaan, jalankan bersihkanJamNyangkutOFF() untuk mengosongkannya.');
}

/** Uji getJadwalHariIni untuk satu nama (default: karyawan pertama di Jadwal). */
function testGetJadwalHariIni(nama) {
  if (!nama) {
    const first = readGrid_(CONFIG.SHEET_JADWAL).employees[0];
    nama = first ? first.nama : '';
  }
  const ss = getSpreadsheet_();
  Logger.log('Nama yang diuji: ' + nama);
  Logger.log('HASIL: ' + JSON.stringify(getJadwalHariIni(nama)));
  Logger.log('Tanggal server sekarang (getDate()): ' + new Date().getDate());
  Logger.log('Waktu server penuh: ' + new Date());
  Logger.log('Spreadsheet yang dipakai: ' + ss.getName() + ' | ID: ' + ss.getId());
}

/** Tampilkan isi sheet SettingJamKerja di log. */
function debugShiftSheet() {
  const ss = getSpreadsheet_();
  Logger.log('Semua nama sheet di spreadsheet ini: ' + JSON.stringify(ss.getSheets().map(s => s.getName())));

  const shiftSheet = ss.getSheetByName(SHEET_SETTING_JAMKERJA);
  Logger.log('Sheet ' + SHEET_SETTING_JAMKERJA + ' ditemukan: ' + (shiftSheet ? 'YA' : 'TIDAK'));
  if (!shiftSheet) return;

  const rawData = shiftSheet.getDataRange().getValues();
  Logger.log('Jumlah baris terbaca: ' + rawData.length);
  rawData.forEach((row, idx) => Logger.log('Baris ' + idx + ': ' + JSON.stringify(row)));
}

/* ============================ PERBAIKAN (MENGUBAH DATA) ============================ */

/**
 * PEMBERSIH: kosongkan JamKerja/JamScan yang terisi padahal Realisasi = OFF di tanggal yang
 * sama. Jalankan HANYA setelah meninjau auditJamNyangkutOFF(). Sheet arsip tidak disentuh.
 */
function bersihkanJamNyangkutOFF() {
  const ss = getSpreadsheet_();
  const size = getActiveGridSize_();
  if (!size) {
    Logger.log('Tidak ada data untuk dibersihkan.');
    return;
  }

  const realGrid = gridRangeOf_(size.realisasiSheet, size).getValues();

  withScriptLock_(() => {
    let dibersihkan = 0;

    [CONFIG.SHEET_JAMKERJA, CONFIG.SHEET_JAMSCAN].forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;

      const range = gridRangeOf_(sheet, size);
      const grid = range.getValues();
      for (let i = 0; i < size.numRows; i++) {
        for (let d = 0; d < size.dayCount; d++) {
          if (cellCode_(realGrid[i][d]) === 'OFF' && grid[i][d] !== '' && grid[i][d] !== null) {
            grid[i][d] = '';
            dibersihkan++;
          }
        }
      }
      range.setValues(grid);
    });

    SpreadsheetApp.flush();
    showAlertQuietly_('Selesai. ' + dibersihkan + ' sel JamKerja/JamScan yang nyangkut (Realisasi=OFF tapi ada jam) sudah dikosongkan.\n\n' +
      'Silakan refresh web app AMS untuk lihat hasilnya.');
  }, 20000);
}

/** Kosongkan SEMUA kolom hari JamKerja & JamScan periode aktif (arsip tidak disentuh). */
function fixJamKerjaPeriodeAktifSekarang() {
  const ss = getSpreadsheet_();
  const size = getActiveGridSize_();
  if (!size) {
    Logger.log('Tidak ada data karyawan/hari untuk dibersihkan.');
    return;
  }

  const dibersihkan = [];
  [CONFIG.SHEET_JAMKERJA, CONFIG.SHEET_JAMSCAN].forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(sheetName + ' belum pernah dibuat sama sekali -- dilewati.');
      return;
    }
    gridRangeOf_(sheet, size).clearContent();
    dibersihkan.push(sheetName);
  });
  SpreadsheetApp.flush();

  showAlertQuietly_(dibersihkan.length
    ? 'Selesai. Kolom hari di sheet ' + dibersihkan.join(' & ') + ' (periode AKTIF saja) sudah dikosongkan.\n\n' +
      'Sheet arsip (Arsip_JamKerja_..., dst) TIDAK disentuh -- data periode sebelumnya tetap aman.\n\n' +
      'Silakan refresh Dashboard / Rekap Karyawan di web app untuk lihat hasilnya.'
    : 'Tidak ada sheet JamKerja/JamScan yang perlu dibersihkan.');
}

/**
 * Isi Rekap_Bulanan dari semua periode yang sudah diarsipkan (untuk periode yang diarsipkan sebelum
 * fitur tren bulanan dipasang, atau untuk memperbaiki snapshot). Aman dijalankan berulang
 * (satu baris per tahun+bulan, yang sudah ada ditimpa). Menggantikan backfillRekapBulananDariArsip_
 * dan fixSnapshotAgustus2026_ (fungsi berakhiran "_" tidak bisa dipilih dari editor; yang kedua
 * khusus Agustus 2026 dan kini tercakup di sini).
 */
function backfillRekapBulananDariArsip() {
  const ss = getSpreadsheet_();
  if (!ss.getSheetByName(ARCHIVE_INDEX_SHEET)) {
    Logger.log('Sheet ArsipIndex belum ada, tidak ada yang bisa di-backfill.');
    return;
  }

  let sukses = 0;
  const gagal = [];

  readArchiveRows_().forEach(row => {
    const periodText = row[0], jadwalName = row[2], realisasiName = row[3];
    if (!periodText || !jadwalName || !realisasiName) return;

    const my = parsePeriodMonthYear_(periodText);
    if (!my) {
      gagal.push(periodText + ' (bulan/tahun tidak terbaca)');
      return;
    }

    const jadwalSheet = ss.getSheetByName(jadwalName);
    const realisasiSheet = ss.getSheetByName(realisasiName);
    if (!jadwalSheet || !realisasiSheet) {
      gagal.push(periodText + ' (sheet arsip tidak ditemukan)');
      return;
    }

    try {
      generateRekapBulananSnapshot_(my.year, my.monthIndex + 1, jadwalSheet, realisasiSheet);
      sukses++;
    } catch (e) {
      gagal.push(periodText + ' (error: ' + e.message + ')');
    }
  });

  showAlertQuietly_('Backfill selesai. ' + sukses + ' periode berhasil disnapshot.' +
    (gagal.length ? '\n\nGagal:\n' + gagal.join('\n') : ''));
}
