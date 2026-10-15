/*************************************************************
 * 02_SheetUtils.gs
 * Helper dasar: akses sheet, lock, dan pembacaan grid
 * Jadwal/Realisasi/JamKerja/JamScan.
 *************************************************************/

function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_(name) {
  const sheet = getSpreadsheet_().getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" tidak ditemukan.');
  return sheet;
}

/** Jalankan fn() di dalam script lock; lock selalu dilepas. */
function withScriptLock_(fn, waitMs) {
  const lock = LockService.getScriptLock();
  lock.waitLock(waitMs || 15000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function toHex_(byteArray) {
  return byteArray.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

/** Batas pemindaian daftar karyawan: jumlah baris yang dibaca dan toleransi baris kosong berturut-turut. */
const EMPLOYEE_SCAN_ROWS_ = 500;
const MAX_CONSECUTIVE_BLANK_ = 5;

/** Baris terakhir yang berisi nama karyawan (toleran MAX_CONSECUTIVE_BLANK_ baris kosong berturut-turut). */
function getLastEmployeeRow_(sheet) {
  const values = sheet.getRange(CONFIG.DATA_START_ROW, 1, EMPLOYEE_SCAN_ROWS_, 2).getValues();
  let last = CONFIG.DATA_START_ROW - 1;
  let blanks = 0;

  for (let i = 0; i < values.length; i++) {
    const nama = values[i][1];
    if (nama !== '' && nama !== null) {
      last = CONFIG.DATA_START_ROW + i;
      blanks = 0;
    } else if (++blanks >= MAX_CONSECUTIVE_BLANK_) {
      break;
    }
  }
  return last;
}

/** Jumlah kolom tanggal = deretan angka berurutan di baris header. */
function getDayColCount_(sheet) {
  const headers = sheet.getRange(CONFIG.HEADER_ROW, CONFIG.FIRST_DAY_COL, 1, 31).getValues()[0];
  let count = 0;
  while (count < headers.length && typeof headers[count] === 'number') count++;
  return count;
}

function readGrid_(sheetName) {
  return readGridFromSheetObj_(getSheet_(sheetName));
}

/**
 * @return {{employees: Array<{no, nama, jabatan, rowIndex, days: string[]}>,
 *           dayNames: string[], dayCount: number, period: *}}
 */
function readGridFromSheetObj_(sheet) {
  const lastRow = getLastEmployeeRow_(sheet);
  const dayCount = getDayColCount_(sheet);
  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const period = sheet.getRange(2, 1).getValue();
  const dayNames = dayCount > 0
    ? sheet.getRange(CONFIG.DAY_NAME_ROW, CONFIG.FIRST_DAY_COL, 1, dayCount).getValues()[0]
    : [];

  if (numRows <= 0 || dayCount <= 0) {
    return { employees: [], dayNames, dayCount, period };
  }

  const info = sheet.getRange(CONFIG.DATA_START_ROW, 1, numRows, 3).getValues();
  const grid = sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.FIRST_DAY_COL, numRows, dayCount).getValues();

  const employees = info.map((row, i) => ({
    no: row[0],
    nama: row[1],
    jabatan: row[2],
    rowIndex: CONFIG.DATA_START_ROW + i,
    days: grid[i].map(v => (v === null ? '' : String(v)))
  }));

  return { employees, dayNames, dayCount, period };
}

/** Nama untuk perbandingan: spasi beruntun jadi satu, tanpa spasi tepi, huruf kecil. */
function normalizeName_(name) {
  return String(name || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

/**
 * Peringatan jika baris karyawan Jadwal dan Realisasi tidak sejajar
 * (penyebab data Realisasi nyasar ke karyawan yang salah).
 * Satu-satunya definisi; salinan lama di Code.gs & 16_RepairAlignment.gs dihapus.
 */
function checkRowAlignment_(jadwalGrid, realisasiGrid) {
  const warnings = [];
  const a = jadwalGrid.employees;
  const b = realisasiGrid.employees;

  if (a.length !== b.length) {
    warnings.push(
      'Jumlah baris karyawan di Jadwal (' + a.length + ') dan Realisasi (' + b.length +
      ') TIDAK SAMA. Ini bisa membuat data Realisasi nyasar ke karyawan yang salah. ' +
      'Jangan insert/hapus baris manual di salah satu sheet saja -- gunakan fitur "Tambah Karyawan", ' +
      'atau jalankan debugNameMismatch() untuk detail lengkap.'
    );
    return warnings;
  }

  for (let i = 0; i < a.length; i++) {
    if (normalizeName_(a[i].nama) !== normalizeName_(b[i].nama)) {
      warnings.push(
        'Baris ke-' + (i + 1) + ': nama di Jadwal ("' + a[i].nama + '") tidak sama dengan ' +
        'nama di Realisasi ("' + b[i].nama + '") pada posisi baris yang sama. ' +
        'Jalankan debugNameMismatch() untuk detail.'
      );
    }
  }
  return warnings;
}
