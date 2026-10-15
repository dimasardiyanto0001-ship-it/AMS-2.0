/*************************************************************
 * 12_Period.gs
 * Nama periode (bulan/tahun dari teks "Periode: OKTOBER 2026"),
 * indeks hari ini, dan alur "Mulai Periode Baru".
 *
 * Alur mulaiPeriodeBaru (Super Admin):
 *   1. arsipkan periode berjalan (13_Archive.gs)
 *   2. snapshot rekap bulanan untuk grafik tren (14_MonthlyTrend.gs)
 *   3. kosongkan Jadwal, Realisasi, JamKerja, JamScan untuk bulan berikutnya
 *************************************************************/

const INDONESIAN_MONTH_NAMES_UPPER_ = [
  'JANUARI', 'FEBRUARI', 'MARET', 'APRIL', 'MEI', 'JUNI',
  'JULI', 'AGUSTUS', 'SEPTEMBER', 'OKTOBER', 'NOVEMBER', 'DESEMBER'
];

/** Nama bulan huruf kecil -> indeks 0..11. */
const INDONESIAN_MONTHS_ = INDONESIAN_MONTH_NAMES_UPPER_.reduce((map, name, i) => {
  map[name.toLowerCase()] = i;
  return map;
}, {});

const INDONESIAN_DAY_ABBR_ = ['Min', 'Sen', 'Sel', 'Rab', 'Kam', 'Jum', 'Sab'];

const MAX_DAY_COLS_ = 31;

/* ============================ PARSING NAMA PERIODE ============================ */

/** @return {{year:number, monthIndex:number}|null} null jika tahun/bulan tidak terbaca. */
function parsePeriodMonthYear_(periodText) {
  const text = String(periodText || '').toLowerCase();
  const yearMatch = text.match(/(20\d{2})/);
  if (!yearMatch) return null;

  let monthIndex = -1;
  Object.keys(INDONESIAN_MONTHS_).forEach(name => {
    if (text.indexOf(name) !== -1) monthIndex = INDONESIAN_MONTHS_[name];
  });
  if (monthIndex === -1) return null;

  return { year: parseInt(yearMatch[1], 10), monthIndex: monthIndex };
}

/**
 * Sampai hari ke berapa (1-based) periode ini sudah "terjadi". Periode yang
 * bukan bulan berjalan (arsip) dianggap sudah lewat semua (dayCount penuh).
 */
function computeTodayDayIndex_(periodText, dayCount) {
  const my = parsePeriodMonthYear_(periodText);
  if (!my) return dayCount;
  const now = new Date();
  const isCurrentMonth = now.getFullYear() === my.year && now.getMonth() === my.monthIndex;
  return isCurrentMonth ? Math.min(now.getDate(), dayCount) : dayCount;
}

function formatCalendarDate_(periodText, dayNumber) {
  const my = parsePeriodMonthYear_(periodText);
  if (!my) return null;
  return Utilities.formatDate(new Date(my.year, my.monthIndex, dayNumber), Session.getScriptTimeZone() || 'GMT+7', 'dd/MM/yyyy');
}

/** Ganti tahun dan bulan pada label lama; jika tak ada nama bulan, tambahkan di belakang. */
function buildNextPeriodLabel_(oldText, newMonthNameUpper, newYear) {
  let result = oldText.replace(/(20\d{2})/, String(newYear));
  const monthPattern = new RegExp(Object.keys(INDONESIAN_MONTHS_).join('|'), 'i');
  return monthPattern.test(result)
    ? result.replace(monthPattern, newMonthNameUpper)
    : result + ' ' + newMonthNameUpper;
}

/* ============================ RESET SHEET ============================ */

/** Tulis label periode, header tanggal/hari bulan baru, dan kosongkan seluruh grid. */
function resetSheetForNewPeriod_(sheet, periodLabel, year, monthIndex, dayCount) {
  sheet.getRange(2, 1).setValue(periodLabel);

  sheet.getRange(CONFIG.HEADER_ROW, CONFIG.FIRST_DAY_COL, 1, MAX_DAY_COLS_).clearContent();
  sheet.getRange(CONFIG.DAY_NAME_ROW, CONFIG.FIRST_DAY_COL, 1, MAX_DAY_COLS_).clearContent();

  const headerValues = [];
  const dayNameValues = [];
  for (let d = 1; d <= dayCount; d++) {
    headerValues.push(d);
    dayNameValues.push(INDONESIAN_DAY_ABBR_[new Date(year, monthIndex, d).getDay()]);
  }
  sheet.getRange(CONFIG.HEADER_ROW, CONFIG.FIRST_DAY_COL, 1, dayCount).setValues([headerValues]);
  sheet.getRange(CONFIG.DAY_NAME_ROW, CONFIG.FIRST_DAY_COL, 1, dayCount).setValues([dayNameValues]);

  const numRows = getLastEmployeeRow_(sheet) - CONFIG.DATA_START_ROW + 1;
  if (numRows > 0) {
    const blankGrid = Array.from({ length: numRows }, () => new Array(MAX_DAY_COLS_).fill(''));
    sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.FIRST_DAY_COL, numRows, MAX_DAY_COLS_).setValues(blankGrid);
  }
}

function resetJamKerjaJamScanForNewPeriod_(periodLabel, year, monthIndex, dayCount) {
  const ss = getSpreadsheet_();
  [CONFIG.SHEET_JAMKERJA, CONFIG.SHEET_JAMSCAN].forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) resetSheetForNewPeriod_(sheet, periodLabel, year, monthIndex, dayCount);
  });
}

/* ============================ MULAI PERIODE BARU ============================ */

function mulaiPeriodeBaru(token) {
  const session = requireSuperAdmin_(token, 'Hanya akun dengan role "Super Admin" yang boleh memulai periode baru.');

  const jadwalSheet = getSheet_(CONFIG.SHEET_JADWAL);
  const currentPeriodText = String(jadwalSheet.getRange(2, 1).getValue() || '').trim();
  if (!currentPeriodText) {
    throw new Error('Sel periode (baris 2, kolom A) di sheet Jadwal masih kosong. Isi dulu sebelum memulai periode baru.');
  }

  const my = parsePeriodMonthYear_(currentPeriodText);
  if (!my) {
    throw new Error('Nama periode saat ini ("' + currentPeriodText + '") tidak bisa dibaca bulan/tahunnya. ' +
      'Pastikan formatnya mengandung nama bulan (misal "Juli") dan tahun 4 digit, lalu coba lagi.');
  }

  archiveCurrentPeriod(token);

  // Snapshot untuk grafik tren. Kegagalan di sini tidak boleh menghalangi periode baru.
  try {
    generateRekapBulananSnapshot_(my.year, my.monthIndex + 1, jadwalSheet, getSheet_(CONFIG.SHEET_REALISASI));
  } catch (e) {
    Logger.log('Gagal snapshot Rekap_Bulanan: ' + e.message);
  }

  let nextMonthIndex = my.monthIndex + 1;
  let nextYear = my.year;
  if (nextMonthIndex > 11) { nextMonthIndex = 0; nextYear++; }
  const nextPeriodLabel = buildNextPeriodLabel_(currentPeriodText, INDONESIAN_MONTH_NAMES_UPPER_[nextMonthIndex], nextYear);
  const nextDayCount = new Date(nextYear, nextMonthIndex + 1, 0).getDate();

  withScriptLock_(() => {
    resetSheetForNewPeriod_(getSheet_(CONFIG.SHEET_JADWAL), nextPeriodLabel, nextYear, nextMonthIndex, nextDayCount);
    resetSheetForNewPeriod_(getSheet_(CONFIG.SHEET_REALISASI), nextPeriodLabel, nextYear, nextMonthIndex, nextDayCount);
    resetJamKerjaJamScanForNewPeriod_(nextPeriodLabel, nextYear, nextMonthIndex, nextDayCount);
    SpreadsheetApp.flush();
  }, 20000);

  const result = getAllDataInternal_();
  result.mulaiPeriodeBaruInfo = {
    periodeLama: currentPeriodText,
    periodeBaru: nextPeriodLabel,
    jumlahHari: nextDayCount
  };

  catatAuditLog_(session, 'Mulai Periode Baru', 'Periode', nextPeriodLabel, 'Dari periode ' + currentPeriodText);
  return result;
}
