/*************************************************************
 * 13_Archive.gs
 * Arsip periode: snapshot sheet (Jadwal, Realisasi, Dashboard, Rekap,
 * JamKerja, JamScan) menjadi sheet tersembunyi "Arsip_<sheet>_<periode>",
 * dicatat di sheet tersembunyi "ArsipIndex".
 *
 * Kolom ArsipIndex (1-based): 1 Periode, 2 PeriodeKey, 3 Jadwal,
 * 4 Realisasi, 5 Dashboard, 6 Rekap, 7 JamKerja, 8 JamScan, 9 DiarsipkanPada.
 *************************************************************/

const ARCHIVE_INDEX_SHEET = 'ArsipIndex';
const ARCHIVE_INDEX_HEADERS = ['Periode', 'PeriodeKey', 'Jadwal', 'Realisasi', 'Dashboard', 'Rekap', 'JamKerja', 'JamScan', 'DiarsipkanPada'];
const ARCHIVE_SOURCE_KEYS_ = ['SHEET_JADWAL', 'SHEET_REALISASI', 'SHEET_DASHBOARD', 'SHEET_REKAP', 'SHEET_JAMKERJA', 'SHEET_JAMSCAN'];
const ARCHIVE_SHEET_NAME_MAX_ = 95;

/* ============================ HELPER ============================ */

/** Teks periode -> kunci aman untuk nama sheet. */
function sanitizePeriodKey_(text) {
  const cleaned = String(text || '').trim()
    .replace(/[\[\]\*\?\/\\:]/g, '')
    .replace(/\s+/g, '_');
  return cleaned || ('Periode_' + Date.now());
}

/** Ubah semua formula jadi nilai tetap (snapshot tidak ikut berubah). */
function freezeSheetFormulas_(sheet) {
  const range = sheet.getDataRange();
  range.setValues(range.getValues());
}

/** @return {{rowNumber:number, row:Array}|null} baris ArsipIndex untuk kunci periode. */
function findArchiveRow_(idxSheet, periodKey) {
  const data = idxSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === periodKey) return { rowNumber: i + 1, row: data[i] };
  }
  return null;
}

function getArchiveIndexSheet_() {
  const idx = getSpreadsheet_().getSheetByName(ARCHIVE_INDEX_SHEET);
  if (!idx) throw new Error('Belum ada data arsip.');
  return idx;
}

/** Sisipkan kolom yang belum ada di ArsipIndex versi lama (kolom 7 JamKerja, kolom 8 JamScan). */
function ensureArsipIndexColumn_(idx, colPos, headerName) {
  if (idx.getRange(1, colPos).getValue() === 'DiarsipkanPada') {
    idx.insertColumnBefore(colPos);
    idx.getRange(1, colPos).setValue(headerName).setFontWeight('bold');
  }
}

function recordArchiveIndex_(periodText, periodKey, archivedNames) {
  const ss = getSpreadsheet_();
  let idx = ss.getSheetByName(ARCHIVE_INDEX_SHEET);
  if (!idx) {
    idx = ss.insertSheet(ARCHIVE_INDEX_SHEET);
    idx.getRange(1, 1, 1, ARCHIVE_INDEX_HEADERS.length).setValues([ARCHIVE_INDEX_HEADERS]).setFontWeight('bold');
    idx.hideSheet();
  }
  ensureArsipIndexColumn_(idx, 7, 'JamKerja');
  ensureArsipIndexColumn_(idx, 8, 'JamScan');

  const rowValues = [
    periodText, periodKey,
    archivedNames.SHEET_JADWAL, archivedNames.SHEET_REALISASI,
    archivedNames.SHEET_DASHBOARD, archivedNames.SHEET_REKAP,
    archivedNames.SHEET_JAMKERJA || '',
    archivedNames.SHEET_JAMSCAN || '',
    new Date()
  ];

  const found = findArchiveRow_(idx, periodKey);
  if (found) idx.getRange(found.rowNumber, 1, 1, rowValues.length).setValues([rowValues]);
  else idx.appendRow(rowValues);
}

/* ============================ ARSIPKAN ============================ */

function archiveCurrentPeriod(token) {
  const session = requireSuperAdmin_(token, 'Hanya akun dengan role "Super Admin" yang boleh mengarsipkan periode.');

  const ss = getSpreadsheet_();
  const periodText = String(getSheet_(CONFIG.SHEET_JADWAL).getRange(2, 1).getValue() || '').trim();
  if (!periodText) {
    throw new Error('Sel periode (baris 2, kolom A) di sheet Jadwal masih kosong. Isi dulu nama periodenya (misal "Periode: JULI 2026") sebelum diarsipkan.');
  }
  const periodKey = sanitizePeriodKey_(periodText);

  withScriptLock_(() => {
    const archivedNames = {};

    ARCHIVE_SOURCE_KEYS_.forEach(key => {
      const srcName = CONFIG[key];
      const src = ss.getSheetByName(srcName);
      if (!src) return; // sheet opsional (JamKerja/JamScan) belum pernah dibuat

      const archiveName = ('Arsip_' + srcName + '_' + periodKey).slice(0, ARCHIVE_SHEET_NAME_MAX_);
      const existing = ss.getSheetByName(archiveName);
      if (existing) ss.deleteSheet(existing);

      const copy = src.copyTo(ss);
      copy.setName(archiveName);
      freezeSheetFormulas_(copy);
      copy.hideSheet();
      archivedNames[key] = archiveName;
    });

    recordArchiveIndex_(periodText, periodKey, archivedNames);
    SpreadsheetApp.flush();
  }, 20000);

  catatAuditLog_(session, 'Arsipkan Periode', 'Periode', periodText, 'Diarsipkan dengan key ' + periodKey);
  return { success: true, period: periodText, periodKey: periodKey };
}

/* ============================ BACA ARSIP ============================ */

/** Daftar arsip, terbaru di atas. */
function getArchivedPeriodsInternal_() {
  const idx = getSpreadsheet_().getSheetByName(ARCHIVE_INDEX_SHEET);
  if (!idx) return [];

  return idx.getDataRange().getValues().slice(1)
    .filter(row => row[0])
    .map(row => ({ period: row[0], key: row[1] }))
    .reverse();
}

function getArchivedPeriods(token) {
  verifySession_(token);
  return getArchivedPeriodsInternal_();
}

function getArchivedDataInternal_(periodKey) {
  const ss = getSpreadsheet_();
  const found = findArchiveRow_(getArchiveIndexSheet_(), periodKey);
  if (!found) throw new Error('Arsip untuk periode ini tidak ditemukan.');
  const row = found.row;

  const jadwalSheet = ss.getSheetByName(row[2]);
  const realisasiSheet = ss.getSheetByName(row[3]);
  const dashboardSheet = ss.getSheetByName(row[4]);
  const rekapSheet = ss.getSheetByName(row[5]);
  if (!jadwalSheet || !realisasiSheet || !dashboardSheet || !rekapSheet) {
    throw new Error('Sheet arsip untuk periode ini tidak lengkap (mungkin terhapus manual).');
  }
  const jamKerjaArchiveName = String(row[6] || '').trim();
  const jamKerjaSheetArsip = jamKerjaArchiveName ? ss.getSheetByName(jamKerjaArchiveName) : null;

  const jadwal = readGridFromSheetObj_(jadwalSheet);
  const realisasi = readGridFromSheetObj_(realisasiSheet);
  const rekapDenganJamKerja = attachJamKerjaToRekap_(computeRekapDariGrid_(jadwal, realisasi), jadwal, jamKerjaSheetArsip);

  return {
    period: row[0],
    periodKey: row[1],
    isArchived: true,
    dayNames: jadwal.dayNames,
    dayCount: jadwal.dayCount,
    jadwal: jadwal.employees,
    realisasi: realisasi.employees,
    dashboard: getDashboardSummaryFromSheet_(rekapDenganJamKerja),
    rekap: attachNikToRekap_(rekapDenganJamKerja),
    trenHarian: computeTrenHarian_(jadwal, realisasi, jamKerjaSheetArsip),
    statusCodes: CONFIG.STATUS_CODES,
    shiftCodes: getShiftCodesDynamic_(),
    lastUpdate: new Date().toISOString()
  };
}

function getArchivedData(periodKey, token) {
  verifySession_(token);
  return getArchivedDataInternal_(periodKey);
}

/** Grid Jadwal dari arsip (dipakai Log Absensi dan getPeriodMeta). */
function readArchivedJadwalGrid_(periodKey) {
  const found = findArchiveRow_(getArchiveIndexSheet_(), periodKey);
  if (!found) throw new Error('Periode arsip tidak ditemukan.');

  const jadwalSheet = getSpreadsheet_().getSheetByName(found.row[2]);
  if (!jadwalSheet) throw new Error('Sheet arsip Jadwal untuk periode ini tidak ditemukan.');
  return readGridFromSheetObj_(jadwalSheet);
}

function getPeriodMetaInternal_(periodKey) {
  const g = (periodKey === 'active' || !periodKey)
    ? readGrid_(CONFIG.SHEET_JADWAL)
    : readArchivedJadwalGrid_(periodKey);

  return {
    period: g.period,
    dayCount: g.dayCount,
    dayNames: g.dayNames,
    employees: g.employees.map(e => ({ nama: e.nama, jabatan: e.jabatan }))
  };
}

function getPeriodMeta(periodKey, token) {
  verifySession_(token);
  return getPeriodMetaInternal_(periodKey);
}

/* ============================ HAPUS ARSIP ============================ */

/** Hanya Super Admin. Periode aktif tidak pernah bisa dihapus lewat fungsi ini. */
function deleteArchivedPeriod(periodKey, token) {
  const session = requireSuperAdmin_(token, 'Hanya akun dengan role "Super Admin" yang boleh menghapus arsip.');
  if (!periodKey || periodKey === 'active') throw new Error('Periode tidak valid untuk dihapus.');

  const ss = getSpreadsheet_();
  const idx = getArchiveIndexSheet_();

  const periodLabel = withScriptLock_(() => {
    const found = findArchiveRow_(idx, periodKey);
    if (!found) throw new Error('Arsip untuk periode ini tidak ditemukan (mungkin sudah terhapus).');

    [2, 3, 4, 5, 6, 7].forEach(col => {
      const name = found.row[col];
      if (!name) return;
      const sheet = ss.getSheetByName(name);
      if (sheet) ss.deleteSheet(sheet);
    });

    idx.deleteRow(found.rowNumber);
    SpreadsheetApp.flush();
    return found.row[0];
  }, 20000);

  catatAuditLog_(session, 'Hapus Arsip Periode', 'Periode', periodLabel, 'Key: ' + periodKey);
  return { success: true, period: periodLabel };
}
