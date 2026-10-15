/*************************************************************
 * 30_AttendanceLog.gs
 * Log absensi per karyawan per hari dalam bentuk file Excel
 * (menu Log Absensi, dan lampiran email dari chatbot).
 * Sumber: Jadwal + Realisasi + JamScan periode aktif atau arsip.
 *************************************************************/

const ATTENDANCE_LOG_HEADER_ = ['No', 'NIK', 'Nama', 'Jabatan', 'Tanggal', 'Hari', 'Shift Terjadwal', 'Kode Realisasi', 'Keterangan Realisasi', 'Jam Masuk', 'Jam Keluar'];
const ATTENDANCE_LOG_EMPTY_ROW_ = '(Tidak ada data untuk filter ini)';
const ATTENDANCE_LOG_FILE_MAX_ = 120;

/* ============================ HELPER ============================ */

/** Kode kehadiran seragam: huruf besar tanpa spasi; "1.0" dianggap "1". */
function normalizeAttendanceCode_(v) {
  if (v === null || v === undefined) return '';
  const s = String(v).trim().toUpperCase();
  return (s === '1' || s === '1.0') ? '1' : s;
}

/** Pecah "07:00 - 15:00" menjadi {masuk, keluar}; "-" dan kosong menjadi ''. */
function splitScanPair_(value) {
  const s = String(value || '').trim();
  if (!s) return { masuk: '', keluar: '' };
  const parts = s.split(' - ');
  const masuk = (parts[0] || '').trim();
  const keluar = (parts[1] || '').trim();
  return {
    masuk: (masuk && masuk !== '-') ? masuk : '',
    keluar: (keluar && keluar !== '-') ? keluar : ''
  };
}

/** Grid JamScan, atau null bila sheet tidak ada / formatnya tidak terbaca. */
function readJamScanGridSafe_(sheet) {
  if (!sheet) return null;
  try {
    return readGridFromSheetObj_(sheet);
  } catch (e) {
    return null;
  }
}

/** Peta nama(huruf kecil) -> objek karyawan dari satu grid. */
function indexEmployeesByName_(grid) {
  const byName = {};
  if (grid) grid.employees.forEach(e => { byName[nameKey_(e.nama)] = e; });
  return byName;
}

/** Grid Jadwal, Realisasi, JamScan untuk periode aktif ('active') atau kunci arsip. */
function readPeriodGrids_(periodKey) {
  if (periodKey === 'active' || !periodKey) {
    return {
      jadwal: readGrid_(CONFIG.SHEET_JADWAL),
      realisasi: readGrid_(CONFIG.SHEET_REALISASI),
      jamScan: readJamScanGridSafe_(getSpreadsheet_().getSheetByName(CONFIG.SHEET_JAMSCAN))
    };
  }

  const ss = getSpreadsheet_();
  const found = findArchiveRow_(getArchiveIndexSheet_(), periodKey);
  if (!found) throw new Error('Periode arsip tidak ditemukan.');
  const row = found.row;

  const jadwalSheet = ss.getSheetByName(row[2]);
  const realisasiSheet = ss.getSheetByName(row[3]);
  if (!jadwalSheet || !realisasiSheet) throw new Error('Sheet arsip untuk periode ini tidak lengkap.');

  const jamScanArchiveName = String(row[7] || '').trim();
  return {
    jadwal: readGridFromSheetObj_(jadwalSheet),
    realisasi: readGridFromSheetObj_(realisasiSheet),
    jamScan: readJamScanGridSafe_(jamScanArchiveName ? ss.getSheetByName(jamScanArchiveName) : null)
  };
}

/* ============================ BANGUN BARIS LOG ============================ */

/**
 * Satu baris per karyawan per hari (tanggal start..end, 1-based). Kolom "No" = urutan
 * karyawan (sama untuk semua baris karyawan itu).
 */
function buildAttendanceLogRows_(grids, employees, start, end, nikMap) {
  const realisasi = grids.realisasi;
  const jadwalByName = indexEmployeesByName_(grids.jadwal);
  const jamScanByName = indexEmployeesByName_(grids.jamScan);
  const periodText = realisasi.period || '';

  const rows = [];
  employees.forEach((emp, i) => {
    const key = nameKey_(emp.nama);
    const jadwalEmp = jadwalByName[key];
    const jamScanEmp = jamScanByName[key];
    const nik = nikMap[key] || '';

    for (let d = start; d <= end; d++) {
      const dayIdx = d - 1;
      const realCode = normalizeAttendanceCode_(emp.days[dayIdx]);
      const shiftCode = jadwalEmp ? normalizeAttendanceCode_(jadwalEmp.days[dayIdx]) : '';
      const realInfo = CONFIG.STATUS_CODES[realCode] || CONFIG.STATUS_CODES[''];
      const shiftInfo = CONFIG.SHIFT_CODES[shiftCode] || { label: 'Belum diisi' };
      const scan = splitScanPair_(jamScanEmp ? jamScanEmp.days[dayIdx] : '');

      rows.push([
        i + 1,
        nik,
        emp.nama,
        emp.jabatan,
        formatCalendarDate_(periodText, d) || d,
        realisasi.dayNames[dayIdx] || '',
        shiftInfo.label,
        realCode || '-',
        realInfo.label,
        scan.masuk,
        scan.keluar
      ]);
    }
  });
  return rows;
}

/* ============================ EXCEL ============================ */

/**
 * @param {string} periodKey 'active' atau kunci arsip
 * @param {string} employeeName nama karyawan, atau 'ALL'/kosong untuk semua
 * @param {number} dayStart tanggal awal (default 1)
 * @param {number} dayEnd tanggal akhir (default akhir periode)
 */
function generateAttendanceLogExcel(periodKey, employeeName, dayStart, dayEnd, token) {
  requireMenuAccess_(token, 'logabsensi');
  const grids = readPeriodGrids_(periodKey);
  const realisasi = grids.realisasi;

  const start = Math.max(1, parseInt(dayStart, 10) || 1);
  const end = Math.min(realisasi.dayCount, parseInt(dayEnd, 10) || realisasi.dayCount);
  if (start > end) throw new Error('Rentang tanggal tidak valid (tanggal awal melebihi tanggal akhir).');

  const oneEmployee = employeeName && employeeName !== 'ALL';
  let employees = realisasi.employees;
  if (oneEmployee) {
    employees = employees.filter(e => nameKey_(e.nama) === nameKey_(employeeName));
    if (!employees.length) throw new Error('Karyawan "' + employeeName + '" tidak ditemukan di periode ini.');
  }

  const dataRows = buildAttendanceLogRows_(grids, employees, start, end, getNikMap_());

  const empPart = oneEmployee ? sanitizePeriodKey_(employeeName) : 'SemuaKaryawan';
  const fileNameBase = ('LogAbsensi_' + empPart + '_' + sanitizePeriodKey_(realisasi.period || 'Periode')).slice(0, ATTENDANCE_LOG_FILE_MAX_);
  const blob = buildExcelBlobFromRows_(fileNameBase, ATTENDANCE_LOG_HEADER_, dataRows);

  return xlsxDownloadResult_(fileNameBase + '.xlsx', blob);
}

/** Buat Excel satu sheet "Log Absensi" lewat spreadsheet sementara. @return {Blob} */
function buildExcelBlobFromRows_(fileNameBase, headerRow, dataRows) {
  const tempSs = SpreadsheetApp.create(fileNameBase + '_TEMP_' + Date.now());
  const fileId = tempSs.getId();

  try {
    const sheet = tempSs.getSheets()[0];
    sheet.setName('Log Absensi');

    const width = headerRow.length;
    const body = dataRows.length ? dataRows : [[ATTENDANCE_LOG_EMPTY_ROW_]];
    const allRows = [headerRow].concat(body).map(r => {
      const padded = r.slice(0, width);
      while (padded.length < width) padded.push('');
      return padded;
    });

    sheet.getRange(1, 1, allRows.length, width).setValues(allRows);
    sheet.getRange(1, 1, 1, width).setFontWeight('bold').setBackground('#f1f5f9');
    sheet.autoResizeColumns(1, width);
    SpreadsheetApp.flush();

    return exportSpreadsheetAsXlsxBlob_(fileId, 'file Excel');
  } finally {
    trashFileQuietly_(fileId);
  }
}
