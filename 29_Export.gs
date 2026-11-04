/*************************************************************
 * 29_Export.gs
 * Export ke PDF/Excel:
 *   exportSheetPdf       sheet Rekap/Dashboard/Jadwal/Realisasi -> PDF (periode aktif atau arsip)
 *   exportSheetExcel     sheet Jadwal/Realisasi -> Excel (periode aktif atau arsip)
 *                        (Jadwal/Realisasi bisa difilter per department)
 *   generateReportExcel  laporan generik (Summary/Detail/Department/Exception)
 *   generateReportPdf    -> Excel / PDF dari kolom+baris yang dikirim frontend
 *
 * Spreadsheet sementara selalu dibuang, termasuk saat error.
 *************************************************************/

/** Parameter export PDF (urutan dipertahankan; gid disisipkan setelah format). */
const PDF_EXPORT_PARAMS_ = [
  'size=A4', 'portrait=false', 'fitw=true', 'scale=4', 'gridlines=true',
  'printtitle=false', 'sheetnames=false', 'pagenumbers=true',
  'top_margin=0.4', 'bottom_margin=0.4', 'left_margin=0.4', 'right_margin=0.4',
  'horizontal_alignment=CENTER'
];

const REPORT_TITLE_MAX_ = 90;
const REPORT_TIME_ZONE_ = 'GMT+7';

/* ============================ EXPORT SHEET KE PDF ============================ */

/** Boleh export: Super Admin & Admin (supervisor) selalu; role lain perlu izin menu export. */
function requireExportAccess_(token) {
  const session = verifySession_(token);
  if (isSuperAdminRole_(session.role) || isAdminRole_(session.role)) return session;
  return requireMenuAccess_(token, 'exportsheetpdf');
}

/** Nama sheet arsip (Jadwal/Realisasi) untuk periodKey arsip; lempar error kalau tidak ada. */
function getArchiveSheetNameForExport_(sheetName, periodKey) {
  const colIdx = { Jadwal: 2, Realisasi: 3, Dashboard: 4, Rekap: 5 }[sheetName];
  if (colIdx === undefined) throw new Error('Sheet "' + sheetName + '" tidak diizinkan untuk diexport.');
  const found = findArchiveRow_(getArchiveIndexSheet_(), periodKey);
  if (!found) throw new Error('Arsip untuk periode ini tidak ditemukan.');
  const name = String(found.row[colIdx] || '').trim();
  if (!name || !getSpreadsheet_().getSheetByName(name)) throw new Error('Sheet arsip "' + sheetName + '" tidak ditemukan (mungkin terhapus manual).');
  return { name: name, period: String(found.row[0] || '') };
}

/**
 * Arsip = nilai beku (rumus sudah dibekukan) sehingga baris rekap harian di bawah tabel (Hadir/Off/Izin/...)
 * masih berisi total SEMUA karyawan walau baris sudah disaring per department. Tulis ulang sebagai rumus
 * COUNTIF atas baris yang tersisa supaya angkanya sesuai dan otomatis mengikuti bila baris ditambah.
 */
function rebuildRecapRowsInExport_(sheet) {
  const lastEmp = getLastEmployeeRow_(sheet);
  const dayCount = getDayColCount_(sheet);
  if (lastEmp < CONFIG.DATA_START_ROW || dayCount <= 0) return;
  const crit = {
    'hadir': ['1', '2'], 'lembur': ['2'], 'off': ['OFF'], 'off (libur)': ['OFF'], 'libur': ['OFF'],
    'izin': ['I'], 'alfa': ['A'], 'sakit': ['S'], 'cuti': ['C'], 'resign': ['R']
  };
  const from = lastEmp + 1;
  const numRows = Math.min(30, sheet.getMaxRows() - from + 1);
  if (numRows <= 0) return;
  const labels = sheet.getRange(from, 1, numRows, 1).getValues();
  const colLetter = c => sheet.getRange(1, c).getA1Notation().replace(/\d+/g, '');
  labels.forEach((l, i) => {
    const key = String(l[0] || '').trim().toLowerCase();
    if (!crit[key]) return;
    const row = from + i;
    const formulas = [];
    for (let d = 0; d < dayCount; d++) {
      const L = colLetter(CONFIG.FIRST_DAY_COL + d);
      const rng = L + CONFIG.DATA_START_ROW + ':' + L + lastEmp;
      formulas.push(crit[key].map(k => 'COUNTIF(' + rng + ',"' + k + '")').join('+'));
    }
    sheet.getRange(row, CONFIG.FIRST_DAY_COL, 1, dayCount).setFormulas([formulas.map(f => '=' + f)]);
    const first = colLetter(CONFIG.FIRST_DAY_COL), last = colLetter(CONFIG.FIRST_DAY_COL + dayCount - 1);
    const totalCell = sheet.getRange(row, CONFIG.FIRST_DAY_COL + dayCount);
    if (typeof totalCell.getValue() === 'number') totalCell.setFormula('=SUM(' + first + row + ':' + last + row + ')');
  });
}

function deptSuffixOf_(department) {
  const d = String(department || '').trim();
  return (d && d !== 'ALL') ? '_' + d.replace(/[^\w]+/g, '') : '';
}

function exportSheetPdf(sheetName, token, department, periodKey) {
  requireExportAccess_(token);
  const isArchive = periodKey && periodKey !== 'active';
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || REPORT_TIME_ZONE_, 'yyyyMMdd_HHmm');
  if (isArchive) {
    const arc = getArchiveSheetNameForExport_(sheetName, String(periodKey));
    const isAgg = (sheetName === 'Dashboard' || sheetName === 'Rekap');
    const blobA = exportArchivedSheetPdfBlob_(arc.name, sheetName, isAgg ? 'ALL' : department);
    return {
      filename: sheetName + '_ARSIP_' + arc.period.replace(/[^\w]+/g, '') + deptSuffixOf_(department) + '_' + dateStr + '.pdf',
      mimeType: MIME_PDF_,
      base64: Utilities.base64Encode(blobA.getBytes())
    };
  }
  const blob = exportSheetPdfBlob_(sheetName, department);
  const deptSuffix = (department && department !== 'ALL') ? '_' + department.replace(/[^\w]+/g, '') : '';

  return {
    filename: sheetName + deptSuffix + '_' + dateStr + '.pdf',
    mimeType: MIME_PDF_,
    base64: Utilities.base64Encode(blob.getBytes())
  };
}

/**
 * Export sheet Jadwal / Realisasi ke Excel (.xlsx), periode aktif ATAU arsip.
 * periodKey 'active'/kosong = periode aktif; selain itu = key arsip. Read-only.
 */
function exportSheetExcel(sheetName, token, department, periodKey) {
  requireExportAccess_(token);
  if (['Jadwal', 'Realisasi'].indexOf(sheetName) === -1) throw new Error('Export Excel hanya untuk Jadwal atau Realisasi.');
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || REPORT_TIME_ZONE_, 'yyyyMMdd_HHmm');
  const isArchive = periodKey && periodKey !== 'active';

  if (isArchive) {
    const arc = getArchiveSheetNameForExport_(sheetName, String(periodKey));
    const src = getSpreadsheet_().getSheetByName(arc.name);
    const dArc = String(department || '').trim();
    return withTempSheetCopy_(src, 'TempXLSX_Arsip_', sheetName,
      copied => {
        copied.showSheet();
        if (dArc && dArc !== 'ALL') {
          keepOnlyDepartmentRows_(copied, dArc);
          if (sheetName === 'Realisasi') rebuildRecapRowsInExport_(copied);
        }
      },
      (copied, fileId) => xlsxDownloadResult_(
        sheetName + '_ARSIP_' + arc.period.replace(/[^\w]+/g, '') + deptSuffixOf_(department) + '_' + dateStr + '.xlsx',
        exportSpreadsheetAsXlsxBlob_(fileId, 'Excel arsip')));
  }

  const deptFilter = String(department || '').trim();
  const isFiltered = deptFilter && deptFilter !== 'ALL';
  return withTempSheetCopy_(getSheet_(sheetName), 'TempXLSX_', sheetName,
    copied => { if (isFiltered) keepOnlyDepartmentRows_(copied, deptFilter); },
    (copied, fileId) => xlsxDownloadResult_(
      sheetName + (isFiltered ? '_' + deptFilter.replace(/[^\w]+/g, '') : '') + '_' + dateStr + '.xlsx',
      exportSpreadsheetAsXlsxBlob_(fileId, 'Excel ' + sheetName)));
}

/** Ambil PDF satu sheet (gid) dari spreadsheet fileId. @return {Blob} */
function fetchSheetPdfBlob_(fileId, gid) {
  const url = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?' +
    ['format=pdf', 'gid=' + gid].concat(PDF_EXPORT_PARAMS_).join('&');

  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) {
    throw new Error('Gagal membuat PDF (kode ' + response.getResponseCode() + ').');
  }
  return response.getBlob();
}

/** PDF sheet; department opsional (hanya Jadwal/Realisasi). Dipakai juga oleh laporan email bulanan. */
function exportSheetPdfBlob_(sheetName, department) {
  if (CONFIG.PDF_ALLOWED_SHEETS.indexOf(sheetName) === -1) {
    throw new Error('Sheet "' + sheetName + '" tidak diizinkan untuk diexport.');
  }

  const deptFilter = String(department || '').trim();
  const isFiltered = deptFilter && deptFilter !== 'ALL';

  // Dashboard & Rekap berisi rumus total lintas-baris -> tidak aman difilter di sini.
  if (isFiltered && (sheetName === CONFIG.SHEET_DASHBOARD || sheetName === CONFIG.SHEET_REKAP)) {
    throw new Error('Export PDF sheet "' + sheetName + '" tidak mendukung filter per department (berisi rumus total gabungan). ' +
      'Gunakan menu Report > Department Report untuk PDF ter-filter.');
  }

  const srcSheet = getSheet_(sheetName);
  if (!isFiltered) return fetchSheetPdfBlob_(getSpreadsheet_().getId(), srcSheet.getSheetId());

  // Filtered: salinan sementara tanpa baris karyawan di luar department.
  return withTempSheetCopy_(srcSheet, 'TempPDF_' + sheetName + '_', sheetName,
    copied => keepOnlyDepartmentRows_(copied, deptFilter),
    (copied, fileId) => fetchSheetPdfBlob_(fileId, copied.getSheetId())
  );
}

/**
 * PDF satu sheet ARSIP (disembunyikan) lewat salinan sementara yang ditampilkan
 * (laporan bulanan untuk periode lampau). copyName = nama sheet pada salinan.
 */
function exportArchivedSheetPdfBlob_(archiveSheetName, copyName, department) {
  const src = getSpreadsheet_().getSheetByName(archiveSheetName);
  if (!src) throw new Error('Sheet arsip "' + archiveSheetName + '" tidak ditemukan.');

  const deptFilter = String(department || '').trim();
  const isFiltered = deptFilter && deptFilter !== 'ALL';
  return withTempSheetCopy_(src, 'TempPDF_Arsip_', copyName,
    copied => { copied.showSheet(); if (isFiltered) keepOnlyDepartmentRows_(copied, deptFilter); },
    (copied, fileId) => fetchSheetPdfBlob_(fileId, copied.getSheetId())
  );
}

/* ============================ LAPORAN GENERIK ============================ */

/** Nama sheet laporan: maks 90 karakter, karakter terlarang diganti "-". */
function reportSheetName_(title) {
  return String(title).substring(0, REPORT_TITLE_MAX_).replace(/[\/\\?*\[\]:]/g, '-');
}

function reportFilename_(reportType, extension) {
  return reportType + '_' + Utilities.formatDate(new Date(), REPORT_TIME_ZONE_, 'yyyyMMdd_HHmm') + '.' + extension;
}

/**
 * Buat spreadsheet sementara berisi satu sheet laporan (header + baris), lalu jalankan
 * produce(fileId, sheet). Spreadsheet sementara selalu dibuang.
 */
function withTempReportSheet_(tempNamePrefix, reportType, title, columns, rows, styleHeader, produce) {
  if (!Array.isArray(columns) || columns.length === 0) throw new Error('Kolom laporan tidak boleh kosong.');
  rows = Array.isArray(rows) ? rows : [];

  const ss = SpreadsheetApp.create(tempNamePrefix + reportType + '_' + new Date().getTime());
  const fileId = ss.getId();
  try {
    const sheet = ss.getActiveSheet();
    sheet.setName(reportSheetName_(title));

    styleHeader(sheet.getRange(1, 1, 1, columns.length).setValues([columns]).setFontWeight('bold'));
    if (rows.length) sheet.getRange(2, 1, rows.length, columns.length).setValues(rows);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, columns.length);
    SpreadsheetApp.flush();

    return produce(fileId, sheet);
  } finally {
    trashFileQuietly_(fileId);
  }
}

/** Generator Excel & PDF generik untuk 4 jenis Report (Summary/Detail/Department/Exception). */
function generateReportExcel(reportType, title, columns, rows, token) {
  requireReportAccess_(reportType, token);

  return withTempReportSheet_('TempReport_', reportType, title, columns, rows,
    headerRange => headerRange.setBackground('#2563eb').setFontColor('#ffffff'),
    fileId => ({
      base64: Utilities.base64Encode(exportSpreadsheetAsXlsxBlob_(fileId, 'laporan Excel').getBytes()),
      filename: reportFilename_(reportType, 'xlsx'),
      mimeType: MIME_XLS_
    })
  );
}

function generateReportPdf(reportType, title, columns, rows, token) {
  requireReportAccess_(reportType, token);

  return withTempReportSheet_('TempReportPDF_', reportType, title, columns, rows,
    headerRange => headerRange,
    (fileId, sheet) => {
      const url = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export' +
        '?format=pdf&size=A4&portrait=false&fitw=true&gridlines=true' +
        '&printtitle=false&sheetnames=false&pagenumbers=true&gid=' + sheet.getSheetId();
      const response = UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
      });
      if (response.getResponseCode() !== 200) {
        throw new Error('Gagal membuat laporan PDF (kode ' + response.getResponseCode() + ').');
      }
      return {
        base64: Utilities.base64Encode(response.getBlob().getBytes()),
        filename: reportFilename_(reportType, 'pdf'),
        mimeType: MIME_PDF_
      };
    }
  );
}
