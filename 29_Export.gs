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
  const labels = sheet.getRange(from, 1, numRows, 2).getValues(); // kolom A (label) + B (legenda punya isi di B)
  const colLetter = c => sheet.getRange(1, c).getA1Notation().replace(/\d+/g, '');
  const first = colLetter(CONFIG.FIRST_DAY_COL), last = colLetter(CONFIG.FIRST_DAY_COL + dayCount - 1);
  let firstRecap = 0, lastRecap = 0;

  for (let i = 0; i < labels.length; i++) {
    const key = String(labels[i][0] || '').trim().toLowerCase();
    if (key === 'legenda:' || key === 'legenda') break; // legenda di bawahnya bukan baris rekap
    const row = from + i;

    if (key === 'jumlah man power') {
      if (!firstRecap) continue;
      const fm = [];
      for (let d = 0; d < dayCount; d++) {
        const L = colLetter(CONFIG.FIRST_DAY_COL + d);
        fm.push('=SUM(' + L + firstRecap + ':' + L + lastRecap + ')');
      }
      sheet.getRange(row, CONFIG.FIRST_DAY_COL, 1, dayCount).setFormulas([fm]);
      const tc = sheet.getRange(row, CONFIG.FIRST_DAY_COL + dayCount);
      if (typeof tc.getValue() === 'number') tc.setFormula('=SUM(' + first + row + ':' + last + row + ')');
      continue;
    }
    if (!crit[key] || String(labels[i][1] || '').trim() !== '') continue;

    const formulas = [];
    for (let d = 0; d < dayCount; d++) {
      const L = colLetter(CONFIG.FIRST_DAY_COL + d);
      const rng = L + CONFIG.DATA_START_ROW + ':' + L + lastEmp;
      formulas.push('=' + crit[key].map(k => 'COUNTIF(' + rng + ',"' + k + '")').join('+'));
    }
    sheet.getRange(row, CONFIG.FIRST_DAY_COL, 1, dayCount).setFormulas([formulas]);
    const totalCell = sheet.getRange(row, CONFIG.FIRST_DAY_COL + dayCount);
    if (typeof totalCell.getValue() === 'number') totalCell.setFormula('=SUM(' + first + row + ':' + last + row + ')');
    if (!firstRecap) firstRecap = row;
    lastRecap = row;
  }
}

/**
 * Kolom "Hari Kerja" (kolom tepat di kanan kolom hari) pada arsip berupa nilai beku, dan kosong untuk
 * karyawan yang ditambah belakangan. Tulis ulang sebagai rumus A1 biasa untuk SEMUA baris karyawan pada
 * salinan ekspor: kode "1" = 1 hari kerja, kode "2" (lembur) = 2 hari. Hanya untuk Realisasi
 * dan hanya bila kolomnya memang berjudul "Hari Kerja".
 */
function restoreRowFormulasInExport_(copied, sheetName) {
  try {
    if (sheetName !== 'Realisasi' && sheetName !== 'Jadwal') return;
    const dayCount = getDayColCount_(copied);
    if (dayCount <= 0) return;
    const col = CONFIG.FIRST_DAY_COL + dayCount;
    if (copied.getLastColumn() < col) return;
    const head = String(copied.getRange(CONFIG.HEADER_ROW, col).getValue() || '').trim().toLowerCase();
    if (head !== 'hari kerja') return;

    const last = getLastEmployeeRow_(copied);
    const numRows = last - CONFIG.DATA_START_ROW + 1;
    if (numRows <= 0) return;
    const names = copied.getRange(CONFIG.DATA_START_ROW, 2, numRows, 1).getValues();
    const letter = c => copied.getRange(1, c).getA1Notation().replace(/\d+/g, '');
    const first = letter(CONFIG.FIRST_DAY_COL), lastL = letter(CONFIG.FIRST_DAY_COL + dayCount - 1);

    names.forEach((n, i) => {
      if (n[0] === '' || n[0] === null) return;
      const r = CONFIG.DATA_START_ROW + i;
      const rng = first + r + ':' + lastL + r;
      copied.getRange(r, col).setFormula(sheetName === 'Jadwal'
        ? '=COUNTA(' + rng + ')-COUNTIF(' + rng + ',"OFF")' // Jadwal: hari berjadwal = sel terisi selain OFF
        : '=COUNTIF(' + rng + ',"1")+COUNTIF(' + rng + ',"2")*2');
    });
  } catch (e) {
    Logger.log('restoreRowFormulasInExport_: ' + e.message);
  }
}

/**
 * Rekap harian sheet Jadwal arsip (Shift Pagi/Siang/Malam/Midle, Jumlah Libur, Total Jadwal) juga beku
 * (total semua karyawan). Tulis ulang sebagai rumus COUNTIF atas baris karyawan yang tersisa.
 * Kode shift untuk tiap baris diambil dari Setting Jam Kerja: label shift yang diawali kata baris itu
 * (mis. "Siang" -> S dan S1). Total Jadwal = jumlah baris shift (tanpa OFF), sama seperti aslinya.
 */
function rebuildJadwalRecapInExport_(sheet) {
  try {
    const lastEmp = getLastEmployeeRow_(sheet);
    const dayCount = getDayColCount_(sheet);
    if (lastEmp < CONFIG.DATA_START_ROW || dayCount <= 0) return;
    const shifts = getShiftCodesDynamic_();
    const codesFor = word => Object.keys(shifts).filter(k => k !== 'OFF' &&
      String(shifts[k].label || '').trim().toLowerCase().indexOf(word) === 0);

    const from = lastEmp + 1;
    const numRows = Math.min(30, sheet.getMaxRows() - from + 1);
    if (numRows <= 0) return;
    const labels = sheet.getRange(from, 1, numRows, 2).getValues();
    const letter = c => sheet.getRange(1, c).getA1Notation().replace(/\d+/g, '');
    const first = letter(CONFIG.FIRST_DAY_COL), lastL = letter(CONFIG.FIRST_DAY_COL + dayCount - 1);
    const shiftRows = [];

    const write = (row, makeFormula) => {
      const fm = [];
      for (let d = 0; d < dayCount; d++) fm.push(makeFormula(letter(CONFIG.FIRST_DAY_COL + d)));
      sheet.getRange(row, CONFIG.FIRST_DAY_COL, 1, dayCount).setFormulas([fm]);
      const tc = sheet.getRange(row, CONFIG.FIRST_DAY_COL + dayCount);
      if (typeof tc.getValue() === 'number') tc.setFormula('=SUM(' + first + row + ':' + lastL + row + ')');
    };

    for (let i = 0; i < labels.length; i++) {
      const key = String(labels[i][0] || '').trim().toLowerCase();
      if (key === 'legenda:' || key === 'legenda') break;
      if (String(labels[i][1] || '').trim() !== '') continue;
      const row = from + i;

      if (key.indexOf('shift ') === 0) {
        const codes = codesFor(key.slice(6).trim());
        if (!codes.length) continue; // tidak ada kode shift yang cocok: biarkan apa adanya
        write(row, L => '=' + codes.map(k => 'COUNTIF(' + L + CONFIG.DATA_START_ROW + ':' + L + lastEmp + ',"' + k + '")').join('+'));
        shiftRows.push(row);
      } else if (key.indexOf('jumlah libur') === 0) {
        write(row, L => '=COUNTIF(' + L + CONFIG.DATA_START_ROW + ':' + L + lastEmp + ',"OFF")');
      } else if (key === 'total jadwal' && shiftRows.length) {
        write(row, L => '=' + shiftRows.map(r => L + r).join('+'));
      }
    }
  } catch (e) {
    Logger.log('rebuildJadwalRecapInExport_: ' + e.message);
  }
}

/**
 * Warna sel jadwal/realisasi pada arsip adalah isian statis (tidak ikut berubah bila isi sel berubah).
 * Hapus warna statis di area hari, lalu pasang aturan warna otomatis (conditional formatting) per kode
 * yang warnanya diambil dari LEGENDA di sheet itu sendiri. Kode shift yang tidak ada di legenda
 * (mis. S1, MD1) memakai warna legenda yang labelnya diawali nama shift tersebut ("Siang 1" -> Siang).
 */
function applyLegendColorsInExport_(sheet) {
  try {
    const lastEmp = getLastEmployeeRow_(sheet);
    const dayCount = getDayColCount_(sheet);
    if (lastEmp < CONFIG.DATA_START_ROW || dayCount <= 0) return;

    const from = lastEmp + 1;
    const numRows = Math.min(40, sheet.getMaxRows() - from + 1);
    if (numRows <= 0) return;
    const block = sheet.getRange(from, 1, numRows, 2);
    const vals = block.getValues();
    const bgs = block.getBackgrounds();
    let start = -1;
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][0] || '').trim().toLowerCase().indexOf('legenda') === 0) { start = i + 1; break; }
    }
    if (start < 0) return;

    const colorOf = {};   // kode -> warna
    const labelColor = []; // [labelHurufKecil, warna]
    const isBlankColor = c => !c || c === '#ffffff';
    for (let i = start; i < vals.length; i++) {
      const code = vals[i][0];
      if (code === '' || code === null) break;
      const color = bgs[i][0];
      if (isBlankColor(color)) continue;
      colorOf[String(code).trim().toUpperCase()] = { code: code, color: color };
      const label = String(vals[i][1] || '').trim().toLowerCase();
      if (label) labelColor.push([label, color]);
    }

    // Kode shift yang tidak ada di legenda: pinjam warna legenda yang labelnya awalan label shift.
    try {
      const shifts = getShiftCodesDynamic_();
      Object.keys(shifts).forEach(k => {
        const key = String(k).toUpperCase();
        if (colorOf[key]) return;
        const lab = String(shifts[k].label || '').trim().toLowerCase();
        const hit = labelColor.find(lc => lab.indexOf(lc[0]) === 0);
        if (hit) colorOf[key] = { code: k, color: hit[1] };
      });
    } catch (e) { /* Setting Jam Kerja tidak terbaca: cukup pakai legenda */ }

    const area = sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.FIRST_DAY_COL, lastEmp - CONFIG.DATA_START_ROW + 1, dayCount);
    area.setBackground(null);

    const rules = sheet.getConditionalFormatRules();
    Object.keys(colorOf).forEach(key => {
      const e = colorOf[key];
      const variants = (typeof e.code === 'number') ? [e.code]
        : (/^\d+$/.test(key) ? [Number(key), key] : [key, key.toLowerCase()].filter((v, i, a) => a.indexOf(v) === i));
      variants.forEach(v => {
        const b = SpreadsheetApp.newConditionalFormatRule().setRanges([area]);
        const rule = (typeof v === 'number' ? b.whenNumberEqualTo(v) : b.whenTextEqualTo(v)).setBackground(e.color).build();
        rules.push(rule);
      });
    });
    sheet.setConditionalFormatRules(rules);
  } catch (err) {
    Logger.log('applyLegendColorsInExport_: ' + err.message);
  }
}

/**
 * Warna sel jadwal/realisasi pada arsip adalah warna beku (tidak mengikuti isi sel). Pada salinan ekspor:
 * baca warna tiap kode dari LEGENDA di bawah tabel, kosongkan warna beku di sel, lalu pasang format
 * bersyarat (kode = isi sel -> warna legenda). Warna lalu otomatis mengikuti bila kode diubah di Excel.
 */
function applyLegendColorRulesInExport_(sheet) {
  try {
    const lastEmp = getLastEmployeeRow_(sheet);
    const dayCount = getDayColCount_(sheet);
    if (lastEmp < CONFIG.DATA_START_ROW || dayCount <= 0) return;

    const from = lastEmp + 1;
    const numRows = Math.min(40, sheet.getMaxRows() - from + 1);
    if (numRows <= 0) return;
    const labels = sheet.getRange(from, 1, numRows, 1).getValues();
    let start = -1;
    for (let i = 0; i < labels.length; i++) {
      if (String(labels[i][0] || '').trim().toLowerCase().indexOf('legenda') === 0) { start = i + 1; break; }
    }
    if (start < 0) return;

    const colors = [];
    for (let i = start; i < labels.length; i++) {
      const code = String(labels[i][0] === null || labels[i][0] === undefined ? '' : labels[i][0]).trim();
      if (!code) break;
      const bg = String(sheet.getRange(from + i, 1).getBackground() || '').toLowerCase();
      if (bg && bg !== '#ffffff') colors.push({ code: code.toUpperCase(), color: bg });
    }
    if (!colors.length) return;

    const grid = sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.FIRST_DAY_COL, lastEmp - CONFIG.DATA_START_ROW + 1, dayCount);
    const tl = grid.getA1Notation().split(':')[0]; // sel kiri-atas, mis. D6 (relatif untuk rumus)

    // Kode yang ada di data tapi tidak ada di legenda (mis. S1, MD1, WD): pakai warna yang paling sering dipakainya.
    const known = {};
    colors.forEach(c => { known[c.code] = true; });
    const vals = grid.getValues(), bgs = grid.getBackgrounds();
    const tally = {};
    vals.forEach((row, i) => row.forEach((v, j) => {
      const code = String(v === null || v === undefined ? '' : v).trim().toUpperCase();
      const bg = String(bgs[i][j] || '').toLowerCase();
      if (!code || known[code] || !bg || bg === '#ffffff') return;
      tally[code] = tally[code] || {};
      tally[code][bg] = (tally[code][bg] || 0) + 1;
    }));
    Object.keys(tally).forEach(code => {
      const best = Object.keys(tally[code]).sort((x, y) => tally[code][y] - tally[code][x])[0];
      colors.push({ code: code, color: best });
    });

    grid.setBackground(null);
    const rules = colors.map(c => SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=UPPER(TRIM(' + tl + '&""))="' + c.code.replace(/"/g, '') + '"')
      .setBackground(c.color)
      .setRanges([grid])
      .build());
    sheet.setConditionalFormatRules(rules.concat(sheet.getConditionalFormatRules()));
  } catch (e) {
    Logger.log('applyLegendColorRulesInExport_: ' + e.message);
  }
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
        if (sheetName === 'Jadwal') rebuildJadwalRecapInExport_(copied);
        restoreRowFormulasInExport_(copied, sheetName);
        colorDayHeaders_(copied); // arsip lama membawa warna header bulan sebelumnya
        applyLegendColorRulesInExport_(copied);
        applyLegendColorsInExport_(copied);
      },
      (copied, fileId) => xlsxDownloadResult_(
        sheetName + '_ARSIP_' + arc.period.replace(/[^\w]+/g, '') + deptSuffixOf_(department) + '_' + dateStr + '.xlsx',
        exportSpreadsheetAsXlsxBlob_(fileId, 'Excel arsip')));
  }

  const deptFilter = String(department || '').trim();
  const isFiltered = deptFilter && deptFilter !== 'ALL';
  return withTempSheetCopy_(getSheet_(sheetName), 'TempXLSX_', sheetName,
    copied => {
      if (isFiltered) keepOnlyDepartmentRows_(copied, deptFilter);
      colorDayHeaders_(copied);
      applyLegendColorRulesInExport_(copied);
    },
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
