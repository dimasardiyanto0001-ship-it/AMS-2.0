/*************************************************************
 * 17_UploadGrid.gs
 * Upload .xlsx untuk menimpa Jadwal atau Realisasi (dicocokkan
 * berdasarkan NAMA karyawan) dan unduh template jadwal kosong.
 * Upload khusus Super Admin. Kode yang tidak dikenal dilewati dan
 * dilaporkan di uploadInfo.invalidCodes (sel tersebut tidak diubah).
 *************************************************************/

const UPLOAD_ERR_XLSX_ = 'Gagal membaca file Excel. Pastikan file berformat .xlsx yang valid.';

/* ============================ UPLOAD ============================ */

function uploadJadwal(base64Data, filename, token) {
  return uploadGridFile_(token, base64Data, filename || 'upload_jadwal.xlsx', {
    roleError: 'Hanya akun dengan role "Super Admin" yang boleh upload jadwal baru.',
    sheetName: CONFIG.SHEET_JADWAL,
    validCodes: () => getShiftCodesDynamic_(),
    auditAction: 'Upload Jadwal'
  });
}

function uploadRealisasi(base64Data, filename, token) {
  return uploadGridFile_(token, base64Data, filename || 'upload_realisasi.xlsx', {
    roleError: 'Hanya akun dengan role "Super Admin" atau "Admin" yang boleh upload realisasi.',
    sheetName: CONFIG.SHEET_REALISASI,
    validCodes: () => CONFIG.STATUS_CODES,
    auditAction: 'Upload Realisasi',
    allowAdmin: true
  });
}

function uploadGridFile_(token, base64Data, filename, opts) {
  const session = opts.allowAdmin
    ? requireUploadAccess_(token, opts.roleError)
    : requireSuperAdmin_(token, opts.roleError);
  if (!base64Data) throw new Error('File tidak ditemukan / kosong.');

  const applyResult = withTempSpreadsheet_(base64Data, filename, '_TEMP_UPLOAD_', MIME_XLSX_, UPLOAD_ERR_XLSX_, fileId => {
    const imported = readImportedGrid_(fileId, opts.sheetName);
    return applyGridImport_(opts.sheetName, imported, opts.validCodes());
  });

  catatAuditLog_(session, opts.auditAction, opts.sheetName, filename,
    applyResult.matched + ' karyawan cocok, ' + applyResult.notFound.length + ' nama tidak ditemukan, ' +
    applyResult.invalidCodes.length + ' kode tidak valid');

  const result = getAllDataInternal_();
  result.uploadInfo = {
    matched: applyResult.matched,
    notFound: applyResult.notFound,
    invalidCodes: applyResult.invalidCodes
  };
  return result;
}

/** Baca grid (nama/jabatan + sel hari) dari sheet pada file upload. */
function readImportedGrid_(fileId, preferredSheetName) {
  const tempSs = SpreadsheetApp.openById(fileId);
  const sheet = tempSs.getSheetByName(preferredSheetName) || tempSs.getSheets()[0];

  const dayCount = getDayColCount_(sheet);
  const numRows = getLastEmployeeRow_(sheet) - CONFIG.DATA_START_ROW + 1;

  if (dayCount <= 0 || numRows <= 0) {
    throw new Error('Format file tidak dikenali. Pastikan strukturnya sama seperti sheet ' + preferredSheetName +
      ' (baris header tanggal di baris ' + CONFIG.HEADER_ROW + ', data karyawan mulai baris ' + CONFIG.DATA_START_ROW + ').');
  }

  return {
    info: sheet.getRange(CONFIG.DATA_START_ROW, 1, numRows, 3).getValues(),
    grid: sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.FIRST_DAY_COL, numRows, dayCount).getValues(),
    dayCount: dayCount
  };
}

/**
 * Tulis hasil import ke sheet target, cocokkan per nama (tanpa beda huruf besar/kecil).
 * Sel kode tidak valid dilewati; sel kosong di file upload MENGOSONGKAN sel target.
 * @return {{matched:number, notFound:string[], invalidCodes:Array<{nama,hariKe,kode}>}}
 */
function applyGridImport_(targetSheetName, imported, validCodeMap) {
  const targetSheet = getSheet_(targetSheetName);
  const numRows = getLastEmployeeRow_(targetSheet) - CONFIG.DATA_START_ROW + 1;
  if (numRows <= 0) throw new Error('Sheet ' + targetSheetName + ' belum berisi karyawan.');
  const dayCountToWrite = Math.min(imported.dayCount, getDayColCount_(targetSheet));

  const currentNames = targetSheet.getRange(CONFIG.DATA_START_ROW, 2, numRows, 1).getValues()
    .map(r => nameKey_(r[0] || ''));

  let matched = 0;
  const notFound = [];
  const invalidCodes = [];

  withScriptLock_(() => {
    const currentValues = dayCountToWrite > 0
      ? targetSheet.getRange(CONFIG.DATA_START_ROW, CONFIG.FIRST_DAY_COL, numRows, dayCountToWrite).getValues()
      : [];
    const touched = new Set();

    imported.info.forEach((row, i) => {
      const nama = String(row[1] || '').trim();
      if (!nama) return;

      const idx = currentNames.indexOf(nama.toLowerCase());
      if (idx === -1) {
        notFound.push(nama);
        return;
      }

      const rowToWrite = currentValues[idx] ? currentValues[idx].slice() : new Array(dayCountToWrite).fill('');
      imported.grid[i].slice(0, dayCountToWrite).forEach((raw, dayIdx) => {
        const kode = (raw === null || raw === undefined) ? '' : String(raw).trim().toUpperCase();
        if (kode === '' || validCodeMap.hasOwnProperty(kode)) {
          rowToWrite[dayIdx] = kode;
        } else {
          invalidCodes.push({ nama: nama, hariKe: dayIdx + 1, kode: kode });
        }
      });

      if (currentValues[idx]) currentValues[idx] = rowToWrite;
      touched.add(idx);
      matched++;
    });

    // Tulis hanya baris yang cocok (baris lain tidak disentuh, formula di sana tetap utuh).
    touched.forEach(idx => {
      targetSheet.getRange(CONFIG.DATA_START_ROW + idx, CONFIG.FIRST_DAY_COL, 1, dayCountToWrite).setValues([currentValues[idx]]);
    });
    SpreadsheetApp.flush();
  }, 20000);

  return { matched: matched, notFound: notFound, invalidCodes: invalidCodes };
}

/* ============================ TEMPLATE ============================ */

/** Template Excel jadwal kosong (opsional hanya satu department) untuk diisi lalu diupload. */
function generateJadwalTemplateExcel(token, department) {
  requireMenuAccess_(token, 'templatejadwal');
  const deptFilter = String(department || '').trim();

  return withTempSheetCopy_(getSheet_(CONFIG.SHEET_JADWAL), 'Template_Jadwal_Kosong_', 'Jadwal',
    copied => {
      if (deptFilter && deptFilter !== 'ALL') keepOnlyDepartmentRows_(copied, deptFilter);

      const numRows = getLastEmployeeRow_(copied) - CONFIG.DATA_START_ROW + 1;
      const dayCount = getDayColCount_(copied);
      if (numRows > 0 && dayCount > 0) {
        const blankGrid = Array.from({ length: numRows }, () => new Array(dayCount).fill(''));
        copied.getRange(CONFIG.DATA_START_ROW, CONFIG.FIRST_DAY_COL, numRows, dayCount).setValues(blankGrid);
      }
    },
    (copied, fileId) => xlsxDownloadResult_('Template_Jadwal_Kosong.xlsx', exportSpreadsheetAsXlsxBlob_(fileId, 'template Excel'))
  );
}

/** Hapus baris karyawan di luar department (dari bawah ke atas supaya indeks tetap valid). */
function keepOnlyDepartmentRows_(sheet, department) {
  const master = getKaryawanMasterMap_();
  const allowed = new Set(Object.keys(master).filter(k => master[k].department === department));

  const lastRow = getLastEmployeeRow_(sheet);
  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  if (numRows <= 0) return;

  const names = sheet.getRange(CONFIG.DATA_START_ROW, 2, numRows, 1).getValues().map(r => nameKey_(r[0] || ''));
  for (let i = names.length - 1; i >= 0; i--) {
    if (names[i] && !allowed.has(names[i])) sheet.deleteRow(CONFIG.DATA_START_ROW + i);
  }
}
