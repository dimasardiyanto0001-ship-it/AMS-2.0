/*************************************************************
 * 26_JamKerja.gs
 * Sheet pendukung tersembunyi "JamKerja" (jam kerja per hari) dan
 * "JamScan" (teks scan masuk-pulang), serta pengayaan baris rekap
 * dengan NIK, jam kerja aktual, dan target jam kerja.
 *
 * Kedua sheet punya tata letak yang sama dengan Realisasi dan
 * disinkronkan ke daftar karyawan Realisasi (nilai lama dipulihkan
 * berdasarkan nama).
 *************************************************************/

/* ============================ SINKRON SHEET ============================ */

function getOrSyncJamKerjaSheet_() {
  return syncAuxGridSheet_(CONFIG.SHEET_JAMKERJA);
}

function getOrSyncJamScanSheet_() {
  return syncAuxGridSheet_(CONFIG.SHEET_JAMSCAN);
}

/** Nilai grid lama per nama (huruf kecil); format sheet lama tidak cocok -> {} (mulai dari kosong). */
function readAuxValuesByName_(sheet) {
  const byName = {};
  try {
    const oldNumRows = getLastEmployeeRow_(sheet) - CONFIG.DATA_START_ROW + 1;
    const oldDayCount = getDayColCount_(sheet);
    if (oldNumRows > 0 && oldDayCount > 0) {
      const oldInfo = sheet.getRange(CONFIG.DATA_START_ROW, 1, oldNumRows, 3).getValues();
      const oldGrid = sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.FIRST_DAY_COL, oldNumRows, oldDayCount).getValues();
      oldInfo.forEach((r, i) => {
        const nama = nameKey_(r[1] || '');
        if (nama) byName[nama] = oldGrid[i];
      });
    }
  } catch (e) { /* sheet lama kosong / formatnya belum sesuai: abaikan */ }
  return byName;
}

/** Buat/segarkan sheet pendukung agar baris & header-nya sama dengan Realisasi. */
function syncAuxGridSheet_(sheetName) {
  const ss = getSpreadsheet_();
  const realisasiSheet = getSheet_(CONFIG.SHEET_REALISASI);
  const dayCount = getDayColCount_(realisasiSheet);
  const numRows = getLastEmployeeRow_(realisasiSheet) - CONFIG.DATA_START_ROW + 1;

  let sheet = ss.getSheetByName(sheetName);
  const isNew = !sheet;
  if (isNew) sheet = ss.insertSheet(sheetName);
  const oldValuesByName = isNew ? {} : readAuxValuesByName_(sheet);

  sheet.clear();
  sheet.getRange(1, 1, 1, 3).setValues([['No', 'Nama', 'Jabatan']]).setFontWeight('bold');
  sheet.getRange(2, 1).setValue(realisasiSheet.getRange(2, 1).getValue());
  sheet.getRange(CONFIG.HEADER_ROW, CONFIG.FIRST_DAY_COL, 1, dayCount)
    .setValues(realisasiSheet.getRange(CONFIG.HEADER_ROW, CONFIG.FIRST_DAY_COL, 1, dayCount).getValues());
  sheet.getRange(CONFIG.DAY_NAME_ROW, CONFIG.FIRST_DAY_COL, 1, dayCount)
    .setValues(realisasiSheet.getRange(CONFIG.DAY_NAME_ROW, CONFIG.FIRST_DAY_COL, 1, dayCount).getValues());

  if (numRows > 0) {
    const info = realisasiSheet.getRange(CONFIG.DATA_START_ROW, 1, numRows, 3).getValues();
    sheet.getRange(CONFIG.DATA_START_ROW, 1, numRows, 3).setValues(info);

    const restoredGrid = info.map(r => {
      const old = oldValuesByName[nameKey_(r[1] || '')];
      if (!(old && old.length)) return new Array(dayCount).fill('');
      const row = old.slice(0, dayCount);
      while (row.length < dayCount) row.push('');
      return row;
    });
    sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.FIRST_DAY_COL, numRows, dayCount).setValues(restoredGrid);
  }

  sheet.hideSheet();
  return sheet;
}

function formatScanPair_(scanMasuk, scanPulang) {
  if (!scanMasuk && !scanPulang) return '';
  return (scanMasuk || '-') + ' - ' + (scanPulang || '-');
}

/* ============================ PENGAYAAN REKAP ============================ */

/** Total target jam kerja per nama (huruf kecil) dari kode shift di grid Jadwal. */
function computeTargetJamKerjaByName_(jadwalGrid) {
  const targetByName = {};
  if (!jadwalGrid) return targetByName;
  const shiftMap = getShiftCodesDynamic_();
  jadwalGrid.employees.forEach(e => {
    let target = 0;
    e.days.forEach(v => {
      const code = cellCode_(v);
      if (shiftMap.hasOwnProperty(code)) target += shiftMap[code].targetJam;
    });
    targetByName[nameKey_(e.nama)] = target;
  });
  return targetByName;
}

function attachNikToRekap_(rekap) {
  const masterMap = getKaryawanMasterMap_();
  return rekap.map(row => {
    const master = masterMap[nameKey_(row.nama || '')];
    return Object.assign({}, row, { nik: master ? master.nik : '' });
  });
}

/**
 * Tambah totalJamKerja, totalHariKerja, targetJamKerja, kekuranganJam ke tiap baris rekap.
 * @param range opsional {from,to}: hanya hari ke-from..to yang dijumlah untuk jam aktual.
 *              (jadwalGrid untuk target harus SUDAH dipotong ke rentang yang sama oleh pemanggil.)
 */
function attachJamKerjaToRekap_(rekapArray, jadwalGrid, jamKerjaSheetOverride, range) {
  const jamKerjaSheet = jamKerjaSheetOverride || getSpreadsheet_().getSheetByName(CONFIG.SHEET_JAMKERJA);
  const targetByName = computeTargetJamKerjaByName_(jadwalGrid);
  const actualByName = {};
  if (jamKerjaSheet) {
    readGridFromSheetObj_(jamKerjaSheet).employees.forEach(e => {
      let totalJam = 0, totalHari = 0;
      const days = range ? e.days.slice(range.from - 1, range.to) : e.days;
      days.forEach(v => {
        const num = Number(v);
        if (v !== '' && v !== null && !isNaN(num)) {
          totalJam += num;
          totalHari++;
        }
      });
      actualByName[nameKey_(e.nama)] = { totalJamKerja: round2_(totalJam), totalHariKerja: totalHari };
    });
  }
  return rekapArray.map(r => {
    const key = nameKey_(r.nama || '');
    const actual = actualByName[key] || { totalJamKerja: 0, totalHariKerja: 0 };
    const target = targetByName.hasOwnProperty(key) ? targetByName[key] : 0;
    return Object.assign({}, r, actual, {
      targetJamKerja: round2_(target),
      kekuranganJam: round2_(target - actual.totalJamKerja)
    });
  });
}

function computeJamKerjaRingkasan_(rekapArray) {
  let totalJamKerja = 0, totalTargetJam = 0;
  rekapArray.forEach(r => {
    totalJamKerja += Number(r.totalJamKerja) || 0;
    totalTargetJam += Number(r.targetJamKerja) || 0;
  });
  return {
    totalJamKerja: round2_(totalJamKerja),
    totalTargetJam: round2_(totalTargetJam),
    totalKekuranganJam: round2_(totalTargetJam - totalJamKerja)
  };
}
