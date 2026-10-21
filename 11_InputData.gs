/*************************************************************
 * 11_InputData.gs
 * Simpan input harian: realisasi kehadiran (butuh izin EDIT menu
 * "input") dan edit jadwal langsung (khusus Super Admin). Keduanya
 * punya jalur satuan (satu sel) dan jalur batch (banyak sel sekaligus).
 *************************************************************/

/** Kode dari klien -> huruf besar tanpa spasi. */
function normalizeKode_(kode) {
  return (kode || '').toString().toUpperCase().trim();
}

/** Nilai yang ditulis ke sel Realisasi: "1" dan "2" disimpan sebagai angka. */
function realisasiCellValue_(kode) {
  if (kode === '1') return 1;
  if (kode === '2') return 2;
  return kode;
}

function assertValidStatusCode_(kode) {
  if (kode !== '' && !CONFIG.STATUS_CODES.hasOwnProperty(kode)) {
    throw new Error('Kode status tidak valid: ' + kode);
  }
}

/** rowIndex/dayIndex harus bilangan bulat. */
function assertIntegerCell_(rowIndex, dayIndex) {
  if (!Number.isInteger(Number(rowIndex)) || !Number.isInteger(Number(dayIndex))) {
    throw new Error('rowIndex dan dayIndex harus berupa angka bulat.');
  }
}

/** Sel harus berada di dalam area data. lastRow/dayCount dibaca SEKALI oleh pemanggil (penting untuk batch). */
function assertCellInRange_(rowIndex, dayIndex, lastRow, dayCount) {
  rowIndex = Number(rowIndex);
  dayIndex = Number(dayIndex);

  if (rowIndex < CONFIG.DATA_START_ROW || rowIndex > lastRow) {
    throw new Error(
      'rowIndex (' + rowIndex + ') di luar jangkauan data karyawan (baris ' +
      CONFIG.DATA_START_ROW + ' s.d. ' + lastRow + ').'
    );
  }
  if (dayIndex < 0 || dayIndex >= dayCount) {
    throw new Error('dayIndex (' + dayIndex + ') di luar jangkauan hari periode ini (0 s.d. ' + (dayCount - 1) + ').');
  }
}

/** Pastikan rowIndex/dayIndex bilangan bulat dan berada di dalam area data. */
function validateCellTarget_(sheet, rowIndex, dayIndex) {
  assertIntegerCell_(rowIndex, dayIndex);
  assertCellInRange_(rowIndex, dayIndex, getLastEmployeeRow_(sheet), getDayColCount_(sheet));
}

/**
 * Validasi daftar sel {rowIndex, dayIndex, kode} tanpa menulis apa pun. Area data dibaca SEKALI.
 * kodeChecker(kode) melempar Error bila kode tidak sah. Sel yang sama dua kali: yang terakhir menang.
 * @return {Array<{row:number, day:number, kode:string}>}
 */
function validateCellUpdates_(sheet, updates, kodeChecker) {
  const lastRow = getLastEmployeeRow_(sheet);
  const dayCount = getDayColCount_(sheet);
  const cells = {};
  updates.forEach(u => {
    if (!u || typeof u !== 'object') throw new Error('Data update tidak valid.');
    assertIntegerCell_(u.rowIndex, u.dayIndex);
    assertCellInRange_(u.rowIndex, u.dayIndex, lastRow, dayCount);
    const kode = normalizeKode_(u.kode);
    kodeChecker(kode);
    const row = Number(u.rowIndex), day = Number(u.dayIndex);
    cells[row + ':' + day] = { row: row, day: day, kode: kode };
  });
  return Object.keys(cells).map(k => cells[k]);
}

/**
 * Tulis sel-sel ke sheet: dikelompokkan per baris, hari yang berurutan digabung
 * menjadi SATU panggilan setValues (jauh lebih cepat daripada setValue per sel).
 * valueOf(kode) menentukan nilai yang ditulis.
 */
function writeGridCells_(sheet, items, valueOf) {
  const byRow = {};
  items.forEach(it => { (byRow[it.row] = byRow[it.row] || []).push(it); });
  Object.keys(byRow).forEach(r => {
    const list = byRow[r].sort((a, b) => a.day - b.day);
    let i = 0;
    while (i < list.length) {
      let j = i;
      while (j + 1 < list.length && list[j + 1].day === list[j].day + 1) j++;
      const run = list.slice(i, j + 1);
      sheet.getRange(Number(r), CONFIG.FIRST_DAY_COL + run[0].day, 1, run.length)
        .setValues([run.map(x => valueOf(x.kode))]);
      i = j + 1;
    }
  });
}

/* ============================ REALISASI ============================ */

function saveRealisasi(rowIndex, dayIndex, kode, token) {
  const session = requireMenuEditAccess_(token, 'input');
  kode = normalizeKode_(kode);
  assertValidStatusCode_(kode);
  dayIndex = Number(dayIndex); // string "5" + 1 akan jadi "51" di log; paksa angka

  withScriptLock_(() => {
    const sheet = getSheet_(CONFIG.SHEET_REALISASI);
    validateCellTarget_(sheet, rowIndex, dayIndex);
    sheet.getRange(rowIndex, CONFIG.FIRST_DAY_COL + dayIndex).setValue(realisasiCellValue_(kode));
    SpreadsheetApp.flush();
  });

  catatAuditLog_(session, 'Edit Realisasi', 'Realisasi',
    'Baris ' + rowIndex + ', Hari ke-' + (dayIndex + 1), 'Kode diubah jadi "' + kode + '"');

  return getAllDataInternal_(true);
}

function saveRealisasiBatch(updates, token) {
  const session = requireMenuEditAccess_(token, 'input');
  if (!Array.isArray(updates)) throw new Error('Data update harus berupa daftar sel.');

  withScriptLock_(() => {
    const sheet = getSheet_(CONFIG.SHEET_REALISASI);

    // Tahap 1: validasi SEMUA sel dulu; satu salah -> tidak ada yang ditulis.
    const items = validateCellUpdates_(sheet, updates, assertValidStatusCode_);

    // Tahap 2: tulis (per baris, hari berurutan digabung).
    writeGridCells_(sheet, items, realisasiCellValue_);
    SpreadsheetApp.flush();
  }, 20000);

  catatAuditLog_(session, 'Edit Realisasi (Batch)', 'Realisasi',
    updates.length + ' sel', 'Update batch ' + updates.length + ' sel realisasi');

  return getAllDataInternal_(true);
}

/* ============================ JADWAL (SUPER ADMIN) ============================ */

/** Ubah satu sel shift di sheet Jadwal (klik sel di web, khusus Super Admin, periode aktif). */
function saveJadwal(rowIndex, dayIndex, kode, token) {
  const session = requireUploadAccess_(token, 'Hanya akun dengan role "Super Admin" atau "Admin" yang boleh mengubah Jadwal.');
  kode = normalizeKode_(kode);
  if (kode !== '' && !getShiftCodesDynamic_().hasOwnProperty(kode)) {
    throw new Error('Kode shift tidak valid: ' + kode);
  }
  dayIndex = Number(dayIndex);

  withScriptLock_(() => {
    const sheet = getSheet_(CONFIG.SHEET_JADWAL);
    validateCellTarget_(sheet, rowIndex, dayIndex);
    sheet.getRange(rowIndex, CONFIG.FIRST_DAY_COL + dayIndex).setValue(kode);
    SpreadsheetApp.flush();
  });

  catatAuditLog_(session, 'Edit Jadwal', 'Jadwal',
    'Baris ' + rowIndex + ', Hari ke-' + (dayIndex + 1), 'Kode diubah jadi "' + kode + '"');

  return getAllDataInternal_();
}

/** Batas sel dalam satu batch jadwal (sebulan penuh ±150 karyawan = ±4.650 sel). */
const JADWAL_BATCH_MAX_ = 6000;

/** Ringkasan satu baris untuk Audit Log: jumlah sel, jumlah karyawan, rincian kode. */
function summarizeJadwalBatch_(items) {
  const rows = {};
  const codes = {};
  items.forEach(it => {
    rows[it.row] = true;
    const label = it.kode === '' ? '(kosong)' : it.kode;
    codes[label] = (codes[label] || 0) + 1;
  });
  const parts = Object.keys(codes)
    .sort((a, b) => (codes[b] - codes[a]) || (a < b ? -1 : 1))
    .map(k => k + '\u00d7' + codes[k]);
  let detail = 'Update batch ' + items.length + ' sel jadwal pada ' + Object.keys(rows).length +
    ' karyawan (' + parts.join(', ') + ')';
  if (detail.length > 300) detail = detail.substring(0, 297) + '...';
  return { total: items.length, detail: detail };
}

/**
 * Ubah BANYAK sel shift sekaligus di sheet Jadwal (khusus Super Admin, periode aktif).
 * Semua-atau-tidak-sama-sekali: satu sel tidak valid = tidak ada yang ditulis.
 * Satu batch dicatat SATU baris di Audit Log (bukan satu baris per sel).
 * @param {Array<{rowIndex:number, dayIndex:number, kode:string}>} updates
 */
function saveJadwalBatch(updates, token) {
  const session = requireUploadAccess_(token, 'Hanya akun dengan role "Super Admin" atau "Admin" yang boleh mengubah Jadwal.');
  if (!Array.isArray(updates)) throw new Error('Data update harus berupa daftar sel.');
  if (updates.length === 0) return getAllDataInternal_();
  if (updates.length > JADWAL_BATCH_MAX_) {
    throw new Error('Terlalu banyak sel dalam satu batch (maks ' + JADWAL_BATCH_MAX_ + ').');
  }

  let summary = null;
  withScriptLock_(() => {
    const sheet = getSheet_(CONFIG.SHEET_JADWAL);
    const shiftCodes = getShiftCodesDynamic_();

    // Tahap 1: validasi SEMUA sel dulu; satu salah -> tidak ada yang ditulis.
    const items = validateCellUpdates_(sheet, updates, kode => {
      if (kode !== '' && !shiftCodes.hasOwnProperty(kode)) throw new Error('Kode shift tidak valid: ' + kode);
    });

    // Tahap 2: tulis (per baris, hari berurutan digabung).
    writeGridCells_(sheet, items, kode => kode);
    SpreadsheetApp.flush();
    summary = summarizeJadwalBatch_(items);
  }, 20000);

  catatAuditLog_(session, 'Edit Jadwal (Batch)', 'Jadwal', summary.total + ' sel', summary.detail);
  return getAllDataInternal_();
}
