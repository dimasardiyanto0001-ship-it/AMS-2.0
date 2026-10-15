/**********************************************************************
 * 23_SettingJamKerja.gs
 * Setting » Jam Kerja + SUMBER KEBENARAN daftar kode shift untuk seluruh
 * sistem (Jadwal, Upload Jadwal, Pola Jadwal, dll).
 * Sheet "SettingJamKerja": [KodeShift, Label, JamMasuk, JamPulang, TargetJamPerHari].
 *
 * Kolom JamMasuk/JamPulang (C-D) SELALU berformat teks ('@'). Bila dibiarkan
 * format default, Sheets mengubah "06:00" jadi nilai waktu dan konversi balik
 * lewat Utilities.formatDate bisa meleset beberapa jam (histori offset
 * Asia/Jakarta). Sebagai teks, "HH:mm" tersimpan apa adanya.
 **********************************************************************/

const SHEET_SETTING_JAMKERJA = 'SettingJamKerja';
const SETTING_JAMKERJA_HEADERS_ = ['KodeShift', 'Label', 'JamMasuk', 'JamPulang', 'TargetJamPerHari'];
const SETTING_JAMKERJA_FORMAT_ROWS_ = 998;   // baris cadangan berformat teks di bawah data
const SHIFT_LABEL_MAX_ = 40;
const TIME_HHMM_RE_ = /^([01]?\d|2[0-3]):([0-5]\d)$/;

// Palet warna tetap, dipakai berputar sesuai urutan baris di sheet.
const SHIFT_COLOR_PALETTE_ = ['#38bdf8', '#fb923c', '#818cf8', '#facc15', '#f472b6', '#34d399', '#f97316', '#a78bfa', '#fb7185', '#22d3ee'];

/* ============================ SHEET ============================ */

/** Paksa kolom C (JamMasuk) & D (JamPulang) bertipe teks pada baris startRow..startRow+numRows-1 (dijepit ke batas sheet). */
function forceTimeColumnsAsText_(sheet, startRow, numRows) {
  const room = sheet.getMaxRows() - startRow + 1;
  const rows = Math.min(numRows, room);
  if (rows <= 0) return;
  sheet.getRange(startRow, 3, rows, 2).setNumberFormat('@');
}

/**
 * Baris data + cadangan sudah berformat teks? Cek baris data yang ada dan baris
 * paling bawah area cadangan, supaya pembacaan biasa tidak menulis ±1000 sel
 * setiap kali (fungsi ini dipanggil di tiap getAllData).
 */
function timeColumnsAlreadyText_(sheet, lastRow, reserveEndRow) {
  const checkRows = [];
  if (lastRow >= 2) {
    sheet.getRange(2, 3, lastRow - 1, 2).getNumberFormats().forEach(r => checkRows.push(r));
  }
  if (reserveEndRow > lastRow) checkRows.push(sheet.getRange(reserveEndRow, 3, 1, 2).getNumberFormats()[0]);
  return checkRows.every(r => r[0] === '@' && r[1] === '@');
}

function getOrCreateSettingJamKerjaSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(SHEET_SETTING_JAMKERJA);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SETTING_JAMKERJA);
    sheet.appendRow(SETTING_JAMKERJA_HEADERS_);
    // Format teks dipasang SEBELUM data ditulis, dan untuk banyak baris ke depan.
    forceTimeColumnsAsText_(sheet, 2, SETTING_JAMKERJA_FORMAT_ROWS_);

    const shiftCodes = CONFIG.SHIFT_CODES || {};
    const targetMap = CONFIG.SHIFT_HOURS_TARGET || {};
    const seedRows = [];
    Object.keys(shiftCodes).forEach(kode => {
      if (kode === 'OFF') return; // OFF tidak pernah masuk sheet
      const info = shiftCodes[kode];
      const label = (info && typeof info === 'object') ? (info.label || kode) : (info || kode);
      seedRows.push([kode, label, '', '', targetMap.hasOwnProperty(kode) ? targetMap[kode] : 8]);
    });
    if (seedRows.length) sheet.getRange(2, 1, seedRows.length, 5).setValues(seedRows);
    return sheet;
  }

  // Migrasi sheet lama: tambah kolom TargetJamPerHari bila belum ada.
  if (sheet.getRange(1, 5).getValue() !== 'TargetJamPerHari') {
    sheet.getRange(1, 5).setValue('TargetJamPerHari');
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const targetMap = CONFIG.SHIFT_HOURS_TARGET || {};
      const kodeCol = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      const defaultTargets = kodeCol.map(r => {
        const kode = String(r[0] || '').trim();
        return [targetMap.hasOwnProperty(kode) ? targetMap[kode] : 8];
      });
      sheet.getRange(2, 5, defaultTargets.length, 1).setValues(defaultTargets);
    }
  }

  // Sheet lama yang kolom C-D-nya sempat berformat waktu: paksa ke teks (baris ada + cadangan).
  const lastRowNow = Math.max(sheet.getLastRow(), 1);
  const reserveEnd = Math.min(lastRowNow + SETTING_JAMKERJA_FORMAT_ROWS_, sheet.getMaxRows());
  if (!timeColumnsAlreadyText_(sheet, lastRowNow, reserveEnd)) {
    forceTimeColumnsAsText_(sheet, 2, reserveEnd - 1);
  }

  return sheet;
}

/* ============================ BACA ============================ */

function readSettingJamKerjaRows_() {
  const sheet = getOrCreateSettingJamKerjaSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  return sheet.getRange(2, 1, lastRow - 1, 5).getValues()
    .filter(row => String(row[0]).trim() !== '')
    .map(row => {
      const target = (row[4] === '' || row[4] === null) ? 8 : Number(row[4]);
      return {
        kodeShift: String(row[0]).trim(),
        label: String(row[1] || '').trim(),
        // Kolom C/D berformat teks, jadi normalnya string "06:00". Cabang Date
        // hanya cadangan untuk baris lama yang sempat tersimpan sebagai waktu
        // (nilainya sudah terlanjur salah; perlu diisi ulang sekali oleh user).
        jamMasuk: row[2] instanceof Date ? Utilities.formatDate(row[2], 'UTC', 'HH:mm') : String(row[2] || ''),
        jamPulang: row[3] instanceof Date ? Utilities.formatDate(row[3], 'UTC', 'HH:mm') : String(row[3] || ''),
        targetJam: isNaN(target) ? 8 : target
      };
    });
}

/** SUMBER KEBENARAN daftar kode shift: pengganti CONFIG.SHIFT_CODES/SHIFT_HOURS_TARGET. */
function getShiftCodesDynamic_() {
  const map = {};
  readSettingJamKerjaRows_().forEach((r, i) => {
    map[r.kodeShift] = {
      label: r.label || r.kodeShift,
      jamMasuk: r.jamMasuk,
      jamPulang: r.jamPulang,
      targetJam: r.targetJam,
      color: SHIFT_COLOR_PALETTE_[i % SHIFT_COLOR_PALETTE_.length]
    };
  });
  map['OFF'] = { label: 'Libur', jamMasuk: '', jamPulang: '', targetJam: 0, color: '#94a3b8' };
  return map;
}

function getJamKerjaSettings(token) {
  verifySession_(token);
  return readSettingJamKerjaRows_();
}

/* ============================ VALIDASI ============================ */

/** "6:00" -> "06:00"; kosong -> ''; selain itu error yang menyebut kodenya. */
function normalizeJamInput_(value, label, kode) {
  const s = String(value === null || value === undefined ? '' : value).trim();
  if (!s) return '';
  const m = s.match(TIME_HHMM_RE_);
  if (!m) throw new Error('Format ' + label + ' untuk shift ' + kode + ' tidak valid ("' + s + '"). Gunakan HH:mm, misal 06:00.');
  return (m[1].length === 1 ? '0' : '') + m[1] + ':' + m[2];
}

/** Target jam per hari: kosong -> 8; harus angka 0..24. */
function normalizeTargetJam_(value, kode) {
  if (value === '' || value === null || value === undefined) return 8;
  const n = Number(value);
  if (isNaN(n) || n < 0 || n > 24) throw new Error('Target jam untuk shift ' + kode + ' tidak valid ("' + value + '"). Isi angka 0 sampai 24.');
  return n;
}

/* ============================ TULIS ============================ */

/** Ubah jam masuk/pulang/target shift yang sudah ada (Super Admin). Semua-atau-tidak-sama-sekali. */
function setJamKerjaSettings(settings, token) {
  const session = requireSuperAdmin_(token, 'Hanya Super Admin yang boleh mengubah Setting Jam Kerja.');
  if (!Array.isArray(settings)) throw new Error('Data setting tidak valid.');

  // Validasi semua dulu sebelum menyentuh sheet.
  const updateMap = {};
  settings.forEach(s => {
    if (!s || !s.kodeShift) return;
    const kode = String(s.kodeShift).trim();
    updateMap[kode] = {
      jamMasuk: normalizeJamInput_(s.jamMasuk, 'jam masuk', kode),
      jamPulang: normalizeJamInput_(s.jamPulang, 'jam pulang', kode),
      targetJam: normalizeTargetJam_(s.targetJam, kode)
    };
  });

  const result = withScriptLock_(() => {
    const sheet = getOrCreateSettingJamKerjaSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return readSettingJamKerjaRows_();

    const kodeCol = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

    // Hanya baris yang diubah yang ditulis (baris lain tidak disentuh sama sekali).
    // Format teks dipasang dulu agar "06:00" tersimpan sebagai teks, bukan waktu.
    kodeCol.forEach((r, idx) => {
      const upd = updateMap[String(r[0]).trim()];
      if (!upd) return;
      const rowNum = idx + 2;
      sheet.getRange(rowNum, 3, 1, 2).setNumberFormat('@');
      sheet.getRange(rowNum, 3, 1, 3).setValues([[upd.jamMasuk, upd.jamPulang, upd.targetJam]]);
    });
    SpreadsheetApp.flush();
    return readSettingJamKerjaRows_();
  }, 15000);

  catatAuditLog_(session, 'UPDATE', 'Setting Jam Kerja', '-', JSON.stringify(settings));
  return result;
}

/** Tambah kode shift baru dari UI (Super Admin). */
function addShiftJamKerja(kodeShift, label, jamMasuk, jamPulang, targetJam, token) {
  const session = requireSuperAdmin_(token, 'Hanya Super Admin yang boleh menambah shift baru.');

  kodeShift = String(kodeShift || '').trim().toUpperCase();
  label = String(label || '').trim();

  if (!kodeShift) throw new Error('Kode shift wajib diisi.');
  if (!/^[A-Z0-9]{1,6}$/.test(kodeShift)) throw new Error('Kode shift hanya boleh huruf/angka, maksimal 6 karakter (misal: SR, N2).');
  if (kodeShift === 'OFF') throw new Error('Kode "OFF" sudah dipakai sistem untuk Libur dan tidak bisa ditambah manual.');
  if (!label) throw new Error('Label shift wajib diisi.');
  if (label.length > SHIFT_LABEL_MAX_) throw new Error('Label shift maksimal ' + SHIFT_LABEL_MAX_ + ' karakter.');

  const jamMasukOk = normalizeJamInput_(jamMasuk, 'jam masuk', kodeShift);
  const jamPulangOk = normalizeJamInput_(jamPulang, 'jam pulang', kodeShift);
  const targetOk = normalizeTargetJam_(targetJam, kodeShift);

  const result = withScriptLock_(() => {
    const sheet = getOrCreateSettingJamKerjaSheet_();
    const lastRow = sheet.getLastRow();
    const existingKode = lastRow >= 2
      ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r => String(r[0]).trim().toUpperCase())
      : [];
    if (existingKode.indexOf(kodeShift) !== -1) throw new Error('Kode shift "' + kodeShift + '" sudah ada.');

    // Format teks pada baris baru SEBELUM diisi, supaya jam tidak dikonversi jadi waktu.
    sheet.getRange(lastRow + 1, 3, 1, 2).setNumberFormat('@');
    sheet.appendRow([kodeShift, guardSheetText_(label, SHIFT_LABEL_MAX_), jamMasukOk, jamPulangOk, targetOk]);
    SpreadsheetApp.flush();
    return readSettingJamKerjaRows_();
  }, 15000);

  catatAuditLog_(session, 'Tambah Shift Baru', 'Setting Jam Kerja', kodeShift, 'Label: ' + label + ', Target: ' + targetOk + ' jam');
  return result;
}
