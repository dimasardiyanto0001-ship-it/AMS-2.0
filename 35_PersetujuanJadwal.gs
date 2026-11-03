/*************************************************************
 * 35_PersetujuanJadwal.gs
 * Alur persetujuan perubahan Jadwal:
 *   Supervisor (role "Admin") mengubah/mengupload/generate Jadwal
 *     -> perubahan TIDAK langsung ditulis, melainkan masuk antrean "Persetujuan Jadwal" (Pending);
 *   HOD (role "Super Admin") Approve -> perubahan baru ditulis ke sheet Jadwal; Reject -> dibuang.
 * Super Admin yang mengubah Jadwal tetap langsung berlaku (tercatat di Audit Log).
 *
 * Penyimpanan (sheet tersembunyi):
 *   PersetujuanJadwal       satu baris per pengajuan (status, pengaju, ringkasan)
 *   PersetujuanJadwalDetail isi perubahan dalam potongan JSON [[nama, hariKe(0..), kodeBaru, kodeLama], ...]
 * Sel dikenali lewat NAMA karyawan (bukan nomor baris), jadi tetap benar walau baris bergeser.
 *************************************************************/

const PJ_SHEET_ = 'PersetujuanJadwal';
const PJ_DETAIL_SHEET_ = 'PersetujuanJadwalDetail';
const PJ_HEADER_ = ['ID', 'Diajukan Pada', 'Username Pengaju', 'Nama Pengaju', 'Sumber', 'Periode', 'Jumlah Sel',
  'Jumlah Karyawan', 'Status', 'Diproses Oleh', 'Diproses Pada', 'Catatan', 'Ringkasan'];
const PJ_COL_ = { ID: 1, WAKTU: 2, USERNAME: 3, NAMA: 4, SUMBER: 5, PERIODE: 6, SEL: 7, KARYAWAN: 8, STATUS: 9, OLEH: 10, PADA: 11, CATATAN: 12, RINGKASAN: 13 };
const PJ_DETAIL_HEADER_ = ['ID', 'Seq', 'Data'];
const PJ_CHUNK_ = 40000;          // batas sel Sheets 50.000 karakter
const PJ_MAX_CELLS_ = 6000;       // sama dengan JADWAL_BATCH_MAX_
const PJ_STATUS_ = { PENDING: 'Pending', APPROVED: 'Approved', REJECTED: 'Rejected', CANCELLED: 'Dibatalkan' };

/* ============================ SHEET ============================ */

function getOrCreatePjSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(PJ_SHEET_);
  if (!sheet) {
    sheet = ss.insertSheet(PJ_SHEET_);
    sheet.getRange(1, 1, 1, PJ_HEADER_.length).setValues([PJ_HEADER_]).setFontWeight('bold').setBackground('#f1f5f9');
    sheet.hideSheet();
  }
  return sheet;
}

function getOrCreatePjDetailSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(PJ_DETAIL_SHEET_);
  if (!sheet) {
    sheet = ss.insertSheet(PJ_DETAIL_SHEET_);
    sheet.getRange(1, 1, 1, PJ_DETAIL_HEADER_.length).setValues([PJ_DETAIL_HEADER_]).setFontWeight('bold').setBackground('#f1f5f9');
    sheet.getRange(1, 3, sheet.getMaxRows(), 1).setNumberFormat('@');
    sheet.hideSheet();
  }
  return sheet;
}

function pjIso_(v) {
  return v instanceof Date ? v.toISOString() : String(v || '');
}

function pjRowToObject_(r) {
  return {
    id: String(r[PJ_COL_.ID - 1]),
    diajukanPada: pjIso_(r[PJ_COL_.WAKTU - 1]),
    username: String(r[PJ_COL_.USERNAME - 1] || ''),
    pengaju: String(r[PJ_COL_.NAMA - 1] || ''),
    sumber: String(r[PJ_COL_.SUMBER - 1] || ''),
    periode: String(r[PJ_COL_.PERIODE - 1] || ''),
    jumlahSel: Number(r[PJ_COL_.SEL - 1]) || 0,
    jumlahKaryawan: Number(r[PJ_COL_.KARYAWAN - 1]) || 0,
    status: String(r[PJ_COL_.STATUS - 1] || ''),
    diprosesOleh: String(r[PJ_COL_.OLEH - 1] || ''),
    diprosesPada: pjIso_(r[PJ_COL_.PADA - 1]),
    catatan: String(r[PJ_COL_.CATATAN - 1] || ''),
    ringkasan: String(r[PJ_COL_.RINGKASAN - 1] || '')
  };
}

/** Semua pengajuan (objek) + nomor baris sheet-nya. */
function readPjRows_() {
  const sheet = getOrCreatePjSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, PJ_HEADER_.length).getValues()
    .map((r, i) => ({ rowNum: i + 2, item: pjRowToObject_(r) }))
    .filter(x => x.item.id);
}

function findPjRow_(id) {
  id = String(id || '').trim();
  if (!id) throw new Error('ID pengajuan wajib diisi.');
  const found = readPjRows_().find(x => x.item.id === id);
  if (!found) throw new Error('Pengajuan jadwal tidak ditemukan.');
  return found;
}

function readPjChanges_(id) {
  const sheet = getOrCreatePjDetailSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const parts = sheet.getRange(2, 1, lastRow - 1, 3).getValues()
    .filter(r => String(r[0]) === id)
    .sort((a, b) => Number(a[1]) - Number(b[1]))
    .map(r => String(r[2]));
  if (!parts.length) return [];
  try {
    return JSON.parse(parts.join(''));
  } catch (e) {
    throw new Error('Isi pengajuan jadwal rusak (JSON tidak terbaca).');
  }
}

/* ============================ AJUKAN ============================ */

/** Peta nameKey -> {nama, rowIndex, days[]} dari sheet Jadwal saat ini. */
function pjCurrentJadwalMap_(jadwalSheet) {
  const map = {};
  readGridFromSheetObj_(jadwalSheet).employees.forEach(e => {
    const key = nameKey_(e.nama || '');
    if (key && !map[key]) map[key] = e;
  });
  return map;
}

/**
 * Simpan perubahan sebagai pengajuan Pending (TIDAK menulis ke sheet Jadwal).
 * changes: [{nama, day (0-based), kode}]. Sel yang nilainya sama dengan Jadwal sekarang dibuang.
 * Hasil: {menunggu:boolean, id, jumlahSel, jumlahKaryawan}. menunggu=false bila tidak ada perubahan nyata.
 */
function ajukanPerubahanJadwal_(session, changes, sumber, catatanRingkas) {
  const jadwalSheet = getSheet_(CONFIG.SHEET_JADWAL);
  const periode = String(jadwalSheet.getRange(2, 1).getValue() || '').trim();
  const current = pjCurrentJadwalMap_(jadwalSheet);
  const dayCount = getDayColCount_(jadwalSheet);
  const shiftCodes = getShiftCodesDynamic_();

  const uniq = {};
  (changes || []).forEach(c => {
    const emp = current[nameKey_(c.nama || '')];
    const day = Number(c.day);
    if (!emp || !isFinite(day) || day < 0 || day >= dayCount) return;
    const kode = normalizeKode_(c.kode);
    if (kode !== '' && !shiftCodes.hasOwnProperty(kode)) throw new Error('Kode shift tidak valid: ' + kode);
    const lama = String(emp.days[day] === null || emp.days[day] === undefined ? '' : emp.days[day]).trim().toUpperCase();
    if (lama === kode) return;                       // tidak ada perubahan nyata
    uniq[nameKey_(emp.nama) + '|' + day] = [String(emp.nama).trim(), day, kode, lama];
  });
  const items = Object.keys(uniq).map(k => uniq[k]);
  if (items.length === 0) return { menunggu: false, id: '', jumlahSel: 0, jumlahKaryawan: 0 };
  if (items.length > PJ_MAX_CELLS_) throw new Error('Terlalu banyak sel dalam satu pengajuan (maks ' + PJ_MAX_CELLS_ + ').');

  const karyawan = {};
  items.forEach(it => { karyawan[nameKey_(it[0])] = true; });
  const jumlahKaryawan = Object.keys(karyawan).length;

  const json = JSON.stringify(items);
  const id = 'PJD' + Date.now() + Math.floor(Math.random() * 1000);
  const ringkasan = (catatanRingkas || '').substring(0, 300);

  withScriptLock_(() => {
    const head = getOrCreatePjSheet_();
    const detail = getOrCreatePjDetailSheet_();
    head.appendRow([id, new Date(), session.username, session.nama || session.username, guardSheetText_(sumber, 120), periode,
      items.length, jumlahKaryawan, PJ_STATUS_.PENDING, '', '', '', guardSheetText_(ringkasan, 300)]);
    const rows = [];
    for (let i = 0, seq = 0; i < json.length; i += PJ_CHUNK_, seq++) rows.push([id, seq, json.substring(i, i + PJ_CHUNK_)]);
    detail.getRange(detail.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
    SpreadsheetApp.flush();
  }, 20000);

  catatAuditLog_(session, 'Ajukan Perubahan Jadwal', 'Jadwal', id,
    sumber + ': ' + items.length + ' sel pada ' + jumlahKaryawan + ' karyawan, menunggu persetujuan HOD');
  return { menunggu: true, id: id, jumlahSel: items.length, jumlahKaryawan: jumlahKaryawan };
}

/** Dari sel bernomor baris (editor Jadwal) ke pengajuan: updates [{rowIndex, dayIndex, kode}]. */
function ajukanJadwalDariSel_(session, updates, sumber) {
  const sheet = getSheet_(CONFIG.SHEET_JADWAL);
  const shiftCodes = getShiftCodesDynamic_();
  const items = validateCellUpdates_(sheet, updates, kode => {
    if (kode !== '' && !shiftCodes.hasOwnProperty(kode)) throw new Error('Kode shift tidak valid: ' + kode);
  });
  const lastRow = getLastEmployeeRow_(sheet);
  const names = sheet.getRange(CONFIG.DATA_START_ROW, 2, lastRow - CONFIG.DATA_START_ROW + 1, 1).getValues();
  const changes = items.map(it => ({
    nama: String(names[it.row - CONFIG.DATA_START_ROW][0] || '').trim(), day: it.day, kode: it.kode
  }));
  return ajukanPerubahanJadwal_(session, changes, sumber, summarizeJadwalBatch_(items).detail);
}

/* ============================ BACA (list, jumlah, detail) ============================ */

/** Daftar pengajuan sebagai JSON string. HOD melihat semua; Supervisor hanya miliknya. */
function getPersetujuanJadwalList(status, token) {
  const session = requireUploadAccess_(token, 'Hanya Super Admin atau Admin yang boleh membuka Persetujuan Jadwal.');
  const isHod = isSuperAdminRole_(session.role);
  status = String(status || 'Pending');
  const list = readPjRows_()
    .map(x => x.item)
    .filter(it => (isHod || it.username === session.username) && (status === 'ALL' || it.status === status))
    .sort((a, b) => (a.diajukanPada < b.diajukanPada ? 1 : -1));
  return JSON.stringify(list);
}

/** Jumlah Pending untuk badge menu (hanya untuk HOD; role lain 0). */
function getPersetujuanJadwalPendingCount(token) {
  const session = verifySession_(token);
  if (!isSuperAdminRole_(session.role)) return { count: 0 };
  return { count: readPjRows_().filter(x => x.item.status === PJ_STATUS_.PENDING).length };
}

/** Rincian perubahan dikelompokkan per karyawan: lama (saat diajukan), baru, dan nilai sekarang di Jadwal. */
function getPersetujuanJadwalDetail(id, token) {
  const session = requireUploadAccess_(token, 'Hanya Super Admin atau Admin yang boleh membuka Persetujuan Jadwal.');
  const found = findPjRow_(id);
  if (!isSuperAdminRole_(session.role) && found.item.username !== session.username) {
    throw new Error('Pengajuan ini bukan milik Anda.');
  }
  const changes = readPjChanges_(found.item.id);
  const current = pjCurrentJadwalMap_(getSheet_(CONFIG.SHEET_JADWAL));
  const group = {};
  changes.forEach(c => {
    const emp = current[nameKey_(c[0])];
    const sekarang = emp ? String(emp.days[c[1]] === null || emp.days[c[1]] === undefined ? '' : emp.days[c[1]]).trim().toUpperCase() : null;
    (group[c[0]] = group[c[0]] || []).push({ hari: c[1] + 1, lama: c[3], baru: c[2], sekarang: sekarang });
  });
  const perKaryawan = Object.keys(group).sort().map(n => ({
    nama: n, perubahan: group[n].sort((a, b) => a.hari - b.hari)
  }));
  return JSON.stringify({ header: found.item, perKaryawan: perKaryawan });
}

/* ============================ PUTUSAN ============================ */

function pjSetStatus_(rowNum, status, oleh, catatan) {
  const sheet = getOrCreatePjSheet_();
  sheet.getRange(rowNum, PJ_COL_.STATUS, 1, 4).setValues([[status, oleh, new Date(), guardSheetText_(catatan || '', 500)]]);
}

/** HOD: setujui -> tulis perubahan ke sheet Jadwal. */
function approvePersetujuanJadwal(id, token) {
  const session = requireSuperAdmin_(token, 'Hanya HOD (Super Admin) yang boleh menyetujui perubahan Jadwal.');
  let hasil = null;

  withScriptLock_(() => {
    const found = findPjRow_(id);
    if (found.item.status !== PJ_STATUS_.PENDING) throw new Error('Pengajuan ini sudah berstatus ' + found.item.status + '.');

    const jadwalSheet = getSheet_(CONFIG.SHEET_JADWAL);
    const periodeSekarang = String(jadwalSheet.getRange(2, 1).getValue() || '').trim();
    if (found.item.periode && periodeSekarang && found.item.periode !== periodeSekarang) {
      throw new Error('Periode Jadwal sudah berganti (' + found.item.periode + ' -> ' + periodeSekarang +
        '). Tolak pengajuan ini dan minta Supervisor mengajukan ulang.');
    }

    const changes = readPjChanges_(found.item.id);
    const current = pjCurrentJadwalMap_(jadwalSheet);
    const dayCount = getDayColCount_(jadwalSheet);
    const shiftCodes = getShiftCodesDynamic_();

    const items = [];
    let dilewati = 0;
    changes.forEach(c => {
      const emp = current[nameKey_(c[0])];
      const kode = normalizeKode_(c[2]);
      if (!emp || c[1] < 0 || c[1] >= dayCount || (kode !== '' && !shiftCodes.hasOwnProperty(kode))) { dilewati++; return; }
      items.push({ row: emp.rowIndex, day: c[1], kode: kode });
    });

    if (items.length) {
      writeGridCells_(jadwalSheet, items, kode => kode);
    }
    pjSetStatus_(found.rowNum, PJ_STATUS_.APPROVED, session.nama || session.username, '');
    SpreadsheetApp.flush();
    hasil = { diterapkan: items.length, dilewati: dilewati, sumber: found.item.sumber, pengaju: found.item.pengaju };
  }, 30000);

  // Di luar lock: syncOffFromJadwal_ memakai lock-nya sendiri.
  syncOffFromJadwal_();
  catatAuditLog_(session, 'Setujui Perubahan Jadwal', 'Jadwal', String(id),
    hasil.diterapkan + ' sel diterapkan, ' + hasil.dilewati + ' dilewati (pengaju: ' + hasil.pengaju + ', ' + hasil.sumber + ')');

  const result = getAllDataInternal_();
  result.persetujuanHasil = hasil;
  return result;
}

/** HOD: tolak (perubahan dibuang, Jadwal tidak berubah). */
function rejectPersetujuanJadwal(id, catatan, token) {
  const session = requireSuperAdmin_(token, 'Hanya HOD (Super Admin) yang boleh menolak perubahan Jadwal.');
  withScriptLock_(() => {
    const found = findPjRow_(id);
    if (found.item.status !== PJ_STATUS_.PENDING) throw new Error('Pengajuan ini sudah berstatus ' + found.item.status + '.');
    pjSetStatus_(found.rowNum, PJ_STATUS_.REJECTED, session.nama || session.username, catatan);
    SpreadsheetApp.flush();
  }, 15000);
  catatAuditLog_(session, 'Tolak Perubahan Jadwal', 'Jadwal', String(id), String(catatan || '').substring(0, 200));
  return { success: true };
}

/** Pengaju membatalkan pengajuannya sendiri selama masih Pending. */
function batalkanPersetujuanJadwal(id, token) {
  const session = requireUploadAccess_(token, 'Hanya Super Admin atau Admin yang boleh membatalkan pengajuan jadwal.');
  withScriptLock_(() => {
    const found = findPjRow_(id);
    if (found.item.username !== session.username && !isSuperAdminRole_(session.role)) throw new Error('Pengajuan ini bukan milik Anda.');
    if (found.item.status !== PJ_STATUS_.PENDING) throw new Error('Pengajuan ini sudah berstatus ' + found.item.status + '.');
    pjSetStatus_(found.rowNum, PJ_STATUS_.CANCELLED, session.nama || session.username, 'Dibatalkan pengaju');
    SpreadsheetApp.flush();
  }, 15000);
  catatAuditLog_(session, 'Batalkan Pengajuan Jadwal', 'Jadwal', String(id), '');
  return { success: true };
}
