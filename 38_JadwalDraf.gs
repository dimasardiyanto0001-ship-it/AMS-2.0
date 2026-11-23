/*************************************************************
 * 38_JadwalDraf.gs
 * Draf Jadwal bulan depan: jadwal bulan berikutnya bisa diunggah SEBELUM periode diganti.
 * Disimpan di sheet tersembunyi "JadwalDraf" dan otomatis diterapkan ke Jadwal saat
 * Super Admin (HOD) menjalankan "Mulai Periode Baru" untuk bulan tersebut.
 * Upload: Super Admin & Admin. Penerapan terjadi hanya lewat aksi HOD, jadi HOD tetap penentu.
 *************************************************************/

const SHEET_JADWAL_DRAF_ = 'JadwalDraf';
const DRAF_COL_DAY1_ = 3;                              // kolom 1 = BulanKey, 2 = Nama, 3..33 = hari 1..31
const DRAF_COL_UPLOADER_ = DRAF_COL_DAY1_ + 31;        // 34
const DRAF_COL_WAKTU_ = DRAF_COL_UPLOADER_ + 1;        // 35
const DRAF_MAX_AHEAD_MONTHS_ = 12;

function getDrafSheet_(createIfMissing) {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(SHEET_JADWAL_DRAF_);
  if (!sh && createIfMissing) {
    sh = ss.insertSheet(SHEET_JADWAL_DRAF_);
    const headers = ['BulanKey', 'Nama'];
    for (let d = 1; d <= 31; d++) headers.push(String(d));
    headers.push('DiunggahOleh', 'Waktu');
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.getRange(1, 1, sh.getMaxRows(), 1).setNumberFormat('@'); // "2026-12" jangan dibaca sebagai tanggal
    sh.hideSheet();
  }
  return sh;
}

/** "2026-12" -> {key, year, monthIndex, label, dayCount} atau lempar error. */
function parseDrafBulanKey_(bulanKey) {
  const m = String(bulanKey || '').trim().match(/^(20\d{2})-(0[1-9]|1[0-2])$/);
  if (!m) throw new Error('Bulan tujuan tidak valid. Pilih bulan dari daftar.');
  const year = Number(m[1]);
  const monthIndex = Number(m[2]) - 1;
  return {
    key: m[1] + '-' + m[2], year: year, monthIndex: monthIndex,
    label: INDONESIAN_MONTH_NAMES_UPPER_[monthIndex] + ' ' + year,
    dayCount: new Date(year, monthIndex + 1, 0).getDate()
  };
}

function getActivePeriodMonth_() {
  const text = String(getSheet_(CONFIG.SHEET_JADWAL).getRange(2, 1).getValue() || '').trim();
  const my = parsePeriodMonthYear_(text);
  if (!my) throw new Error('Periode aktif ("' + text + '") tidak bisa dibaca bulan/tahunnya.');
  return my;
}

/** Bulan-bulan yang boleh dipilih sebagai tujuan draf (setelah periode aktif). */
function getDrafMonthOptions_() {
  const my = getActivePeriodMonth_();
  const out = [];
  for (let i = 1; i <= 6; i++) {
    const idx = my.year * 12 + my.monthIndex + i;
    const year = Math.floor(idx / 12), mi = idx % 12;
    out.push({ key: year + '-' + ('0' + (mi + 1)).slice(-2), label: INDONESIAN_MONTH_NAMES_UPPER_[mi] + ' ' + year });
  }
  return out;
}

function listDrafInternal_() {
  const sh = getDrafSheet_(false);
  if (!sh || sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, DRAF_COL_WAKTU_).getValues();
  const byKey = {};
  rows.forEach(r => {
    const key = String(r[0] || '').trim();
    if (!key) return;
    const e = byKey[key] || (byKey[key] = { key: key, jumlahKaryawan: 0, diunggahOleh: '', waktu: '' });
    e.jumlahKaryawan++;
    e.diunggahOleh = String(r[DRAF_COL_UPLOADER_ - 1] || '');
    const w = r[DRAF_COL_WAKTU_ - 1];
    e.waktu = w instanceof Date ? Utilities.formatDate(w, 'Asia/Jakarta', 'dd/MM/yyyy HH:mm') : String(w || '');
  });
  return Object.keys(byKey).sort().map(k => {
    const e = byKey[k];
    try { e.label = parseDrafBulanKey_(k).label; } catch (err) { e.label = k; }
    return e;
  });
}

/** Daftar draf + pilihan bulan tujuan untuk UI. */
function getJadwalDrafList(token) {
  verifySession_(token);
  return { draf: listDrafInternal_(), bulanTujuan: getDrafMonthOptions_() };
}

function hapusJadwalDraf(bulanKey, token) {
  const session = requireUploadAccess_(token, 'Hanya Super Admin atau Admin yang boleh menghapus draf jadwal.');
  const t = parseDrafBulanKey_(bulanKey);
  const n = deleteDrafRows_(t.key);
  catatAuditLog_(session, 'Hapus Draf Jadwal', SHEET_JADWAL_DRAF_, t.label, n + ' baris dihapus');
  return { draf: listDrafInternal_(), bulanTujuan: getDrafMonthOptions_() };
}

function deleteDrafRows_(key) {
  const sh = getDrafSheet_(false);
  if (!sh || sh.getLastRow() < 2) return 0;
  const keys = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  let deleted = 0;
  for (let i = keys.length - 1; i >= 0; i--) { // dari bawah supaya indeks tetap valid
    if (String(keys[i][0]).trim() === key) { sh.deleteRow(i + 2); deleted++; }
  }
  return deleted;
}

function uploadJadwalDraf(base64Data, filename, bulanKey, token) {
  const session = requireUploadAccess_(token, 'Hanya akun Super Admin atau Admin yang boleh upload jadwal bulan depan.');
  if (!base64Data) throw new Error('File tidak ditemukan / kosong.');
  const target = parseDrafBulanKey_(bulanKey);
  const active = getActivePeriodMonth_();
  const aheadIdx = (target.year * 12 + target.monthIndex) - (active.year * 12 + active.monthIndex);
  if (aheadIdx < 1) throw new Error('Bulan tujuan harus setelah periode aktif. Untuk periode aktif gunakan tombol Upload Jadwal.');
  if (aheadIdx > DRAF_MAX_AHEAD_MONTHS_) throw new Error('Bulan tujuan terlalu jauh (maks ' + DRAF_MAX_AHEAD_MONTHS_ + ' bulan ke depan).');

  const validCodes = getShiftCodesDynamic_();
  const info = withTempSpreadsheet_(base64Data, filename || 'draf_jadwal.xlsx', '_TEMP_DRAF_', MIME_XLSX_, UPLOAD_ERR_XLSX_, fileId => {
    return buildDrafRows_(readImportedGrid_(fileId, CONFIG.SHEET_JADWAL), validCodes, target.dayCount);
  });
  if (info.rows.length === 0) {
    throw new Error('Tidak ada karyawan di file yang cocok dengan daftar karyawan. Pastikan nama sama dengan di AMS.');
  }

  const sh = getDrafSheet_(true);
  const waktu = new Date();
  withScriptLock_(() => {
    deleteDrafRows_(target.key);                  // unggahan baru menggantikan draf bulan itu
    const width = DRAF_COL_WAKTU_;
    const values = info.rows.map(r => {
      const line = new Array(width).fill('');
      line[0] = target.key;
      line[1] = r.nama;
      r.days.forEach((c, i) => { line[DRAF_COL_DAY1_ - 1 + i] = c; });
      line[DRAF_COL_UPLOADER_ - 1] = session.nama || session.username || '';
      line[DRAF_COL_WAKTU_ - 1] = waktu;
      return line;
    });
    const start = Math.max(sh.getLastRow(), 1) + 1;
    sh.getRange(start, 1, values.length, 1).setNumberFormat('@');
    sh.getRange(start, 1, values.length, width).setValues(values);
    SpreadsheetApp.flush();
  }, 20000);

  catatAuditLog_(session, 'Upload Draf Jadwal', SHEET_JADWAL_DRAF_, target.label,
    filename + ': ' + info.rows.length + ' karyawan cocok, ' + info.notFound.length + ' nama tidak ditemukan, ' + info.invalidCodes.length + ' kode tidak valid');

  return {
    bulan: target.label,
    matched: info.rows.length,
    notFound: info.notFound,
    invalidCodes: info.invalidCodes,
    hariDipotong: info.hariDipotong,
    draf: listDrafInternal_(),
    bulanTujuan: getDrafMonthOptions_()
  };
}

/** Cocokkan baris file dengan karyawan di Jadwal; kode tidak valid dikosongkan dan dilaporkan. */
function buildDrafRows_(imported, validCodes, targetDayCount) {
  const sheet = getSheet_(CONFIG.SHEET_JADWAL);
  const numRows = getLastEmployeeRow_(sheet) - CONFIG.DATA_START_ROW + 1;
  const names = numRows > 0 ? sheet.getRange(CONFIG.DATA_START_ROW, 2, numRows, 1).getValues().map(r => String(r[0] || '').trim()) : [];
  const nameIdx = {};
  names.forEach((n, i) => { if (n) nameIdx[nameKey_(n)] = n; });

  const rows = [], notFound = [], invalidCodes = [], seen = {};
  imported.info.forEach((row, i) => {
    const nama = String(row[1] || '').trim();
    if (!nama) return;
    const canonical = nameIdx[nameKey_(nama)];
    if (!canonical) { notFound.push(nama); return; }
    if (seen[canonical]) return;                         // nama ganda di file: ambil baris pertama
    seen[canonical] = true;
    const days = [];
    for (let d = 0; d < targetDayCount; d++) {
      const raw = d < imported.dayCount ? imported.grid[i][d] : '';
      const kode = (raw === null || raw === undefined) ? '' : String(raw).trim().toUpperCase();
      if (kode === '' || validCodes.hasOwnProperty(kode)) days.push(kode);
      else { days.push(''); invalidCodes.push({ nama: nama, hariKe: d + 1, kode: kode }); }
    }
    rows.push({ nama: canonical, days: days });
  });
  return { rows: rows, notFound: notFound, invalidCodes: invalidCodes, hariDipotong: Math.max(0, imported.dayCount - targetDayCount) };
}

/**
 * Terapkan draf bulan (year, monthIndex) ke sheet Jadwal yang SUDAH direset ke periode tersebut.
 * Dipanggil dari mulaiPeriodeBaru. Draf yang diterapkan dihapus. Tidak melempar error ke pemanggil.
 */
function terapkanJadwalDrafUntukPeriode_(year, monthIndex, dayCount) {
  const key = year + '-' + ('0' + (monthIndex + 1)).slice(-2);
  const out = { ada: false, matched: 0, notFound: [] };
  try {
    const sh = getDrafSheet_(false);
    if (!sh || sh.getLastRow() < 2) return out;
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, DRAF_COL_WAKTU_).getValues()
      .filter(r => String(r[0]).trim() === key);
    if (!data.length) return out;
    out.ada = true;

    const jadwal = getSheet_(CONFIG.SHEET_JADWAL);
    const numRows = getLastEmployeeRow_(jadwal) - CONFIG.DATA_START_ROW + 1;
    if (numRows <= 0) return out;
    const names = jadwal.getRange(CONFIG.DATA_START_ROW, 2, numRows, 1).getValues().map(r => nameKey_(r[0] || ''));
    const width = Math.min(dayCount, 31);

    withScriptLock_(() => {
      const current = jadwal.getRange(CONFIG.DATA_START_ROW, CONFIG.FIRST_DAY_COL, numRows, width).getValues();
      const touched = {};
      data.forEach(r => {
        const nama = String(r[1] || '').trim();
        const idx = names.indexOf(nameKey_(nama));
        if (idx === -1) { out.notFound.push(nama); return; }
        current[idx] = r.slice(DRAF_COL_DAY1_ - 1, DRAF_COL_DAY1_ - 1 + width).map(v => String(v === null || v === undefined ? '' : v).trim().toUpperCase());
        touched[idx] = true;
        out.matched++;
      });
      Object.keys(touched).forEach(i => {
        jadwal.getRange(CONFIG.DATA_START_ROW + Number(i), CONFIG.FIRST_DAY_COL, 1, width).setValues([current[i]]);
      });
      SpreadsheetApp.flush();
    }, 20000);

    deleteDrafRows_(key);
  } catch (e) {
    Logger.log('Gagal menerapkan draf jadwal: ' + e.message);
    out.error = e.message;
  }
  return out;
}
