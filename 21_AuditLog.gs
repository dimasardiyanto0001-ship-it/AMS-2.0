/*************************************************************
 * 21_AuditLog.gs
 * Audit Log: mencatat & membaca histori aksi admin/user.
 * Sheet "AuditLog": [Timestamp, Username, Nama, Role, Aksi, Modul, Target, Detail].
 *
 * Cara pakai dari modul lain:
 *   catatAuditLog_(session, 'Approve', 'Pengajuan Izin', id, 'Disetujui admin');
 * session = { username, nama, role } dari verifySession_(token).
 *************************************************************/

const AUDIT_LOG_SHEET_NAME_ = 'AuditLog';
const AUDIT_LOG_HEADER_ = ['Timestamp', 'Username', 'Nama', 'Role', 'Aksi', 'Modul', 'Target', 'Detail'];
const AUDIT_LOG_DEFAULT_LIMIT_ = 300;
const AUDIT_LOG_MAX_LIMIT_ = 5000;

/** Pastikan sheet AuditLog ada dan header-nya benar. Aman dipanggil berkali-kali. */
function ensureAuditLogSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(AUDIT_LOG_SHEET_NAME_);
  if (!sheet) {
    sheet = ss.insertSheet(AUDIT_LOG_SHEET_NAME_);
    sheet.appendRow(AUDIT_LOG_HEADER_);
    return sheet;
  }
  const firstRow = sheet.getRange(1, 1, 1, AUDIT_LOG_HEADER_.length).getValues()[0];
  const isHeaderValid = firstRow[0] === 'Timestamp' && firstRow[1] === 'Username';
  if (!isHeaderValid) {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, AUDIT_LOG_HEADER_.length).setValues([AUDIT_LOG_HEADER_]);
  }
  return sheet;
}

/**
 * Catat satu baris histori. Dibungkus try-catch: bila pencatatan gagal, aksi
 * utama pengguna (mis. simpan jadwal) TETAP berhasil dan hanya masuk Logger.
 * Teks dari pengguna (nama, catatan, JSON setting) dijaga agar tidak menjadi
 * rumus dan tidak melewati batas panjang sel.
 */
function catatAuditLog_(session, aksi, modul, target, detail) {
  try {
    const sheet = ensureAuditLogSheet_();
    sheet.appendRow([
      new Date(),
      guardSheetText_((session && session.username) || ''),
      guardSheetText_((session && session.nama) || ''),
      guardSheetText_((session && session.role) || ''),
      guardSheetText_(aksi || ''),
      guardSheetText_(modul || ''),
      guardSheetText_(target || ''),
      guardSheetText_(detail || '')
    ]);
  } catch (e) {
    Logger.log('Gagal mencatat audit log: ' + e.message);
  }
}

/** "yyyy-MM-dd" -> Date jam 00:00 WAKTU LOKAL script; null bila kosong; error bila format salah. */
function parseAuditFilterDate_(text, label) {
  if (!text) return null;
  const m = String(text).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) throw new Error('Format ' + label + ' tidak valid (harus yyyy-MM-dd).');
  return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
}

/**
 * Super Admin saja. filter = { modul: 'ALL'|string, tanggalMulai: 'yyyy-MM-dd'|null,
 * tanggalAkhir: 'yyyy-MM-dd'|null, limit: number }. Hasil diurutkan terbaru di atas.
 */
function getAuditLog(token, filter) {
  requireSuperAdmin_(token, 'Hanya Super Admin yang boleh melihat Audit Log.');

  filter = filter || {};
  const sheet = ensureAuditLogSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { logs: [], modulList: [], total: 0 };

  const data = sheet.getRange(2, 1, lastRow - 1, AUDIT_LOG_HEADER_.length).getValues();

  const tglMulai = parseAuditFilterDate_(filter.tanggalMulai, 'tanggal mulai');
  const tglAkhir = parseAuditFilterDate_(filter.tanggalAkhir, 'tanggal akhir');
  if (tglAkhir) tglAkhir.setHours(23, 59, 59, 999);

  const modulSet = {};
  const rows = [];

  data.forEach(function (r) {
    const ts = r[0], modul = r[5];

    if (modul) modulSet[modul] = true;
    if (filter.modul && filter.modul !== 'ALL' && modul !== filter.modul) return;
    if (tglMulai && ts < tglMulai) return;
    if (tglAkhir && ts > tglAkhir) return;

    rows.push({
      timestamp: ts instanceof Date ? ts.toISOString() : ts,
      username: r[1],
      nama: r[2],
      role: r[3],
      aksi: r[4],
      modul: modul,
      target: r[6],
      detail: r[7]
    });
  });

  rows.sort(function (a, b) { return new Date(b.timestamp) - new Date(a.timestamp); });

  const limit = Math.min(parseInt(filter.limit, 10) || AUDIT_LOG_DEFAULT_LIMIT_, AUDIT_LOG_MAX_LIMIT_);
  return {
    logs: rows.slice(0, limit),
    modulList: Object.keys(modulSet).sort(),
    total: rows.length
  };
}
