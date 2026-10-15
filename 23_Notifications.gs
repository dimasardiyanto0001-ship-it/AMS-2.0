/*************************************************************
 * 23_Notifications.gs
 * Notifikasi lonceng: aksi terakhir di AuditLog + Pengajuan Izin
 * berstatus Pending.
 *
 * Perbaikan v2:
 *  - Kolom dibaca lewat NAMA header (bukan posisi tetap). Versi lama
 *    mengasumsikan [Waktu, User, Aksi, Modul, Detail], padahal AuditLog
 *    ditulis [Timestamp, Username, Nama, Role, Aksi, Modul, Target, Detail],
 *    sehingga judul notifikasi berisi Role, bukan Aksi/Modul.
 *  - Kolom tanggal pengajuan memakai 'DiajukanPada' (sebelumnya
 *    'TanggalPengajuan' yang tidak ada, jadi selalu jatuh ke waktu "sekarang").
 *  - Fungsi mati getNotifications_ (memanggil validateToken_ yang tidak ada)
 *    dihapus.
 *************************************************************/

const NOTIF_AUDIT_ROWS_ = 30;
const NOTIF_MAX_ITEMS_ = 50;

function getNotifications(token) {
  verifySession_(token);
  return getNotificationsInternal_();
}

function getNotificationsInternal_() {
  const ss = getSpreadsheet_();
  const items = [].concat(auditNotifications_(ss), pengajuanNotifications_(ss));
  items.sort((a, b) => b.timestamp - a.timestamp);
  return { success: true, notifications: items.slice(0, NOTIF_MAX_ITEMS_) };
}

function auditNotifications_(ss) {
  const sheet = ss.getSheetByName('AuditLog');
  if (!sheet || sheet.getLastRow() < 2) return [];

  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const col = name => header.indexOf(name);
  const iTime = col('Timestamp'), iUser = col('Username'), iAksi = col('Aksi'), iModul = col('Modul'), iDetail = col('Detail');

  const lastRow = sheet.getLastRow();
  const startRow = Math.max(2, lastRow - NOTIF_AUDIT_ROWS_ + 1);
  const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getValues();

  const out = [];
  data.forEach((row, i) => {
    if (iTime < 0 || !row[iTime]) return;
    const user = iUser >= 0 ? row[iUser] : '';
    const detail = iDetail >= 0 ? row[iDetail] : '';
    out.push({
      id: 'audit_' + (startRow + i),
      type: 'audit',
      title: (iModul >= 0 ? row[iModul] : '') + ' — ' + (iAksi >= 0 ? row[iAksi] : ''),
      message: (user ? user + ': ' : '') + (detail || ''),
      timestamp: new Date(row[iTime]).getTime()
    });
  });
  return out;
}

function pengajuanNotifications_(ss) {
  const sheet = ss.getSheetByName('PengajuanIzin');
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const iStatus = header.indexOf('Status');
  const iNama = header.indexOf('Nama');
  const iTgl = header.indexOf('DiajukanPada');
  if (iStatus < 0 || iNama < 0) return [];

  const out = [];
  for (let r = 1; r < data.length; r++) {
    if (data[r][iStatus] !== 'Pending' || !data[r][iNama]) continue;
    out.push({
      id: 'pengajuan_' + r,
      type: 'pengajuan',
      title: 'Pengajuan Izin Baru',
      message: data[r][iNama] + ' mengajukan izin',
      timestamp: iTgl >= 0 && data[r][iTgl] ? new Date(data[r][iTgl]).getTime() : Date.now()
    });
  }
  return out;
}
