/*************************************************************
 * 40_NotifikasiJadwal.gs
 * Notifikasi email untuk alur "upload jadwal bulan depan" (draf):
 *   1. Pengingat ke Admin  : H-7 dan H-3 sebelum periode aktif berakhir, hanya untuk Admin
 *                            yang departemennya belum punya draf bulan depan.
 *   2. Konfirmasi ke HOD   : setiap kali Admin selesai upload draf.
 *   3. Ringkasan ke HOD    : H-1 (sekali), daftar departemen sudah / belum upload.
 *
 * Alamat email diisi di Setting > Account (kolom Email, sheet Users).
 * Pasang pemicu harian sekali dari editor: jalankan setupPengingatJadwalTrigger().
 * Uji coba kirim: jalankan ujiKirimNotifikasiJadwal() (kirim ke semua Super Admin).
 * Opsional: Script Property APP_URL = alamat aplikasi, ditampilkan sebagai tautan di email.
 *
 * Kegagalan kirim email TIDAK pernah menggagalkan upload; hanya dicatat di log.
 *************************************************************/

const NOTIF_REMIND_DAYS_ = [7, 3];      // pengingat ke Admin: sekian hari sebelum akhir periode
const NOTIF_SUMMARY_DAY_ = 1;           // ringkasan ke HOD: sekian hari sebelum akhir periode
const NOTIF_HOUR_ = 8;                  // jam kirim pemicu harian
const NOTIF_SENT_PROP_ = 'NOTIF_JADWAL_SENT';
const NOTIF_TZ_ = 'Asia/Jakarta';

/* ============================ UTIL ============================ */

function isValidEmail_(s) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s || '').trim());
}

/** Semua akun aktif dengan role & departemen & email (huruf asli). */
function listNotifUsers_() {
  const sheet = getSheet_(CONFIG.SHEET_USERS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, USERS_COL.EMAIL).getValues()
    .map(r => ({
      username: String(r[USERS_COL.USERNAME - 1] || '').trim(),
      nama: String(r[USERS_COL.NAMA - 1] || '').trim(),
      role: resolveRoleName_(r[USERS_COL.ROLE - 1]),
      status: String(r[USERS_COL.STATUS - 1] || 'Aktif'),
      department: String(r[USERS_COL.DEPARTMENT - 1] || '').trim(),
      email: String(r[USERS_COL.EMAIL - 1] || '').trim()
    }))
    .filter(u => u.username && u.status !== 'Nonaktif');
}

function notifHodEmails_(users) {
  return users.filter(u => isSuperAdminRole_(u.role) && isValidEmail_(u.email)).map(u => u.email);
}

/** Admin yang terikat departemen (yang bisa dipantau drafnya). */
function notifScopedAdmins_(users) {
  return users.filter(u => u.role === 'Admin' && u.department);
}

/** Himpunan deptKey_ departemen yang sudah punya baris draf untuk bulan tujuan. */
function drafDeptCoverage_(bulanKey) {
  const covered = {};
  const sh = getDrafSheet_(false);
  if (!sh || sh.getLastRow() < 2) return covered;
  const master = getKaryawanMasterMap_();
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => {
    if (String(r[0] || '').trim() !== bulanKey) return;
    const m = master[String(r[1] || '').trim().toLowerCase()];
    if (m && m.department) covered[deptKey_(m.department)] = true;
  });
  return covered;
}

/** Bulan tujuan draf (bulan setelah periode aktif) dan sisa hari sampai periode aktif berakhir. */
function notifTargetInfo_(now) {
  const active = getActivePeriodMonth_();
  const nextIdx = active.year * 12 + active.monthIndex + 1;
  const ny = Math.floor(nextIdx / 12), nm = nextIdx % 12;
  const key = ny + '-' + ('0' + (nm + 1)).slice(-2);
  const t = Utilities.formatDate(now, NOTIF_TZ_, 'yyyy-MM-dd').split('-').map(Number);
  const todayUtc = Date.UTC(t[0], t[1] - 1, t[2]);
  const endUtc = Date.UTC(active.year, active.monthIndex + 1, 0);        // hari terakhir bulan periode aktif
  return {
    key: key,
    label: INDONESIAN_MONTH_NAMES_UPPER_[nm] + ' ' + ny,
    activeLabel: INDONESIAN_MONTH_NAMES_UPPER_[active.monthIndex] + ' ' + active.year,
    daysLeft: Math.round((endUtc - todayUtc) / 86400000)
  };
}

function notifAlreadySent_(id) {
  try { return (PropertiesService.getScriptProperties().getProperty(NOTIF_SENT_PROP_) || '').split('\n').indexOf(id) !== -1; }
  catch (e) { return false; }
}

function notifMarkSent_(id) {
  const props = PropertiesService.getScriptProperties();
  const list = (props.getProperty(NOTIF_SENT_PROP_) || '').split('\n').filter(Boolean);
  list.push(id);
  props.setProperty(NOTIF_SENT_PROP_, list.slice(-40).join('\n'));   // simpan 40 terakhir saja
}

function notifSend_(to, subject, bodyHtml) {
  const brand = getBrandConfig_();
  const appUrl = getSetting_('APP_URL', '');
  const link = appUrl ? '<p><a href="' + brandEsc_(appUrl) + '" style="color:#2563eb;">Buka ' + brandEsc_(brand.appName) + '</a></p>' : '';
  const html = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#0f172a;max-width:560px;">' +
    bodyHtml + link +
    '<hr style="border:none;border-top:1px solid #e2e8f0;"><p style="color:#64748b;font-size:12px;">Email otomatis dari ' +
    brandEsc_(brand.appName) + ' - ' + brandEsc_(brand.companyName) + '. Jangan membalas email ini.</p></div>';
  MailApp.sendEmail({ to: to, subject: subject, htmlBody: html, name: brand.appName });
}

/* ============================ 1 & 3: PEMICU HARIAN ============================ */

/** Dipanggil pemicu harian. Aman dijalankan berulang (ada penanda "sudah terkirim"). */
function kirimNotifikasiJadwalHarian_() {
  const info = notifTargetInfo_(new Date());
  const users = listNotifUsers_();
  const hasil = { pengingat: 0, ringkasan: 0 };

  const covered = drafDeptCoverage_(info.key);
  const admins = notifScopedAdmins_(users);

  if (NOTIF_REMIND_DAYS_.indexOf(info.daysLeft) !== -1) {
    admins.forEach(a => {
      if (covered[deptKey_(a.department)] || !isValidEmail_(a.email)) return;
      const id = info.key + '|H' + info.daysLeft + '|' + a.username.toLowerCase();
      if (notifAlreadySent_(id)) return;
      try {
        notifSend_(a.email,
          'Pengingat: upload jadwal ' + info.label + ' (' + a.department + ')',
          '<p>Halo ' + brandEsc_(a.nama) + ',</p>' +
          '<p>Periode <b>' + brandEsc_(info.activeLabel) + '</b> berakhir dalam <b>' + info.daysLeft + ' hari</b>. ' +
          'Jadwal <b>' + brandEsc_(info.label) + '</b> untuk departemen <b>' + brandEsc_(a.department) + '</b> belum diupload.</p>' +
          '<p>Silakan upload lewat menu <b>Jadwal Kerja &rarr; Upload Jadwal &rarr; Jadwal bulan depan</b>.</p>');
        notifMarkSent_(id);
        hasil.pengingat++;
      } catch (e) { Logger.log('Gagal kirim pengingat ke ' + a.username + ': ' + e.message); }
    });
  }

  if (info.daysLeft <= NOTIF_SUMMARY_DAY_ && info.daysLeft >= -3) {
    const id = info.key + '|ringkasan';
    const hod = notifHodEmails_(users);
    if (!notifAlreadySent_(id) && hod.length) {
      try {
        notifSend_(hod.join(','), 'Ringkasan upload jadwal ' + info.label + ' sebelum Mulai Periode Baru',
          buildRingkasanDrafHtml_(info, admins, covered));
        notifMarkSent_(id);
        hasil.ringkasan = 1;
      } catch (e) { Logger.log('Gagal kirim ringkasan HOD: ' + e.message); }
    }
  }
  Logger.log('Notifikasi jadwal: ' + JSON.stringify(hasil) + ' (sisa ' + info.daysLeft + ' hari, target ' + info.key + ')');
  return hasil;
}

function buildRingkasanDrafHtml_(info, admins, covered) {
  const depts = {};
  admins.forEach(a => { (depts[deptKey_(a.department)] = depts[deptKey_(a.department)] || { nama: a.department, admins: [] }).admins.push(a.nama); });
  const keys = Object.keys(depts).sort();
  const sudah = keys.filter(k => covered[k]), belum = keys.filter(k => !covered[k]);
  const li = list => list.length
    ? '<ul>' + list.map(k => '<li>' + brandEsc_(depts[k].nama) + ' <span style="color:#64748b;">(' + brandEsc_(depts[k].admins.join(', ')) + ')</span></li>').join('') + '</ul>'
    : '<p style="color:#64748b;">-</p>';
  return '<p>Periode <b>' + brandEsc_(info.activeLabel) + '</b> berakhir ' + (info.daysLeft > 0 ? 'dalam ' + info.daysLeft + ' hari' : 'hari ini / sudah lewat') + '. ' +
    'Status draf jadwal <b>' + brandEsc_(info.label) + '</b>:</p>' +
    '<p><b style="color:#16a34a;">Sudah upload (' + sudah.length + ')</b></p>' + li(sudah) +
    '<p><b style="color:#dc2626;">Belum upload (' + belum.length + ')</b></p>' + li(belum) +
    '<p>Periksa draf lewat tombol <b>Lihat</b> di Jadwal Kerja, lalu tekan <b>Mulai Periode Baru</b> bila sudah sesuai.</p>';
}

/* ============================ 2: KONFIRMASI UPLOAD KE HOD ============================ */

/** Dipanggil dari uploadJadwalDraf. Tidak pernah melempar error. */
function notifUploadDrafKeHod_(session, target, jumlahKaryawan, filename) {
  try {
    const hod = notifHodEmails_(listNotifUsers_());
    if (!hod.length) return;
    // Jangan kirim ke pengunggah sendiri bila dia Super Admin.
    const self = listNotifUsers_().filter(u => u.username.toLowerCase() === String(session.username || '').toLowerCase())[0];
    const to = hod.filter(e => !(self && isSuperAdminRole_(self.role) && e === self.email));
    if (!to.length) return;
    const dept = session.department || '(semua departemen)';
    notifSend_(to.join(','), 'Jadwal ' + target.label + ' diupload: ' + dept,
      '<p><b>' + brandEsc_(session.nama || session.username) + '</b> mengupload draf jadwal.</p>' +
      '<table style="border-collapse:collapse;">' +
      '<tr><td style="padding:4px 12px 4px 0;color:#64748b;">Departemen</td><td><b>' + brandEsc_(dept) + '</b></td></tr>' +
      '<tr><td style="padding:4px 12px 4px 0;color:#64748b;">Bulan tujuan</td><td><b>' + brandEsc_(target.label) + '</b></td></tr>' +
      '<tr><td style="padding:4px 12px 4px 0;color:#64748b;">Karyawan</td><td><b>' + jumlahKaryawan + '</b></td></tr>' +
      '<tr><td style="padding:4px 12px 4px 0;color:#64748b;">File</td><td>' + brandEsc_(filename || '-') + '</td></tr></table>' +
      '<p>Periksa lewat tombol <b>Lihat</b> pada panel draf di Jadwal Kerja.</p>');
  } catch (e) {
    Logger.log('Gagal kirim konfirmasi upload ke HOD: ' + e.message);
  }
}

/* ============================ AKUN: EMAIL ============================ */

function setUserEmail(username, email, token) {
  const session = requireSuperAdmin_(token, 'Hanya Super Admin yang boleh mengatur email akun.');
  getUsersSheet_();
  username = String(username || '').trim();
  email = String(email || '').trim();
  if (email && !isValidEmail_(email)) throw new Error('Format email tidak valid.');

  withScriptLock_(() => {
    const sheet = getUsersSheet_();
    const found = findUserRow_(sheet, username);
    if (!found) throw new Error('Username tidak ditemukan.');
    sheet.getRange(found.rowNumber, USERS_COL.EMAIL).setValue(email);
    SpreadsheetApp.flush();
  });

  catatAuditLog_(session, 'Ubah Email User', 'Akun', username, 'Email: ' + (email || '(dikosongkan)'));
  return getUsersList(token);
}

/* ============================ PEMASANGAN & UJI (jalankan dari editor) ============================ */

function setupPengingatJadwalTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'kirimNotifikasiJadwalHarian_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('kirimNotifikasiJadwalHarian_').timeBased().everyDays(1).atHour(NOTIF_HOUR_).create();
  Logger.log('Pemicu pengingat jadwal terpasang: setiap hari jam ' + ('0' + NOTIF_HOUR_).slice(-2) + ':00.');
}

function ujiKirimNotifikasiJadwal() {
  const users = listNotifUsers_();
  const hod = notifHodEmails_(users);
  if (!hod.length) throw new Error('Belum ada Super Admin dengan email. Isi dulu di Setting > Account.');
  const info = notifTargetInfo_(new Date());
  notifSend_(hod.join(','), '[UJI] Notifikasi jadwal AMS',
    '<p>Ini email uji. Notifikasi jadwal berfungsi.</p><p>Periode aktif: <b>' + brandEsc_(info.activeLabel) +
    '</b>, sisa ' + info.daysLeft + ' hari. Bulan tujuan draf: <b>' + brandEsc_(info.label) + '</b>.</p>' +
    '<p>Admin terikat departemen: ' + notifScopedAdmins_(users).length + ' (dengan email: ' +
    notifScopedAdmins_(users).filter(a => isValidEmail_(a.email)).length + ').</p>');
  Logger.log('Email uji terkirim ke: ' + hod.join(', ') + ' | sisa kuota harian: ' + MailApp.getRemainingDailyQuota());
}
