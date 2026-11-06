/*************************************************************
 * 37_SetupInstalasi.gs
 * Menyiapkan SALINAN spreadsheet menjadi instalasi bersih untuk pelanggan baru.
 *
 * PERINGATAN: siapkanInstalasiBaru() MENGHAPUS data karyawan, jadwal, realisasi, arsip, log,
 * dan akun. Jangan pernah dijalankan di spreadsheet produksi. Pengamannya: harus ada Script
 * Property IZIN_RESET_INSTALASI = YA (diisi manual di salinan). Properti dihapus otomatis setelah jalan.
 *
 * cekKesiapanInstalasi() aman dijalankan kapan saja (hanya membaca).
 *************************************************************/

const SETUP_GUARD_PROP_ = 'IZIN_RESET_INSTALASI';
const SETUP_LOG_SHEETS_DELETE_ = ['AuditLog', 'PengajuanIzin', 'PersetujuanJadwal', 'PersetujuanJadwalDetail', 'ArsipIndex'];
const SETUP_DATA_SHEETS_CLEAR_ = ['DataKaryawan', 'DataReliever', 'LogAbsenOnlineReliever', 'BackupReliever'];

function setupRandomPassword_() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 12) + 'A9';
}

/** Hapus baris karyawan pada satu sheet grid, sisakan 1 baris kosong supaya rumus tetap utuh. */
function setupClearEmployeeBlock_(sheet) {
  if (!sheet) return 0;
  const last = getLastEmployeeRow_(sheet);
  if (last < CONFIG.DATA_START_ROW) return 0;
  const count = last - CONFIG.DATA_START_ROW + 1;
  if (count > 1) sheet.deleteRows(CONFIG.DATA_START_ROW + 1, count - 1);
  const cols = Math.max(sheet.getLastColumn(), 3);
  const formulas = sheet.getRange(CONFIG.DATA_START_ROW, 1, 1, cols).getFormulas()[0];
  // kosongkan sel non-rumus (nomor, nama, jabatan, isi tanggal); rumus (mis. Hari Kerja) dipertahankan
  for (let c = 0; c < cols; c++) {
    if (!formulas[c]) sheet.getRange(CONFIG.DATA_START_ROW, c + 1).clearContent();
  }
  return count;
}

function siapkanInstalasiBaru() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(SETUP_GUARD_PROP_) !== 'YA') {
    throw new Error('Dibatalkan demi keamanan. Fungsi ini MENGHAPUS data. Jika ini memang SALINAN untuk pelanggan baru, ' +
      'isi Script Property ' + SETUP_GUARD_PROP_ + ' = YA (Project Settings > Script Properties), lalu jalankan lagi.');
  }
  const ss = getSpreadsheet_();
  const log = [];

  // 1. Arsip & log: hapus sheet
  ss.getSheets().forEach(sh => {
    const n = sh.getName();
    if (/^Arsip_/.test(n) || SETUP_LOG_SHEETS_DELETE_.indexOf(n) >= 0) {
      if (ss.getSheets().length > 1) { ss.deleteSheet(sh); log.push('hapus sheet ' + n); }
    }
  });

  // 2. Data master & log: kosongkan baris data, sisakan header (baris 1)
  SETUP_DATA_SHEETS_CLEAR_.forEach(n => {
    const sh = ss.getSheetByName(n);
    if (sh && sh.getLastRow() > 1) { sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(sh.getLastColumn(), 1)).clearContent(); log.push('kosongkan ' + n); }
  });

  // 3. Grid karyawan
  [CONFIG.SHEET_JADWAL, CONFIG.SHEET_REALISASI, CONFIG.SHEET_REKAP, CONFIG.SHEET_JAMKERJA, CONFIG.SHEET_JAMSCAN].forEach(n => {
    const sh = ss.getSheetByName(n);
    const jumlah = setupClearEmployeeBlock_(sh);
    if (sh) log.push('kosongkan ' + jumlah + ' baris karyawan di ' + n);
  });

  // 4. Lokasi absen: kembalikan placeholder
  const lok = ss.getSheetByName(SHEET_SETTING_LOKASI_ABSEN);
  if (lok) { lok.getRange(2, 1, 1, 3).setValues([[0, 0, 100]]); log.push('reset lokasi absen'); }

  // 5. Akun: ganti dengan satu Super Admin baru, password acak
  const users = ss.getSheetByName(CONFIG.SHEET_USERS);
  if (users) ss.deleteSheet(users);
  const password = setupRandomPassword_();
  const sheet = ss.insertSheet(CONFIG.SHEET_USERS);
  sheet.getRange(1, 1, 1, 6).setValues([['Username', 'PasswordHash', 'Nama', 'Role', 'Salt', 'Status']]).setFontWeight('bold');
  const salt = generateSalt_();
  sheet.getRange(2, 1, 1, 6).setValues([['admin', hashPasswordWithSalt_(password, salt), 'Administrator', SUPER_ADMIN_ROLE, salt, 'Aktif']]);
  sheet.setColumnWidths(1, 6, 160);
  log.push('akun admin dibuat ulang');

  // 6. Brand, rahasia lama, sesi
  props.deleteProperty(BRAND_PROP_KEY_);
  props.deleteProperty(SETUP_GUARD_PROP_);
  try { CacheService.getScriptCache().remove(BRAND_CACHE_KEY_); } catch (e) {}
  log.push('pengaturan perusahaan direset ke bawaan');

  Logger.log(log.join('\n'));
  Logger.log('=== LOGIN AWAL ===\nUsername: admin\nPassword: ' + password + '\nSimpan, lalu segera ganti password setelah login.');
  Logger.log('Langkah berikut: isi GEMINI_API_KEY, REPORT_EMAIL, LEGACY_PASSWORD_SALT (opsional) di Script Properties, lalu deploy ulang sebagai Web App.');
  return { log: log, username: 'admin', password: password };
}

/** Hanya membaca. Menampilkan apa saja yang masih perlu diisi pada instalasi ini. */
function cekKesiapanInstalasi() {
  const props = PropertiesService.getScriptProperties();
  const ss = getSpreadsheet_();
  const out = [];
  const ok = (cond, okText, todoText) => out.push((cond ? '[OK]    ' : '[ISI]   ') + (cond ? okText : todoText));

  ok(!!props.getProperty('GEMINI_API_KEY'), 'GEMINI_API_KEY terisi', 'GEMINI_API_KEY belum diisi (Asisten AI tidak aktif)');
  ok(!!props.getProperty('REPORT_EMAIL'), 'REPORT_EMAIL terisi', 'REPORT_EMAIL belum diisi (laporan bulanan tidak terkirim)');
  const brand = readBrandRaw_();
  ok(brand.companyName !== BRAND_DEFAULTS_.companyName, 'Nama perusahaan sudah diganti: ' + brand.companyName, 'Nama perusahaan masih bawaan (Setting > Perusahaan)');
  const lok = ss.getSheetByName(SHEET_SETTING_LOKASI_ABSEN);
  const v = lok ? lok.getRange(2, 1, 1, 2).getValues()[0] : [0, 0];
  ok(Number(v[0]) !== 0 || Number(v[1]) !== 0, 'Lokasi absen online sudah diisi', 'Lokasi absen online belum diisi (sheet tersembunyi SettingLokasiAbsen, baris 2)');
  const users = ss.getSheetByName(CONFIG.SHEET_USERS);
  ok(!!users && users.getLastRow() >= 2, 'Sheet Users ada', 'Sheet Users belum ada (jalankan setupUsersSheet atau siapkanInstalasiBaru)');
  const nik = ss.getSheetByName(NIK_SHEET_NAME);
  ok(!!nik && nik.getLastRow() >= 2, 'DataKaryawan berisi karyawan', 'DataKaryawan masih kosong (tambah karyawan / impor)');
  const hl = ss.getSheetByName(HARI_LIBUR_SHEET_);
  ok(!!hl && hl.getLastRow() >= 1, 'Sheet HariLibur ada', 'Sheet HariLibur belum dibuat (opsional; kolom A = tanggal libur nasional)');
  Logger.log(out.join('\n'));
  return out;
}
