/*************************************************************
 * AMS v2 - Attendance Monitoring System
 * Departemen Housekeeping PT Swadharma Primautama
 *
 * 00_Config.gs
 * Semua konstanta global. Tidak ada fungsi bisnis di file ini.
 *
 * Urutan file (Apps Script memuat berdasarkan nama):
 *   00_Config      konstanta
 *   01_Router      doGet / doPost + tabel route
 *   02_SheetUtils  akses sheet & pembacaan grid
 *   03_Auth        login, session, hash password
 *   04_Access      user, role, hak akses (RBAC)
 *   1x_...         modul data (menyusul bertahap)
 *
 * SETUP AWAL (sekali): jalankan setupUsersSheet() dari editor.
 * RAHASIA (opsional tapi disarankan): isi di
 *   Project Settings > Script Properties
 *     LEGACY_PASSWORD_SALT, REPORT_EMAIL, GEMINI_API_KEY
 *************************************************************/

const CONFIG = {
  // ---- Nama sheet ----
  SHEET_JADWAL: 'Jadwal',
  SHEET_REALISASI: 'Realisasi',
  SHEET_DASHBOARD: 'Dashboard',
  SHEET_REKAP: 'Rekap',
  SHEET_USERS: 'Users',
  SHEET_JAMKERJA: 'JamKerja',
  SHEET_JAMSCAN: 'JamScan',
  SHEET_POLASHIFT: 'PolaShift',
  SHEET_DATARELIEVER: 'DataReliever',
  SHEET_LOGABSENRELIEVER: 'LogAbsenOnlineReliever',
  SHEET_BACKUPRELIEVER: 'BackupReliever',
  SHEET_ROLES: 'Roles',
  SHEET_ROLE_PERMISSIONS: 'RolePermissions',

  // ---- Layout grid Jadwal / Realisasi ----
  HEADER_ROW: 4,
  DAY_NAME_ROW: 5,
  DATA_START_ROW: 6,
  FIRST_DAY_COL: 4,

  // ---- Session & laporan ----
  SESSION_DURATION_SEC: 6 * 60 * 60, // 6 jam = batas maksimum CacheService
  REPORT_LOW_ATTENDANCE_THRESHOLD: 0.8,

  // Kode status REALISASI (kehadiran harian)
  STATUS_CODES: {
    '1':   { label: 'Hadir',                  color: '#22c55e' },
    '2':   { label: 'Lembur (Gantiin Shift)', color: '#c026d3' },
    'OFF': { label: 'Libur (OFF)',            color: '#94a3b8' },
    'I':   { label: 'Izin',                   color: '#3b82f6' },
    'S':   { label: 'Sakit',                  color: '#f59e0b' },
    'C':   { label: 'Cuti',                   color: '#a855f7' },
    'A':   { label: 'Alfa',                   color: '#ef4444' },
    'R':   { label: 'Resign',                 color: '#6b7280' },
    '':    { label: 'Belum diisi',            color: '#e5e7eb' }
  },

  // Kode SHIFT (jadwal). "MD" di sini = Midle, BUKAN Alfa.
  SHIFT_CODES: {
    'P':  { label: 'Pagi',    color: '#38bdf8' },
    'S':  { label: 'Siang',   color: '#fb923c' },
    'M':  { label: 'Malam',   color: '#818cf8' },
    'MD': { label: 'Midle',   color: '#facc15' },
    'WD': { label: 'Weekend', color: '#f472b6' },
    'OFF': { label: 'Libur',  color: '#94a3b8' }
  },

  // Target jam kerja per kode shift
  SHIFT_HOURS_TARGET: { 'P': 8, 'S': 8, 'M': 8, 'MD': 8, 'WD': 5, 'OFF': 0 },

  // Jam lembur >= nilai ini -> Realisasi ditandai "2"
  LEMBUR_MIN_HOURS: 5,

  // Whitelist sheet yang boleh diexport PDF
  PDF_ALLOWED_SHEETS: ['Rekap', 'Dashboard', 'Jadwal', 'Realisasi']
};

const NIK_SHEET_NAME = 'DataKaryawan';

const FINGER_REQUIRED_COLUMNS = [
  'No. ID', 'Nama', 'Tanggal', 'Jam Kerja',
  'Scan Masuk', 'Scan Pulang', 'Jml Jam Kerja', 'Lembur'
];

/* ---- Auth ---- */
const AUTH_HASH_ITERATIONS = 2000;
const AUTH_MAX_FAILED_ATTEMPTS = 5;
const AUTH_LOCKOUT_SEC = 15 * 60;
const USERS_COL = { USERNAME: 1, HASH: 2, NAMA: 3, ROLE: 4, SALT: 5, STATUS: 6 };

/* ---- Role ---- */
const SUPER_ADMIN_ROLE = 'Super Admin';
const LEGACY_ROLE_MAP = { 'admin': 'Super Admin', 'input': 'Staff' };

/* ---- Daftar menu (sumber kebenaran RBAC & sidebar) ---- */
const MENU_LIST = [
  { key: 'dashboard',          label: 'Dashboard' },
  { key: 'jadwal',             label: 'Jadwal Kerja' },
  { key: 'input',              label: 'Input Realisasi' },
  { key: 'rekap',              label: 'Rekap Karyawan' },
  { key: 'logabsensi',         label: 'Log Absensi' },
  { key: 'daftarkaryawan',     label: 'Daftar Karyawan' },
  { key: 'reportsummary',      label: 'Summary Report' },
  { key: 'reportdetail',       label: 'Detail Report' },
  { key: 'reportdept',         label: 'Department Report' },
  { key: 'reportexception',    label: 'Exception Report' },
  { key: 'settingaccount',     label: 'Setting - Account' },
  { key: 'settinghakakses',    label: 'Setting - Hak Akses' },
  { key: 'settingjamkerja',    label: 'Setting - Jam Kerja' },
  { key: 'settingjadwalkerja', label: 'Setting - Jadwal Kerja' },
  { key: 'settingassignjadwal', label: 'Setting - Assign Karyawan ke Jadwal' },
  { key: 'settinglibur',       label: 'Setting - Hari Libur' },
  { key: 'settingfingerprint', label: 'Setting - Integrasi Fingerprint' },
  { key: 'settingauditlog',    label: 'Setting - Audit Log' },
  { key: 'exportsheetpdf',     label: 'Export PDF Sheet (Rekap/Dashboard/Jadwal/Realisasi)' },
  { key: 'templatejadwal',     label: 'Download Template Jadwal' },
  { key: 'pengajuanizin',      label: 'Pengajuan Izin' }
];

const REPORT_TYPE_TO_MENU = {
  summary: 'reportsummary',
  detail: 'reportdetail',
  department: 'reportdept',
  exception: 'reportexception'
};

/* ---- Rahasia / pengaturan dari Script Properties ---- */
function getSetting_(key, fallback) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  return (v !== null && v !== '') ? v : (fallback === undefined ? '' : fallback);
}

/** Salt hash lama. Hanya dipakai untuk migrasi akun yang belum punya salt per-user. */
function getLegacyPasswordSalt_() {
  return getSetting_('LEGACY_PASSWORD_SALT', 'servanda-ventura-2026');
}

function getReportEmail_() {
  return getSetting_('REPORT_EMAIL', 'dimas.ardiyanto@wisma46.com');
}
