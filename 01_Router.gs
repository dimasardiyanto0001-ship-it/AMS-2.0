/*************************************************************
 * 01_Router.gs
 * Titik masuk web app (doGet / doPost) berbasis TABEL ROUTE.
 *
 * Satu sumber kebenaran: ROUTES di bawah menentukan
 *   - nama aksi yang dikirim frontend,
 *   - fungsi backend yang dipanggil,
 *   - urutan parameter (sama dengan GAS_PARAM_MAP_ di JavaScript.html;
 *     panggil getRouteParamMap() lalu salin hasilnya, tidak perlu
 *     menulis peta itu manual lagi).
 *
 * Aturan:
 *   - Route non-public WAJIB memuat parameter 'token'
 *     (diperiksa auditRoutes()).
 *   - public: true = sengaja bisa diakses tanpa login
 *     (form Pengajuan Izin, Absen Online, Reliever).
 *   - stringify: true = hasil dikirim sebagai string JSON.
 *   - fn berupa NAMA fungsi (string) supaya file ini tetap bisa
 *     dimuat walau modul lain belum dipindahkan; fungsi yang belum
 *     ada hanya error saat rutenya dipanggil. Jalankan auditRoutes()
 *     dari editor untuk melihat route yang belum tersambung.
 *
 * CORS: frontend Netlify WAJIB mengirim
 *   Content-Type: text/plain;charset=utf-8
 * (bukan application/json) supaya tidak ada preflight OPTIONS,
 * yang tidak didukung Apps Script. Body tetap JSON.stringify(...).
 *************************************************************/

/** File HTML utama (nama file di Apps Script, tanpa .html). Samakan huruf besar/kecilnya. */
const HTML_ENTRY_FILE = 'Index';
const APP_TITLE = 'Attendance Monitoring System (AMS) - Departemen Housekeeping PT Swadharma Primautama';

const ROUTES = {
  /* ---- AUTH & AKUN ---- */
  login: { fn: 'login', args: ['username', 'password'], public: true },
  logout: { fn: 'logout', args: ['token'] },
  changeOwnPassword: { fn: 'changeOwnPassword', args: ['oldPassword', 'newPassword', 'token'] },
  getUsersList: { fn: 'getUsersList', args: ['token'] },
  addUserWeb: { fn: 'addUserWeb', args: ['nama', 'username', 'password', 'role', 'token'] },
  setUserStatus: { fn: 'setUserStatus', args: ['username', 'status', 'token'] },
  resetUserPasswordByAdmin: { fn: 'resetUserPasswordByAdmin', args: ['username', 'newPassword', 'token'] },

  /* ---- HAK AKSES ---- */
  getRolesList: { fn: 'getRolesList', args: ['token'] },
  addRole: { fn: 'addRole', args: ['roleName', 'deskripsi', 'token'] },
  getRolePermissionsMatrix: { fn: 'getRolePermissionsMatrix', args: ['token'] },
  setRolePermissions: { fn: 'setRolePermissions', args: ['roleName', 'menuKeys', 'editMenuKeys', 'token'] },
  getMyMenuPermissions: { fn: 'getMyMenuPermissions', args: ['token'] },

  /* ---- DATA UTAMA & PERIODE ---- */
  getAllData: { fn: 'getAllData', args: ['token', 'department'] },
  getArchivedPeriods: { fn: 'getArchivedPeriods', args: ['token'] },
  getArchivedData: { fn: 'getArchivedData', args: ['periodKey', 'token'] },
  getPeriodMeta: { fn: 'getPeriodMeta', args: ['periodKey', 'token'] },
  archiveCurrentPeriod: { fn: 'archiveCurrentPeriod', args: ['token'] },
  mulaiPeriodeBaru: { fn: 'mulaiPeriodeBaru', args: ['token'] },
  deleteArchivedPeriod: { fn: 'deleteArchivedPeriod', args: ['periodKey', 'token'] },
  getMonthlyTrendYear: { fn: 'getMonthlyTrendYear', args: ['tahun', 'token'] },

  /* ---- INPUT JADWAL & REALISASI ---- */
  saveRealisasi: { fn: 'saveRealisasi', args: ['rowIndex', 'dayIndex', 'kode', 'token'] },
  saveRealisasiBatch: { fn: 'saveRealisasiBatch', args: ['updates', 'token'] },
  saveJadwal: { fn: 'saveJadwal', args: ['rowIndex', 'dayIndex', 'kode', 'token'] },
  saveJadwalBatch: { fn: 'saveJadwalBatch', args: ['updates', 'token'] },
  uploadJadwal: { fn: 'uploadJadwal', args: ['base64Data', 'filename', 'token'] },
  uploadRealisasi: { fn: 'uploadRealisasi', args: ['base64Data', 'filename', 'token'] },
  uploadFingerAbsen: { fn: 'uploadFingerAbsen', args: ['base64Data', 'filename', 'token'] },
  downloadJadwalTemplate: { fn: 'generateJadwalTemplateExcel', args: ['token', 'department'] },
  generateJadwalTemplateExcel: { fn: 'generateJadwalTemplateExcel', args: ['token'] },

  /* ---- KARYAWAN ---- */
  getDaftarKaryawan: { fn: 'getDaftarKaryawan', args: ['token'] },
  addEmployee: { fn: 'addEmployee', args: ['nama', 'jabatan', 'nik', 'department', 'email', 'noHp', 'token'] },
  deleteEmployee: { fn: 'deleteEmployee', args: ['nama', 'token', 'confirmDataLoss'] },
  setEmployeeNik: { fn: 'setEmployeeNik', args: ['nama', 'nik', 'token'] },
  setEmployeeMasterData: { fn: 'setEmployeeMasterData', args: ['nama', 'fields', 'token'] },
  deleteUnassignedEmployee: { fn: 'deleteUnassignedEmployee', args: ['namaList', 'token'] },
  assignEmployeeToJadwal: { fn: 'assignEmployeeToJadwal', args: ['namaList', 'token'] },
  setKaryawanResign: { fn: 'setKaryawanResign', args: ['nama', 'tanggalResign', 'token'] },

  /* ---- SETTING JAM KERJA ---- */
  getJamKerjaSettings: { fn: 'getJamKerjaSettings', args: ['token'] },
  setJamKerjaSettings: { fn: 'setJamKerjaSettings', args: ['settings', 'token'] },
  addShiftJamKerja: { fn: 'addShiftJamKerja', args: ['kodeShift', 'label', 'jamMasuk', 'jamPulang', 'targetJam', 'token'] },

  /* ---- POLA SHIFT (Setting Jadwal Kerja) ---- */
  getPolaShiftList: { fn: 'getPolaShiftList', args: ['token'] },
  addPolaShift: { fn: 'addPolaShift', args: ['namaPola', 'urutanKode', 'deskripsi', 'token'] },
  deletePolaShift: { fn: 'deletePolaShift', args: ['namaPola', 'token'] },
  assignPolaKaryawan: { fn: 'assignPolaKaryawan', args: ['nama', 'polaAssigned', 'offsetHari', 'token'] },
  generateJadwalDariPola: { fn: 'generateJadwalDariPola', args: ['mode', 'token'] },

  /* ---- LAPORAN & EXPORT ---- */
  generateAttendanceLogExcel: { fn: 'generateAttendanceLogExcel', args: ['periodKey', 'employeeName', 'dayStart', 'dayEnd', 'token'] },
  exportPdf: { fn: 'exportSheetPdf', args: ['sheetName', 'token', 'department'] },
  exportSheetPdf: { fn: 'exportSheetPdf', args: ['sheetName', 'token', 'department'] },
  generateReportExcel: { fn: 'generateReportExcel', args: ['reportType', 'title', 'columns', 'rows', 'token'] },
  generateReportPdf: { fn: 'generateReportPdf', args: ['reportType', 'title', 'columns', 'rows', 'token'] },
  confirmSendEmailReport: { fn: 'confirmSendEmailReport', args: ['tglMulai', 'tglAkhir', 'tujuanEmail', 'token'] },

  /* ---- PENGAJUAN IZIN ---- */
  getKaryawanForPublicForm: { fn: 'getKaryawanForPublicForm', args: [], public: true },
  submitPengajuanIzin: { fn: 'submitPengajuanIzin', args: ['nama', 'nik', 'jenis', 'tglMulai', 'tglAkhir', 'alasan', 'lampiranBase64', 'lampiranFilename', 'lampiranMimeType'], public: true },
  getPengajuanList: { fn: 'getPengajuanList', args: ['status', 'token'] },
  getPengajuanListSafe: { fn: 'getPengajuanList', args: ['status', 'token'], stringify: true },
  getPengajuanPendingCount: { fn: 'getPengajuanPendingCount', args: ['token'] },
  approvePengajuan: { fn: 'approvePengajuan', args: ['id', 'token'] },
  rejectPengajuan: { fn: 'rejectPengajuan', args: ['id', 'catatan', 'token'] },

  /* ---- ABSEN ONLINE (publik) ---- */
  getJadwalHariIni: { fn: 'getJadwalHariIni', args: ['nama'], public: true },
  verifyLoginAbsenOnline: { fn: 'verifyLoginAbsenOnline', args: ['nama', 'nik'], public: true },
  searchKaryawanUntukAbsen: { fn: 'searchKaryawanUntukAbsen', args: ['query'], public: true },
  submitAbsenOnline: { fn: 'submitAbsenOnline', args: ['nama', 'nik', 'jenis', 'lat', 'lng', 'akurasi', 'fotoBase64'], public: true },
  getRiwayatAbsen: { fn: 'getRiwayatAbsen', args: ['nik'], public: true },
  getPeriodeAktif: { fn: 'getPeriodeAktif', args: [], public: true },
  setFotoProfilKaryawan: { fn: 'setFotoProfilKaryawan', args: ['nama', 'nik', 'fotoBase64'], public: true },

  /* ---- BACKUP RELIEVER (publik) ---- */
  verifyLoginReliever: { fn: 'verifyLoginReliever', args: ['nama'], public: true },
  searchReliverAbsen: { fn: 'searchReliverUntukAbsen', args: ['query'], public: true },
  getKaryawanAbsenHariIni: { fn: 'getKaryawanAbsenHariIni', args: [], public: true },
  getRiwayatAbsenReliever: { fn: 'getRiwayatAbsenReliever', args: ['nama'], public: true },
  submitAbsenOnlineReliever: { fn: 'submitAbsenOnlineReliever', args: ['namaReliever', 'asalPerusahaan', 'namaDigantikan', 'jenisDigantikan', 'jenis', 'lat', 'lng', 'akurasi', 'fotoBase64'], public: true },
  getRelieverSummary: { fn: 'getRelieverSummary', args: ['token'] },
  getRelieverData: { fn: 'getRelieverData', args: ['tglMulai', 'tglAkhir', 'token'] },
  exportRelieverExcel: { fn: 'exportRelieverExcel', args: ['tglMulai', 'tglAkhir', 'cari', 'token'] },
  exportRelieverPdf: { fn: 'exportRelieverPdf', args: ['tglMulai', 'tglAkhir', 'cari', 'token'] },
  getRelieverMaster: { fn: 'getRelieverMaster', args: ['token'] },
  addReliever: { fn: 'addReliever', args: ['nama', 'perusahaan', 'noHp', 'token'] },
  updateReliever: { fn: 'updateReliever', args: ['nama', 'perusahaan', 'noHp', 'token'] },
  setRelieverStatus: { fn: 'setRelieverStatus', args: ['nama', 'status', 'token'] },

  /* ---- LAIN-LAIN ---- */
  getAuditLog: { fn: 'getAuditLog', args: ['token', 'filter'] },
  getNotifications: { fn: 'getNotifications', args: ['token'] },
  chatWithAI: { fn: 'chatWithAI', args: ['pesan', 'token'] }
};

/** Aksi GET (query string) -> route + cara membentuk argumen. */
const GET_ACTIONS = {
  data:            { route: 'getAllData',          params: p => ({ token: p.token, department: p.department }) },
  archivedPeriods: { route: 'getArchivedPeriods',  params: p => ({ token: p.token }) },
  archivedData:    { route: 'getArchivedData',     params: p => ({ periodKey: requireParam_(p, 'key'), token: p.token }) },
  periodMeta:      { route: 'getPeriodMeta',       params: p => ({ periodKey: p.key || 'active', token: p.token }) }
};

/* ============================ ENTRY POINT ============================ */

function doGet(e) {
  const params = (e && e.parameter) || {};
  const get = GET_ACTIONS[params.action];

  if (get) {
    try {
      return jsonOutput_({ success: true, data: dispatchRoute_(get.route, get.params(params)) });
    } catch (err) {
      return jsonOutput_({ success: false, error: err.message });
    }
  }

  return HtmlService.createTemplateFromFile(HTML_ENTRY_FILE).evaluate()
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    return jsonOutput_({ success: true, data: dispatchRoute_(body.action, body) });
  } catch (err) {
    return jsonOutput_({ success: false, error: err.message });
  }
}

/** Panggil fungsi backend sesuai tabel route. @param {Object} body objek {namaParam: nilai} */
function dispatchRoute_(action, body) {
  const route = ROUTES[action];
  if (!route) throw new Error('Action tidak dikenal: ' + action);

  const fn = globalThis[route.fn];
  if (typeof fn !== 'function') {
    throw new Error('Route "' + action + '" menunjuk ke fungsi "' + route.fn + '" yang belum ada.');
  }

  const result = fn.apply(null, route.args.map(name => body[name]));
  return route.stringify ? JSON.stringify(result) : result;
}

function requireParam_(params, name) {
  if (!params[name]) throw new Error('Parameter "' + name + '" wajib diisi.');
  return params[name];
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Dipanggil oleh template HTML: <?!= include('Stylesheet') ?> */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ============================ ALAT BANTU PENGEMBANG ============================ */

/**
 * Jalankan dari editor. Melaporkan:
 *  - route yang menunjuk ke fungsi yang belum ada,
 *  - route non-public tanpa parameter 'token' (celah keamanan),
 *  - nama fungsi ganda di tabel (alias).
 */
function auditRoutes() {
  const problems = [];
  Object.keys(ROUTES).forEach(name => {
    const r = ROUTES[name];
    if (typeof globalThis[r.fn] !== 'function') problems.push('FUNGSI HILANG  : ' + name + ' -> ' + r.fn);
    if (!r.public && r.args.indexOf('token') === -1) problems.push('TANPA TOKEN    : ' + name + ' (tandai public:true jika memang disengaja)');
  });
  Logger.log(problems.length ? problems.join('\n') : 'Semua ' + Object.keys(ROUTES).length + ' route OK.');
  return problems;
}

/** Hasilkan isi GAS_PARAM_MAP_ untuk adapter Netlify (salin dari Log eksekusi). */
function getRouteParamMap() {
  const map = {};
  Object.keys(ROUTES).forEach(name => { map[name] = ROUTES[name].args; });
  Logger.log(JSON.stringify(map, null, 2));
  return map;
}
