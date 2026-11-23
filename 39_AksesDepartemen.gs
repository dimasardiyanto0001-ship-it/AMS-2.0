/*************************************************************
 * 39_AksesDepartemen.gs
 * Pembatasan akses per departemen. Setiap akun non-Super Admin boleh dikaitkan ke SATU departemen
 * (kolom "Departemen" di sheet Users). Akun itu hanya melihat dan mengubah karyawan departemennya.
 * Akun tanpa departemen (dan Super Admin/HOD) tidak dibatasi.
 *
 * Titik pusat: verifySession_() mengisi CURRENT_SCOPE_DEPT_ pada setiap request, sehingga
 * getAllDataInternal_, arsip, upload, draf, dst otomatis ikut terbatas.
 *************************************************************/

var CURRENT_SCOPE_DEPT_ = '';                 // departemen pembatas untuk request ini ('' = tidak dibatasi)
const SCOPE_CACHE_PREFIX_ = 'scope_dept_';
const SCOPE_CACHE_SEC_ = 60;

function deptKey_(s) {
  return String(s === null || s === undefined ? '' : s).replace(/\s+/g, ' ').trim().toLowerCase();
}

/** Departemen akun dari sheet Users (cache 60 dtk agar tidak membaca sheet tiap request). */
function readUserDepartment_(username) {
  const cache = CacheService.getScriptCache();
  const key = SCOPE_CACHE_PREFIX_ + String(username || '').toLowerCase();
  const cached = cache.get(key);
  if (cached !== null && cached !== undefined) return cached === '-' ? '' : cached;
  const found = findUserRow_(getSheet_(CONFIG.SHEET_USERS), username);
  const dept = found ? String(found.row[USERS_COL.DEPARTMENT - 1] || '').trim() : '';
  cache.put(key, dept || '-', SCOPE_CACHE_SEC_);
  return dept;
}

/** Departemen pembatas untuk sesi ini; '' untuk Super Admin atau akun tanpa departemen. */
function resolveScopeDept_(session) {
  if (!session || isSuperAdminRole_(session.role)) return '';
  return readUserDepartment_(session.username);
}

/** Departemen yang dipakai: pembatas akun bila ada, selain itu yang diminta klien. */
function effectiveDept_(requested) {
  return CURRENT_SCOPE_DEPT_ || requested;
}

function isScopedRequest_() {
  return !!CURRENT_SCOPE_DEPT_;
}

/** True bila karyawan boleh disentuh request ini. Karyawan yang tidak ada di master dianggap di luar scope. */
function nameInScope_(nama, master) {
  if (!CURRENT_SCOPE_DEPT_) return true;
  const m = (master || getKaryawanMasterMap_())[String(nama || '').trim().toLowerCase()];
  return !!m && deptKey_(m.department) === deptKey_(CURRENT_SCOPE_DEPT_);
}

/** Lempar error bila ada nama di luar departemen akun. */
function assertNamesInScope_(names) {
  if (!CURRENT_SCOPE_DEPT_) return;
  const master = getKaryawanMasterMap_();
  const out = [];
  (names || []).forEach(n => { if (!nameInScope_(n, master) && out.indexOf(n) === -1) out.push(n); });
  if (out.length) {
    throw new Error('Akses ditolak: karyawan berikut bukan dari departemen ' + CURRENT_SCOPE_DEPT_ + ': ' +
      out.slice(0, 5).join(', ') + (out.length > 5 ? ' (+' + (out.length - 5) + ' lainnya)' : '') + '.');
  }
}

/** Periksa nama karyawan pada baris-baris sheet grid (nomor baris sheet) terhadap departemen akun. */
function assertRowsInScope_(sheet, rowNumbers) {
  if (!CURRENT_SCOPE_DEPT_) return;
  const first = CONFIG.DATA_START_ROW;
  const last = getLastEmployeeRow_(sheet);
  if (last < first) return;
  const names = sheet.getRange(first, 2, last - first + 1, 1).getValues();
  const picked = [];
  rowNumbers.forEach(r => {
    const i = Number(r) - first;
    if (i >= 0 && i < names.length) picked.push(String(names[i][0] || '').trim());
  });
  assertNamesInScope_(picked);
}

/**
 * Saring hasil readImportedGrid_ ke karyawan departemen akun. Baris di luar departemen dibuang
 * dan namanya dikembalikan di `ditolak` supaya dilaporkan, bukan ditulis.
 */
function filterImportedByScope_(imported) {
  if (!CURRENT_SCOPE_DEPT_) return { imported: imported, ditolak: [] };
  const master = getKaryawanMasterMap_();
  const info = [], grid = [], ditolak = [];
  imported.info.forEach((row, i) => {
    const nama = String(row[1] || '').trim();
    if (!nama) return;
    if (nameInScope_(nama, master)) { info.push(row); grid.push(imported.grid[i]); }
    else if (master[nama.toLowerCase()]) ditolak.push(nama);   // ada di AMS tapi departemen lain
    // nama yang tidak dikenal sama sekali tetap lewat sebagai "tidak ditemukan"
    else { info.push(row); grid.push(imported.grid[i]); }
  });
  return { imported: { info: info, grid: grid, dayCount: imported.dayCount }, ditolak: ditolak };
}

/** Daftar departemen yang ada di master karyawan (untuk dropdown akun). */
function listDepartmentsInMaster_() {
  const master = getKaryawanMasterMap_();
  const seen = {}, out = [];
  Object.keys(master).forEach(k => {
    const d = String(master[k].department || '').trim();
    if (d && !seen[deptKey_(d)]) { seen[deptKey_(d)] = true; out.push(d); }
  });
  return out.sort();
}

function getDepartmentOptions(token) {
  requireMenuAccess_(token, 'settingaccount');
  return listDepartmentsInMaster_();
}

/** Atur departemen satu akun. Hanya Super Admin. Kosong = tidak dibatasi. */
function setUserDepartment(username, department, token) {
  const session = requireSuperAdmin_(token, 'Hanya Super Admin yang boleh mengatur departemen akun.');
  getUsersSheet_();
  username = String(username || '').trim();
  department = String(department || '').trim();

  let canonical = '';
  if (department && department !== 'ALL') {
    const hit = listDepartmentsInMaster_().filter(d => deptKey_(d) === deptKey_(department))[0];
    if (!hit) throw new Error('Departemen "' + department + '" tidak ada di Daftar Karyawan.');
    canonical = hit;
  }

  withScriptLock_(() => {
    const sheet = getUsersSheet_();
    const found = findUserRow_(sheet, username);
    if (!found) throw new Error('Username tidak ditemukan.');
    if (isSuperAdminRole_(resolveRoleName_(found.row[USERS_COL.ROLE - 1])) && canonical) {
      throw new Error('Akun Super Admin tidak dapat dibatasi ke satu departemen.');
    }
    sheet.getRange(found.rowNumber, USERS_COL.DEPARTMENT).setValue(canonical);
    SpreadsheetApp.flush();
  });

  CacheService.getScriptCache().remove(SCOPE_CACHE_PREFIX_ + username.toLowerCase());
  catatAuditLog_(session, 'Ubah Departemen User', 'Akun', username, 'Departemen: ' + (canonical || '(semua)'));
  return getUsersList(token);
}
