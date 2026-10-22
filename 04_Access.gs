/*************************************************************
 * 04_Access.gs
 * Role, kelola user (Setting > Account), dan hak akses granular
 * (Setting > Hak Akses). Sheet: Users, Roles, RolePermissions.
 *
 * Role sistem: Super Admin (akses penuh) / Admin / Staff.
 * Role lama "admin"/"input" otomatis dipetakan ke Super Admin/Staff.
 *************************************************************/

/* ============================ HELPER ROLE ============================ */

const ZERO_WIDTH_RE_ = /[\u200B\u200C\u200D\uFEFF\u00A0]/g;

/** Kunci pembanding role: tanpa karakter tak terlihat, spasi dirapikan, huruf kecil. */
function normalizeRoleKey_(name) {
  return String(name || '').replace(ZERO_WIDTH_RE_, '').replace(/\s+/g, ' ').trim().toLowerCase();
}

/** Alias role lama -> role baru, atau null jika bukan alias (case-SENSITIVE: "Admin" bukan "admin"). */
function resolveLegacyRole_(rawRole) {
  const cleaned = String(rawRole || '').replace(ZERO_WIDTH_RE_, '').trim();
  return Object.prototype.hasOwnProperty.call(LEGACY_ROLE_MAP, cleaned) ? LEGACY_ROLE_MAP[cleaned] : null;
}

function resolveRoleName_(rawRole) {
  const trimmed = String(rawRole || '').trim();
  return resolveLegacyRole_(trimmed) || trimmed;
}

function isSuperAdminRole_(rawRole) {
  return normalizeRoleKey_(resolveRoleName_(rawRole)) === normalizeRoleKey_(SUPER_ADMIN_ROLE);
}

function requireSuperAdmin_(token, pesanError) {
  const session = verifySession_(token);
  if (!isSuperAdminRole_(session.role)) throw new Error(pesanError);
  return session;
}

/** True jika role = "Admin" (supervisor). Role lama huruf kecil "admin" sudah dipetakan ke Super Admin. */
function isAdminRole_(rawRole) {
  return normalizeRoleKey_(resolveRoleName_(rawRole)) === 'admin';
}

/** Upload realisasi / finger: boleh Super Admin dan Admin (supervisor). */
function requireUploadAccess_(token, pesanError) {
  const session = verifySession_(token);
  if (isSuperAdminRole_(session.role) || isAdminRole_(session.role)) return session;
  throw new Error(pesanError);
}

/**
 * Untuk fungsi yang HANYA boleh dijalankan pemilik script dari editor Apps Script
 * (setup, debug, perbaikan data). Fungsi global tanpa garis bawah di akhir nama bisa
 * dipanggil siapa pun yang membuka web app lewat google.script.run, jadi fungsi seperti
 * itu wajib memanggil ini di baris pertama. Pengunjung web app (anonim) punya
 * getActiveUser() kosong, sedangkan dari editor keduanya sama dengan email pemilik.
 */
function isScriptOwnerRun_() {
  const active = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  const owner = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  return !!active && !!owner && active === owner;
}

/** Lempar error bila pemanggil bukan pemilik script (lihat penjelasan di atas). */
function requireScriptOwner_() {
  const active = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  const owner = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  if (!active || !owner || active !== owner) {
    throw new Error('Fungsi ini hanya boleh dijalankan pemilik script dari editor Apps Script.');
  }
}

/* ============================ SHEET ROLES & PERMISSIONS ============================ */

const ROLE_DESC_ = {
  'Super Admin': 'Akses penuh ke seluruh sistem (tidak bisa diubah)',
  'Admin': 'Akses administratif - hak akses diatur manual lewat matrix',
  'Staff': 'Akses operasional harian'
};

function mapColumnA_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const range = sheet.getRange(2, 1, lastRow - 1, 1);
  const vals = range.getValues();
  let changed = false;
  vals.forEach(r => {
    const mapped = resolveLegacyRole_(r[0]);
    if (mapped) { r[0] = mapped; changed = true; }
  });
  if (changed) range.setValues(vals);
}

function dedupeRows_(sheet, numCols, keyFn) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  const seen = new Set();
  const out = [];
  data.forEach(row => {
    const key = keyFn(row);
    if (!key || seen.has(key)) return;
    seen.add(key);
    out.push(row);
  });
  if (out.length !== data.length) {
    sheet.getRange(2, 1, lastRow - 1, numCols).clearContent();
    if (out.length) sheet.getRange(2, 1, out.length, numCols).setValues(out);
  }
}

function createRolesSheet_(ss) {
  const sheet = ss.insertSheet(CONFIG.SHEET_ROLES);
  sheet.getRange(1, 1, 1, 2).setValues([['RoleName', 'Deskripsi']]).setFontWeight('bold');
  Object.keys(ROLE_DESC_).forEach(name => sheet.appendRow([name, ROLE_DESC_[name]]));
  sheet.hideSheet();
  return sheet;
}

function createPermissionsSheet_(ss) {
  const sheet = ss.insertSheet(CONFIG.SHEET_ROLE_PERMISSIONS);
  sheet.getRange(1, 1, 1, 4).setValues([['RoleName', 'MenuKey', 'CanView', 'CanEdit']]).setFontWeight('bold');
  const rows = [];
  MENU_LIST.forEach(m => {
    const isSetting = m.key.indexOf('setting') === 0;
    rows.push([SUPER_ADMIN_ROLE, m.key, true, true]);
    rows.push(['Admin', m.key, false, false]);
    rows.push(['Staff', m.key, !isSetting, !isSetting]);
  });
  sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  sheet.hideSheet();
  return sheet;
}

/**
 * Pastikan sheet Roles & RolePermissions ada, termigrasi, bebas duplikat, dan
 * punya baris untuk setiap kombinasi role x menu. Dibungkus lock.
 * @return {{rolesSheet: Sheet, permSheet: Sheet}}
 */
function ensureRolesSheets_() {
  const ss = getSpreadsheet_();
  return withScriptLock_(() => {
    let rolesSheet = ss.getSheetByName(CONFIG.SHEET_ROLES);
    let permSheet = ss.getSheetByName(CONFIG.SHEET_ROLE_PERMISSIONS);
    const permIsNew = !permSheet;

    if (!rolesSheet) {
      rolesSheet = createRolesSheet_(ss);
    } else {
      mapColumnA_(rolesSheet);
      dedupeRows_(rolesSheet, 2, row => normalizeRoleKey_(row[0]));
      const names = readColumn_(rolesSheet, 1).map(normalizeRoleKey_);
      if (names.indexOf('admin') === -1) rolesSheet.appendRow(['Admin', ROLE_DESC_['Admin']]);
    }

    if (!permSheet) {
      permSheet = createPermissionsSheet_(ss);
    } else {
      mapColumnA_(permSheet);
      if (permSheet.getRange(1, 4).getValue() !== 'CanEdit') {
        permSheet.getRange(1, 4).setValue('CanEdit').setFontWeight('bold');
        const last = permSheet.getLastRow();
        if (last >= 2) permSheet.getRange(2, 4, last - 1, 1).setValues(permSheet.getRange(2, 3, last - 1, 1).getValues());
      }
    }

    if (!permIsNew) {
      dedupeRows_(permSheet, 4, row => {
        const role = normalizeRoleKey_(row[0]);
        const menu = String(row[1] || '').trim();
        return role && menu ? role + '||' + menu : '';
      });
      fillMissingPermissionRows_(rolesSheet, permSheet);
    }
    return { rolesSheet, permSheet };
  });
}

function readColumn_(sheet, col) {
  const last = sheet.getLastRow();
  return last >= 2 ? sheet.getRange(2, col, last - 1, 1).getValues().map(r => String(r[0])) : [];
}

function fillMissingPermissionRows_(rolesSheet, permSheet) {
  const last = permSheet.getLastRow();
  const existing = last >= 2 ? permSheet.getRange(2, 1, last - 1, 2).getValues() : [];
  const have = new Set(existing.map(r => r[0] + '||' + r[1]));
  const roleNames = Array.from(new Set(rolesSheet.getLastRow() >= 2 ? readColumn_(rolesSheet, 1) : existing.map(r => String(r[0]))));
  const missing = [];
  roleNames.forEach(role => MENU_LIST.forEach(m => {
    if (!have.has(role + '||' + m.key)) missing.push([role, m.key, false, false]);
  }));
  if (missing.length) permSheet.getRange(permSheet.getLastRow() + 1, 1, missing.length, 4).setValues(missing);
}

function readPermissionRows_() {
  const sheets = ensureRolesSheets_();
  const last = sheets.permSheet.getLastRow();
  return { sheets, rows: last >= 2 ? sheets.permSheet.getRange(2, 1, last - 1, 4).getValues() : [] };
}

/* ============================ CEK AKSES ============================ */

function checkMenuPermission_(token, menuKey, colIndex, label) {
  const session = verifySession_(token);
  if (isSuperAdminRole_(session.role)) return session;

  const roleName = resolveRoleName_(session.role);
  const row = readPermissionRows_().rows.find(r => String(r[0]) === roleName && String(r[1]) === menuKey);
  if (!(row && row[colIndex])) {
    throw new Error('Role "' + roleName + '" tidak punya izin ' + label + ' fitur ini (menu: ' + menuKey +
      '). Hubungi Super Admin untuk mengaktifkan lewat Setting > Hak Akses.');
  }
  return session;
}

/** Wajib punya izin LIHAT menu tersebut. */
function requireMenuAccess_(token, menuKey) {
  return checkMenuPermission_(token, menuKey, 2, 'akses');
}

/** Wajib punya izin EDIT menu tersebut (aksi tulis). */
function requireMenuEditAccess_(token, menuKey) {
  return checkMenuPermission_(token, menuKey, 3, 'EDIT untuk');
}

function requireReportAccess_(reportType, token) {
  return requireMenuAccess_(token, REPORT_TYPE_TO_MENU[reportType] || 'reportsummary');
}

function getMyMenuPermissions(token) {
  const session = verifySession_(token);
  if (isSuperAdminRole_(session.role)) {
    const all = MENU_LIST.map(m => m.key);
    return { view: all, edit: all };
  }
  const roleName = resolveRoleName_(session.role);
  const rows = readPermissionRows_().rows.filter(r => String(r[0]) === roleName);

  if (rows.length === 0) {
    const fallback = MENU_LIST.filter(m => m.key.indexOf('setting') !== 0).map(m => m.key);
    return { view: fallback, edit: fallback };
  }
  return {
    view: rows.filter(r => !!r[2]).map(r => String(r[1])),
    edit: rows.filter(r => !!r[3]).map(r => String(r[1]))
  };
}

/* ============================ SETTING > ACCOUNT ============================ */

function getUsersList(token) {
  const sheet = getUsersSheet_();
  requireMenuAccess_(token, 'settingaccount');
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, 6).getValues().map(r => ({
    username: r[0], nama: r[2], role: resolveRoleName_(r[3]), status: r[5] || 'Aktif'
  }));
}

function addUserWeb(nama, username, password, role, token) {
  getUsersSheet_();
  const session = requireSuperAdmin_(token, 'Hanya Super Admin yang boleh mengelola akun pengguna.');

  nama = String(nama || '').trim();
  username = String(username || '').trim();
  role = String(role || '').trim() || 'Staff';
  if (!nama || !username || !password) throw new Error('Nama, username, dan password wajib diisi.');
  if (String(password).length < 6) throw new Error('Password minimal 6 karakter.');

  const validRoles = readColumn_(ensureRolesSheets_().rolesSheet, 1).map(normalizeRoleKey_);
  if (validRoles.indexOf(normalizeRoleKey_(role)) === -1) {
    throw new Error('Role "' + role + '" tidak dikenal / belum terdaftar. Pilih role dari dropdown yang tersedia.');
  }

  withScriptLock_(() => {
    const sheet = getUsersSheet_();
    if (findUserRow_(sheet, username)) throw new Error('Username "' + username + '" sudah dipakai.');
    const salt = generateSalt_();
    sheet.appendRow([username, hashPasswordWithSalt_(password, salt), nama, role, salt, 'Aktif']);
    SpreadsheetApp.flush();
  });

  catatAuditLog_(session, 'Tambah User', 'Akun', username, 'Nama: ' + nama + ', Role: ' + role);
  return getUsersList(token);
}

function setUserStatus(username, status, token) {
  getUsersSheet_();
  const session = requireSuperAdmin_(token, 'Hanya Super Admin yang boleh mengelola akun pengguna.');

  status = status === 'Nonaktif' ? 'Nonaktif' : 'Aktif';
  username = String(username || '').trim();
  if (username.toLowerCase() === String(session.username).toLowerCase() && status === 'Nonaktif') {
    throw new Error('Anda tidak bisa menonaktifkan akun Anda sendiri.');
  }

  withScriptLock_(() => {
    const sheet = getUsersSheet_();
    const found = findUserRow_(sheet, username);
    if (!found) throw new Error('Username tidak ditemukan.');
    sheet.getRange(found.rowNumber, USERS_COL.STATUS).setValue(status);
    SpreadsheetApp.flush();
  });

  catatAuditLog_(session, 'Ubah Status User', 'Akun', username, 'Status diubah jadi ' + status);
  return getUsersList(token);
}

function resetUserPasswordByAdmin(username, newPassword, token) {
  getUsersSheet_();
  const session = requireSuperAdmin_(token, 'Hanya Super Admin yang boleh mengelola akun pengguna.');
  if (!newPassword || String(newPassword).length < 6) throw new Error('Password baru minimal 6 karakter.');
  updateUserPassword_(username, newPassword);
  catatAuditLog_(session, 'Reset Password User', 'Akun', username, 'Password direset oleh admin');
  return { success: true };
}

/* ============================ SETTING > HAK AKSES ============================ */

function getRolesList(token) {
  requireSuperAdmin_(token, 'Hanya Super Admin yang boleh melihat daftar role.');
  const { rolesSheet } = ensureRolesSheets_();
  const last = rolesSheet.getLastRow();
  if (last < 2) return [];
  return rolesSheet.getRange(2, 1, last - 1, 2).getValues().map(r => ({ roleName: r[0], deskripsi: r[1] }));
}

function addRole(roleName, deskripsi, token) {
  const session = requireSuperAdmin_(token, 'Hanya Super Admin yang boleh menambah role.');
  roleName = String(roleName || '').trim();
  if (!roleName) throw new Error('Nama role wajib diisi.');
  if (normalizeRoleKey_(roleName) === normalizeRoleKey_(SUPER_ADMIN_ROLE)) {
    throw new Error('Nama role "' + SUPER_ADMIN_ROLE + '" sudah dipakai sistem.');
  }

  withScriptLock_(() => {
    const { rolesSheet, permSheet } = ensureRolesSheets_();
    if (readColumn_(rolesSheet, 1).map(normalizeRoleKey_).indexOf(normalizeRoleKey_(roleName)) !== -1) {
      throw new Error('Role "' + roleName + '" sudah ada.');
    }
    rolesSheet.appendRow([roleName, String(deskripsi || '').trim()]);
    const rows = MENU_LIST.map(m => [roleName, m.key, false, false]);
    permSheet.getRange(permSheet.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
    SpreadsheetApp.flush();
  });

  catatAuditLog_(session, 'Tambah Role', 'Hak Akses', roleName, deskripsi || '-');
  return getRolesList(token);
}

function getRolePermissionsMatrix(token) {
  requireSuperAdmin_(token, 'Hanya Super Admin yang boleh melihat matrix hak akses.');
  const { rows } = readPermissionRows_();
  return {
    menus: MENU_LIST,
    roles: getRolesList(token),
    permissions: rows.map(r => ({ roleName: r[0], menuKey: r[1], canView: !!r[2], canEdit: !!r[3] }))
  };
}

function setRolePermissions(roleName, menuKeys, editMenuKeys, token) {
  const session = requireSuperAdmin_(token, 'Hanya Super Admin yang boleh mengubah hak akses.');
  roleName = String(roleName || '').trim();
  if (!roleName) throw new Error('Role tidak valid.');
  if (normalizeRoleKey_(roleName) === normalizeRoleKey_(SUPER_ADMIN_ROLE)) {
    throw new Error('Hak akses role "' + SUPER_ADMIN_ROLE + '" tidak bisa diubah (selalu penuh).');
  }
  menuKeys = menuKeys || [];
  editMenuKeys = editMenuKeys || [];

  withScriptLock_(() => {
    const { sheets, rows } = readPermissionRows_();
    let changed = false;
    rows.forEach(r => {
      if (String(r[0]) !== roleName) return;
      const canView = menuKeys.indexOf(String(r[1])) !== -1;
      const canEdit = editMenuKeys.indexOf(String(r[1])) !== -1;
      if (r[2] !== canView) { r[2] = canView; changed = true; }
      if (r[3] !== canEdit) { r[3] = canEdit; changed = true; }
    });
    if (changed) sheets.permSheet.getRange(2, 1, rows.length, 4).setValues(rows);
    SpreadsheetApp.flush();
  });

  catatAuditLog_(session, 'Ubah Hak Akses', 'Hak Akses', roleName,
    'View: ' + menuKeys.length + ' menu, Edit: ' + editMenuKeys.length + ' menu');
  return getRolePermissionsMatrix(token);
}
