/*************************************************************
 * 03_Auth.gs
 * Setup akun awal, hashing password, rate limiting, login/logout,
 * dan verifikasi session. Session disimpan di CacheService.
 *************************************************************/

/* ============================ SETUP ============================ */

/** Jalankan SEKALI dari editor. Membuat sheet Users + akun admin default. */
function setupUsersSheet() {
  const ss = getSpreadsheet_();
  if (ss.getSheetByName(CONFIG.SHEET_USERS)) {
    throw new Error('Sheet "Users" sudah ada. Hapus manual dulu jika ingin setup ulang.');
  }
  const sheet = ss.insertSheet(CONFIG.SHEET_USERS);
  sheet.getRange(1, 1, 1, 6).setValues([['Username', 'PasswordHash', 'Nama', 'Role', 'Salt', 'Status']]).setFontWeight('bold');

  const salt = generateSalt_();
  sheet.getRange(2, 1, 1, 6).setValues([[
    'admin', hashPasswordWithSalt_('admin123', salt), 'Administrator', SUPER_ADMIN_ROLE, salt, 'Aktif'
  ]]);
  sheet.setColumnWidths(1, 6, 160);

  Logger.log('Sheet "Users" dibuat. Login: admin / admin123 (Super Admin). SEGERA ganti password.');
}

/* ============================ HASHING ============================ */

function generateSalt_() {
  return Utilities.getUuid() + Utilities.getUuid();
}

function hashPasswordWithSalt_(password, salt) {
  let value = password + '|' + salt;
  for (let i = 0; i < AUTH_HASH_ITERATIONS; i++) {
    value = toHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value + '|' + salt));
  }
  return value;
}

/** Skema hash lama (SHA-256 + salt global). Hanya untuk migrasi saat login. */
function hashPasswordLegacy_(password) {
  return toHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + getLegacyPasswordSalt_()));
}

/** Cocokkan password dengan baris user (skema baru jika ada salt, jika tidak skema lama). */
function passwordMatches_(password, storedHash, storedSalt) {
  return storedSalt
    ? hashPasswordWithSalt_(password, storedSalt) === storedHash
    : hashPasswordLegacy_(password) === storedHash;
}

/* ============================ RATE LIMIT ============================ */

function failedAttemptsKey_(username) {
  return 'loginfail_' + String(username).toLowerCase();
}

function getFailedAttempts_(username) {
  const raw = CacheService.getScriptCache().get(failedAttemptsKey_(username));
  return raw ? JSON.parse(raw) : { count: 0 };
}

function registerFailedAttempt_(username) {
  const state = getFailedAttempts_(username);
  state.count = (state.count || 0) + 1;
  CacheService.getScriptCache().put(failedAttemptsKey_(username), JSON.stringify(state), AUTH_LOCKOUT_SEC);
}

function clearFailedAttempts_(username) {
  CacheService.getScriptCache().remove(failedAttemptsKey_(username));
}

/* ============================ USERS SHEET ============================ */

/** Pastikan header Salt (kol 5) dan Status (kol 6) ada; isi Status kosong dengan "Aktif". */
function ensureUsersSheetColumns_(sheet) {
  if (sheet.getRange(1, USERS_COL.SALT).getValue() !== 'Salt') {
    sheet.getRange(1, USERS_COL.SALT).setValue('Salt').setFontWeight('bold');
  }
  if (sheet.getRange(1, USERS_COL.STATUS).getValue() !== 'Status') {
    sheet.getRange(1, USERS_COL.STATUS).setValue('Status').setFontWeight('bold');
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const range = sheet.getRange(2, USERS_COL.STATUS, lastRow - 1, 1);
      range.setValues(range.getValues().map(r => [r[0] || 'Aktif']));
    }
  }
}

/** Migrasi role lama ("admin"/"input") di sheet Users. Idempotent. */
function migrateUserRoles_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const range = sheet.getRange(2, USERS_COL.ROLE, lastRow - 1, 1);
  const vals = range.getValues();
  let changed = false;
  vals.forEach(r => {
    const mapped = resolveLegacyRole_(r[0]);
    if (mapped) { r[0] = mapped; changed = true; }
  });
  if (changed) range.setValues(vals);
}

/** Buka sheet Users dalam keadaan skema & role sudah termigrasi. */
function getUsersSheet_() {
  const sheet = getSheet_(CONFIG.SHEET_USERS);
  ensureUsersSheetColumns_(sheet);
  migrateUserRoles_(sheet);
  return sheet;
}

/** Cari baris user (case-insensitive). @return {{rowNumber:number, row:Array}|null} */
function findUserRow_(sheet, username) {
  const data = sheet.getDataRange().getValues();
  const target = String(username).toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === target) return { rowNumber: i + 1, row: data[i] };
  }
  return null;
}

function setUserPassword_(sheet, rowNumber, newPassword) {
  const salt = generateSalt_();
  sheet.getRange(rowNumber, USERS_COL.HASH).setValue(hashPasswordWithSalt_(newPassword, salt));
  sheet.getRange(rowNumber, USERS_COL.SALT).setValue(salt);
}

/* ============================ API AKUN (editor) ============================ */

function addUser(username, password, nama, role) {
  const sheet = getUsersSheet_();
  const salt = generateSalt_();
  sheet.appendRow([username, hashPasswordWithSalt_(password, salt), nama, role || 'Staff', salt, 'Aktif']);
  Logger.log('User "%s" berhasil ditambahkan.', username);
}

function updateUserPassword(username, newPassword) {
  if (!username || !newPassword) throw new Error('Username dan password baru wajib diisi.');
  const sheet = getUsersSheet_();
  const found = findUserRow_(sheet, username);
  if (!found) throw new Error('Username "' + username + '" tidak ditemukan di sheet Users.');
  setUserPassword_(sheet, found.rowNumber, newPassword);
  clearFailedAttempts_(username);
  return { success: true, username: found.row[0] };
}

/* ============================ LOGIN / SESSION ============================ */

function login(username, password) {
  if (!username || !password) throw new Error('Username dan password wajib diisi.');

  if (getFailedAttempts_(username).count >= AUTH_MAX_FAILED_ATTEMPTS) {
    throw new Error('Terlalu banyak percobaan login gagal untuk akun ini. Coba lagi dalam beberapa menit.');
  }

  const sheet = getUsersSheet_();
  const found = findUserRow_(sheet, username);

  if (!found) {
    registerFailedAttempt_(username);
    throw new Error('Username atau password salah.');
  }

  const row = found.row;
  if ((row[USERS_COL.STATUS - 1] || 'Aktif') === 'Nonaktif') {
    throw new Error('Akun ini telah dinonaktifkan. Hubungi admin.');
  }

  const hadSalt = !!row[USERS_COL.SALT - 1];
  if (!passwordMatches_(password, row[USERS_COL.HASH - 1], row[USERS_COL.SALT - 1])) {
    registerFailedAttempt_(username);
    throw new Error('Username atau password salah.');
  }
  if (!hadSalt) setUserPassword_(sheet, found.rowNumber, password); // upgrade skema hash

  clearFailedAttempts_(username);
  const token = Utilities.getUuid();
  const session = {
    username: row[USERS_COL.USERNAME - 1],
    nama: row[USERS_COL.NAMA - 1],
    role: resolveRoleName_(row[USERS_COL.ROLE - 1])
  };
  CacheService.getScriptCache().put('session_' + token, JSON.stringify(session), CONFIG.SESSION_DURATION_SEC);
  return Object.assign({ token }, session);
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('session_' + token);
  return { success: true };
}

function verifySession_(token) {
  if (!token) throw new Error('Anda harus login terlebih dahulu untuk mengubah data.');
  const raw = CacheService.getScriptCache().get('session_' + token);
  if (!raw) throw new Error('Sesi login sudah habis (maks 6 jam). Silakan login ulang.');
  return JSON.parse(raw);
}

/** Alias lama yang masih dipakai banyak modul. */
function requireLogin_(token) {
  return verifySession_(token);
}

function changeOwnPassword(oldPassword, newPassword, token) {
  const session = verifySession_(token);
  if (!oldPassword || !newPassword) throw new Error('Password lama dan password baru wajib diisi.');
  if (String(newPassword).length < 6) throw new Error('Password baru minimal 6 karakter.');

  const sheet = getUsersSheet_();
  const found = findUserRow_(sheet, session.username);
  if (!found) throw new Error('Akun tidak ditemukan.');

  if (!passwordMatches_(oldPassword, found.row[USERS_COL.HASH - 1], found.row[USERS_COL.SALT - 1])) {
    throw new Error('Password lama yang Anda masukkan salah.');
  }
  setUserPassword_(sheet, found.rowNumber, newPassword);
  return { success: true };
}
