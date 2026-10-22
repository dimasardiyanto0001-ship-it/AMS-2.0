/*************************************************************
 * 98_SecurityCheck.gs
 * Alat diagnosis keamanan (AMAN: hanya membaca, tidak mengubah data, tidak menampilkan nilai
 * rahasia). Jalankan manual dari editor lalu lihat Log eksekusi.
 *
 *   periksaRahasiaScriptProperties()  apakah rahasia sudah diisi di Script Properties
 *   periksaAkunBelumTermigrasi()      akun yang masih memakai hash lama (tanpa salt per-user)
 *************************************************************/

const SECRET_PROPERTY_KEYS_ = ['LEGACY_PASSWORD_SALT', 'REPORT_EMAIL', 'GEMINI_API_KEY'];

/** Status terisi/kosong tiap rahasia (nilainya TIDAK ditampilkan). @return {Object} kunci -> boolean */
function periksaRahasiaScriptProperties() {
  requireScriptOwner_();
  const status = {};
  SECRET_PROPERTY_KEYS_.forEach(key => {
    status[key] = !!getSetting_(key, '');
    Logger.log(key + ': ' + (status[key] ? 'terisi' : 'KOSONG'));
  });
  return status;
}

/**
 * Username akun yang belum punya salt per-user (kolom Salt kosong). Akun ini masih diverifikasi
 * dengan LEGACY_PASSWORD_SALT dan baru "naik" ke hash per-user saat pemiliknya login berikutnya.
 * Selama daftar ini tidak kosong, jangan menghapus LEGACY_PASSWORD_SALT.
 * @return {string[]}
 */
function periksaAkunBelumTermigrasi() {
  requireScriptOwner_();
  const rows = getSheet_(CONFIG.SHEET_USERS).getDataRange().getValues().slice(1);
  const legacy = rows
    .filter(r => String(r[USERS_COL.USERNAME - 1] || '').trim() && !String(r[USERS_COL.SALT - 1] || '').trim())
    .map(r => String(r[USERS_COL.USERNAME - 1]).trim());

  Logger.log('Akun total: ' + rows.filter(r => String(r[USERS_COL.USERNAME - 1] || '').trim()).length +
    ' | belum termigrasi ke salt per-user: ' + legacy.length);
  legacy.forEach(u => Logger.log('  - ' + u));
  return legacy;
}
