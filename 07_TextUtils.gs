/*************************************************************
 * 07_TextUtils.gs
 * Utilitas teks kecil yang dipakai banyak modul (email, laporan).
 *************************************************************/

/** Escape &, <, > untuk disisipkan ke HTML email. null/undefined -> ''. */
function escapeHtmlServer_(str) {
  if (str === null || str === undefined) return '';
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/**
 * Teks aman untuk ditulis ke sel lewat appendRow/setValues:
 *  - berawalan = + - @ diberi apostrof di depan, supaya Sheets menyimpannya sebagai
 *    TEKS (bukan rumus; input pengguna seperti "=IMPORTXML(...)" tidak boleh jalan);
 *  - dipotong ke maxLen (batas sel Sheets 50.000 karakter; default 5.000).
 * Selain string dikembalikan apa adanya (kecuali null/undefined -> '').
 */
function guardSheetText_(value, maxLen) {
  if (value === null || value === undefined) return '';
  if (typeof value !== 'string') return value;
  let s = value;
  const limit = maxLen || 5000;
  if (s.length > limit) s = s.substring(0, limit - 1) + '…';
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
