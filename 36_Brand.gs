/*************************************************************
 * 36_Brand.gs
 * Pengaturan Perusahaan (white-label): nama aplikasi, perusahaan,
 * departemen, logo. Disimpan di Script Properties (BRAND_CONFIG, JSON).
 * Nilai bawaan = identitas lama, jadi tampilan tidak berubah sebelum diubah.
 *************************************************************/

const BRAND_PROP_KEY_ = 'BRAND_CONFIG';
const BRAND_CACHE_KEY_ = 'brand_config_v1';
const BRAND_DEFAULTS_ = {
  appName: 'Attendance Monitoring System',
  appShort: 'AMS',
  companyName: 'PT Swadharma Primautama',
  departmentName: 'Housekeeping',
  logoUrl: ''
};

function readBrandRaw_() {
  const out = {};
  Object.keys(BRAND_DEFAULTS_).forEach(k => { out[k] = BRAND_DEFAULTS_[k]; });
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(BRAND_PROP_KEY_);
    if (raw) {
      const o = JSON.parse(raw);
      Object.keys(BRAND_DEFAULTS_).forEach(k => {
        if (typeof o[k] === 'string' && (o[k].trim() !== '' || k === 'logoUrl' || k === 'departmentName')) out[k] = o[k];
      });
    }
  } catch (e) { /* pakai bawaan */ }
  return out;
}

/** Konfigurasi brand + teks turunan. Dipakai backend & dikirim ke frontend. */
function getBrandConfig_() {
  const b = readBrandRaw_();
  const dept = b.departmentName ? ('Departemen ' + b.departmentName) : '';
  b.departmentLabel = dept;
  b.orgLine = [dept, b.companyName].filter(Boolean).join(' - ');
  b.fullTitle = b.appName + ' (' + b.appShort + ')' + (b.orgLine ? ' - ' + b.orgLine : '');
  return b;
}

/** Publik (dibutuhkan layar login & halaman absen sebelum login). Hanya nama/logo, bukan rahasia. */
function getBrandConfig() {
  return getBrandConfig_();
}

function cleanBrandField_(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u001f<>]/g, '').trim().slice(0, max);
}

function saveBrandConfig(config, token) {
  const session = requireSuperAdmin_(token, 'Hanya Super Admin yang boleh mengubah Pengaturan Perusahaan.');
  config = config || {};
  const appName = cleanBrandField_(config.appName, 60);
  const appShort = cleanBrandField_(config.appShort, 12);
  const companyName = cleanBrandField_(config.companyName, 80);
  if (!appName || !appShort || !companyName) throw new Error('Nama aplikasi, singkatan, dan nama perusahaan wajib diisi.');
  const logoUrl = cleanBrandField_(config.logoUrl, 400);
  if (logoUrl && !/^https:\/\/[^\s"'()]+$/i.test(logoUrl)) throw new Error('Logo harus berupa alamat https:// atau dikosongkan.');
  const data = {
    appName: appName, appShort: appShort, companyName: companyName,
    departmentName: cleanBrandField_(config.departmentName, 60),
    logoUrl: logoUrl
  };
  PropertiesService.getScriptProperties().setProperty(BRAND_PROP_KEY_, JSON.stringify(data));
  catatAuditLog_(session, 'Ubah Pengaturan Perusahaan', 'Pengaturan', companyName, appName + ' / ' + data.departmentName);
  return getBrandConfig_();
}

/** Escape HTML sederhana untuk dipakai di template email/laporan. */
function brandEsc_(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
