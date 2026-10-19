/*************************************************************
 * 34_RelieverMaster.gs
 * Daftar Reliever di dalam sistem (menu Kepegawaian > Daftar Reliever).
 * Pengganti input manual di sheet DataReliever (Nama | Perusahaan | NoHP | Status).
 *
 *   getRelieverMaster   daftar semua reliever        (izin LIHAT menu 'daftarreliever')
 *   addReliever         daftarkan reliever baru      (izin EDIT)  -> Status = Aktif
 *   updateReliever      ubah Perusahaan / No HP      (izin EDIT)  -> Nama = kunci, tidak diubah
 *   setRelieverStatus   Aktif / Nonaktif             (izin EDIT)
 *
 * Tidak ada hapus: riwayat absen (LogAbsenOnlineReliever, BackupReliever) memakai nama sebagai
 * kunci, jadi reliever yang sudah tidak dipakai cukup dinonaktifkan.
 * Memakai dari 24_BackupReliever.gs: getOrCreateDataRelieverSheet_, isRelieverNonaktif_,
 * RELIEVER_FIELD_MAKS_.
 *************************************************************/

const RELIEVER_MENU_KEY_ = 'daftarreliever';
const RELIEVER_HP_MAKS_ = 20;
const RELIEVER_HP_RE_ = /^[0-9+\-\s()]*$/;
const RELIEVER_STATUS_AKTIF_ = 'Aktif';
const RELIEVER_STATUS_NONAKTIF_ = 'Nonaktif';

function relieverMasterCell_(v) {
  return (v === null || v === undefined) ? '' : String(v).trim();
}

/** Semua baris DataReliever yang punya nama: [{row, nama, perusahaan, noHp, status}]. */
function bacaRelieverMaster_() {
  const sheet = getOrCreateDataRelieverSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const out = [];
  sheet.getRange(2, 1, lastRow - 1, 4).getValues().forEach(function (r, i) {
    const nama = relieverMasterCell_(r[0]);
    if (!nama) return;
    out.push({
      row: i + 2,
      nama: nama,
      perusahaan: relieverMasterCell_(r[1]),
      noHp: relieverMasterCell_(r[2]),
      status: isRelieverNonaktif_(r[3]) ? RELIEVER_STATUS_NONAKTIF_ : RELIEVER_STATUS_AKTIF_
    });
  });
  return out;
}

function validasiInputReliever_(nama, perusahaan, noHp, namaWajib) {
  nama = relieverMasterCell_(nama).replace(/\s+/g, ' ');
  perusahaan = relieverMasterCell_(perusahaan).replace(/\s+/g, ' ');
  noHp = relieverMasterCell_(noHp);

  if (namaWajib && !nama) throw new Error('Nama reliever wajib diisi.');
  if (!perusahaan) throw new Error('Perusahaan wajib diisi.');
  if (nama.length > RELIEVER_FIELD_MAKS_ || perusahaan.length > RELIEVER_FIELD_MAKS_) {
    throw new Error('Nama atau perusahaan terlalu panjang (maks ' + RELIEVER_FIELD_MAKS_ + ' karakter).');
  }
  if (noHp.length > RELIEVER_HP_MAKS_ || !RELIEVER_HP_RE_.test(noHp)) {
    throw new Error('No HP hanya boleh berisi angka, spasi, +, - dan tanda kurung (maks ' + RELIEVER_HP_MAKS_ + ' karakter).');
  }
  return { nama: nama, perusahaan: perusahaan, noHp: noHp };
}

function getRelieverMaster(token) {
  requireMenuAccess_(token, RELIEVER_MENU_KEY_);
  return { rows: bacaRelieverMaster_() };
}

function addReliever(nama, perusahaan, noHp, token) {
  const session = requireMenuEditAccess_(token, RELIEVER_MENU_KEY_);
  const data = validasiInputReliever_(nama, perusahaan, noHp, true);

  withScriptLock_(function () {
    const key = nameKey_(data.nama);
    if (bacaRelieverMaster_().some(function (r) { return nameKey_(r.nama) === key; })) {
      throw new Error('Reliever "' + data.nama + '" sudah terdaftar.');
    }
    const sheet = getOrCreateDataRelieverSheet_();
    const row = Math.max(sheet.getLastRow(), 1) + 1;
    if (row > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 20);
    sheet.getRange(row, 3).setNumberFormat('@'); // 0 di depan nomor HP jangan hilang
    sheet.getRange(row, 1, 1, 4).setValues([[
      guardSheetText_(data.nama, RELIEVER_FIELD_MAKS_),
      guardSheetText_(data.perusahaan, RELIEVER_FIELD_MAKS_),
      data.noHp,
      RELIEVER_STATUS_AKTIF_
    ]]);
    SpreadsheetApp.flush();
  }, 20000);

  catatAuditLog_(session, 'Tambah Reliever', 'Reliever', data.nama, 'Perusahaan: ' + data.perusahaan);
  return { rows: bacaRelieverMaster_() };
}

function updateReliever(nama, perusahaan, noHp, token) {
  const session = requireMenuEditAccess_(token, RELIEVER_MENU_KEY_);
  const data = validasiInputReliever_(nama, perusahaan, noHp, true);

  withScriptLock_(function () {
    const key = nameKey_(data.nama);
    const found = bacaRelieverMaster_().find(function (r) { return nameKey_(r.nama) === key; });
    if (!found) throw new Error('Reliever "' + data.nama + '" tidak ditemukan.');
    const sheet = getOrCreateDataRelieverSheet_();
    sheet.getRange(found.row, 3).setNumberFormat('@');
    sheet.getRange(found.row, 2, 1, 2).setValues([[guardSheetText_(data.perusahaan, RELIEVER_FIELD_MAKS_), data.noHp]]);
    SpreadsheetApp.flush();
  }, 20000);

  catatAuditLog_(session, 'Ubah Reliever', 'Reliever', data.nama, 'Perusahaan: ' + data.perusahaan);
  return { rows: bacaRelieverMaster_() };
}

function setRelieverStatus(nama, status, token) {
  const session = requireMenuEditAccess_(token, RELIEVER_MENU_KEY_);
  nama = relieverMasterCell_(nama);
  status = relieverMasterCell_(status);
  if (!nama) throw new Error('Nama reliever wajib diisi.');
  if (status !== RELIEVER_STATUS_AKTIF_ && status !== RELIEVER_STATUS_NONAKTIF_) {
    throw new Error('Status harus "Aktif" atau "Nonaktif".');
  }

  withScriptLock_(function () {
    const key = nameKey_(nama);
    const found = bacaRelieverMaster_().find(function (r) { return nameKey_(r.nama) === key; });
    if (!found) throw new Error('Reliever "' + nama + '" tidak ditemukan.');
    getOrCreateDataRelieverSheet_().getRange(found.row, 4).setValue(status);
    SpreadsheetApp.flush();
  }, 20000);

  catatAuditLog_(session, status === RELIEVER_STATUS_AKTIF_ ? 'Aktifkan Reliever' : 'Nonaktifkan Reliever', 'Reliever', nama, 'Status: ' + status);
  return { rows: bacaRelieverMaster_() };
}
