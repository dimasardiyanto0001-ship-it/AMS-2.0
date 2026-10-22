/*************************************************************
 * 27_EmployeeData.gs
 * Sheet "DataKaryawan" (master data karyawan) dan menu
 * Kepegawaian > Daftar Karyawan.
 *
 * Kolom (1-based): lihat EMP_COL. Nama menjadi kunci (tanpa beda
 * huruf besar/kecil) di seluruh sistem.
 *************************************************************/

const STATUS_BELUM_TERJADWAL = 'Belum Terjadwal';
const STATUS_AKTIF_TERJADWAL = 'Aktif Terjadwal';

const EMP_COL = {
  NAMA: 1, NIK: 2, DEPARTMENT: 3, EMAIL: 4, STATUS_KARYAWAN: 5,
  TGL_REGISTRASI: 6, JABATAN: 7, STATUS_ASSIGNMENT: 8, NO_HP: 9, FOTO_PROFIL: 10,
  POLA_ASSIGNED: 11, OFFSET_HARI: 12   // pola rotasi shift (18_PolaJadwalGenerator.gs)
};
const EMP_HEADERS_ = ['Nama', 'NIK', 'Department', 'Email', 'Status Karyawan', 'Tanggal Registrasi', 'Jabatan', 'Status Assignment', 'No HP', 'FotoProfil', 'PolaAssigned', 'OffsetHari'];
const EMP_NUM_COLS_ = EMP_HEADERS_.length;

/** Field yang boleh diubah lewat setEmployeeMasterData -> nomor kolom. */
const EMP_EDITABLE_FIELDS_ = {
  nik: EMP_COL.NIK,
  department: EMP_COL.DEPARTMENT,
  email: EMP_COL.EMAIL,
  statusKaryawan: EMP_COL.STATUS_KARYAWAN,
  jabatan: EMP_COL.JABATAN,
  noHp: EMP_COL.NO_HP,
  fotoProfilUrl: EMP_COL.FOTO_PROFIL
};

/* ============================ SHEET ============================ */

/** Buka (atau buat) DataKaryawan dengan header yang benar dan kolom No HP berformat teks. */
function getOrCreateNikSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(NIK_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(NIK_SHEET_NAME);

  const lastCol = Math.max(sheet.getLastColumn(), EMP_NUM_COLS_);
  const current = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  let headerChanged = false;
  EMP_HEADERS_.forEach((header, i) => {
    if (current[i] !== header) {
      sheet.getRange(1, i + 1).setValue(header);
      headerChanged = true;
    }
  });
  if (headerChanged) {
    sheet.getRange(1, 1, 1, EMP_NUM_COLS_).setFontWeight('bold').setBackground('#f1f5f9');
    sheet.setColumnWidths(1, EMP_NUM_COLS_, 160);
  }

  // No HP harus berformat teks supaya angka 0 di depan tidak hilang. Format satu kolom penuh
  // ditulis hanya bila perlu (header baru, atau baris paling bawah belum berformat teks,
  // mis. setelah sheet bertambah baris) -- fungsi ini dipanggil di setiap pembacaan data.
  const maxRows = sheet.getMaxRows();
  if (headerChanged || sheet.getRange(maxRows, EMP_COL.NO_HP).getNumberFormat() !== '@') {
    sheet.getRange(1, EMP_COL.NO_HP, maxRows, 1).setNumberFormat('@');
  }
  return sheet;
}

/** Baris data (tanpa header) sebanyak numCols kolom pertama; [] bila kosong. */
function readEmployeeRows_(sheet, numCols) {
  const lastRow = sheet.getLastRow();
  return lastRow < 2 ? [] : sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
}

/** Indeks (0-based) baris karyawan di hasil readEmployeeRows_, atau -1. */
function indexOfEmployee_(rows, nama) {
  const target = nameKey_(nama || '');
  return rows.findIndex(r => nameKey_(r[0] || '') === target);
}

function cellText_(value) {
  return String(value || '').trim();
}

/* ============================ BACA ============================ */

/** Peta nama(huruf kecil) -> {nik, department, email, statusKaryawan, tanggalRegistrasi, fotoProfilUrl}. */
function getKaryawanMasterMap_() {
  const map = {};
  readEmployeeRows_(getOrCreateNikSheet_(), EMP_NUM_COLS_).forEach(row => {
    const nama = cellText_(row[EMP_COL.NAMA - 1]);
    if (!nama) return;
    map[nama.toLowerCase()] = {
      nik: cellText_(row[EMP_COL.NIK - 1]),
      department: cellText_(row[EMP_COL.DEPARTMENT - 1]),
      email: cellText_(row[EMP_COL.EMAIL - 1]),
      statusKaryawan: cellText_(row[EMP_COL.STATUS_KARYAWAN - 1]),
      tanggalRegistrasi: row[EMP_COL.TGL_REGISTRASI - 1] || '',
      fotoProfilUrl: cellText_(row[EMP_COL.FOTO_PROFIL - 1])
    };
  });
  return map;
}

/** Peta nama(huruf kecil) -> NIK. */
function getNikMap_() {
  const master = getKaryawanMasterMap_();
  const nikMap = {};
  Object.keys(master).forEach(nama => { nikMap[nama] = master[nama].nik; });
  return nikMap;
}

function getDaftarKaryawanInternal_() {
  return readEmployeeRows_(getOrCreateNikSheet_(), EMP_NUM_COLS_)
    .filter(r => r[0] && String(r[0]).trim() !== '')
    .map(r => {
      let tglRegistrasi = r[EMP_COL.TGL_REGISTRASI - 1] || '';
      if (tglRegistrasi instanceof Date) {
        tglRegistrasi = Utilities.formatDate(tglRegistrasi, Session.getScriptTimeZone() || 'GMT+7', 'dd/MM/yyyy');
      }
      return {
        nik: String(r[EMP_COL.NIK - 1] || ''),
        nama: cellText_(r[EMP_COL.NAMA - 1]),
        department: String(r[EMP_COL.DEPARTMENT - 1] || ''),
        jabatan: String(r[EMP_COL.JABATAN - 1] || ''),
        email: String(r[EMP_COL.EMAIL - 1] || ''),
        statusKaryawan: String(r[EMP_COL.STATUS_KARYAWAN - 1] || '') || 'Aktif',
        tanggalRegistrasi: tglRegistrasi,
        statusAssignment: String(r[EMP_COL.STATUS_ASSIGNMENT - 1] || '') || STATUS_BELUM_TERJADWAL,
        noHp: String(r[EMP_COL.NO_HP - 1] || ''),
        fotoProfilUrl: String(r[EMP_COL.FOTO_PROFIL - 1] || ''),
        polaAssigned: cellText_(r[EMP_COL.POLA_ASSIGNED - 1]),
        offsetHari: parseInt(r[EMP_COL.OFFSET_HARI - 1], 10) || 0
      };
    });
}

function getDaftarKaryawan(token) {
  verifySession_(token);
  return getDaftarKaryawanInternal_();
}

/* ============================ UBAH ============================ */

/** Ubah field master data satu karyawan. Perlu izin EDIT menu "daftarkaryawan". */
function setEmployeeMasterData(nama, fields, token) {
  const session = requireMenuEditAccess_(token, 'daftarkaryawan');
  nama = String(nama || '').trim();
  if (!nama) throw new Error('Nama karyawan wajib diisi.');
  fields = fields || {};

  withScriptLock_(() => {
    const sheet = getOrCreateNikSheet_();
    const idx = indexOfEmployee_(readEmployeeRows_(sheet, 1), nama);
    if (sheet.getLastRow() < 2) throw new Error('Data karyawan tidak ditemukan.');
    if (idx === -1) throw new Error('Karyawan "' + nama + '" tidak ditemukan di sheet DataKaryawan.');

    const row = idx + 2;
    Object.keys(EMP_EDITABLE_FIELDS_).forEach(field => {
      if (fields.hasOwnProperty(field)) {
        sheet.getRange(row, EMP_EDITABLE_FIELDS_[field]).setValue(guardSheetText_(String(fields[field] || '').trim(), 150));
      }
    });
    SpreadsheetApp.flush();
  }, 15000);

  const result = getDaftarKaryawanInternal_();
  catatAuditLog_(session, 'Edit Data Karyawan', 'Karyawan', nama, 'Field diubah: ' + Object.keys(fields).join(', '));
  return result;
}

/**
 * Isi/ubah NIK satu karyawan (tombol "Isi NIK"). Mengembalikan state penuh
 * karena handler frontend langsung merender ulang dari hasilnya.
 */
function setEmployeeNik(nama, nik, token) {
  setEmployeeMasterData(nama, { nik: nik }, token);
  return getAllDataInternal_();
}
