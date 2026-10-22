/*************************************************************
 * 28_EmployeeLifecycle.gs
 * Siklus karyawan, semuanya khusus Super Admin:
 *   addEmployee              daftar baru di DataKaryawan (Belum Terjadwal)
 *   assignEmployeeToJadwal   masukkan ke Jadwal/Realisasi/Rekap periode aktif
 *   deleteEmployee           keluarkan dari Jadwal/Realisasi/Rekap/JamKerja/JamScan
 *                            (kembali Belum Terjadwal; data master tetap ada)
 *   deleteUnassignedEmployee hapus permanen dari DataKaryawan (hanya yang belum terjadwal)
 *   setKaryawanResign        BELUM DIIMPLEMENTASIKAN (lihat catatan di fungsinya)
 *************************************************************/

/** Status assignment; sel kosong dianggap "Belum Terjadwal" (sama seperti tampilan Daftar Karyawan). */
function assignmentStatusOf_(rawValue) {
  return cellText_(rawValue) || STATUS_BELUM_TERJADWAL;
}

/* ============================ TAMBAH ============================ */

function addEmployee(nama, jabatan, nik, department, email, noHp, token) {
  const session = requireUploadAccess_(token, 'Hanya akun dengan role "Super Admin" atau "Admin" yang boleh menambah karyawan.');
  nama = String(nama || '').trim();
  jabatan = String(jabatan || '').trim();
  nik = String(nik || '').trim();
  department = String(department || '').trim();
  email = String(email || '').trim();
  noHp = String(noHp || '').replace(/[\s\-().]/g, '');
  if (!nama) throw new Error('Nama karyawan wajib diisi.');
  if (nama.length > 100 || /[\u0000-\u001f<>]/.test(nama) || /^[=+\-@]/.test(nama)) {
    throw new Error('Nama tidak valid (maks. 100 karakter; tanpa < > dan tidak boleh diawali = + - @).');
  }
  if (/[\u0000-\u001f]/.test(jabatan + department + nik)) throw new Error('Jabatan/department/NIK mengandung karakter tidak valid.');
  jabatan = guardSheetText_(jabatan, 100);
  department = guardSheetText_(department, 100);
  nik = guardSheetText_(nik, 50);
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error('Format email tidak valid.');
  }
  if (noHp && !/^\+?\d{8,15}$/.test(noHp)) {
    throw new Error('Nomor telepon tidak valid (hanya angka, 8-15 digit).');
  }

  withScriptLock_(() => {
    const nikSheet = getOrCreateNikSheet_();
    if (indexOfEmployee_(readEmployeeRows_(nikSheet, 1), nama) !== -1) {
      throw new Error('Karyawan dengan nama "' + nama + '" sudah terdaftar.');
    }

    const row = new Array(STATUS_ASSIGNMENT_COL_COUNT_).fill('');
    row[EMP_COL.NAMA - 1] = nama;
    row[EMP_COL.NIK - 1] = nik;
    row[EMP_COL.DEPARTMENT - 1] = department;
    row[EMP_COL.EMAIL - 1] = guardSheetText_(email, 150);
    row[EMP_COL.NO_HP - 1] = noHp;
    row[EMP_COL.STATUS_KARYAWAN - 1] = 'Aktif';
    row[EMP_COL.TGL_REGISTRASI - 1] = new Date();
    row[EMP_COL.JABATAN - 1] = jabatan;
    row[EMP_COL.STATUS_ASSIGNMENT - 1] = STATUS_BELUM_TERJADWAL;
    nikSheet.appendRow(row);
    SpreadsheetApp.flush();
  }, 20000);

  const result = getAllDataInternal_();
  result.addEmployeeInfo = { nama: nama, jabatan: jabatan, department: department, status: STATUS_BELUM_TERJADWAL };

  catatAuditLog_(session, 'Tambah Karyawan', 'Karyawan', nama, 'Jabatan: ' + jabatan + ', Department: ' + department);
  return result;
}

/** Karyawan baru mengisi kolom 1..9 (sampai No HP); FotoProfil dan kolom pola dibiarkan kosong. */
const STATUS_ASSIGNMENT_COL_COUNT_ = EMP_COL.NO_HP;

/* ============================ ASSIGN KE JADWAL ============================ */

function setAssignmentStatus_(nikSheet, nama, statusBaru) {
  const idx = indexOfEmployee_(readEmployeeRows_(nikSheet, EMP_COL.STATUS_ASSIGNMENT), nama);
  if (idx !== -1) nikSheet.getRange(idx + 2, EMP_COL.STATUS_ASSIGNMENT).setValue(statusBaru);
}

/** Sisipkan satu baris karyawan baru (No, Nama, Jabatan, sel hari kosong) setelah baris afterRow. */
function insertEmployeeRowAfter_(sheet, afterRow, newNo, nama, jabatan, dayCount) {
  const newRow = afterRow + 1;
  sheet.insertRowAfter(afterRow);
  sheet.getRange(newRow, 1, 1, 3).setValues([[newNo, nama, jabatan]]);
  if (dayCount > 0) sheet.getRange(newRow, CONFIG.FIRST_DAY_COL, 1, dayCount).setValue('');
}

/**
 * Assign satu/beberapa karyawan (array nama) ke periode berjalan: sisipkan baris di
 * Jadwal, Realisasi (kosong) dan Rekap (format disalin dari baris terakhir), lalu ubah
 * status jadi "Aktif Terjadwal". Yang sudah terjadwal, sudah ada di Jadwal, atau tidak
 * terdaftar dilewati dan dilaporkan di assignInfo.skipped.
 */
function assignEmployeeToJadwal(namaList, token) {
  const session = requireSuperAdmin_(token, 'Hanya akun dengan role "Super Admin" yang boleh assign karyawan ke Jadwal.');
  if (!Array.isArray(namaList)) namaList = [namaList];
  namaList = namaList.map(n => String(n || '').trim()).filter(Boolean);
  if (namaList.length === 0) throw new Error('Tidak ada karyawan dipilih.');

  const assigned = [];
  const skipped = [];

  withScriptLock_(() => {
    const nikSheet = getOrCreateNikSheet_();
    const jadwalSheet = getSheet_(CONFIG.SHEET_JADWAL);
    const realisasiSheet = getSheet_(CONFIG.SHEET_REALISASI);
    const rekapSheet = getSheet_(CONFIG.SHEET_REKAP);

    namaList.forEach(nama => {
      const rows = readEmployeeRows_(nikSheet, EMP_COL.STATUS_ASSIGNMENT);
      const idx = indexOfEmployee_(rows, nama);
      if (idx === -1) {
        skipped.push({ nama: nama, alasan: 'tidak ditemukan' });
        return;
      }

      const status = assignmentStatusOf_(rows[idx][EMP_COL.STATUS_ASSIGNMENT - 1]);
      if (status !== STATUS_BELUM_TERJADWAL) {
        skipped.push({ nama: nama, alasan: 'sudah ' + status });
        return;
      }
      if (findEmployeeRowByName_(jadwalSheet, nama) || findEmployeeRowByName_(realisasiSheet, nama)) {
        skipped.push({ nama: nama, alasan: 'sudah ada di sheet Jadwal/Realisasi' });
        return;
      }

      const jabatan = String(rows[idx][EMP_COL.JABATAN - 1] || '');
      const lastRow = getLastEmployeeRow_(jadwalSheet);
      const newNo = lastRow - CONFIG.DATA_START_ROW + 2;
      const newRow = lastRow + 1;
      const dayCount = getDayColCount_(jadwalSheet);

      insertEmployeeRowAfter_(jadwalSheet, lastRow, newNo, nama, jabatan, dayCount);
      insertEmployeeRowAfter_(realisasiSheet, lastRow, newNo, nama, jabatan, dayCount);

      rekapSheet.insertRowAfter(lastRow);
      rekapSheet.getRange(lastRow, 1, 1, 12).copyTo(rekapSheet.getRange(newRow, 1, 1, 12));
      rekapSheet.getRange(newRow, 1, 1, 3).setValues([[newNo, nama, jabatan]]);

      setAssignmentStatus_(nikSheet, nama, STATUS_AKTIF_TERJADWAL);
      assigned.push(nama);
    });

    SpreadsheetApp.flush();
  }, 30000);

  const result = getAllDataInternal_();
  result.assignInfo = { assigned: assigned, skipped: skipped };

  catatAuditLog_(session, 'Assign Karyawan ke Jadwal', 'Karyawan', assigned.join(', ') || '-',
    assigned.length + ' di-assign, ' + skipped.length + ' dilewati');
  return result;
}

/* ============================ KELUARKAN DARI JADWAL ============================ */

/** Baris karyawan (nama persis, abaikan spasi tepi) di Jadwal/Realisasi/Rekap/JamKerja/JamScan; null jika tidak ada. */
function findEmployeeRowByName_(sheet, nama) {
  const lastRow = getLastEmployeeRow_(sheet);
  if (lastRow < CONFIG.DATA_START_ROW) return null;

  const values = sheet.getRange(CONFIG.DATA_START_ROW, 2, lastRow - CONFIG.DATA_START_ROW + 1, 1).getValues();
  const target = String(nama).trim();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === target) return CONFIG.DATA_START_ROW + i;
  }
  return null;
}

/** Rapikan kolom "No" supaya urut 1,2,3,... mengikuti baris yang terisi Nama. */
function renumberEmployeeNoColumn_(sheet) {
  const numRows = getLastEmployeeRow_(sheet) - CONFIG.DATA_START_ROW + 1;
  if (numRows <= 0) return;

  const names = sheet.getRange(CONFIG.DATA_START_ROW, 2, numRows, 1).getValues();
  let n = 1;
  const newNos = names.map(r => (r[0] !== '' && r[0] !== null) ? [n++] : ['']);
  sheet.getRange(CONFIG.DATA_START_ROW, 1, numRows, 1).setValues(newNos);
}

/**
 * Hapus karyawan dari Jadwal/Realisasi/Rekap/JamKerja/JamScan (data master di DataKaryawan
 * tetap, statusnya kembali "Belum Terjadwal"). Bila Realisasi-nya sudah terisi dan force
 * tidak true, ditolak: pesan error HARUS tetap memuat kata "Realisasi" karena frontend
 * mengenali pesan itu untuk menampilkan tombol "Hapus Paksa".
 */
function deleteEmployee(nama, token, force) {
  const session = requireUploadAccess_(token, 'Hanya akun dengan role "Super Admin" atau "Admin" yang boleh menghapus karyawan.');
  nama = String(nama || '').trim();
  if (!nama) throw new Error('Nama karyawan tidak boleh kosong.');

  withScriptLock_(() => {
    const ss = getSpreadsheet_();
    const jadwalSheet = getSheet_(CONFIG.SHEET_JADWAL);
    const realisasiSheet = getSheet_(CONFIG.SHEET_REALISASI);
    const rekapSheet = getSheet_(CONFIG.SHEET_REKAP);
    const nikSheet = getOrCreateNikSheet_();

    const realisasiRow = findEmployeeRowByName_(realisasiSheet, nama);
    if (realisasiRow && !force) {
      const dayCount = getDayColCount_(realisasiSheet);
      if (dayCount > 0) {
        const values = realisasiSheet.getRange(realisasiRow, CONFIG.FIRST_DAY_COL, 1, dayCount).getValues()[0];
        if (values.some(v => String(v || '').trim() !== '')) {
          throw new Error(
            'Karyawan "' + nama + '" masih memiliki data Realisasi terisi pada periode ini. ' +
            'Klik "Hapus Paksa" untuk tetap menghapus beserta data Realisasinya.'
          );
        }
      }
    }

    const sheetsToClean = [jadwalSheet, realisasiSheet, rekapSheet]
      .concat([CONFIG.SHEET_JAMKERJA, CONFIG.SHEET_JAMSCAN].map(name => ss.getSheetByName(name)).filter(Boolean));

    sheetsToClean.forEach(sheet => {
      const row = findEmployeeRowByName_(sheet, nama);
      if (row) {
        sheet.deleteRow(row);
        renumberEmployeeNoColumn_(sheet);
      }
    });

    setAssignmentStatus_(nikSheet, nama, STATUS_BELUM_TERJADWAL);
    SpreadsheetApp.flush();
  }, 30000);

  const result = getAllDataInternal_();
  catatAuditLog_(session, 'Hapus Karyawan dari Jadwal', 'Karyawan', nama, force ? 'Hapus paksa (data Realisasi ikut dihapus)' : 'Dikeluarkan dari Jadwal');
  return result;
}

/** Hapus PERMANEN dari DataKaryawan, hanya yang berstatus Belum Terjadwal (yang lain dilewati). */
function deleteUnassignedEmployee(namaList, token) {
  const session = requireUploadAccess_(token, 'Hanya akun dengan role "Super Admin" atau "Admin" yang boleh menghapus karyawan.');
  if (!namaList || !namaList.length) throw new Error('Tidak ada karyawan yang dipilih untuk dihapus.');

  const namaSet = new Set(namaList.map(n => nameKey_(n)));
  const deleted = [];
  const skipped = [];

  withScriptLock_(() => {
    const sheet = getOrCreateNikSheet_();
    if (sheet.getLastRow() < 2) throw new Error('Tidak ada data karyawan di sheet DataKaryawan.');

    const data = readEmployeeRows_(sheet, EMP_COL.STATUS_ASSIGNMENT);
    for (let i = data.length - 1; i >= 0; i--) { // dari bawah supaya nomor baris tetap valid
      const nama = cellText_(data[i][0]);
      if (!namaSet.has(nama.toLowerCase())) continue;

      if (assignmentStatusOf_(data[i][EMP_COL.STATUS_ASSIGNMENT - 1]) !== STATUS_BELUM_TERJADWAL) {
        skipped.push({ nama: nama, alasan: 'sudah ter-assign ke jadwal' });
        continue;
      }
      sheet.deleteRow(i + 2);
      deleted.push(nama);
    }
    SpreadsheetApp.flush();
  }, 20000);

  const result = getAllDataInternal_();
  result.deleteUnassignedInfo = { deleted: deleted, skipped: skipped };

  catatAuditLog_(session, 'Hapus Karyawan', 'Karyawan', deleted.join(', ') || '-',
    deleted.length + ' dihapus, ' + skipped.length + ' dilewati');
  return result;
}

/* ============================ RESIGN (BELUM JADI) ============================ */

/**
 * CATATAN: versi sebelumnya (21_Resign.gs) hanyalah kerangka: tidak menulis apa pun ke
 * sheet, tetapi tetap mengembalikan {success:true} dan mencatat "Tandai Resign" di Audit
 * Log. Tidak ada tombol di frontend yang memakainya. Agar tidak memberi kesan data sudah
 * berubah, fungsi ini sekarang menolak dengan pesan jelas. Rancangan fitur (kode "R" di
 * Jadwal/Realisasi setelah tanggal resign, dan status di DataKaryawan) perlu diputuskan dulu.
 */
function setKaryawanResign(nama, tanggalResignStr, token) {
  requireSuperAdmin_(token, 'Hanya akun dengan role "Super Admin" yang boleh menandai resign.');
  if (isNaN(new Date(tanggalResignStr).getTime())) throw new Error('Format tanggal resign tidak valid.');
  throw new Error('Fitur "Tandai Resign" belum selesai diimplementasikan sehingga tidak ada data yang diubah. ' +
    'Untuk saat ini gunakan "Hapus Karyawan" (keluarkan dari Jadwal).');
}
