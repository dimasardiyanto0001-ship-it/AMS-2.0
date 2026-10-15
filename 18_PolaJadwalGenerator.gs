/*************************************************************
 * 18_PolaJadwalGenerator.gs (v1.1 - FIX ALIGN KE HARI MINGGU KALENDER)
 * MODUL: Setting Jadwal Kerja -> generate isi sheet "Jadwal"
 * otomatis dari Pola Rotasi Shift (bukan lewat upload Excel).
 *
 * FIX v1.1:
 *  Sebelumnya urutan kode pola diputar berdasarkan "hari ke-N dalam
 *  periode" (dayNumber - 1 + offsetHari) % urutan.length. Ini membuat
 *  posisi OFF pada pola seperti "P,P,P,P,P,P,OFF" jatuh di hari yang
 *  BERBEDA setiap bulan, tergantung tanggal 1 periode jatuh di hari
 *  apa (bisa Rabu, Kamis, dst) -- bukan selalu Sabtu/Minggu seperti
 *  yang diharapkan untuk pola "6 kerja 1 libur".
 *
 *  Sekarang, KHUSUS untuk pola yang panjangnya PERSIS 7 kode (artinya
 *  pola mingguan), perhitungan indeks diganti memakai HARI-DALAM-MINGGU
 *  KALENDER ASLI: Minggu = indeks 0, Senin = 1, Selasa = 2, ...,
 *  Sabtu = indeks 6 (sama seperti Date.getDay() bawaan JavaScript).
 *
 *  Artinya, kalau pola diketik "P,P,P,P,P,P,OFF" (7 kode, OFF di
 *  posisi terakhir/indeks 6), maka dengan offsetHari = 0, OFF akan
 *  SELALU jatuh di hari Sabtu, setiap minggu, setiap bulan -- tidak
 *  peduli tanggal 1 periode jatuh di hari apa.
 *
 *  offsetHari (di tab "Assign Pola ke Karyawan") sekarang punya arti
 *  BARU khusus untuk pola 7-kode: "geser posisi pola sebanyak N hari
 *  relatif terhadap Minggu". Contoh cepat:
 *    - offsetHari = 0  -> indeks 6 (posisi terakhir pola) = hari Sabtu
 *    - offsetHari = 1  -> indeks 6 = hari Minggu (geser 1 hari)
 *  Kalau kamu mau OFF jatuh di Minggu (bukan Sabtu), taruh kode OFF
 *  di posisi ke-6 (indeks 5, sebelum posisi terakhir) pada urutan
 *  kode, ATAU cukup set offsetHari = 6 supaya pola bergeser 6 hari.
 *  Cara termudah: buat pola dengan OFF di posisi yang kamu mau, lalu
 *  coba generate untuk 1 periode dan lihat hasilnya -- tinggal
 *  sesuaikan offsetHari per karyawan sampai OFF jatuh di hari yang
 *  diinginkan (setelah pas, tidak perlu diubah lagi bulan-bulan
 *  berikutnya karena sekarang berbasis kalender, bukan hari ke-N).
 *
 *  Pola yang PANJANGNYA BUKAN 7 (mis. rotasi 4 hari kerja 3 libur non-
 *  mingguan) TETAP memakai logika lama (hari ke-N dalam periode),
 *  karena pola semacam itu memang tidak dimaksudkan mengikuti kalender
 *  mingguan.
 *
 * CARA PASANG (3 langkah):
 *  1. Tambahkan satu baris ini ke object CONFIG di 00_Config.gs:
 *       SHEET_POLASHIFT: 'PolaShift',
 *
 *  2. Tambahkan case baru ini ke switch(action) di doPost() (01_WebApp.gs):
 *       case 'getPolaShiftList': result = getPolaShiftList(body.token); break;
 *       case 'addPolaShift': result = addPolaShift(body.namaPola, body.urutanKode, body.deskripsi, body.token); break;
 *       case 'deletePolaShift': result = deletePolaShift(body.namaPola, body.token); break;
 *       case 'assignPolaKaryawan': result = assignPolaKaryawan(body.nama, body.polaAssigned, body.offsetHari, body.token); break;
 *       case 'generateJadwalDariPola': result = generateJadwalDariPola(body.mode, body.token); break;
 *
 *  3. Tempel file ini apa adanya sebagai file baru di Apps Script Editor
 *     (menimpa file 18_PolaJadwalGenerator.gs yang lama).
 *
 * CATATAN TENTANG SHEET "Hari Libur":
 *  Fungsi getHariLiburDaysInPeriode_() di bawah mengasumsikan sheet
 *  bernama "Hari Libur" dengan kolom A = tanggal (Date object atau
 *  dd/MM/yyyy) untuk SELURUH tahun. Kalau struktur sheet Hari Libur
 *  Anda beda, cukup sesuaikan isi fungsi itu saja -- bagian lain
 *  tidak perlu diubah.
 *************************************************************/

/* ============================ SHEET PolaShift (CRUD) ============================ */

function ensurePolaShiftSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_POLASHIFT);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_POLASHIFT);
    sheet.getRange(1, 1, 1, 3).setValues([['NamaPola', 'UrutanKode', 'Deskripsi']]);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    sheet.setColumnWidths(1, 3, 220);
  }
  return sheet;
}

/** Validasi urutan kode terhadap CONFIG.SHIFT_CODES sebelum disimpan. */
function validateUrutanKodePola_(urutanKodeStr) {
  const kodeArray = String(urutanKodeStr || '')
    .split(',')
    .map(k => k.trim().toUpperCase())
    .filter(k => k !== '');

  if (kodeArray.length === 0) {
    throw new Error('Urutan kode pola tidak boleh kosong. Contoh: P,P,P,P,P,P,OFF');
  }
  const shiftMap = getShiftCodesDynamic_();
  kodeArray.forEach(kode => {
    if (!shiftMap.hasOwnProperty(kode)) {
      throw new Error('Kode shift "' + kode + '" tidak dikenal. Kode valid: ' + Object.keys(shiftMap).join(', '));
    }
  });
  return kodeArray;
}

function getPolaShiftList(token) {
  requireMenuAccess_(token, 'settingjadwalkerja');
  const sheet = ensurePolaShiftSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 3).getValues()
    .filter(r => String(r[0] || '').trim() !== '')
    .map(r => ({ namaPola: r[0], urutanKode: r[1], deskripsi: r[2] }));
}

function addPolaShift(namaPola, urutanKode, deskripsi, token) {
  const session = requireMenuEditAccess_(token, 'settingjadwalkerja');
  namaPola = String(namaPola || '').trim();
  if (!namaPola) throw new Error('Nama pola wajib diisi.');
  validateUrutanKodePola_(urutanKode); // lempar error kalau ada kode tidak valid

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = ensurePolaShiftSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r => String(r[0]).trim().toLowerCase());
      if (existing.indexOf(namaPola.toLowerCase()) !== -1) {
        throw new Error('Pola dengan nama "' + namaPola + '" sudah ada.');
      }
    }
    sheet.appendRow([namaPola, String(urutanKode || '').trim().toUpperCase(), String(deskripsi || '').trim()]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  catatAuditLog_(session, 'Tambah Pola Shift', 'PolaShift', namaPola, 'Urutan: ' + urutanKode);
  return getPolaShiftList(token);
}

function deletePolaShift(namaPola, token) {
  const session = requireMenuEditAccess_(token, 'settingjadwalkerja');
  namaPola = String(namaPola || '').trim();

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = ensurePolaShiftSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('Belum ada pola yang tersimpan.');
    const names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    let found = false;
    for (let i = names.length - 1; i >= 0; i--) {
      if (String(names[i][0]).trim().toLowerCase() === namaPola.toLowerCase()) {
        sheet.deleteRow(i + 2);
        found = true;
        break;
      }
    }
    if (!found) throw new Error('Pola "' + namaPola + '" tidak ditemukan.');
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  catatAuditLog_(session, 'Hapus Pola Shift', 'PolaShift', namaPola, '-');
  return getPolaShiftList(token);
}

/* ============================ ASSIGN POLA KE KARYAWAN ============================ */
/* Menambah 2 kolom baru di sheet DataKaryawan: PolaAssigned, OffsetHari.
   Kolom lama (1-8) TIDAK diubah urutannya -- aman untuk data yang sudah ada. */

function ensureDataKaryawanPolaColumns_(sheet) {
  const wanted9 = 'PolaAssigned';
  const wanted10 = 'OffsetHari';
  const current = sheet.getRange(1, 9, 1, 2).getValues()[0];
  let changed = false;
  if (current[0] !== wanted9) { sheet.getRange(1, 9).setValue(wanted9); changed = true; }
  if (current[1] !== wanted10) { sheet.getRange(1, 10).setValue(wanted10); changed = true; }
  if (changed) {
    sheet.getRange(1, 9, 1, 2).setFontWeight('bold').setBackground('#f1f5f9');
    sheet.setColumnWidths(9, 2, 160);
  }
}

function assignPolaKaryawan(nama, polaAssigned, offsetHari, token) {
  const session = requireMenuEditAccess_(token, 'settingjadwalkerja');
  nama = String(nama || '').trim();
  polaAssigned = String(polaAssigned || '').trim();
  offsetHari = parseInt(offsetHari, 10) || 0;
  if (!nama) throw new Error('Nama karyawan wajib diisi.');

  if (polaAssigned) {
    const validPola = getPolaShiftList(token).map(p => p.namaPola.toLowerCase());
    if (validPola.indexOf(polaAssigned.toLowerCase()) === -1) {
      throw new Error('Pola "' + polaAssigned + '" tidak ditemukan. Buat dulu di daftar Pola Shift.');
    }
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getOrCreateNikSheet_();
    ensureDataKaryawanPolaColumns_(sheet);
    const lastRow = sheet.getLastRow();
    const names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    let rowNum = -1;
    for (let i = 0; i < names.length; i++) {
      if (String(names[i][0]).trim().toLowerCase() === nama.toLowerCase()) { rowNum = i + 2; break; }
    }
    if (rowNum === -1) throw new Error('Karyawan "' + nama + '" tidak ditemukan di DataKaryawan.');

    sheet.getRange(rowNum, 9, 1, 2).setValues([[polaAssigned, offsetHari]]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  catatAuditLog_(session, 'Assign Pola Shift', 'Karyawan', nama, 'Pola: ' + (polaAssigned || '-') + ', Offset: ' + offsetHari);
  return { success: true };
}

function getPolaAssignmentMap_() {
  const sheet = getOrCreateNikSheet_();
  ensureDataKaryawanPolaColumns_(sheet);
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow < 2) return map;
  const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  data.forEach(row => {
    const nama = String(row[0] || '').trim();
    if (!nama) return;
    map[nama.toLowerCase()] = {
      polaAssigned: String(row[8] || '').trim(),
      offsetHari: parseInt(row[9], 10) || 0
    };
  });
  return map;
}

/* ============================ HARI LIBUR (opsional, sesuaikan struktur) ============================ */
/* Asumsi: sheet "Hari Libur" kolom A = tanggal (Date atau dd/MM/yyyy).
   Kalau struktur sheet Anda beda, cukup ganti isi fungsi ini saja. */
function getHariLiburDaysInPeriode_(periodText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Hari Libur');
  if (!sheet) return new Set(); // sheet belum ada -> anggap tidak ada libur nasional

  const my = parsePeriodMonthYear_(periodText);
  if (!my) return new Set();

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Set();

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const dayNumbers = new Set();
  values.forEach(r => {
    const v = r[0];
    let d;
    if (v instanceof Date) {
      d = v;
    } else {
      const s = String(v || '').trim();
      const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
      if (!m) return;
      d = new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));
    }
    if (d.getFullYear() === my.year && d.getMonth() === my.monthIndex) {
      dayNumbers.add(d.getDate());
    }
  });
  return dayNumbers;
}

/**
 * BARU (v1.1): indeks hari-dalam-minggu kalender asli untuk tanggal
 * ke-`dayNumber` pada periode `my` (year, monthIndex).
 * 0 = Minggu, 1 = Senin, 2 = Selasa, 3 = Rabu, 4 = Kamis, 5 = Jumat, 6 = Sabtu.
 * (Sama persis dengan Date.prototype.getDay() bawaan JavaScript.)
 */
function getWeekdayIndexOfDay_(my, dayNumber) {
  return new Date(my.year, my.monthIndex, dayNumber).getDay();
}

/* ============================ GENERATE JADWAL DARI POLA (INTI FITUR) ============================ */

/**
 * mode = 'kosongSaja'  -> hanya isi sel yang masih kosong di Jadwal (AMAN, default).
 * mode = 'timpaSemua'  -> timpa semua sel karyawan yang punya pola assigned (DESTRUKTIF).
 * Hanya Super Admin (sama seperti saveJadwal/uploadJadwal).
 */
function generateJadwalDariPola(mode, token) {
  const session = verifySession_(token);
  if (!isSuperAdminRole_(session.role)) {
    throw new Error('Hanya akun dengan role "Super Admin" yang boleh generate Jadwal dari Pola.');
  }
  mode = (mode === 'timpaSemua') ? 'timpaSemua' : 'kosongSaja';

  const jadwalSheet = getSheet_(CONFIG.SHEET_JADWAL);
  const lastRow = getLastEmployeeRow_(jadwalSheet);
  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const dayCount = getDayColCount_(jadwalSheet);
  if (numRows <= 0 || dayCount <= 0) {
    throw new Error('Sheet Jadwal belum punya data karyawan / kolom hari untuk periode ini.');
  }

  const periodText = String(jadwalSheet.getRange(2, 1).getValue() || '').trim();
  const hariLiburSet = getHariLiburDaysInPeriode_(periodText);

  // BARU (v1.1): parse tahun/bulan periode sekali di sini, dipakai untuk
  // menghitung hari-dalam-minggu kalender asli tiap tanggal.
  const myPeriod = parsePeriodMonthYear_(periodText);

  const polaMapByName = {}; // namaPola (lowercase) -> array kode
  getPolaShiftList(token).forEach(p => {
    polaMapByName[p.namaPola.toLowerCase()] = validateUrutanKodePola_(p.urutanKode);
  });

  const assignMap = getPolaAssignmentMap_();

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const info = jadwalSheet.getRange(CONFIG.DATA_START_ROW, 1, numRows, 3).getValues(); // No, Nama, Jabatan
    const currentGrid = jadwalSheet.getRange(CONFIG.DATA_START_ROW, CONFIG.FIRST_DAY_COL, numRows, dayCount).getValues();

    let karyawanDigenerate = 0;
    let selDiisi = 0;
    let selDilewati = 0;
    const tanpaPola = [];

    for (let i = 0; i < numRows; i++) {
      const nama = String(info[i][1] || '').trim();
      if (!nama) continue;

      const assign = assignMap[nama.toLowerCase()];
      if (!assign || !assign.polaAssigned) {
        tanpaPola.push(nama);
        continue;
      }

      const urutan = polaMapByName[assign.polaAssigned.toLowerCase()];
      if (!urutan) {
        tanpaPola.push(nama + ' (pola "' + assign.polaAssigned + '" tidak ditemukan)');
        continue;
      }

      // BARU (v1.1): kalau pola persis 7 kode DAN periode berhasil di-parse,
      // anggap ini pola MINGGUAN -> sejajarkan ke hari kalender asli
      // (Minggu=0 ... Sabtu=6), supaya OFF konsisten jatuh di hari yang
      // sama tiap minggu, tiap bulan. Pola dengan panjang lain (mis. 4
      // atau 5 kode) tetap pakai logika lama (hari ke-N dalam periode).
      const useWeeklyAlignment = (urutan.length === 7 && !!myPeriod);

      karyawanDigenerate++;
      for (let d = 0; d < dayCount; d++) {
        const sudahTerisi = currentGrid[i][d] !== '' && currentGrid[i][d] !== null && currentGrid[i][d] !== undefined;
        if (sudahTerisi && mode === 'kosongSaja') {
          selDilewati++;
          continue;
        }

        const dayNumber = d + 1;
        let kode;
        if (hariLiburSet.has(dayNumber)) {
          kode = 'OFF'; // tanggal merah nasional -> selalu OFF, override pola
        } else if (useWeeklyAlignment) {
          const weekdayIdx = getWeekdayIndexOfDay_(myPeriod, dayNumber); // 0=Minggu..6=Sabtu
          const idxUrutan = (weekdayIdx + assign.offsetHari) % 7;
          kode = urutan[idxUrutan];
        } else {
          const idxUrutan = (dayNumber - 1 + assign.offsetHari) % urutan.length;
          kode = urutan[idxUrutan];
        }

        currentGrid[i][d] = kode;
        selDiisi++;
      }
    }

    jadwalSheet.getRange(CONFIG.DATA_START_ROW, CONFIG.FIRST_DAY_COL, numRows, dayCount).setValues(currentGrid);
    SpreadsheetApp.flush();

    // Sinkronkan OFF ke Realisasi seperti biasa (fungsi ini sudah ada di 04_DataApi.gs)
    syncOffFromJadwal_();

    catatAuditLog_(
      session, 'Generate Jadwal dari Pola', 'Jadwal', periodText,
      'Mode: ' + mode + ', ' + karyawanDigenerate + ' karyawan, ' + selDiisi + ' sel diisi, ' + selDilewati + ' sel dilewati'
    );

    const result = getAllDataInternal_();
    result.generateJadwalInfo = {
      mode: mode,
      karyawanDigenerate: karyawanDigenerate,
      selDiisi: selDiisi,
      selDilewati: selDilewati,
      karyawanTanpaPola: tanpaPola
    };
    return result;
  } finally {
    lock.releaseLock();
  }
}