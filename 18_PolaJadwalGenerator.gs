/*************************************************************
 * 18_PolaJadwalGenerator.gs
 * Menu Setting Jadwal Kerja: daftar Pola Rotasi Shift, penetapan pola
 * ke karyawan, dan generate isi sheet "Jadwal" dari pola tersebut
 * (alternatif dari upload Excel).
 *
 * Sheet:
 *  - PolaShift   : NamaPola | UrutanKode | Deskripsi   (CONFIG.SHEET_POLASHIFT)
 *  - DataKaryawan: kolom PolaAssigned dan OffsetHari (EMP_COL.POLA_ASSIGNED,
 *                  EMP_COL.OFFSET_HARI = kolom 11-12; didefinisikan di 27_EmployeeData.gs).
 *                  Kolom 9-10 adalah No HP dan FotoProfil, JANGAN dipakai untuk pola.
 *
 * Aturan generate:
 *  - Pola PERSIS 7 kode = pola mingguan, disejajarkan ke hari kalender asli:
 *    indeks 0 = Minggu ... 6 = Sabtu (seperti Date.getDay()). Indeks yang dipakai
 *    = (harikalender + offsetHari) mod 7. Dengan offsetHari 0, kode ke-7 selalu
 *    jatuh pada hari Sabtu, ke-1 pada Minggu.
 *  - Pola dengan panjang lain berputar menurut "hari ke-N dalam periode":
 *    (N - 1 + offsetHari) mod panjang.
 *  - Tanggal yang ada di sheet "Hari Libur" selalu diisi OFF.
 *  - offsetHari boleh negatif.
 *
 * Satu kali setelah pindah ke versi ini (hanya bila AssignPola pernah dipakai
 * pada struktur lama): jalankan migrasiPolaDariKolom9dan10() dari editor
 * (lihat fungsinya di bawah).
 *************************************************************/

const POLA_NAMA_MAKS_ = 60;
const POLA_DESKRIPSI_MAKS_ = 200;
const POLA_URUTAN_MAKS_KODE_ = 62;
const OFFSET_HARI_MAKS_ = 366;
const SHEET_HARI_LIBUR_ = 'Hari Libur';

/* ============================ SHEET PolaShift ============================ */

function ensurePolaShiftSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(CONFIG.SHEET_POLASHIFT);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_POLASHIFT);
    sheet.getRange(1, 1, 1, 3).setValues([['NamaPola', 'UrutanKode', 'Deskripsi']]).setFontWeight('bold');
    sheet.setColumnWidths(1, 3, 220);
  }
  return sheet;
}

/** Urutan kode "P,P,OFF" -> ['P','P','OFF'] (huruf besar). Error bila kosong, terlalu panjang, atau ada kode tidak dikenal. */
function validateUrutanKodePola_(urutanKodeStr) {
  const kodeArray = String(urutanKodeStr || '')
    .split(',')
    .map(k => k.trim().toUpperCase())
    .filter(k => k !== '');

  if (kodeArray.length === 0) {
    throw new Error('Urutan kode pola tidak boleh kosong. Contoh: P,P,P,P,P,P,OFF');
  }
  if (kodeArray.length > POLA_URUTAN_MAKS_KODE_) {
    throw new Error('Urutan kode pola maksimal ' + POLA_URUTAN_MAKS_KODE_ + ' kode.');
  }
  const shiftMap = getShiftCodesDynamic_();
  kodeArray.forEach(kode => {
    if (!shiftMap.hasOwnProperty(kode)) {
      throw new Error('Kode shift "' + kode + '" tidak dikenal. Kode valid: ' + Object.keys(shiftMap).join(', '));
    }
  });
  return kodeArray;
}

/** Baris pola apa adanya dari sheet: [{namaPola, urutanKode, deskripsi}] (nama kosong dilewati). */
function readPolaShiftRows_() {
  const sheet = ensurePolaShiftSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 3).getValues()
    .filter(r => String(r[0] === null || r[0] === undefined ? '' : r[0]).trim() !== '')
    .map(r => ({ namaPola: String(r[0]).trim(), urutanKode: String(r[1] === null || r[1] === undefined ? '' : r[1]), deskripsi: String(r[2] === null || r[2] === undefined ? '' : r[2]) }));
}

function getPolaShiftList(token) {
  requireMenuAccess_(token, 'settingjadwalkerja');
  return readPolaShiftRows_();
}

function addPolaShift(namaPola, urutanKode, deskripsi, token) {
  const session = requireMenuEditAccess_(token, 'settingjadwalkerja');
  namaPola = String(namaPola || '').trim();
  if (!namaPola) throw new Error('Nama pola wajib diisi.');
  if (namaPola.length > POLA_NAMA_MAKS_) throw new Error('Nama pola maksimal ' + POLA_NAMA_MAKS_ + ' karakter.');
  const kodeArray = validateUrutanKodePola_(urutanKode);
  deskripsi = String(deskripsi || '').trim().substring(0, POLA_DESKRIPSI_MAKS_);

  withScriptLock_(() => {
    const sheet = ensurePolaShiftSheet_();
    const key = nameKey_(namaPola);
    if (readPolaShiftRows_().some(p => nameKey_(p.namaPola) === key)) {
      throw new Error('Pola dengan nama "' + namaPola + '" sudah ada.');
    }
    // Sel berformat teks: nama seperti "2" atau "=A1" tersimpan apa adanya, tidak jadi angka/tanggal/rumus.
    const row = Math.max(sheet.getLastRow(), 1) + 1;
    const range = sheet.getRange(row, 1, 1, 3);
    range.setNumberFormat('@');
    range.setValues([[namaPola, kodeArray.join(','), deskripsi]]);
    SpreadsheetApp.flush();
  }, 15000);

  catatAuditLog_(session, 'Tambah Pola Shift', 'PolaShift', namaPola, 'Urutan: ' + kodeArray.join(','));
  return readPolaShiftRows_();
}

function deletePolaShift(namaPola, token) {
  const session = requireMenuEditAccess_(token, 'settingjadwalkerja');
  namaPola = String(namaPola || '').trim();

  withScriptLock_(() => {
    const sheet = ensurePolaShiftSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('Belum ada pola yang tersimpan.');
    const key = nameKey_(namaPola);
    const names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    let rowNum = -1;
    for (let i = names.length - 1; i >= 0; i--) {
      if (nameKey_(names[i][0] === null || names[i][0] === undefined ? '' : names[i][0]) === key) { rowNum = i + 2; break; }
    }
    if (rowNum === -1) throw new Error('Pola "' + namaPola + '" tidak ditemukan.');
    sheet.deleteRow(rowNum);
    SpreadsheetApp.flush();
  }, 15000);

  // Penetapan pola karyawan yang memakai pola ini tidak dihapus; saat generate mereka
  // dilaporkan di karyawanTanpaPola sampai diberi pola lain.
  const assignMap = getPolaAssignmentMap_();
  const dipakai = Object.keys(assignMap).filter(k => nameKey_(assignMap[k].polaAssigned) === nameKey_(namaPola)).length;
  catatAuditLog_(session, 'Hapus Pola Shift', 'PolaShift', namaPola, dipakai ? 'Masih ditetapkan ke ' + dipakai + ' karyawan' : '-');
  return readPolaShiftRows_();
}

/* ============================ ASSIGN POLA KE KARYAWAN ============================ */

/** Menulis PolaAssigned/OffsetHari di kolom 11-12 DataKaryawan (EMP_COL), bukan 9-10 (No HP/FotoProfil). */
function assignPolaKaryawan(nama, polaAssigned, offsetHari, token) {
  const session = requireMenuEditAccess_(token, 'settingjadwalkerja');
  nama = String(nama || '').trim();
  polaAssigned = String(polaAssigned || '').trim();
  offsetHari = parseInt(offsetHari, 10) || 0;
  if (!nama) throw new Error('Nama karyawan wajib diisi.');
  if (Math.abs(offsetHari) > OFFSET_HARI_MAKS_) {
    throw new Error('Offset hari harus antara -' + OFFSET_HARI_MAKS_ + ' dan ' + OFFSET_HARI_MAKS_ + '.');
  }

  withScriptLock_(() => {
    // Nama pola disimpan persis seperti di daftar Pola Shift (huruf besar/kecil), supaya
    // dropdown di layar bisa memilihnya kembali.
    if (polaAssigned) {
      const key = nameKey_(polaAssigned);
      const pola = readPolaShiftRows_().find(p => nameKey_(p.namaPola) === key);
      if (!pola) throw new Error('Pola "' + polaAssigned + '" tidak ditemukan. Buat dulu di daftar Pola Shift.');
      polaAssigned = pola.namaPola;
    }

    const sheet = getOrCreateNikSheet_();
    const idx = indexOfEmployee_(readEmployeeRows_(sheet, 1), nama);
    if (idx === -1) throw new Error('Karyawan "' + nama + '" tidak ditemukan di DataKaryawan.');

    const row = idx + 2;
    const polaCell = sheet.getRange(row, EMP_COL.POLA_ASSIGNED);
    polaCell.setNumberFormat('@');
    polaCell.setValue(polaAssigned);
    sheet.getRange(row, EMP_COL.OFFSET_HARI).setValue(offsetHari);
    SpreadsheetApp.flush();
  }, 15000);

  catatAuditLog_(session, 'Assign Pola Shift', 'Karyawan', nama, 'Pola: ' + (polaAssigned || '-') + ', Offset: ' + offsetHari);
  return { success: true };
}

/** Peta nama(huruf kecil) -> {polaAssigned, offsetHari}. Bila nama ganda, baris PERTAMA yang dipakai (sama dengan assignPolaKaryawan). */
function getPolaAssignmentMap_() {
  const map = {};
  readEmployeeRows_(getOrCreateNikSheet_(), EMP_NUM_COLS_).forEach(row => {
    const nama = cellText_(row[EMP_COL.NAMA - 1]);
    if (!nama) return;
    const key = nameKey_(nama);
    if (map.hasOwnProperty(key)) return;
    map[key] = {
      polaAssigned: cellText_(row[EMP_COL.POLA_ASSIGNED - 1]),
      offsetHari: parseInt(row[EMP_COL.OFFSET_HARI - 1], 10) || 0
    };
  });
  return map;
}

/* ============================ HARI LIBUR ============================ */

/**
 * Sheet "Hari Libur": kolom A = tanggal untuk seluruh tahun, baris 2 ke bawah.
 * Format yang dikenal: sel tanggal, "dd/MM/yyyy", "dd-MM-yyyy", "yyyy-MM-dd" (tahun 2 digit = 20xx).
 * Hasil: Set nomor tanggal (1..31) yang libur pada bulan periode.
 */
function getHariLiburDaysInPeriode_(periodText) {
  const sheet = getSpreadsheet_().getSheetByName(SHEET_HARI_LIBUR_);
  if (!sheet) return new Set();

  const my = parsePeriodMonthYear_(periodText);
  if (!my) return new Set();

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Set();

  const dayNumbers = new Set();
  sheet.getRange(2, 1, lastRow - 1, 1).getValues().forEach(r => {
    const d = parseTanggalLibur_(r[0]);
    if (d && d.getFullYear() === my.year && d.getMonth() === my.monthIndex) dayNumbers.add(d.getDate());
  });
  return dayNumbers;
}

function parseTanggalLibur_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  const s = String(v === null || v === undefined ? '' : v).trim();
  let y, mo, dd;
  let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) { dd = +m[1]; mo = +m[2]; y = +m[3]; }
  else if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/))) { y = +m[1]; mo = +m[2]; dd = +m[3]; }
  else return null;
  if (y < 100) y += 2000;
  const d = new Date(y, mo - 1, dd);
  // 31/02 jangan diam-diam menjadi 3 Maret.
  return (d.getFullYear() === y && d.getMonth() === mo - 1 && d.getDate() === dd) ? d : null;
}

/** 0 = Minggu ... 6 = Sabtu untuk tanggal ke-dayNumber pada periode my. */
function getWeekdayIndexOfDay_(my, dayNumber) {
  return new Date(my.year, my.monthIndex, dayNumber).getDay();
}

function modPositif_(n, m) {
  return ((n % m) + m) % m;
}

/* ============================ GENERATE JADWAL DARI POLA ============================ */

/**
 * mode 'kosongSaja' (default): hanya sel Jadwal yang masih kosong.
 * mode 'timpaSemua': timpa semua sel karyawan yang punya pola (merusak isian manual).
 * Hanya Super Admin. Hasil: state penuh (getAllDataInternal_) + generateJadwalInfo.
 */
function generateJadwalDariPola(mode, token) {
  const session = requireSuperAdmin_(token, 'Hanya akun dengan role "Super Admin" yang boleh generate Jadwal dari Pola.');
  mode = (mode === 'timpaSemua') ? 'timpaSemua' : 'kosongSaja';

  const jadwalSheet = getSheet_(CONFIG.SHEET_JADWAL);
  const periodText = String(jadwalSheet.getRange(2, 1).getValue() || '').trim();

  const info = { mode: mode, karyawanDigenerate: 0, selDiisi: 0, selDilewati: 0, karyawanTanpaPola: [] };

  withScriptLock_(() => {
    const lastRow = getLastEmployeeRow_(jadwalSheet);
    const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
    const dayCount = getDayColCount_(jadwalSheet);
    if (numRows <= 0 || dayCount <= 0) {
      throw new Error('Sheet Jadwal belum punya data karyawan / kolom hari untuk periode ini.');
    }

    const hariLiburSet = getHariLiburDaysInPeriode_(periodText);
    const myPeriod = parsePeriodMonthYear_(periodText);

    // Pola dengan kode yang sudah tidak valid hanya mempengaruhi karyawan yang memakainya.
    const polaByName = {};
    readPolaShiftRows_().forEach(p => {
      try { polaByName[nameKey_(p.namaPola)] = { urutan: validateUrutanKodePola_(p.urutanKode) }; }
      catch (e) { polaByName[nameKey_(p.namaPola)] = { error: e.message }; }
    });
    const assignMap = getPolaAssignmentMap_();

    const names = jadwalSheet.getRange(CONFIG.DATA_START_ROW, 2, numRows, 1).getValues();
    const gridRange = jadwalSheet.getRange(CONFIG.DATA_START_ROW, CONFIG.FIRST_DAY_COL, numRows, dayCount);
    const grid = gridRange.getValues();

    for (let i = 0; i < numRows; i++) {
      const nama = String(names[i][0] === null || names[i][0] === undefined ? '' : names[i][0]).trim();
      if (!nama) continue;

      const assign = assignMap[nameKey_(nama)];
      if (!assign || !assign.polaAssigned) { info.karyawanTanpaPola.push(nama); continue; }

      const pola = polaByName[nameKey_(assign.polaAssigned)];
      if (!pola) { info.karyawanTanpaPola.push(nama + ' (pola "' + assign.polaAssigned + '" tidak ditemukan)'); continue; }
      if (pola.error) { info.karyawanTanpaPola.push(nama + ' (pola "' + assign.polaAssigned + '" tidak valid: ' + pola.error + ')'); continue; }

      const urutan = pola.urutan;
      const mingguan = (urutan.length === 7 && !!myPeriod);

      info.karyawanDigenerate++;
      for (let d = 0; d < dayCount; d++) {
        const cell = grid[i][d];
        const terisi = cell !== null && cell !== undefined && String(cell).trim() !== '';
        if (terisi && mode === 'kosongSaja') { info.selDilewati++; continue; }

        const dayNumber = d + 1;
        let kode;
        if (hariLiburSet.has(dayNumber)) {
          kode = 'OFF';
        } else if (mingguan) {
          kode = urutan[modPositif_(getWeekdayIndexOfDay_(myPeriod, dayNumber) + assign.offsetHari, 7)];
        } else {
          kode = urutan[modPositif_(dayNumber - 1 + assign.offsetHari, urutan.length)];
        }
        grid[i][d] = kode;
        info.selDiisi++;
      }
    }

    if (info.selDiisi > 0) {
      gridRange.setValues(grid);
      SpreadsheetApp.flush();
    }
  }, 20000);

  // Di luar lock: syncOffFromJadwal_ memakai lock-nya sendiri.
  syncOffFromJadwal_();

  catatAuditLog_(
    session, 'Generate Jadwal dari Pola', 'Jadwal', periodText,
    'Mode: ' + mode + ', ' + info.karyawanDigenerate + ' karyawan, ' + info.selDiisi + ' sel diisi, ' + info.selDilewati + ' sel dilewati'
  );

  const result = getAllDataInternal_();
  result.generateJadwalInfo = {
    mode: info.mode,
    karyawanDigenerate: info.karyawanDigenerate,
    selDiisi: info.selDiisi,
    selDilewati: info.selDilewati,
    karyawanTanpaPola: info.karyawanTanpaPola
  };
  return result;
}

/* ============================ MIGRASI (SEKALI JALAN) ============================ */

/**
 * Versi lama modul ini menulis PolaAssigned/OffsetHari ke kolom 9-10 DataKaryawan, padahal di
 * struktur sekarang kolom itu adalah No HP dan FotoProfil. Fungsi ini memindahkan data pola
 * yang masih tertinggal di kolom 9-10 ke kolom 11-12.
 *
 * Aturan (hanya baris yang cocok yang dipindah, lainnya tidak disentuh):
 *  - kolom 9 berisi nama yang ada di sheet PolaShift DAN kolom 11 masih kosong -> pola dipindah;
 *  - bila itu terjadi dan kolom 10 berisi ANGKA -> itu OffsetHari, ikut dipindah.
 * Nomor HP dan URL foto tidak pernah cocok dengan aturan ini. No HP/FotoProfil yang sudah
 * telanjur tertimpa pola TIDAK bisa dipulihkan otomatis: isi ulang dari sumbernya.
 *
 * Panggil migrasiPolaDariKolom9dan10() = uji coba (hanya mencatat di Log).
 * Panggil migrasiPolaDariKolom9dan10(true) = benar-benar memindahkan.
 */
function migrasiPolaDariKolom9dan10(jalankan) {
  const sheet = getOrCreateNikSheet_();
  const polaNames = {};
  readPolaShiftRows_().forEach(p => { polaNames[nameKey_(p.namaPola)] = p.namaPola; });

  const rows = readEmployeeRows_(sheet, EMP_NUM_COLS_);
  const dipindah = [];
  rows.forEach((r, i) => {
    const kolom9 = cellText_(r[EMP_COL.NO_HP - 1]);
    if (!kolom9 || !polaNames.hasOwnProperty(nameKey_(kolom9))) return;
    if (cellText_(r[EMP_COL.POLA_ASSIGNED - 1])) return;
    const kolom10 = r[EMP_COL.FOTO_PROFIL - 1];
    dipindah.push({ baris: i + 2, nama: cellText_(r[EMP_COL.NAMA - 1]), pola: polaNames[nameKey_(kolom9)], offset: typeof kolom10 === 'number' ? kolom10 : null });
  });

  if (jalankan === true && dipindah.length) {
    withScriptLock_(() => {
      dipindah.forEach(x => {
        const polaCell = sheet.getRange(x.baris, EMP_COL.POLA_ASSIGNED);
        polaCell.setNumberFormat('@');
        polaCell.setValue(x.pola);
        if (x.offset !== null) {
          sheet.getRange(x.baris, EMP_COL.OFFSET_HARI).setValue(x.offset);
          sheet.getRange(x.baris, EMP_COL.FOTO_PROFIL).setValue('');
        }
        sheet.getRange(x.baris, EMP_COL.NO_HP).setValue('');
      });
      SpreadsheetApp.flush();
    }, 20000);
  }

  Logger.log((jalankan === true ? 'DIPINDAH: ' : 'UJI COBA (tidak ada yang diubah): ') + dipindah.length + ' baris.\n' +
    dipindah.map(x => '  baris ' + x.baris + ' ' + x.nama + ' -> pola "' + x.pola + '", offset ' + (x.offset === null ? '(kosong)' : x.offset)).join('\n'));
  return dipindah;
}
