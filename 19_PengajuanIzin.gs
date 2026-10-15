/*************************************************************
 * 19_PengajuanIzin.gs  (REBUILD)
 * -----------------------------------------------------------
 * Form pengajuan Izin/Sakit/Cuti dari karyawan (PUBLIK, tanpa
 * login admin) + approval oleh admin + auto-tandai Alfa harian
 * lewat trigger terjadwal.
 *
 * Sheet "PengajuanIzin" (hidden), kolom:
 * ID | Nama | NIK | Jenis | TglMulai | TglAkhir | Alasan |
 * LampiranURL | Status | DiajukanPada | DiprosesOleh |
 * DiprosesPada | CatatanAdmin
 *
 * Status: 'Pending' | 'Approved' | 'Rejected'
 *
 * CATATAN KEAMANAN: submitPengajuanIzin() & getKaryawanForPublicForm()
 * SENGAJA tidak butuh token (dipanggil dari halaman publik terpisah,
 * form-izin.html di Netlify). Identitas diverifikasi dengan mencocokkan
 * Nama + NIK terhadap sheet DataKaryawan.
 *
 * CATATAN PERBAIKAN (rebuild ini):
 * Versi sebelumnya sempat macet di UI ("Memuat..." tidak pernah
 * selesai) karena getPengajuanList() memperbaiki baris ber-ID kosong
 * dengan menulis SATU-PER-SATU ke sheet di dalam loop
 * (sheet.getRange(row,1).setValue(id)) — kalau ada banyak baris
 * berID kosong, ini bisa sangat lambat. Sekarang perbaikan ID
 * dikumpulkan dulu lalu ditulis SEKALI lewat setValues (batch),
 * jadi jauh lebih cepat dan tidak lagi berpotensi macet.
 *************************************************************/

const PENGAJUAN_SHEET_NAME_ = 'PengajuanIzin';
const PENGAJUAN_HEADER_ = ['ID', 'Nama', 'NIK', 'Jenis', 'TglMulai', 'TglAkhir', 'Alasan',
  'LampiranURL', 'Status', 'DiajukanPada', 'DiprosesOleh', 'DiprosesPada', 'CatatanAdmin'];
const PENGAJUAN_JENIS_KODE_ = { 'Izin': 'I', 'Sakit': 'S', 'Cuti': 'C' };
const PENGAJUAN_LAMPIRAN_FOLDER_NAME_ = 'AMS_Lampiran_Pengajuan';

/* ============================ SETUP SHEET ============================ */

/** Ambil (atau buat kalau belum ada) sheet PengajuanIzin di spreadsheet AKTIF script ini. */
function getOrCreatePengajuanSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PENGAJUAN_SHEET_NAME_);
  if (!sheet) {
    sheet = ss.insertSheet(PENGAJUAN_SHEET_NAME_);
    sheet.getRange(1, 1, 1, PENGAJUAN_HEADER_.length)
      .setValues([PENGAJUAN_HEADER_])
      .setFontWeight('bold')
      .setBackground('#f1f5f9');
    sheet.setColumnWidths(1, PENGAJUAN_HEADER_.length, 150);
    sheet.hideSheet();
  }
  return sheet;
}

/* ============================ PUBLIK (TANPA LOGIN) ============================ */

/** Daftar nama karyawan aktif (dari sheet Jadwal periode aktif), buat dropdown di form publik. */
function getKaryawanForPublicForm() {
  const jadwalSheet = getSheet_(CONFIG.SHEET_JADWAL);
  const lastRow = getLastEmployeeRow_(jadwalSheet);
  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  if (numRows <= 0) return [];
  return jadwalSheet.getRange(CONFIG.DATA_START_ROW, 2, numRows, 1).getValues()
    .map(r => String(r[0] || '').trim())
    .filter(n => n !== '')
    .sort();
}

/**
 * Terima pengajuan baru dari form publik. TIDAK butuh token — identitas
 * diverifikasi dengan mencocokkan nama+NIK terhadap sheet DataKaryawan.
 * Lampiran WAJIB (base64), disimpan ke folder Drive khusus dan link-nya
 * dicatat di sheet.
 */
function submitPengajuanIzin(nama, nik, jenis, tglMulai, tglAkhir, alasan, lampiranBase64, lampiranFilename, lampiranMimeType) {
  nama = String(nama || '').trim();
  nik = String(nik || '').trim();
  jenis = String(jenis || '').trim();
  alasan = String(alasan || '').trim();

  if (!nama) throw new Error('Nama wajib dipilih.');
  if (!nik) throw new Error('NIK wajib diisi.');
  if (!PENGAJUAN_JENIS_KODE_.hasOwnProperty(jenis)) {
    throw new Error('Jenis pengajuan tidak valid. Harus Izin, Sakit, atau Cuti.');
  }
  if (!tglMulai || !tglAkhir) throw new Error('Tanggal mulai dan tanggal akhir wajib diisi.');
  const dMulai = new Date(tglMulai);
  const dAkhir = new Date(tglAkhir);
  if (isNaN(dMulai.getTime()) || isNaN(dAkhir.getTime())) throw new Error('Format tanggal tidak valid.');
  if (dMulai.getTime() > dAkhir.getTime()) throw new Error('Tanggal mulai tidak boleh setelah tanggal akhir.');
  if (!alasan) throw new Error('Alasan wajib diisi.');
  if (!lampiranBase64) throw new Error('Lampiran wajib diupload (mis. surat dokter untuk Sakit, atau dokumen pendukung lainnya).');

  // Verifikasi Nama + NIK cocok dengan data karyawan (getNikMap_ dari 12_Employee.gs)
  const nikMap = getNikMap_(); // nama(lowercase) -> nik
  const nikTerdaftar = nikMap[nama.toLowerCase()];
  if (!nikTerdaftar) {
    throw new Error('NIK untuk nama "' + nama + '" belum terdaftar di sistem. Hubungi admin untuk melengkapi data karyawan terlebih dahulu.');
  }
  if (String(nikTerdaftar).trim() !== nik) {
    throw new Error('Nama dan NIK tidak cocok dengan data karyawan terdaftar. Periksa kembali.');
  }

  const lampiranUrl = saveLampiranToDrive_(lampiranBase64, lampiranFilename, lampiranMimeType);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getOrCreatePengajuanSheet_();
    const id = generatePengajuanId_();
    sheet.appendRow([
      id, nama, nik, jenis,
      Utilities.formatDate(dMulai, 'GMT+7', 'yyyy-MM-dd'),
      Utilities.formatDate(dAkhir, 'GMT+7', 'yyyy-MM-dd'),
      alasan, lampiranUrl, 'Pending', new Date(), '', '', ''
    ]);
    SpreadsheetApp.flush();
    return { success: true, id: id };
  } finally {
    lock.releaseLock();
  }
}

function generatePengajuanId_() {
  return 'PGJ' + Date.now() + Math.floor(Math.random() * 1000);
}

function getOrCreateLampiranFolder_() {
  const folders = DriveApp.getFoldersByName(PENGAJUAN_LAMPIRAN_FOLDER_NAME_);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(PENGAJUAN_LAMPIRAN_FOLDER_NAME_);
}

function saveLampiranToDrive_(base64Data, filename, mimeType) {
  const folder = getOrCreateLampiranFolder_();
  const blob = Utilities.newBlob(
    Utilities.base64Decode(base64Data),
    mimeType || 'application/octet-stream',
    filename || ('lampiran_' + Date.now())
  );
  const file = folder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    // Domain Google Workspace kadang membatasi sharing publik; link tetap
    // tersimpan (bisa dibuka oleh yang punya akses di dalam organisasi),
    // ini tidak menggagalkan proses submit.
    Logger.log('Gagal set sharing publik utk lampiran: ' + e.message);
  }
  return file.getUrl();
}

/* ============================ ADMIN (BUTUH LOGIN) ============================ */

/**
 * Daftar pengajuan, opsional filter status ('Pending'/'Approved'/'Rejected'/'ALL').
 * PERBAIKAN: perbaikan baris berID kosong (data lama/rusak) sekarang dikumpulkan
 * dan ditulis SEKALI (batch setValues), bukan satu-per-satu di dalam loop —
 * ini yang sebelumnya bisa membuat UI macet di "Memuat..." kalau datanya banyak.
 */
function getPengajuanList(status, token) {
  const session = verifySession_(token);
  if (   !isSuperAdminRole_(session.role)) throw new Error('Hanya admin yang boleh melihat daftar pengajuan.');

  const sheet = getOrCreatePengajuanSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const range = sheet.getRange(2, 1, lastRow - 1, PENGAJUAN_HEADER_.length);
  const data = range.getValues();

  const rawStatus = String(status || '').trim();
  const filterStatus = (!rawStatus || rawStatus.toUpperCase() === 'ALL' || rawStatus.toUpperCase() === 'SEMUA')
    ? 'ALL' : rawStatus;

  const result = [];
  let idsFixed = false;

  data.forEach((r, i) => {
    const nama = String(r[1] || '').trim();
    if (!nama) return; // baris benar-benar kosong, lewati

    // Perbaikan ID kosong (data lama/rusak): generate ID baru DI MEMORI dulu,
    // lalu ditulis balik ke `data` supaya batch-write di akhir bisa menuliskannya
    // sekaligus (tidak menulis ke sheet satu-per-satu di sini).
    let id = String(r[0] || '').trim();
    if (!id) {
      id = generatePengajuanId_() + i;
      data[i][0] = id;
      idsFixed = true;
    }

    const rowStatus = String(r[8] || '').trim();
    if (filterStatus !== 'ALL' && rowStatus !== filterStatus) return;

    
    // PENTING: Sheets bisa otomatis mengonversi string tanggal "yyyy-MM-dd" yang
// ditulis submitPengajuanIzin() menjadi tipe Date asli saat disimpan. Kalau ini
// terjadi, mengembalikan objek Date apa adanya akan membuat JSON.stringify()
// memanggil .toISOString() (konversi ke UTC) -> tanggal tampak mundur 1 hari
// dan selalu berjam 17:00:00.000Z (karena WIB = UTC+7). Maka di sini selalu
// diformat ulang jadi string yyyy-MM-dd di zona waktu GMT+7, apa pun tipe aslinya.
const tglMulaiStr = r[4] instanceof Date ? Utilities.formatDate(r[4], 'GMT+7', 'yyyy-MM-dd') : String(r[4] || '');
const tglAkhirStr = r[5] instanceof Date ? Utilities.formatDate(r[5], 'GMT+7', 'yyyy-MM-dd') : String(r[5] || '');

result.push({
  id: id, nama: r[1], nik: r[2], jenis: r[3],
  tglMulai: tglMulaiStr, tglAkhir: tglAkhirStr, alasan: r[6], lampiranUrl: r[7],
  status: rowStatus,
  diajukanPada: r[9] instanceof Date ? r[9].toISOString() : String(r[9] || ''),
  diprosesOleh: r[10], diprosesPada: r[11] instanceof Date ? r[11].toISOString() : String(r[11] || ''),
  catatanAdmin: r[12]
});
});

  if (idsFixed) {
    // Satu kali tulis balik seluruh blok data (termasuk ID yang baru diperbaiki).
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      range.setValues(data);
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
  }

  return result.reverse(); // terbaru duluan
  }

/**
 * Wrapper "aman" dari getPengajuanList: frontend (loadPengajuanList_ di
 * JavaScript.html) memanggil fungsi INI, dan mengharapkan hasilnya berupa
 * TEKS JSON (bukan array/object mentah), karena di sisi frontend hasilnya
 * di-parse manual pakai JSON.parse().
 */
function getPengajuanListSafe(status, token) {
  const result = getPengajuanList(status, token);
  return JSON.stringify(result);
}

function findPengajuanRow_(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r => String(r[0]));
  const idx = ids.indexOf(String(id));
  return idx === -1 ? -1 : idx + 2;
}

/** Setujui pengajuan: tulis kode I/S/C ke Realisasi utk rentang tanggalnya, lalu tandai status Approved. */
function approvePengajuan(id, token) {
  const session = verifySession_(token);
  if (   !isSuperAdminRole_(session.role)) throw new Error('Hanya admin yang boleh menyetujui pengajuan.');

  const sheet = getOrCreatePengajuanSheet_();
  const rowNum = findPengajuanRow_(sheet, id);
  if (rowNum === -1) throw new Error('Pengajuan tidak ditemukan.');

  const row = sheet.getRange(rowNum, 1, 1, PENGAJUAN_HEADER_.length).getValues()[0];
  const [, nama, , jenis, tglMulai, tglAkhir, , , status] = row;
  if (status !== 'Pending') throw new Error('Pengajuan ini sudah diproses sebelumnya (status: ' + status + ').');

  const applyResult = applyPengajuanToRealisasi_(nama, jenis, tglMulai, tglAkhir);

  sheet.getRange(rowNum, 9).setValue('Approved');
  sheet.getRange(rowNum, 11).setValue(session.username);
  sheet.getRange(rowNum, 12).setValue(new Date());
  SpreadsheetApp.flush();

  kirimEmailNotifikasiPengajuan_(nama, jenis, tglMulai, tglAkhir, 'Approved', '');

  // Balikin daftar Pending saja (bukan ALL) supaya baris yang baru
  // di-approve otomatis hilang dari tabel "Menunggu Persetujuan" di frontend.
  const result = getPengajuanList('Pending', token);
  return { list: result, applyInfo: applyResult };
}

/** Tolak pengajuan, dengan catatan alasan penolakan. */
function rejectPengajuan(id, catatan, token) {
  const session = verifySession_(token);
  if (   !isSuperAdminRole_(session.role)) throw new Error('Hanya admin yang boleh menolak pengajuan.');

  const sheet = getOrCreatePengajuanSheet_();
  const rowNum = findPengajuanRow_(sheet, id);
  if (rowNum === -1) throw new Error('Pengajuan tidak ditemukan.');

  const status = sheet.getRange(rowNum, 9).getValue();
  if (status !== 'Pending') throw new Error('Pengajuan ini sudah diproses sebelumnya (status: ' + status + ').');

  sheet.getRange(rowNum, 9).setValue('Rejected');
  sheet.getRange(rowNum, 11).setValue(session.username);
  sheet.getRange(rowNum, 12).setValue(new Date());
  sheet.getRange(rowNum, 13).setValue(String(catatan || '').trim());
  SpreadsheetApp.flush();

  const rowFull = sheet.getRange(rowNum, 1, 1, PENGAJUAN_HEADER_.length).getValues()[0];
  kirimEmailNotifikasiPengajuan_(rowFull[1], rowFull[3], rowFull[4], rowFull[5], 'Rejected', catatan);

  return getPengajuanList('Pending', token);
}

/** Hitung ringan jumlah pengajuan Pending saja — untuk badge notifikasi di sidebar. */
function getPengajuanPendingCount(token) {
  const session = verifySession_(token);
  if (   !isSuperAdminRole_(session.role)) return { count: 0 };

  const sheet = getOrCreatePengajuanSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { count: 0 };

  const namaValues = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  const statusValues = sheet.getRange(2, 9, lastRow - 1, 1).getValues();

  let count = 0;
  for (let i = 0; i < statusValues.length; i++) {
    const nama = String(namaValues[i][0] || '').trim();
    const status = String(statusValues[i][0] || '').trim();
    if (nama && status === 'Pending') count++;
  }
  return { count: count };
}

/**
 * Tulis kode I/S/C ke sheet Realisasi (periode AKTIF saja) utk rentang
 * tglMulai..tglAkhir milik satu karyawan. Tanggal yang jatuh di LUAR
 * bulan/tahun periode aktif dilaporkan sebagai outOfRangeDays (tidak
 * ditulis, karena sheet Realisasi cuma sejumlah dayCount kolom).
 */
function applyPengajuanToRealisasi_(nama, jenis, tglMulaiStr, tglAkhirStr) {
  const kode = PENGAJUAN_JENIS_KODE_[jenis];
  if (!kode) throw new Error('Jenis pengajuan "' + jenis + '" tidak dikenali.');

  const realisasiSheet = getSheet_(CONFIG.SHEET_REALISASI);
  const periodText = String(realisasiSheet.getRange(2, 1).getValue() || '').trim();
  const my = parsePeriodMonthYear_(periodText); // dari 10_AttendanceLog.gs
  if (!my) {
    throw new Error('Periode aktif ("' + periodText + '") tidak bisa dibaca bulan/tahunnya, tidak bisa menerapkan pengajuan ke Realisasi.');
  }

  const lastRow = getLastEmployeeRow_(realisasiSheet);
  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const names = realisasiSheet.getRange(CONFIG.DATA_START_ROW, 2, numRows, 1).getValues()
    .map(r => String(r[0] || '').trim().toLowerCase());
  const idx = names.indexOf(String(nama).trim().toLowerCase());
  if (idx === -1) {
    throw new Error('Karyawan "' + nama + '" tidak ditemukan di sheet Realisasi periode aktif.');
  }
  const rowIndex = CONFIG.DATA_START_ROW + idx;
  const dayCount = getDayColCount_(realisasiSheet);

  const dMulai = new Date(tglMulaiStr);
  const dAkhir = new Date(tglAkhirStr);
  const appliedDays = [];
  const outOfRangeDays = [];

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    let cursor = new Date(dMulai);
    let safety = 0;
    while (cursor.getTime() <= dAkhir.getTime() && safety < 62) {
      safety++;
      if (cursor.getFullYear() === my.year && cursor.getMonth() === my.monthIndex) {
        const dayNum = cursor.getDate();
        if (dayNum >= 1 && dayNum <= dayCount) {
          const col = CONFIG.FIRST_DAY_COL + (dayNum - 1);
          realisasiSheet.getRange(rowIndex, col).setValue(kode);
          appliedDays.push(dayNum);
        } else {
          outOfRangeDays.push(Utilities.formatDate(cursor, 'GMT+7', 'dd/MM/yyyy'));
        }
      } else {
        outOfRangeDays.push(Utilities.formatDate(cursor, 'GMT+7', 'dd/MM/yyyy'));
      }
      cursor.setDate(cursor.getDate() + 1);
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  return { appliedDays: appliedDays, outOfRangeDays: outOfRangeDays };
}

/** Kirim email notifikasi hasil pengajuan (Approved/Rejected) ke karyawan, kalau emailnya terdaftar. */
function kirimEmailNotifikasiPengajuan_(nama, jenis, tglMulai, tglAkhir, statusBaru, catatanAdmin) {
  try {
    const tglMulaiFmt = tglMulai instanceof Date ? Utilities.formatDate(tglMulai, 'GMT+7', 'dd/MM/yyyy') : String(tglMulai);
    const tglAkhirFmt = tglAkhir instanceof Date ? Utilities.formatDate(tglAkhir, 'GMT+7', 'dd/MM/yyyy') : String(tglAkhir);
    kirimEmailNotifikasiPengajuan_(nama, jenis, tglMulaiFmt, tglAkhirFmt, 'Approved', '');
    const masterMap = getKaryawanMasterMap_();
    const master = masterMap[String(nama).trim().toLowerCase()];
    const email = master ? String(master.email || '').trim() : '';
    if (!email) {
      Logger.log('kirimEmailNotifikasiPengajuan_: email untuk "' + nama + '" belum terdaftar, notifikasi dilewati.');
      return;
    }

    const isApproved = (statusBaru === 'Approved');
    const subject = (isApproved ? '✅ Pengajuan ' : '❌ Pengajuan ') + jenis + ' Anda ' + (isApproved ? 'Disetujui' : 'Ditolak');

    let htmlBody = '' +
      '<div style="font-family:Arial,Helvetica,sans-serif;max-width:600px;margin:0 auto;color:#1e293b;">' +
        '<h2 style="color:' + (isApproved ? '#16a34a' : '#ef4444') + ';margin-bottom:4px;">' +
          (isApproved ? '✅ Pengajuan Disetujui' : '❌ Pengajuan Ditolak') +
        '</h2>' +
        '<p>Halo <strong>' + escapeHtmlServer_(nama) + '</strong>,</p>' +
        '<p>Pengajuan <strong>' + escapeHtmlServer_(jenis) + '</strong> Anda untuk tanggal ' +
          '<strong>' + escapeHtmlServer_(tglMulaiFmt) + '</strong> s/d <strong>' + escapeHtmlServer_(tglAkhirFmt) + '</strong> ' +
          'telah <strong>' + (isApproved ? 'DISETUJUI' : 'DITOLAK') + '</strong>.</p>';

    if (!isApproved && catatanAdmin) {
      htmlBody += '<p style="background:#fef2f2;border-left:4px solid #ef4444;padding:10px 14px;">' +
        '<strong>Catatan dari admin:</strong><br>' + escapeHtmlServer_(catatanAdmin) + '</p>';
    }

    htmlBody += '' +
        '<p style="color:#94a3b8;font-size:12px;margin-top:24px;">Email ini dikirim otomatis oleh Attendance Monitoring System (AMS) - Departemen Housekeeping PT Swadharma Primautama.</p>' +
      '</div>';

    MailApp.sendEmail({ to: email, subject: subject, htmlBody: htmlBody });
  } catch (e) {
    // Jangan sampai gagal kirim email menggagalkan proses approve/reject.
    Logger.log('kirimEmailNotifikasiPengajuan_ gagal: ' + e.message);
  }
}

/* ============================ AUTO-TANDAI ALFA (TRIGGER HARIAN) ============================ */

/**
 * Dipanggil otomatis tiap hari (lihat setupAutoAlfaTrigger()). Untuk
 * TANGGAL HARI INI: kalau seorang karyawan ada jadwal (bukan OFF/kosong)
 * tapi sel Realisasi-nya masih kosong, DAN tidak ada pengajuan Izin/
 * Sakit/Cuti (Pending ATAU Approved) yang mencakup tanggal ini, maka sel
 * itu otomatis ditandai "A" (Alfa).
 */
function autoTandaiAlfaHarian_() {
  const jadwalSheet = getSheet_(CONFIG.SHEET_JADWAL);
  const realisasiSheet = getSheet_(CONFIG.SHEET_REALISASI);
  const periodText = String(realisasiSheet.getRange(2, 1).getValue() || '').trim();
  const my = parsePeriodMonthYear_(periodText);
  if (!my) {
    Logger.log('autoTandaiAlfaHarian_: periode aktif tidak terbaca, dilewati.');
    return;
  }

  const today = new Date();
  if (today.getFullYear() !== my.year || today.getMonth() !== my.monthIndex) {
    Logger.log('autoTandaiAlfaHarian_: hari ini di luar bulan periode aktif, dilewati.');
    return;
  }
  const dayNum = today.getDate();

  const lastRow = getLastEmployeeRow_(jadwalSheet);
  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const dayCount = getDayColCount_(jadwalSheet);
  if (numRows <= 0 || dayNum > dayCount) return;

  const col = CONFIG.FIRST_DAY_COL + (dayNum - 1);

  const jadwalCol = jadwalSheet.getRange(CONFIG.DATA_START_ROW, col, numRows, 1).getValues();
  const realisasiCol = realisasiSheet.getRange(CONFIG.DATA_START_ROW, col, numRows, 1).getValues();
  const names = realisasiSheet.getRange(CONFIG.DATA_START_ROW, 2, numRows, 1).getValues();

  // Kumpulkan nama yang punya pengajuan Pending/Approved mencakup HARI INI.
  const namaTerlindungi = getNamaDenganPengajuanAktifPadaTanggal_(today);

  let ditandai = 0;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    for (let i = 0; i < numRows; i++) {
      const jadwalKode = String(jadwalCol[i][0] || '').trim().toUpperCase();
      const realKosong = (realisasiCol[i][0] === '' || realisasiCol[i][0] === null || realisasiCol[i][0] === undefined);
      const nama = String(names[i][0] || '').trim();

      if (jadwalKode && jadwalKode !== 'OFF' && realKosong && nama && !namaTerlindungi.has(nama.toLowerCase())) {
        realisasiSheet.getRange(CONFIG.DATA_START_ROW + i, col).setValue('A');
        ditandai++;
      }
    }
    if (ditandai > 0) SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  Logger.log('autoTandaiAlfaHarian_: ' + ditandai + ' karyawan ditandai Alfa untuk tanggal ' +
    Utilities.formatDate(today, 'GMT+7', 'dd/MM/yyyy') + '.');
}

/** Set nama(lowercase) yang punya pengajuan Pending/Approved yang mencakup tanggal tertentu. */
function getNamaDenganPengajuanAktifPadaTanggal_(targetDate) {
  const set = new Set();
  const sheet = getOrCreatePengajuanSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return set;

  const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues(); // sampai kolom Status
  const targetStr = Utilities.formatDate(targetDate, 'GMT+7', 'yyyy-MM-dd');

  data.forEach(r => {
    const nama = String(r[1] || '').trim().toLowerCase();
    const tglMulai = String(r[4] || '');
    const tglAkhir = String(r[5] || '');
    const status = String(r[8] || '');
    if (!nama || (status !== 'Pending' && status !== 'Approved')) return;
    if (tglMulai <= targetStr && targetStr <= tglAkhir) set.add(nama);
  });

  return set;
}

/**
 * JALANKAN SEKALI SAJA dari Apps Script Editor untuk memasang trigger
 * otomatis yang menandai Alfa setiap hari jam 23:00. Aman dijalankan
 * berkali-kali (trigger lama dengan handler sama dihapus dulu).
 */
function setupAutoAlfaTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'autoTandaiAlfaHarian_') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('autoTandaiAlfaHarian_')
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .create();

  Logger.log('Trigger auto-Alfa berhasil dipasang: setiap hari jam 23:00.');
}

/**
 * Utilitas debug manual: jalankan LANGSUNG dari editor (tombol ▶) untuk
 * ngecek isi sheet PengajuanIzin apa adanya, tanpa lewat token/session/
 * web app sama sekali. Hasilnya muncul di panel Log Eksekusi. Berguna
 * kalau modul ini bermasalah lagi di kemudian hari.
 */
function debugPengajuanRaw() {
  const sheet = getOrCreatePengajuanSheet_();
  const lastRow = sheet.getLastRow();
  Logger.log('Spreadsheet aktif: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl());
  Logger.log('Nama sheet: ' + sheet.getName() + ', lastRow: ' + lastRow);
  if (lastRow < 2) { Logger.log('Tidak ada data.'); return; }
  const data = sheet.getRange(2, 1, lastRow - 1, PENGAJUAN_HEADER_.length).getValues();
  data.forEach((r, i) => {
    Logger.log('Baris ' + (i + 2) + ': ID=[' + r[0] + '] Nama=[' + r[1] + '] Status=[' + r[8] + ']');
  });
}