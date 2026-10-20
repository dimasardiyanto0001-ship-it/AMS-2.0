/*************************************************************
 * 19_PengajuanIzin.gs
 * Form pengajuan Izin/Sakit/Cuti dari karyawan (PUBLIK, tanpa login)
 * + persetujuan oleh admin + auto-tandai Alfa harian (trigger).
 *
 * Sheet "PengajuanIzin" (tersembunyi), kolom:
 *   ID | Nama | NIK | Jenis | TglMulai | TglAkhir | Alasan |
 *   LampiranURL | Status | DiajukanPada | DiprosesOleh |
 *   DiprosesPada | CatatanAdmin
 * Status: 'Pending' | 'Approved' | 'Rejected'
 *
 * KEAMANAN: submitPengajuanIzin() dan getKaryawanForPublicForm()
 * SENGAJA tanpa token (dipanggil dari form-izin.html di Netlify).
 * Identitas = Nama + NIK yang dicocokkan dengan sheet DataKaryawan.
 * Karena publik, semua teks bebas (Alasan) dijaga agar tidak bisa
 * menjadi rumus Sheets, dan ukuran masukan dibatasi.
 *************************************************************/

const PENGAJUAN_SHEET_NAME_ = 'PengajuanIzin';
const PENGAJUAN_HEADER_ = ['ID', 'Nama', 'NIK', 'Jenis', 'TglMulai', 'TglAkhir', 'Alasan',
  'LampiranURL', 'Status', 'DiajukanPada', 'DiprosesOleh', 'DiprosesPada', 'CatatanAdmin'];
const PENGAJUAN_JENIS_KODE_ = { 'Izin': 'I', 'Sakit': 'S', 'Cuti': 'C' };
const PENGAJUAN_LAMPIRAN_FOLDER_NAME_ = 'AMS_Lampiran_Pengajuan';
const PENGAJUAN_TZ_ = 'GMT+7';
const PENGAJUAN_ALASAN_MAX_ = 1000;
const PENGAJUAN_CATATAN_MAX_ = 500;
const PENGAJUAN_RENTANG_MAX_HARI_ = 62;                 // batas yang sama dengan penerapan ke Realisasi
const PENGAJUAN_LAMPIRAN_MAX_BYTES_ = 10 * 1024 * 1024; // 10 MB setelah didekode

/* ============================ SETUP SHEET ============================ */

/** Ambil (atau buat bila belum ada) sheet PengajuanIzin di spreadsheet aktif. */
function getOrCreatePengajuanSheet_() {
  const ss = getSpreadsheet_();
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

/* ============================ HELPER ============================ */

function generatePengajuanId_() {
  return 'PGJ' + Date.now() + Math.floor(Math.random() * 1000);
}

/**
 * Isi sel tanggal -> 'yyyy-MM-dd'. Sheets sering mengubah teks "2026-10-05" jadi
 * tanggal asli saat disimpan; tanggal asli diformat di GMT+7 (bukan ISO/UTC,
 * yang membuat tanggal mundur 1 hari), teks dipakai apa adanya.
 */
function pengajuanDateStr_(value) {
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? '' : Utilities.formatDate(value, PENGAJUAN_TZ_, 'yyyy-MM-dd');
  }
  return String(value === null || value === undefined ? '' : value).trim();
}

/** 'yyyy-MM-dd' -> 'dd/MM/yyyy' untuk tampilan; teks lain dikembalikan apa adanya. */
function formatTanggalPengajuan_(value) {
  const s = pengajuanDateStr_(value);
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? m[3] + '/' + m[2] + '/' + m[1] : s;
}

/** 'yyyy-MM-dd' -> {y, m(1-12), d}, atau null bila bukan tanggal yang sah. */
function parsePengajuanYmd_(str) {
  const m = String(str || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;
  const y = parseInt(m[1], 10), mo = parseInt(m[2], 10), d = parseInt(m[3], 10);
  const check = new Date(y, mo - 1, d);
  if (check.getFullYear() !== y || check.getMonth() !== mo - 1 || check.getDate() !== d) return null;
  return { y: y, m: mo, d: d };
}

/** Super Admin dan Admin (supervisor); tidak lewat matrix Hak Akses. */
function requirePengajuanAdmin_(token, pesan) {
  return requireUploadAccess_(token, pesan);
}

/* ============================ PUBLIK (TANPA LOGIN) ============================ */

/** Daftar nama karyawan di sheet Jadwal periode aktif (untuk pilihan nama di form publik). */
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
 * Terima pengajuan baru dari form publik. Lampiran WAJIB (base64), disimpan ke
 * folder Drive khusus dan tautannya dicatat di sheet.
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

  // Batas masukan (form publik).
  if (alasan.length > PENGAJUAN_ALASAN_MAX_) throw new Error('Alasan terlalu panjang (maksimal ' + PENGAJUAN_ALASAN_MAX_ + ' karakter).');
  const jumlahHari = Math.round((dAkhir.getTime() - dMulai.getTime()) / 86400000) + 1;
  if (jumlahHari > PENGAJUAN_RENTANG_MAX_HARI_) {
    throw new Error('Rentang pengajuan maksimal ' + PENGAJUAN_RENTANG_MAX_HARI_ + ' hari. Pisahkan menjadi beberapa pengajuan.');
  }
  const lampiranData = stripDataUrlPrefix_(lampiranBase64);
  if (Math.floor(lampiranData.length * 3 / 4) > PENGAJUAN_LAMPIRAN_MAX_BYTES_) {
    throw new Error('Ukuran lampiran maksimal ' + (PENGAJUAN_LAMPIRAN_MAX_BYTES_ / 1024 / 1024) + ' MB.');
  }

  // Verifikasi Nama + NIK terhadap DataKaryawan.
  const nikTerdaftar = getNikMap_()[nama.toLowerCase()];
  if (!nikTerdaftar) {
    throw new Error('NIK untuk nama "' + nama + '" belum terdaftar di sistem. Hubungi admin untuk melengkapi data karyawan terlebih dahulu.');
  }
  if (String(nikTerdaftar).trim() !== nik) {
    throw new Error('Nama dan NIK tidak cocok dengan data karyawan terdaftar. Periksa kembali.');
  }

  const lampiranUrl = saveLampiranToDrive_(lampiranData, lampiranFilename, lampiranMimeType);

  return withScriptLock_(() => {
    const sheet = getOrCreatePengajuanSheet_();
    const id = generatePengajuanId_();
    sheet.appendRow([
      id, nama, nik, jenis,
      Utilities.formatDate(dMulai, PENGAJUAN_TZ_, 'yyyy-MM-dd'),
      Utilities.formatDate(dAkhir, PENGAJUAN_TZ_, 'yyyy-MM-dd'),
      guardSheetText_(alasan, PENGAJUAN_ALASAN_MAX_), lampiranUrl, 'Pending', new Date(), '', '', ''
    ]);
    SpreadsheetApp.flush();
    return { success: true, id: id };
  }, 15000);
}

/** Buang awalan "data:...;base64," bila ada (tanda tangan Base64 mentah dibiarkan). */
function stripDataUrlPrefix_(text) {
  return String(text).replace(/^data:[^;,]*;base64,/, '');
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
    // Domain Workspace kadang melarang sharing publik; tautan tetap tersimpan
    // (bisa dibuka oleh yang punya akses), pengajuan tidak digagalkan.
    Logger.log('Gagal set sharing publik utk lampiran: ' + e.message);
  }
  return file.getUrl();
}

/* ============================ ADMIN (BUTUH LOGIN) ============================ */

/**
 * Daftar pengajuan, opsional filter status ('Pending'/'Approved'/'Rejected'/'ALL'),
 * terbaru di atas. Baris ber-ID kosong (data lama/rusak) diberi ID baru dan
 * ditulis SEKALI (hanya kolom ID) setelah semua baris dibaca.
 */
function getPengajuanList(status, token) {
  requirePengajuanAdmin_(token, 'Hanya admin yang boleh melihat daftar pengajuan.');

  const sheet = getOrCreatePengajuanSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, PENGAJUAN_HEADER_.length).getValues();

  const rawStatus = String(status || '').trim();
  const filterStatus = (!rawStatus || rawStatus.toUpperCase() === 'ALL' || rawStatus.toUpperCase() === 'SEMUA')
    ? 'ALL' : rawStatus;

  const result = [];
  const idColumn = data.map(r => [r[0]]);
  let idsFixed = false;

  data.forEach((r, i) => {
    const nama = String(r[1] || '').trim();
    if (!nama) return; // baris benar-benar kosong

    let id = String(r[0] || '').trim();
    if (!id) {
      id = generatePengajuanId_() + i;
      idColumn[i][0] = id;
      idsFixed = true;
    }

    const rowStatus = String(r[8] || '').trim();
    if (filterStatus !== 'ALL' && rowStatus !== filterStatus) return;

    result.push({
      id: id, nama: r[1], nik: r[2], jenis: r[3],
      tglMulai: pengajuanDateStr_(r[4]), tglAkhir: pengajuanDateStr_(r[5]), alasan: r[6], lampiranUrl: r[7],
      status: rowStatus,
      diajukanPada: r[9] instanceof Date ? r[9].toISOString() : String(r[9] || ''),
      diprosesOleh: r[10], diprosesPada: r[11] instanceof Date ? r[11].toISOString() : String(r[11] || ''),
      catatanAdmin: r[12]
    });
  });

  if (idsFixed) {
    withScriptLock_(() => {
      sheet.getRange(2, 1, idColumn.length, 1).setValues(idColumn);
      SpreadsheetApp.flush();
    }, 10000);
  }

  return result.reverse();
}

/**
 * Pembungkus: frontend (loadPengajuanList_) memanggil fungsi INI dan
 * mengharapkan TEKS JSON yang di-parse sendiri.
 */
function getPengajuanListSafe(status, token) {
  return JSON.stringify(getPengajuanList(status, token));
}

function findPengajuanRow_(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r => String(r[0]));
  const idx = ids.indexOf(String(id));
  return idx === -1 ? -1 : idx + 2;
}

/**
 * Setujui pengajuan: tulis kode I/S/C ke Realisasi untuk rentang tanggalnya,
 * lalu tandai Approved. Pemeriksaan status, penulisan Realisasi, dan penandaan
 * berjalan di dalam satu lock, jadi dua admin yang menyetujui bersamaan tidak
 * bisa memproses pengajuan yang sama dua kali.
 */
function approvePengajuan(id, token) {
  const session = requirePengajuanAdmin_(token, 'Hanya admin yang boleh menyetujui pengajuan.');

  const done = withScriptLock_(() => {
    const sheet = getOrCreatePengajuanSheet_();
    const rowNum = findPengajuanRow_(sheet, id);
    if (rowNum === -1) throw new Error('Pengajuan tidak ditemukan.');

    const row = sheet.getRange(rowNum, 1, 1, PENGAJUAN_HEADER_.length).getValues()[0];
    const nama = row[1], jenis = row[3], status = row[8];
    if (status !== 'Pending') throw new Error('Pengajuan ini sudah diproses sebelumnya (status: ' + status + ').');

    const tglMulai = pengajuanDateStr_(row[4]);
    const tglAkhir = pengajuanDateStr_(row[5]);
    const applyResult = applyPengajuanToRealisasi_(nama, jenis, tglMulai, tglAkhir);

    sheet.getRange(rowNum, 9).setValue('Approved');
    sheet.getRange(rowNum, 11, 1, 2).setValues([[session.username, new Date()]]);
    SpreadsheetApp.flush();
    return { nama: nama, jenis: jenis, tglMulai: tglMulai, tglAkhir: tglAkhir, applyResult: applyResult };
  }, 20000);

  catatAuditLog_(session, 'Approve', 'Pengajuan Izin', id,
    done.nama + ' - ' + done.jenis + ' ' + done.tglMulai + ' s/d ' + done.tglAkhir +
    '; hari diterapkan: ' + done.applyResult.appliedDays.length +
    (done.applyResult.outOfRangeDays.length ? '; di luar periode aktif: ' + done.applyResult.outOfRangeDays.length : ''));

  kirimEmailNotifikasiPengajuan_(done.nama, done.jenis, done.tglMulai, done.tglAkhir, 'Approved', '');

  // Kembalikan daftar Pending supaya baris yang baru disetujui hilang dari tabel.
  return { list: getPengajuanList('Pending', token), applyInfo: done.applyResult };
}

/** Tolak pengajuan, dengan catatan alasan penolakan. */
function rejectPengajuan(id, catatan, token) {
  const session = requirePengajuanAdmin_(token, 'Hanya admin yang boleh menolak pengajuan.');
  const catatanBersih = String(catatan || '').trim().substring(0, PENGAJUAN_CATATAN_MAX_);

  const done = withScriptLock_(() => {
    const sheet = getOrCreatePengajuanSheet_();
    const rowNum = findPengajuanRow_(sheet, id);
    if (rowNum === -1) throw new Error('Pengajuan tidak ditemukan.');

    const status = sheet.getRange(rowNum, 9).getValue();
    if (status !== 'Pending') throw new Error('Pengajuan ini sudah diproses sebelumnya (status: ' + status + ').');

    sheet.getRange(rowNum, 9).setValue('Rejected');
    sheet.getRange(rowNum, 11, 1, 2).setValues([[session.username, new Date()]]);
    sheet.getRange(rowNum, 13).setValue(guardSheetText_(catatanBersih, PENGAJUAN_CATATAN_MAX_));
    SpreadsheetApp.flush();

    const rowFull = sheet.getRange(rowNum, 1, 1, PENGAJUAN_HEADER_.length).getValues()[0];
    return { nama: rowFull[1], jenis: rowFull[3], tglMulai: pengajuanDateStr_(rowFull[4]), tglAkhir: pengajuanDateStr_(rowFull[5]) };
  }, 15000);

  catatAuditLog_(session, 'Reject', 'Pengajuan Izin', id,
    done.nama + ' - ' + done.jenis + ' ' + done.tglMulai + ' s/d ' + done.tglAkhir +
    (catatanBersih ? '; catatan: ' + catatanBersih : ''));

  kirimEmailNotifikasiPengajuan_(done.nama, done.jenis, done.tglMulai, done.tglAkhir, 'Rejected', catatanBersih);

  return getPengajuanList('Pending', token);
}

/** Jumlah pengajuan Pending (badge notifikasi sidebar). Non-admin: 0. */
function getPengajuanPendingCount(token) {
  const session = verifySession_(token);
  if (!isSuperAdminRole_(session.role) && !isAdminRole_(session.role)) return { count: 0 };

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
 * Tulis kode I/S/C ke Realisasi (periode AKTIF saja) untuk rentang tglMulai..tglAkhir
 * ('yyyy-MM-dd') milik satu karyawan. Tanggal di luar bulan/tahun periode aktif
 * dilaporkan di outOfRangeDays (tidak ditulis).
 * DIPANGGIL DI DALAM LOCK (approvePengajuan); fungsi ini tidak mengambil lock sendiri.
 * Semua pemeriksaan dilakukan sebelum ada sel yang ditulis.
 */
function applyPengajuanToRealisasi_(nama, jenis, tglMulaiStr, tglAkhirStr) {
  const kode = PENGAJUAN_JENIS_KODE_[jenis];
  if (!kode) throw new Error('Jenis pengajuan "' + jenis + '" tidak dikenali.');

  const realisasiSheet = getSheet_(CONFIG.SHEET_REALISASI);
  const periodText = String(realisasiSheet.getRange(2, 1).getValue() || '').trim();
  const my = parsePeriodMonthYear_(periodText);
  if (!my) {
    throw new Error('Periode aktif ("' + periodText + '") tidak bisa dibaca bulan/tahunnya, tidak bisa menerapkan pengajuan ke Realisasi.');
  }

  const lastRow = getLastEmployeeRow_(realisasiSheet);
  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  const names = numRows > 0
    ? realisasiSheet.getRange(CONFIG.DATA_START_ROW, 2, numRows, 1).getValues().map(r => String(r[0] || '').trim().toLowerCase())
    : [];
  const idx = names.indexOf(String(nama).trim().toLowerCase());
  if (idx === -1) {
    throw new Error('Karyawan "' + nama + '" tidak ditemukan di sheet Realisasi periode aktif.');
  }
  const rowIndex = CONFIG.DATA_START_ROW + idx;
  const dayCount = getDayColCount_(realisasiSheet);

  const mulai = parsePengajuanYmd_(tglMulaiStr);
  const akhir = parsePengajuanYmd_(tglAkhirStr);
  if (!mulai || !akhir) {
    throw new Error('Tanggal pada pengajuan ini tidak valid ("' + tglMulaiStr + '" s/d "' + tglAkhirStr + '"). Tolak pengajuan dan minta karyawan mengajukan ulang.');
  }
  const startTime = new Date(mulai.y, mulai.m - 1, mulai.d).getTime();
  const endTime = new Date(akhir.y, akhir.m - 1, akhir.d).getTime();
  const totalHari = Math.round((endTime - startTime) / 86400000) + 1;
  if (totalHari > PENGAJUAN_RENTANG_MAX_HARI_) {
    throw new Error('Rentang pengajuan ' + totalHari + ' hari melebihi batas ' + PENGAJUAN_RENTANG_MAX_HARI_ + ' hari. Tolak pengajuan dan minta karyawan memecahnya.');
  }

  const appliedDays = [];
  const outOfRangeDays = [];
  for (let k = 0; k < totalHari; k++) {
    const cursor = new Date(mulai.y, mulai.m - 1, mulai.d + k);
    const inPeriod = cursor.getFullYear() === my.year && cursor.getMonth() === my.monthIndex;
    const dayNum = cursor.getDate();
    if (inPeriod && dayNum >= 1 && dayNum <= dayCount) appliedDays.push(dayNum);
    else outOfRangeDays.push(Utilities.formatDate(cursor, PENGAJUAN_TZ_, 'dd/MM/yyyy'));
  }

  // Hari yang diterapkan selalu berurutan (satu bulan), jadi cukup satu blok sel.
  if (appliedDays.length) {
    const firstCol = CONFIG.FIRST_DAY_COL + (appliedDays[0] - 1);
    realisasiSheet.getRange(rowIndex, firstCol, 1, appliedDays.length)
      .setValues([appliedDays.map(() => kode)]);
  }
  SpreadsheetApp.flush();

  return { appliedDays: appliedDays, outOfRangeDays: outOfRangeDays };
}

/**
 * Kirim email hasil pengajuan (Approved/Rejected) ke karyawan bila emailnya terdaftar.
 * Kegagalan kirim tidak boleh menggagalkan approve/reject.
 */
function kirimEmailNotifikasiPengajuan_(nama, jenis, tglMulai, tglAkhir, statusBaru, catatanAdmin) {
  try {
    const tglMulaiFmt = formatTanggalPengajuan_(tglMulai);
    const tglAkhirFmt = formatTanggalPengajuan_(tglAkhir);

    const master = getKaryawanMasterMap_()[String(nama).trim().toLowerCase()];
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
    Logger.log('kirimEmailNotifikasiPengajuan_ gagal: ' + e.message);
  }
}

/* ============================ AUTO-TANDAI ALFA (TRIGGER HARIAN) ============================ */

/**
 * Dipanggil otomatis tiap hari (lihat setupAutoAlfaTrigger). Untuk TANGGAL HARI INI:
 * karyawan yang berjadwal (bukan OFF/kosong), sel Realisasi-nya masih kosong, dan
 * tidak punya pengajuan Pending/Approved yang mencakup hari ini, ditandai "A".
 * Sel dibaca di dalam lock, jadi isian yang masuk sesaat sebelumnya tidak tertimpa.
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
  const namaTerlindungi = getNamaDenganPengajuanAktifPadaTanggal_(today);

  const ditandai = withScriptLock_(() => {
    const jadwalCol = jadwalSheet.getRange(CONFIG.DATA_START_ROW, col, numRows, 1).getValues();
    const realisasiCol = realisasiSheet.getRange(CONFIG.DATA_START_ROW, col, numRows, 1).getValues();
    const names = realisasiSheet.getRange(CONFIG.DATA_START_ROW, 2, numRows, 1).getValues();

    let count = 0;
    for (let i = 0; i < numRows; i++) {
      const jadwalKode = String(jadwalCol[i][0] || '').trim().toUpperCase();
      const realKosong = (realisasiCol[i][0] === '' || realisasiCol[i][0] === null || realisasiCol[i][0] === undefined);
      const nama = String(names[i][0] || '').trim();

      if (jadwalKode && jadwalKode !== 'OFF' && realKosong && nama && !namaTerlindungi.has(nama.toLowerCase())) {
        realisasiSheet.getRange(CONFIG.DATA_START_ROW + i, col).setValue('A');
        count++;
      }
    }
    if (count > 0) SpreadsheetApp.flush();
    return count;
  }, 20000);

  Logger.log('autoTandaiAlfaHarian_: ' + ditandai + ' karyawan ditandai Alfa untuk tanggal ' +
    Utilities.formatDate(today, PENGAJUAN_TZ_, 'dd/MM/yyyy') + '.');
}

/** Set nama (huruf kecil) yang punya pengajuan Pending/Approved yang mencakup tanggal tertentu. */
function getNamaDenganPengajuanAktifPadaTanggal_(targetDate) {
  const set = new Set();
  const sheet = getOrCreatePengajuanSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return set;

  const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues(); // sampai kolom Status
  const targetStr = Utilities.formatDate(targetDate, PENGAJUAN_TZ_, 'yyyy-MM-dd');

  data.forEach(r => {
    const nama = String(r[1] || '').trim().toLowerCase();
    const tglMulai = pengajuanDateStr_(r[4]);
    const tglAkhir = pengajuanDateStr_(r[5]);
    const status = String(r[8] || '');
    if (!nama || (status !== 'Pending' && status !== 'Approved')) return;
    if (tglMulai <= targetStr && targetStr <= tglAkhir) set.add(nama);
  });

  return set;
}

/** JALANKAN SEKALI dari editor: pasang trigger yang menandai Alfa tiap hari jam 23:00. Aman diulang. */
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
 * Alat debug manual (jalankan dari editor): tampilkan isi sheet PengajuanIzin apa
 * adanya di Log Eksekusi, tanpa token/session/web app.
 */
function debugPengajuanRaw() {
  const sheet = getOrCreatePengajuanSheet_();
  const lastRow = sheet.getLastRow();
  Logger.log('Spreadsheet aktif: ' + getSpreadsheet_().getUrl());
  Logger.log('Nama sheet: ' + sheet.getName() + ', lastRow: ' + lastRow);
  if (lastRow < 2) { Logger.log('Tidak ada data.'); return; }
  const data = sheet.getRange(2, 1, lastRow - 1, PENGAJUAN_HEADER_.length).getValues();
  data.forEach((r, i) => {
    Logger.log('Baris ' + (i + 1 + 1) + ': ID=[' + r[0] + '] Nama=[' + r[1] + '] Status=[' + r[8] + ']');
  });
}
