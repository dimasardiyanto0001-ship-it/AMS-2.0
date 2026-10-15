/*************************************************************
 * 20_AbsenOnline.gs
 * Backend halaman publik "Absen Online": absen GPS (submitAbsenOnline)
 * + riwayat absen (getRiwayatAbsen). Tanpa token — verifikasi Nama+NIK.
 *************************************************************/

const SHEET_LOG_ABSEN_ONLINE = 'LogAbsenOnline';
const SHEET_SETTING_LOKASI_ABSEN = 'SettingLokasiAbsen';

function getOrCreateSettingLokasiAbsenSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_SETTING_LOKASI_ABSEN);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SETTING_LOKASI_ABSEN);
    sheet.getRange(1, 1, 1, 3).setValues([['Latitude', 'Longitude', 'RadiusMeter']]).setFontWeight('bold');
    // NILAI PLACEHOLDER -- WAJIB diganti manual di baris 2 ke koordinat kantor sebenarnya.
    sheet.getRange(2, 1, 1, 3).setValues([[0, 0, 100]]);
    sheet.getRange('E1').setValue('Isi Latitude/Longitude kantor & radius toleransi (meter) di baris 2.');
    sheet.hideSheet();
  }
  return sheet;
}

function getSettingLokasiAbsen_() {
  const sheet = getOrCreateSettingLokasiAbsenSheet_();
  const row = sheet.getRange(2, 1, 1, 3).getValues()[0];
  return { lat: Number(row[0]) || 0, lng: Number(row[1]) || 0, radius: Number(row[2]) || 100 };
}

function haversineDistanceMeter_(lat1, lng1, lat2, lng2) {
  const R = 6371000;
  const toRad = d => d * Math.PI / 180;
  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLng / 2) ** 2;
  return Math.round(R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a)));
}

function getOrCreateLogAbsenOnlineSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_LOG_ABSEN_ONLINE);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_LOG_ABSEN_ONLINE);
    sheet.getRange(1, 1, 1, 11).setValues([[
      'Timestamp', 'Tanggal', 'Nama', 'NIK', 'Jenis', 'Jam', 'Latitude', 'Longitude', 'JarakMeter', 'Status', 'FotoUrl'
    ]]).setFontWeight('bold');
    sheet.hideSheet();
  } else {
    ensureLogAbsenOnlineHasFotoColumn_(sheet);
  }
  return sheet;
}

function ensureLogAbsenOnlineHasFotoColumn_(sheet) {
  const header = sheet.getRange(1, 11).getValue();
  if (header !== 'FotoUrl') {
    sheet.getRange(1, 11).setValue('FotoUrl').setFontWeight('bold');
  }
}

function verifyKaryawanByNamaNik_(nama, nik) {
  const masterMap = getKaryawanMasterMap_();
  const master = masterMap[String(nama || '').trim().toLowerCase()];
  if (!master) throw new Error('Nama karyawan tidak ditemukan.');
  if (String(master.nik || '').trim() !== String(nik || '').trim()) {
    throw new Error('NIK tidak sesuai dengan data karyawan.');
  }
  return master;
}

function verifyLoginAbsenOnline(nama, nik) {
  const master = verifyKaryawanByNamaNik_(nama, nik);
  return { department: master.department || '-', fotoProfilUrl: master.fotoProfilUrl || '' };
}

function submitAbsenOnline(nama, nik, jenis, lat, lng, akurasi, fotoBase64) {
  nama = String(nama || '').trim();
  nik = String(nik || '').trim();
  jenis = String(jenis || '').trim();
  lat = Number(lat); lng = Number(lng);

  if (!nama || !nik) throw new Error('Nama dan NIK wajib diisi.');
  if (jenis !== 'Masuk' && jenis !== 'Pulang') throw new Error('Jenis absen tidak valid.');
  if (isNaN(lat) || isNaN(lng)) throw new Error('Data lokasi GPS tidak valid.');
  if (!fotoBase64) throw new Error('Foto selfie wajib diambil sebelum absen.');

  verifyKaryawanByNamaNik_(nama, nik);

  const lokasi = getSettingLokasiAbsen_();
  const jarak = haversineDistanceMeter_(lat, lng, lokasi.lat, lokasi.lng);
  const status = jarak <= lokasi.radius ? 'Valid' : 'Di Luar Radius';

  const now = new Date();
  const tz = Session.getScriptTimeZone() || 'GMT+7';
  const jamStr = Utilities.formatDate(now, tz, 'HH:mm');
  const tanggalStr = Utilities.formatDate(now, tz, 'dd/MM/yyyy');

  // Simpan foto selfie ke Drive, folder per bulan. Kalau gagal upload,
  // absen tetap tercatat (jangan sampai karyawan gagal absen gara-gara
  // masalah jaringan/Drive), tapi FotoUrl dikosongkan.
  let fotoUrl = '';
  try {
    const folder = getOrCreateFotoAbsenBulanFolder_(now.getFullYear(), now.getMonth() + 1);
    const namaFile = nama.replace(/[^\w]+/g, '_') + '_' + jenis + '_' +
      Utilities.formatDate(now, tz, 'yyyyMMdd_HHmmss') + '.jpg';
    const rawBase64 = fotoBase64.replace(/^data:image\/\w+;base64,/, '');
    const bytes = Utilities.base64Decode(rawBase64);
    const blob = Utilities.newBlob(bytes, 'image/jpeg', namaFile);
    const file = folder.createFile(blob);
    fotoUrl = file.getUrl();
  } catch (e) {
    Logger.log('Gagal simpan foto absen untuk ' + nama + ': ' + e.message);
  }

  // Bila periode aktif BUKAN bulan berjalan (mis. tanggal 1 sebelum "Mulai Periode Baru"),
  // absen tetap dicatat di LogAbsenOnline tetapi TIDAK ditulis ke Realisasi/JamScan/JamKerja:
  // kolom tanggalnya milik bulan lalu. Admin memprosesnya setelah periode diganti.
  let periodeBelumDiganti = false;
  let namaPeriode = '';

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    getOrCreateLogAbsenOnlineSheet_().appendRow([
      now, tanggalStr, nama, nik, jenis, jamStr, lat, lng, jarak, status, fotoUrl
    ]);

    const realisasiSheet = getSheet_(CONFIG.SHEET_REALISASI);
    const periodText = String(realisasiSheet.getRange(2, 1).getValue() || '').trim();
    const my = parsePeriodMonthYear_(periodText);
    if (my && (now.getFullYear() !== my.year || now.getMonth() !== my.monthIndex)) {
      periodeBelumDiganti = true;
      namaPeriode = periodText.replace(/^periode\s*:?\s*/i, '');
    }

    const lastRow = getLastEmployeeRow_(realisasiSheet);
    const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
    const names = realisasiSheet.getRange(CONFIG.DATA_START_ROW, 2, numRows, 1).getValues()
      .map(r => String(r[0] || '').trim().toLowerCase());
    const rowIdx = periodeBelumDiganti ? -1 : names.indexOf(nama.toLowerCase());

    if (rowIdx !== -1) {
      const dayCount = getDayColCount_(realisasiSheet);
      const dayIdx = now.getDate() - 1;
      if (dayIdx >= 0 && dayIdx < dayCount) {
        const rowNum = CONFIG.DATA_START_ROW + rowIdx;
        const col = CONFIG.FIRST_DAY_COL + dayIdx;
        const currentVal = realisasiSheet.getRange(rowNum, col).getValue();
        const currentCode = (currentVal === null || currentVal === undefined) ? '' : String(currentVal).trim().toUpperCase();
        const isProtected = ['S', 'I', 'C', 'A'].includes(currentCode);
        if (!isProtected && currentCode !== '1' && currentCode !== '2') {
          realisasiSheet.getRange(rowNum, col).setValue(1);
        }

        let jamScanSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_JAMSCAN);
        if (!jamScanSheet) jamScanSheet = getOrSyncJamScanSheet_();
        const scanCell = jamScanSheet.getRange(rowNum, col);
        const existing = splitScanPair_(scanCell.getValue());
        const masuk = jenis === 'Masuk' ? jamStr : existing.masuk;
        const keluar = jenis === 'Pulang' ? jamStr : existing.keluar;
        scanCell.setValue(formatScanPair_(masuk, keluar));

        // === TAMBAHAN: tulis jam kerja terhitung ke sheet JamKerja saat absen Pulang ===
        if (jenis === 'Pulang' && masuk && keluar) {
          const jamKerjaJam = hitungDurasiJamKerja_(masuk, keluar);
          if (jamKerjaJam > 0) {
            let jamKerjaSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_JAMKERJA);
            if (!jamKerjaSheet) jamKerjaSheet = getOrSyncJamKerjaSheet_();
            jamKerjaSheet.getRange(rowNum, col).setValue(Math.round(jamKerjaJam * 100) / 100);
          }
        }
               // === AKHIR TAMBAHAN ===
      }
    }

    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  const hasil = { status: status, jam: jamStr, jarak: jarak };
  if (periodeBelumDiganti) { hasil.periodeBelumDiganti = true; hasil.periode = namaPeriode; }
  return hasil;
}
function getRiwayatAbsen(nik) {
  nik = String(nik || '').trim();
  if (!nik) throw new Error('NIK wajib diisi.');

  const sheet = getOrCreateLogAbsenOnlineSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  return sheet.getRange(2, 1, lastRow - 1, 10).getValues()
    .filter(r => String(r[3] || '').trim() === nik)
    .map(r => ({
      tanggal: formatTanggalAbsen_(r[1]),
      jenis: r[4],
      jam: formatJam_(r[5]),
      status: r[9]
    }))
    .reverse();
}

function formatTanggalAbsen_(value) {
  if (value === null || value === undefined || value === '') return '-';
  if (value instanceof Date) {
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    return Utilities.formatDate(value, tz, 'dd/MM/yyyy');
  }
  return String(value).trim();
}
/**
 * Cari karyawan by Nama (hanya) untuk autocomplete Absen Online.
 * Mengembalikan maksimal 10 hasil paling relevan.
 * Format tiap hasil: { nama, nik, label } dengan label = Nama saja
 */
function searchKaryawanUntukAbsen(query) {
  query = String(query || '').trim().toLowerCase();
  if (!query) return [];

  const sheet = getOrCreateNikSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // kolom A=Nama, B=NIK
  const results = [];

  data.forEach(row => {
    const nama = String(row[0] || '').trim();
    const nik = String(row[1] || '').trim();
    if (!nama) return;
    const namaLower = nama.toLowerCase();

    let score = -1;
    if (namaLower === query) score = 100;
    else if (namaLower.startsWith(query)) score = 80;
    else if (namaLower.includes(query)) score = 50;

    if (score > -1) results.push({ nama: nama, nik: nik, score: score });
  });

  results.sort((a, b) => b.score - a.score || a.nama.localeCompare(b.nama));
  return results.slice(0, 10).map(r => ({ nama: r.nama, nik: r.nik, label: r.nama }));
}
function formatJam_(value) {
  if (value === null || value === undefined || value === '') return '-';
  if (value instanceof Date) {
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    return Utilities.formatDate(value, tz, 'HH:mm');
  }
  return String(value).trim();
}
const FOTO_ABSEN_ROOT_FOLDER_NAME_ = 'Foto Absen Online - AMS Housekeeping';
const RETENSI_FOTO_ABSEN_BULAN_ = 3;

function getOrCreateFotoAbsenRootFolder_() {
  const it = DriveApp.getFoldersByName(FOTO_ABSEN_ROOT_FOLDER_NAME_);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(FOTO_ABSEN_ROOT_FOLDER_NAME_);
}

function getOrCreateFotoAbsenBulanFolder_(tahun, bulan) {
  const root = getOrCreateFotoAbsenRootFolder_();
  const namaFolder = tahun + '-' + String(bulan).padStart(2, '0');
  const it = root.getFoldersByName(namaFolder);
  if (it.hasNext()) return it.next();
  return root.createFolder(namaFolder);
}
function hapusFotoAbsenLama_() {
  const root = getOrCreateFotoAbsenRootFolder_();
  const batasWaktu = new Date();
  batasWaktu.setMonth(batasWaktu.getMonth() - RETENSI_FOTO_ABSEN_BULAN_);
  const batasAwalBulan = new Date(batasWaktu.getFullYear(), batasWaktu.getMonth(), 1);

  const folders = root.getFolders();
  let totalDihapus = 0;

  while (folders.hasNext()) {
    const folder = folders.next();
    const match = folder.getName().match(/^(\d{4})-(\d{2})$/);
    if (!match) continue;

    const folderDate = new Date(parseInt(match[1], 10), parseInt(match[2], 10) - 1, 1);
    if (folderDate < batasAwalBulan) {
      const files = folder.getFiles();
      while (files.hasNext()) {
        files.next().setTrashed(true);
        totalDihapus++;
      }
      folder.setTrashed(true);
    }
  }

  Logger.log('hapusFotoAbsenLama_: ' + totalDihapus + ' foto (folder > ' + RETENSI_FOTO_ABSEN_BULAN_ + ' bulan) dipindah ke trash.');
  tandaiFotoTerhapusDiSheet_(batasAwalBulan);
}

function tandaiFotoTerhapusDiSheet_(batasWaktu) {
  const sheet = getOrCreateLogAbsenOnlineSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  let changed = false;

  data.forEach((row, i) => {
    const timestamp = row[0];
    const fotoUrl = row[10];
    if (timestamp instanceof Date && timestamp < batasWaktu && fotoUrl &&
        fotoUrl !== '[Foto dihapus - retensi 3 bulan]') {
      data[i][10] = '[Foto dihapus - retensi 3 bulan]';
      changed = true;
    }
  });

  if (changed) {
    sheet.getRange(2, 1, data.length, 11).setValues(data);
    SpreadsheetApp.flush();
  }
}

/** Jalankan SEKALI dari editor untuk memasang trigger bulanan. */
function setupTriggerHapusFotoAbsenLama() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'hapusFotoAbsenLama_') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('hapusFotoAbsenLama_')
    .timeBased()
    .onMonthDay(1)
    .atHour(3)
    .create();

  Logger.log('Trigger hapus foto absen lama terpasang: tiap tanggal 1 jam 03:00.');
}
function testRiwayat() {
  const sheet = getOrCreateNikSheet_();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const nikNurul = data.find(r => String(r[0]).toLowerCase().includes('nurul'));
  Logger.log('NIK ditemukan: ' + JSON.stringify(nikNurul));
  
  if (nikNurul) {
    const hasil = getRiwayatAbsen(String(nikNurul[1]));
    Logger.log('HASIL RIWAYAT: ' + JSON.stringify(hasil, null, 2));
  }
}
function hitungDurasiJamKerja_(masukStr, keluarStr) {
  const toMinutes = s => {
    const parts = String(s).split(':');
    const h = parseInt(parts[0], 10) || 0;
    const m = parseInt(parts[1], 10) || 0;
    return h * 60 + m;
  };
  let diff = toMinutes(keluarStr) - toMinutes(masukStr);
  if (diff < 0) diff += 24 * 60; // jaga-jaga shift malam lewat tengah malam
  return diff / 60;
}
const FOTO_PROFIL_ROOT_FOLDER_NAME_ = 'Foto Profil Karyawan - AMS Housekeeping';

function getOrCreateFotoProfilFolder_() {
  const it = DriveApp.getFoldersByName(FOTO_PROFIL_ROOT_FOLDER_NAME_);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(FOTO_PROFIL_ROOT_FOLDER_NAME_);
}

function setFotoProfilKaryawan(nama, nik, fotoBase64) {
  nama = String(nama || '').trim();
  nik = String(nik || '').trim();
  if (!nama || !nik) throw new Error('Nama dan NIK wajib diisi.');
  if (!fotoBase64) throw new Error('Foto wajib dipilih.');

  verifyKaryawanByNamaNik_(nama, nik);

  const folder = getOrCreateFotoProfilFolder_();
  const namaFile = nama.replace(/[^\w]+/g, '_') + '_profil_' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+7', 'yyyyMMdd_HHmmss') + '.jpg';
  const rawBase64 = fotoBase64.replace(/^data:image\/\w+;base64,/, '');
  const bytes = Utilities.base64Decode(rawBase64);
  const blob = Utilities.newBlob(bytes, 'image/jpeg', namaFile);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const fotoUrl = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w400';

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getOrCreateNikSheet_();
    const lastRow = sheet.getLastRow();
    const names = sheet.getRange(2, 1, lastRow - 1, 1).getValues()
      .map(r => String(r[0] || '').trim().toLowerCase());
    const idx = names.indexOf(nama.toLowerCase());
    if (idx === -1) throw new Error('Karyawan tidak ditemukan di DataKaryawan.');
    sheet.getRange(idx + 2, 10).setValue(fotoUrl); // kolom J = FotoProfil
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  return { fotoProfilUrl: fotoUrl };
}