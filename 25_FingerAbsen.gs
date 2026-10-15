/*************************************************************
 * 25_FingerAbsen.gs
 * Import absen dari mesin finger print (.xls/.xlsx). Hasilnya:
 *   - Realisasi : 1 (hadir) atau 2 (lembur >= LEMBUR_MIN_HOURS jam);
 *                 kode S/I/C/A yang sudah ada TIDAK ditimpa
 *   - JamKerja  : jumlah jam kerja per hari
 *   - JamScan   : "scan masuk - scan pulang"
 * Pencocokan karyawan: NIK (No. ID) dulu; kalau NIK tidak ada/ganda,
 * dicocokkan lewat nama. Khusus Super Admin.
 *************************************************************/

const FINGER_PROTECTED_CODES_ = ['S', 'I', 'C', 'A'];
const UPLOAD_ERR_FINGER_ = 'Gagal membaca file absen. Pastikan file berformat .xls / .xlsx yang valid dari mesin finger.';

function uploadFingerAbsen(base64Data, filename, token) {
  requireSuperAdmin_(token, 'Hanya akun dengan role "Super Admin" yang boleh upload data absen finger.');
  if (!base64Data) throw new Error('File tidak ditemukan / kosong.');
  filename = filename || 'upload_finger.xls';

  const sourceMime = filename.toLowerCase().endsWith('.xls') ? MIME_XLS_ : MIME_XLSX_;
  const applyResult = withTempSpreadsheet_(base64Data, filename, '_TEMP_UPLOAD_FINGER_', sourceMime, UPLOAD_ERR_FINGER_, fileId =>
    applyFingerRowsToRealisasi_(readFingerRows_(fileId))
  );

  const result = getAllDataInternal_();
  result.uploadFingerInfo = applyResult;
  return result;
}

/* ============================ PARSING FILE ============================ */

function readFingerRows_(fileId) {
  const sheet = SpreadsheetApp.openById(fileId).getSheets()[0];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) throw new Error('File absen finger kosong / tidak ada data.');

  const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
  const col = {};
  FINGER_REQUIRED_COLUMNS.forEach(name => {
    const idx = headerRow.indexOf(name);
    if (idx === -1) {
      throw new Error('Kolom "' + name + '" tidak ditemukan di file. Pastikan format file sesuai export mesin finger (header: ' +
        FINGER_REQUIRED_COLUMNS.join(', ') + ').');
    }
    col[name] = idx;
  });

  const text = (r, name) => String(r[col[name]] || '').trim();

  return sheet.getRange(2, 1, lastRow - 1, lastCol).getValues()
    .filter(r => text(r, 'Nama') !== '')
    .map(r => ({
      nik: text(r, 'No. ID'),
      nama: text(r, 'Nama'),
      dayOfMonth: extractDayOfMonth_(r[col['Tanggal']]),
      shiftLabel: text(r, 'Jam Kerja'),
      scanMasuk: formatTimeValue_(r[col['Scan Masuk']]),
      scanPulang: formatTimeValue_(r[col['Scan Pulang']]),
      jamKerjaJam: parseHmsToHours_(r[col['Jml Jam Kerja']]),
      lemburJam: parseHmsToHours_(r[col['Lembur']])
    }));
}

/** Tanggal (Date atau teks dd/MM/yyyy, dd-MM-yy) -> tanggal dalam bulan (1..31). */
function extractDayOfMonth_(value) {
  if (value instanceof Date) return value.getDate();
  const s = String(value || '').trim();
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) return parseInt(m[1], 10);
  throw new Error('Format tanggal tidak dikenali di file finger: "' + s + '". Harus dd/MM/yyyy.');
}

function formatTimeValue_(value) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone() || 'GMT+7', 'HH:mm');
  return String(value).trim();
}

/** "08:30" -> 8.5 jam; kosong/tidak terbaca -> 0. */
function parseHmsToHours_(value) {
  const s = formatTimeValue_(value);
  if (!s) return 0;
  const parts = s.split(':');
  if (parts.length < 2) return 0;
  return (parseInt(parts[0], 10) || 0) + ((parseInt(parts[1], 10) || 0) / 60);
}

/* ============================ TERAPKAN KE SHEET ============================ */

/** NIK -> daftar nama (huruf kecil) di DataKaryawan; NIK kosong dilewati. */
function groupNamesByNik_(nikMap) {
  const nikToNames = {};
  Object.keys(nikMap).forEach(nameLower => {
    const nik = nikMap[nameLower];
    if (!nik) return;
    (nikToNames[nik] = nikToNames[nik] || []).push(nameLower);
  });
  return nikToNames;
}

/** Indeks baris Realisasi untuk satu baris finger: NIK unik dulu, lalu nama; -1 jika tidak ada. */
function resolveFingerRowIdx_(row, nikToIdx, names) {
  const fingerNik = String(row.nik || '').trim();
  if (fingerNik && nikToIdx.hasOwnProperty(fingerNik)) return nikToIdx[fingerNik];
  return names.indexOf(row.nama.toLowerCase());
}

function applyFingerRowsToRealisasi_(rows) {
  const realisasiSheet = getSheet_(CONFIG.SHEET_REALISASI);
  const numRows = getLastEmployeeRow_(realisasiSheet) - CONFIG.DATA_START_ROW + 1;
  const dayCount = getDayColCount_(realisasiSheet);
  if (numRows <= 0 || dayCount <= 0) throw new Error('Sheet Realisasi belum berisi karyawan/hari. Isi dulu sebelum upload finger.');

  const gridRange = sheet => sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.FIRST_DAY_COL, numRows, dayCount);

  const names = realisasiSheet.getRange(CONFIG.DATA_START_ROW, 2, numRows, 1).getValues().map(r => nameKey_(r[0] || ''));
  const currentRealisasi = gridRange(realisasiSheet).getValues();

  // NIK yang dipakai lebih dari satu nama dianggap ambigu -> pencocokan jatuh ke nama.
  const nikMap = getNikMap_();
  const nikToNames = groupNamesByNik_(nikMap);
  const duplicateNiks = Object.keys(nikToNames).filter(nik => nikToNames[nik].length > 1);
  const duplicateNikSet = new Set(duplicateNiks);

  const nikToIdx = {};
  names.forEach((nameLower, i) => {
    const nik = nikMap[nameLower];
    if (nik && !duplicateNikSet.has(nik)) nikToIdx[String(nik).trim()] = i;
  });

  const jamKerjaSheet = getOrSyncJamKerjaSheet_();
  const jamKerjaGrid = gridRange(jamKerjaSheet).getValues();
  const jamScanSheet = getOrSyncJamScanSheet_();
  const jamScanGrid = gridRange(jamScanSheet).getValues();

  // Kosongkan JamKerja/JamScan pada rentang tanggal (min..max) tiap karyawan yang ada di file,
  // supaya tanggal "hilang" di dalam rentang (OFF/tidak scan) tidak menyisakan jam lama.
  const dayRangeByIdx = {};
  rows.forEach(row => {
    const idx = resolveFingerRowIdx_(row, nikToIdx, names);
    if (idx === -1) return;
    const d = row.dayOfMonth;
    const range = dayRangeByIdx[idx];
    if (!range) dayRangeByIdx[idx] = { min: d, max: d };
    else { range.min = Math.min(range.min, d); range.max = Math.max(range.max, d); }
  });
  Object.keys(dayRangeByIdx).forEach(idx => {
    const range = dayRangeByIdx[idx];
    for (let d = range.min; d <= range.max; d++) {
      if (d - 1 < 0 || d - 1 >= dayCount) continue;
      jamKerjaGrid[idx][d - 1] = '';
      jamScanGrid[idx][d - 1] = '';
    }
  });

  let filled = 0, filledOvertime = 0, overwritten = 0, skippedAlreadyFilled = 0;
  const notFound = [];
  const outOfRangeDay = [];
  const duplicateNikWarnings = duplicateNiks.map(nik =>
    'NIK "' + nik + '" dipakai oleh lebih dari satu nama di sheet DataKaryawan (' + nikToNames[nik].join(', ') +
    ') — pencocokan untuk NIK ini otomatis dialihkan ke pencocokan nama.'
  );

  withScriptLock_(() => {
    rows.forEach(row => {
      const idx = resolveFingerRowIdx_(row, nikToIdx, names);
      if (idx === -1) {
        const fingerNik = String(row.nik || '').trim();
        const label = row.nama + (fingerNik ? ' (ID:' + fingerNik + ')' : '');
        if (notFound.indexOf(label) === -1) notFound.push(label);
        return;
      }

      const dayIdx = row.dayOfMonth - 1;
      if (dayIdx < 0 || dayIdx >= dayCount) {
        outOfRangeDay.push(row.nama + ' (tanggal ' + row.dayOfMonth + ')');
        return;
      }

      const currentRaw = currentRealisasi[idx][dayIdx];
      const currentCode = (currentRaw === null || currentRaw === undefined) ? '' : String(currentRaw).trim().toUpperCase();

      if (FINGER_PROTECTED_CODES_.indexOf(currentCode) === -1) {
        const isLembur = row.lemburJam >= CONFIG.LEMBUR_MIN_HOURS;
        currentRealisasi[idx][dayIdx] = isLembur ? 2 : 1;
        filled++;
        if (isLembur) filledOvertime++;
        if (currentCode !== '') overwritten++;
      } else {
        skippedAlreadyFilled++;
      }

      jamKerjaGrid[idx][dayIdx] = row.jamKerjaJam ? round2_(row.jamKerjaJam) : '';
      jamScanGrid[idx][dayIdx] = formatScanPair_(row.scanMasuk, row.scanPulang);
    });

    gridRange(realisasiSheet).setValues(currentRealisasi);
    gridRange(jamKerjaSheet).setValues(jamKerjaGrid);
    gridRange(jamScanSheet).setValues(jamScanGrid);
    SpreadsheetApp.flush();
  }, 20000);

  return { filled, filledOvertime, overwritten, skippedAlreadyFilled, notFound, outOfRangeDay, duplicateNik: duplicateNikWarnings };
}
