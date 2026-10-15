/*************************************************************
 * 10_DataApi.gs
 * Agregasi data untuk web app: getAllData, ringkasan dashboard,
 * tren harian, dan sinkronisasi otomatis OFF Jadwal -> Realisasi.
 * Mendukung filter opsional per Department.
 *
 * Performa:
 *  - syncOffFromJadwal_ dibatasi max 1x per THROTTLE_SYNC_SEC.
 *  - hasil getAllDataInternal_ di-cache per department selama
 *    CACHE_ALLDATA_SEC. User lain bisa melihat data tersimpan dengan
 *    jeda sampai CACHE_ALLDATA_SEC detik (self-refresh di poll berikutnya).
 *************************************************************/

const THROTTLE_SYNC_SEC = 45;
const CACHE_ALLDATA_SEC = 8;
const CACHE_ALLDATA_PREFIX_ = 'alldata_v1_';
const OFFSYNC_THROTTLE_KEY_ = 'offsync_throttle';

/* ============================ HELPER ============================ */

function nameKey_(name) {
  return String(name).trim().toLowerCase();
}

function round2_(n) {
  return Math.round(n * 100) / 100;
}

function filterGridByNames_(grid, allowedSet) {
  return Object.assign({}, grid, {
    employees: grid.employees.filter(e => allowedSet.has(nameKey_(e.nama)))
  });
}

function isNameAllowed_(nama, allowedSet) {
  return !allowedSet || allowedSet.has(nameKey_(nama));
}

/** Kode sel -> huruf besar tanpa spasi; null/undefined -> ''. */
function cellCode_(val) {
  return (val || '').toString().trim().toUpperCase();
}

/* ============================ ENTRY ============================ */

function getAllData(token, department) {
  verifySession_(token);
  return getCachedAllData_(department);
}

function getCachedAllData_(department) {
  const cache = CacheService.getScriptCache();
  const key = CACHE_ALLDATA_PREFIX_ + String(department || 'ALL');
  const cached = cache.get(key);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* cache korup, hitung ulang */ }
  }
  const result = getAllDataInternal_(false, department);
  try {
    cache.put(key, JSON.stringify(result), CACHE_ALLDATA_SEC);
  } catch (e) {
    // Terlalu besar untuk CacheService (~100KB/key): lewati cache, tidak fatal.
  }
  return result;
}

function maybeSyncOffFromJadwal_() {
  const cache = CacheService.getScriptCache();
  if (cache.get(OFFSYNC_THROTTLE_KEY_)) return;
  cache.put(OFFSYNC_THROTTLE_KEY_, '1', THROTTLE_SYNC_SEC);
  syncOffFromJadwal_();
}

/* ============================ AGREGASI UTAMA ============================ */

function getAllDataInternal_(skipSync, department) {
  if (!skipSync) maybeSyncOffFromJadwal_();

  const ss = getSpreadsheet_();
  const jadwalFull = readGrid_(CONFIG.SHEET_JADWAL);
  const realisasiFull = readGrid_(CONFIG.SHEET_REALISASI);
  const jamKerjaSheet = ss.getSheetByName(CONFIG.SHEET_JAMKERJA);
  const jamScanSheet = ss.getSheetByName(CONFIG.SHEET_JAMSCAN);
  const jamScanFull = jamScanSheet ? readGridFromSheetObj_(jamScanSheet) : { employees: [] };

  const deptFilter = String(department || '').trim();
  const isFiltered = !!deptFilter && deptFilter !== 'ALL';

  let allowedNames = null;
  let jadwal = jadwalFull, realisasi = realisasiFull, jamScan = jamScanFull;

  if (isFiltered) {
    const master = getKaryawanMasterMap_();
    allowedNames = new Set(Object.keys(master).filter(k => master[k].department === deptFilter));
    jadwal = filterGridByNames_(jadwalFull, allowedNames);
    realisasi = filterGridByNames_(realisasiFull, allowedNames);
    jamScan = filterGridByNames_(jamScanFull, allowedNames);
  }

  const rekapAll = attachJamKerjaToRekap_(attachNikToRekap_(getRekap_()), jadwalFull);
  const rekap = allowedNames ? rekapAll.filter(r => allowedNames.has(nameKey_(r.nama))) : rekapAll;

  return {
    period: jadwalFull.period,
    dayNames: jadwalFull.dayNames,
    dayCount: jadwalFull.dayCount,
    jadwal: jadwal.employees,
    realisasi: realisasi.employees,
    dashboard: computeDashboardSummaryFromRekapLengkap_(rekap),
    rekap: rekap,
    trenHarian: computeTrenHarian_(jadwal, realisasi, jamKerjaSheet, allowedNames),
    jamKerjaRingkasan: computeJamKerjaRingkasan_(rekap),
    statusCodes: CONFIG.STATUS_CODES,
    shiftCodes: getShiftCodesDynamic_(),
    alignmentWarnings: checkRowAlignment_(jadwalFull, realisasiFull),
    departmentFilter: isFiltered ? deptFilter : null,
    lastUpdate: new Date().toISOString(),
    jamScan: jamScan.employees
  };
}

/* ============================ TREN HARIAN ============================ */

/** Kode Realisasi -> nama deret tren (selain hadir/lembur yang ditangani khusus). */
const TREN_CODE_TO_SERIES_ = { 'A': 'alfa', 'I': 'izin', 'S': 'sakit', 'C': 'cuti', 'OFF': 'off' };

function computeTrenHarian_(jadwal, realisasi, jamKerjaSheet, allowedNamesSet) {
  const dayCount = realisasi.dayCount;
  const dayNames = realisasi.dayNames;
  const zeros = () => new Array(dayCount).fill(0);

  const s = { hadir: zeros(), lembur: zeros(), alfa: zeros(), izin: zeros(), sakit: zeros(), cuti: zeros(), off: zeros() };
  const planCount = zeros();

  realisasi.employees.forEach(emp => {
    if (!isNameAllowed_(emp.nama, allowedNamesSet)) return;
    emp.days.forEach((val, i) => {
      if (i >= dayCount) return;
      const code = cellCode_(val);
      if (code === '1') {
        s.hadir[i]++;
      } else if (code === '2') {
        s.hadir[i]++;
        s.lembur[i]++;
      } else if (TREN_CODE_TO_SERIES_[code]) {
        s[TREN_CODE_TO_SERIES_[code]][i]++;
      }
    });
  });

  jadwal.employees.forEach(emp => {
    if (!isNameAllowed_(emp.nama, allowedNamesSet)) return;
    emp.days.forEach((val, i) => {
      if (i >= dayCount) return;
      const code = cellCode_(val);
      if (code !== 'OFF' && code !== '') planCount[i]++;
    });
  });

  return {
    dayNames, dayCount,
    hadir: s.hadir, lembur: s.lembur, alfa: s.alfa, izin: s.izin, sakit: s.sakit, cuti: s.cuti, off: s.off,
    planCount,
    actualCount: s.hadir,
    totalJamKerjaPerHari: computeJamKerjaTrend_(dayCount, jamKerjaSheet, allowedNamesSet),
    todayDayIndex: computeTodayDayIndex_(realisasi.period, dayCount)
  };
}

function computeJamKerjaTrend_(dayCount, jamKerjaSheet, allowedNamesSet) {
  const total = new Array(dayCount).fill(0);
  if (!jamKerjaSheet) return total;

  readGridFromSheetObj_(jamKerjaSheet).employees.forEach(emp => {
    if (!isNameAllowed_(emp.nama, allowedNamesSet)) return;
    emp.days.forEach((val, i) => {
      if (i >= dayCount) return;
      const num = Number(val);
      if (val !== '' && val !== null && !isNaN(num)) total[i] += num;
    });
  });
  return total.map(round2_);
}

/* ============================ RINGKASAN DASHBOARD ============================ */

function getDashboardSummary_() {
  try {
    const rekapLengkap = attachJamKerjaToRekap_(getRekap_(), readGrid_(CONFIG.SHEET_JADWAL));
    return computeDashboardSummaryFromRekapLengkap_(rekapLengkap);
  } catch (err) {
    Logger.log('Error pada getDashboardSummary_: ' + err.toString());
    return {};
  }
}

function getDashboardSummaryFromSheet_(rekapArsip) {
  return computeDashboardSummaryFromRekapLengkap_(rekapArsip || []);
}

/**
 * Hitung ringkasan dashboard. Mengembalikan DUA set kunci (camelCase untuk
 * frontend baru, "HURUF BESAR" untuk tampilan/arsip lama) - keduanya masih dipakai.
 */
function computeDashboardSummaryFromRekapLengkap_(rekapLengkap) {
  try {
    const sum = { jadwal: 0, hadir: 0, alfa: 0, izin: 0, sakit: 0, cuti: 0, off: 0, jam: 0, target: 0 };
    rekapLengkap.forEach(r => {
      sum.jadwal += Number(r.jadwal) || 0;
      sum.hadir += Number(r.hadir) || 0;
      sum.alfa += Number(r.alfa) || 0;
      sum.izin += Number(r.izin) || 0;
      sum.sakit += Number(r.sakit) || 0;
      sum.cuti += Number(r.cuti) || 0;
      sum.off += Number(r.off) || 0;
      sum.jam += Number(r.totalJamKerja) || 0;
      sum.target += Number(r.targetJamKerja) || 0;
    });

    const totalKaryawan = rekapLengkap.length;
    const totalTidakHadir = sum.alfa + sum.izin + sum.sakit + sum.cuti;
    const persen = sum.jadwal > 0 ? Math.round((sum.hadir / sum.jadwal) * 100) + '%' : '0%';
    const totalJam = round2_(sum.jam);
    const targetJam = round2_(sum.target);
    const kekuranganJam = round2_(sum.target - sum.jam) + ' jam';

    return {
      totalKaryawan: totalKaryawan,
      totalHariKerjaTerjadwal: sum.jadwal,
      totalHadirAktual: sum.hadir,
      rataRataKehadiran: persen,
      persenKehadiran: persen,
      totalTidakHadir: totalTidakHadir,
      totalOff: sum.off,
      alfa: sum.alfa,
      izin: sum.izin,
      sakit: sum.sakit,
      cuti: sum.cuti,
      totalJamKerja: totalJam,
      targetJamKerja: targetJam,
      kekuranganJam: kekuranganJam,

      'TOTAL KARYAWAN': totalKaryawan,
      'TOTAL HARI KERJA TERJADWAL': sum.jadwal,
      'TOTAL HADIR AKTUAL': sum.hadir,
      'RATA-RATA KEHADIRAN': persen,
      'TOTAL TIDAK HADIR (ALFA/IZIN/SAKIT)': totalTidakHadir,
      'TOTAL OFF (LIBUR)': sum.off,
      'TOTAL JAM KERJA': totalJam,
      'TARGET JAM KERJA': targetJam,
      'KEKURANGAN JAM': kekuranganJam
    };
  } catch (err) {
    Logger.log('Error pada computeDashboardSummaryFromRekapLengkap_: ' + err.toString());
    return {};
  }
}

/* ============================ SINKRON OFF ============================ */

/**
 * Isi otomatis "OFF" di sel Realisasi yang MASIH KOSONG jika di Jadwal
 * baris+hari yang sama = OFF. Sel yang sudah berisi tidak pernah ditimpa.
 */
function syncOffFromJadwal_() {
  const jadwalSheet = getSheet_(CONFIG.SHEET_JADWAL);
  const realisasiSheet = getSheet_(CONFIG.SHEET_REALISASI);

  const numRows = Math.min(getLastEmployeeRow_(jadwalSheet), getLastEmployeeRow_(realisasiSheet)) - CONFIG.DATA_START_ROW + 1;
  if (numRows <= 0) return { updated: 0 };

  const dayCount = Math.min(getDayColCount_(jadwalSheet), getDayColCount_(realisasiSheet));
  if (dayCount <= 0) return { updated: 0 };

  return withScriptLock_(() => {
    const rangeOf = sheet => sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.FIRST_DAY_COL, numRows, dayCount);
    const jadwalGrid = rangeOf(jadwalSheet).getValues();
    const realisasiRange = rangeOf(realisasiSheet);
    const realisasiGrid = realisasiRange.getValues();

    let updated = 0;
    for (let i = 0; i < numRows; i++) {
      for (let j = 0; j < dayCount; j++) {
        const v = realisasiGrid[i][j];
        const kosong = v === '' || v === null || v === undefined;
        if (kosong && cellCode_(jadwalGrid[i][j]) === 'OFF') {
          realisasiGrid[i][j] = 'OFF';
          updated++;
        }
      }
    }

    if (updated > 0) {
      realisasiRange.setValues(realisasiGrid);
      SpreadsheetApp.flush();
      Logger.log('syncOffFromJadwal_: ' + updated + ' sel Realisasi otomatis diisi "OFF" sesuai Jadwal.');
    }
    return { updated: updated };
  }, 10000);
}
