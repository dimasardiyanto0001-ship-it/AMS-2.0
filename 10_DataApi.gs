/*************************************************************
 * 10_DataApi.gs
 * Agregasi data untuk web app: getAllData, ringkasan dashboard,
 * tren harian, dan sinkronisasi otomatis OFF Jadwal -> Realisasi.
 * Mendukung filter opsional per Department.
 *
 * BARU (filter rentang tanggal + reliever):
 * - getAllData(token, department, dayFrom, dayTo): dayFrom/dayTo =
 *   nomor hari (1..jumlah hari periode). Kosong = seluruh periode.
 * - Yang ikut terfilter rentang: rekap, dashboard, trenHarian,
 *   jamKerjaRingkasan, reliever. Grid jadwal/realisasi/jamScan TETAP
 *   penuh supaya tab Jadwal & Input Realisasi (yang memakai dayIndex)
 *   tidak bergeser.
 * - Respons ditambah: rangeInfo { year, monthIndex, dayCount, dayFrom,
 *   dayTo, active } dan reliever { jumlah, hariTugas, daftar[] }.
 *
 * Performa:
 * - syncOffFromJadwal_ dibatasi max 1x per THROTTLE_SYNC_SEC.
 * - hasil getAllDataInternal_ di-cache per department + rentang selama
 *   CACHE_ALLDATA_SEC. User lain bisa melihat data tersimpan dengan
 *   jeda sampai CACHE_ALLDATA_SEC detik (self-refresh di poll berikutnya).
 *************************************************************/

const THROTTLE_SYNC_SEC = 45;
const CACHE_ALLDATA_SEC = 8;
const CACHE_ALLDATA_PREFIX_ = 'alldata_v2_';   // v2: kunci cache kini memuat rentang tanggal
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

/* ============================ RENTANG TANGGAL ============================ */

/**
 * Ubah dayFrom/dayTo (nomor hari, 1-based) menjadi {from, to, count}, dijepit ke
 * 1..dayCount dan ditukar bila terbalik. Kosong/tidak valid, atau meliputi seluruh
 * periode -> null (artinya: tanpa filter).
 */
function resolveDayRange_(dayFrom, dayTo, dayCount) {
  const f = parseInt(dayFrom, 10);
  const t = parseInt(dayTo, 10);
  if (isNaN(f) || isNaN(t) || !(dayCount > 0)) return null;
  let from = Math.max(1, Math.min(f, dayCount));
  let to = Math.max(1, Math.min(t, dayCount));
  if (from > to) { const x = from; from = to; to = x; }
  if (from === 1 && to === dayCount) return null;
  return { from: from, to: to, count: to - from + 1 };
}

/** Salinan grid dengan kolom hari dipotong ke rentang (karyawan & urutan baris tetap). */
function sliceGridDays_(grid, range) {
  if (!range) return grid;
  return Object.assign({}, grid, {
    dayNames: (grid.dayNames || []).slice(range.from - 1, range.to),
    dayCount: range.count,
    employees: grid.employees.map(e => Object.assign({}, e, {
      days: e.days.slice(range.from - 1, range.to)
    }))
  });
}

/** Geser indeks "hari ini" ke koordinat rentang; di luar rentang -> -1. */
function shiftTodayDayIndex_(idx, range) {
  if (!range || typeof idx !== 'number' || idx < 0) return idx;
  const rel = idx - (range.from - 1);
  return (rel >= 0 && rel < range.count) ? rel : -1;
}

/**
 * Rekap per karyawan dari DUA grid (dicocokkan per POSISI BARIS, sama seperti getRekap_).
 * Dipakai saat ada filter rentang; tanpa filter tetap memakai getRekap_().
 * Rank 1 = % kehadiran paling rendah (sama dengan getRekap_).
 */
function buildRekapPositional_(jadwal, realisasi) {
  const dayCount = Math.min(jadwal.dayCount, realisasi.dayCount);
  const numEmployees = Math.min(jadwal.employees.length, realisasi.employees.length);
  const rows = [];
  for (let i = 0; i < numEmployees; i++) {
    const j = jadwal.employees[i];
    const r = realisasi.employees[i];
    const tally = tallyAttendance_(j.days.slice(0, dayCount), r.days.slice(0, dayCount));
    const persen = tally.jadwal > 0 ? (tally.hadir / tally.jadwal) : 0;
    rows.push(makeRekapRow_(r.no || j.no || (i + 1), r.nama || j.nama, r.jabatan || j.jabatan, tally, persen));
  }
  rows.slice().sort((a, b) => a.persenKehadiran - b.persenKehadiran)
    .forEach((row, idx) => { row.rank = idx + 1; });
  return rows;
}

/* ============================ RELIEVER ============================ */

/**
 * Ringkasan reliever dari sheet BackupReliever (baris = reliever, kolom = tanggal,
 * isi sel = nama karyawan yang digantikan; terisi saat reliever absen "Masuk").
 *  - jumlah    : jumlah reliever BERBEDA yang bertugas pada rentang
 *  - hariTugas : total hari tugas semua reliever (1 sel terisi = 1 hari)
 * Bila filter Department aktif, hanya sel yang karyawan-digantikannya ada di department itu.
 */
function computeRelieverSummary_(range, allowedNamesSet) {
  const empty = { jumlah: 0, hariTugas: 0, daftar: [] };
  try {
    const sheet = getSpreadsheet_().getSheetByName(CONFIG.SHEET_BACKUPRELIEVER);
    if (!sheet) return empty;
    const grid = readGridFromSheetObj_(sheet);
    const from = range ? range.from : 1;
    const to = range ? range.to : grid.dayCount;
    const daftar = [];
    let hariTugas = 0;
    grid.employees.forEach(e => {
      let hari = 0;
      e.days.forEach((val, i) => {
        const day = i + 1;
        if (day < from || day > to) return;
        const digantikan = String(val || '').trim();
        if (!digantikan) return;
        if (!isNameAllowed_(digantikan, allowedNamesSet)) return;
        hari++;
      });
      if (hari > 0) {
        // kolom ke-3 sheet BackupReliever = Perusahaan (terbaca sebagai "jabatan" oleh readGrid)
        daftar.push({ nama: e.nama, perusahaan: e.jabatan || '', hari: hari });
        hariTugas += hari;
      }
    });
    daftar.sort((a, b) => b.hari - a.hari || String(a.nama).localeCompare(String(b.nama)));
    return { jumlah: daftar.length, hariTugas: hariTugas, daftar: daftar };
  } catch (err) {
    Logger.log('computeRelieverSummary_: ' + err.toString());
    return empty;
  }
}

/* ============================ ENTRY ============================ */

function getAllData(token, department, dayFrom, dayTo) {
  verifySession_(token);
  return getCachedAllData_(department, dayFrom, dayTo);
}

function getCachedAllData_(department, dayFrom, dayTo) {
  const cache = CacheService.getScriptCache();
  const key = CACHE_ALLDATA_PREFIX_ + String(department || 'ALL') + '_' +
    String(dayFrom || '') + '-' + String(dayTo || '');
  const cached = cache.get(key);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* cache korup, hitung ulang */ }
  }
  const result = getAllDataInternal_(false, department, dayFrom, dayTo);
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

function getAllDataInternal_(skipSync, department, dayFrom, dayTo) {
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

  // ---- Rentang tanggal (opsional) ----
  const fullDayCount = Math.min(jadwalFull.dayCount, realisasiFull.dayCount);
  const range = resolveDayRange_(dayFrom, dayTo, fullDayCount);

  let rekapAll;
  if (range) {
    const jRange = sliceGridDays_(jadwalFull, range);
    const rRange = sliceGridDays_(realisasiFull, range);
    rekapAll = attachJamKerjaToRekap_(
      attachNikToRekap_(buildRekapPositional_(jRange, rRange)), jRange, null, range);
  } else {
    rekapAll = attachJamKerjaToRekap_(attachNikToRekap_(getRekap_()), jadwalFull);
  }
  const rekap = allowedNames ? rekapAll.filter(r => allowedNames.has(nameKey_(r.nama))) : rekapAll;

  const jadwalAgg = range ? sliceGridDays_(jadwal, range) : jadwal;
  const realisasiAgg = range ? sliceGridDays_(realisasi, range) : realisasi;

  const my = parsePeriodMonthYear_(String(jadwalFull.period || '').trim());

  return {
    period: jadwalFull.period,
    dayNames: jadwalFull.dayNames,
    dayCount: jadwalFull.dayCount,
    jadwal: jadwal.employees,
    realisasi: realisasi.employees,
    dashboard: computeDashboardSummaryFromRekapLengkap_(rekap),
    rekap: rekap,
    trenHarian: computeTrenHarian_(jadwalAgg, realisasiAgg, jamKerjaSheet, allowedNames, range, fullDayCount),
    jamKerjaRingkasan: computeJamKerjaRingkasan_(rekap),
    statusCodes: CONFIG.STATUS_CODES,
    shiftCodes: getShiftCodesDynamic_(),
    alignmentWarnings: checkRowAlignment_(jadwalFull, realisasiFull),
    departmentFilter: isFiltered ? deptFilter : null,
    lastUpdate: new Date().toISOString(),
    jamScan: jamScan.employees,
    rangeInfo: {
      year: my ? my.year : null,
      monthIndex: my ? my.monthIndex : null,
      dayCount: fullDayCount,
      dayFrom: range ? range.from : 1,
      dayTo: range ? range.to : fullDayCount,
      active: !!range
    },
    reliever: computeRelieverSummary_(range, allowedNames)
  };
}

/* ============================ TREN HARIAN ============================ */

/** Kode Realisasi -> nama deret tren (selain hadir/lembur yang ditangani khusus). */
const TREN_CODE_TO_SERIES_ = { 'A': 'alfa', 'I': 'izin', 'S': 'sakit', 'C': 'cuti', 'OFF': 'off' };

/**
 * @param range         opsional {from,to,count}: jadwal & realisasi SUDAH dipotong ke rentang ini
 * @param fullDayCount  jumlah hari periode penuh (untuk menghitung indeks "hari ini")
 */
function computeTrenHarian_(jadwal, realisasi, jamKerjaSheet, allowedNamesSet, range, fullDayCount) {
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
    totalJamKerjaPerHari: computeJamKerjaTrend_(dayCount, jamKerjaSheet, allowedNamesSet, range ? range.from - 1 : 0),
    todayDayIndex: shiftTodayDayIndex_(computeTodayDayIndex_(realisasi.period, fullDayCount || dayCount), range)
  };
}

/** @param offset indeks kolom hari pertama yang diambil (0 = dari awal periode). */
function computeJamKerjaTrend_(dayCount, jamKerjaSheet, allowedNamesSet, offset) {
  const total = new Array(dayCount).fill(0);
  if (!jamKerjaSheet) return total;
  const off = offset || 0;
  readGridFromSheetObj_(jamKerjaSheet).employees.forEach(emp => {
    if (!isNameAllowed_(emp.nama, allowedNamesSet)) return;
    emp.days.forEach((val, i) => {
      const idx = i - off;
      if (idx < 0 || idx >= dayCount) return;
      const num = Number(val);
      if (val !== '' && val !== null && !isNaN(num)) total[idx] += num;
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
