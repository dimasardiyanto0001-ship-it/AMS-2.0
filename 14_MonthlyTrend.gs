/*************************************************************
 * 14_MonthlyTrend.gs   (pengganti 20_TrendComparison.gs)
 * Tren kehadiran bulanan untuk grafik Dashboard.
 *
 * Sheet "Rekap_Bulanan" (1 baris per tahun+bulan):
 *   Tahun | Bulan | TotalHadir | TotalAlfa | TotalIzin | TotalSakit |
 *   TotalCuti | TotalOff | TotalKaryawan | RataPersenKehadiran
 * RataPersenKehadiran disimpan sebagai pecahan 0..1.
 *************************************************************/

const SHEET_REKAP_BULANAN = 'Rekap_Bulanan';
const TARGET_KEHADIRAN_PERSEN = 95;
const REKAP_BULANAN_HEADERS = ['Tahun', 'Bulan', 'TotalHadir', 'TotalAlfa', 'TotalIzin', 'TotalSakit', 'TotalCuti', 'TotalOff', 'TotalKaryawan', 'RataPersenKehadiran'];
const BULAN_LABEL_ = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];

/** Hitung total satu periode dari grid Jadwal+Realisasi. @return {Array} satu baris Rekap_Bulanan */
function buildRekapBulananRow_(tahun, bulan, jadwalGrid, realisasiGrid) {
  const rekap = computeRekapDariGrid_(jadwalGrid, realisasiGrid);
  const sum = key => rekap.reduce((total, r) => total + (Number(r[key]) || 0), 0);
  const totalKaryawan = rekap.length;
  const rataPersen = totalKaryawan ? sum('persenKehadiran') / totalKaryawan : 0;

  return [tahun, bulan, sum('hadir'), sum('alfa'), sum('izin'), sum('sakit'), sum('cuti'), sum('off'), totalKaryawan, rataPersen];
}

/**
 * Simpan snapshot rekap satu bulan (dipanggil saat periode ditutup).
 * Anti-duplikat: tahun+bulan yang sama ditimpa.
 */
function generateRekapBulananSnapshot_(tahun, bulan, jadwalSheet, realisasiSheet) {
  const ss = getSpreadsheet_();
  const rowData = buildRekapBulananRow_(tahun, bulan, readGridFromSheetObj_(jadwalSheet), readGridFromSheetObj_(realisasiSheet));

  let sheet = ss.getSheetByName(SHEET_REKAP_BULANAN);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_REKAP_BULANAN);
    sheet.appendRow(REKAP_BULANAN_HEADERS);
  }

  const existing = sheet.getDataRange().getValues();
  const dupRow = existing.findIndex((r, i) => i > 0 && r[0] == tahun && r[1] == bulan);
  if (dupRow > 0) sheet.getRange(dupRow + 1, 1, 1, rowData.length).setValues([rowData]);
  else sheet.appendRow(rowData);
}

/** Tren 1 tahun (persen hadir per bulan) + KPI untuk grafik Dashboard. */
function getMonthlyTrendYear(tahun, token) {
  verifySession_(token);

  const emptyKpi = { targetKehadiran: TARGET_KEHADIRAN_PERSEN, persenKehadiranBulanIni: null, statusTerhadapTarget: 'tidak ada data' };
  const sheet = getSpreadsheet_().getSheetByName(SHEET_REKAP_BULANAN);
  if (!sheet) return { labels: BULAN_LABEL_, persenHadir: new Array(12).fill(null), kpi: emptyKpi };

  const data = sheet.getDataRange().getValues();
  const header = data.shift();
  const col = {
    tahun: header.indexOf('Tahun'),
    bulan: header.indexOf('Bulan'),
    persen: header.indexOf('RataPersenKehadiran')
  };

  const persenHadir = new Array(12).fill(null);
  data.forEach(row => {
    if (row[col.tahun] != tahun) return;
    const bulanIndex = Number(row[col.bulan]) - 1;
    if (!(bulanIndex >= 0 && bulanIndex < 12)) return;
    persenHadir[bulanIndex] = row[col.persen] !== '' ? Math.round(Number(row[col.persen]) * 1000) / 10 : null;
  });

  let lastIndex = -1;
  persenHadir.forEach((v, i) => { if (v !== null) lastIndex = i; });
  const kpiBulanIni = lastIndex >= 0 ? persenHadir[lastIndex] : null;

  return {
    labels: BULAN_LABEL_,
    persenHadir: persenHadir,
    kpi: {
      persenKehadiranBulanIni: kpiBulanIni,
      targetKehadiran: TARGET_KEHADIRAN_PERSEN,
      statusTerhadapTarget: kpiBulanIni === null
        ? 'tidak ada data'
        : (kpiBulanIni >= TARGET_KEHADIRAN_PERSEN ? 'tercapai' : 'di bawah target')
    }
  };
}
