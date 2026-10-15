/*************************************************************
 * 32_JadwalHariIni.gs
 * Jadwal shift satu karyawan untuk HARI INI (route publik halaman Absen Online,
 * tanpa login). Hasil: {namaShift, jamMasuk, jamPulang, keterangan}; bila belum ada
 * jadwal, namaShift "-" dan alasannya di keterangan (ditampilkan apa adanya di halaman).
 *
 * PERIODE HARUS SESUAI: bulan & tahun periode aktif (sel A2 sheet Jadwal) harus sama dengan
 * bulan & tahun hari ini. Bila beda (mis. tanggal 1 sebelum "Mulai Periode Baru"), jawabannya
 * "Periode belum diganti ..." dan BUKAN shift pada tanggal yang sama di bulan lalu. Bila nama
 * periode tidak bisa dibaca bulan/tahunnya, pengecekan dilewati (perilaku lama).
 *************************************************************/

const HEADER_SCAN_ROWS_ = 10;
const HEADER_MIN_DAY_NUMBERS_ = 15;

function jadwalHariIniEmpty_(keterangan, namaShift) {
  return { namaShift: namaShift || '-', jamMasuk: '-', jamPulang: '-', keterangan: keterangan };
}

/** Indeks baris header tanggal (>= 15 angka 1..31 di 10 baris pertama), atau -1. */
function findDayHeaderRowIdx_(data) {
  for (let r = 0; r < Math.min(data.length, HEADER_SCAN_ROWS_); r++) {
    const dayNumbers = data[r].filter(v => typeof v === 'number' && v >= 1 && v <= 31).length;
    if (dayNumbers >= HEADER_MIN_DAY_NUMBERS_) return r;
  }
  return -1;
}

function getJadwalHariIni(nama) {
  const ss = getSpreadsheet_();
  const jadwalSheet = ss.getSheetByName(CONFIG.SHEET_JADWAL);
  const shiftSheet = ss.getSheetByName(SHEET_SETTING_JAMKERJA);
  if (!jadwalSheet || !shiftSheet) return jadwalHariIniEmpty_('Sheet Jadwal/SettingJamKerja tidak ditemukan');

  const data = jadwalSheet.getDataRange().getValues();

  const headerRowIdx = findDayHeaderRowIdx_(data);
  if (headerRowIdx === -1) return jadwalHariIniEmpty_('Format header Jadwal tidak dikenali');
  const header = data[headerRowIdx];

  // Periode aktif (A2) harus bulan & tahun yang sama dengan hari ini.
  const today = new Date();
  const periodText = String((data[1] && data[1][0]) || '').trim();
  const my = parsePeriodMonthYear_(periodText);
  if (my && (today.getFullYear() !== my.year || today.getMonth() !== my.monthIndex)) {
    const namaPeriode = periodText.replace(/^periode\s*:?\s*/i, '');
    // Pesan dibuat pendek: halaman menampilkannya dengan huruf besar di kartu jadwal.
    return jadwalHariIniEmpty_('Periode belum diganti (' + namaPeriode + ')');
  }

  const dateColIdx = header.indexOf(today.getDate());
  if (dateColIdx === -1) return jadwalHariIniEmpty_('Tanggal hari ini tidak ada di periode aktif');

  let namaColIdx = header.findIndex(v => String(v).toLowerCase().includes('nama'));
  if (namaColIdx === -1) namaColIdx = 1;

  const target = String(nama).trim().toLowerCase();
  let kodeShift = null;
  for (let r = headerRowIdx + 1; r < data.length; r++) {
    if (String(data[r][namaColIdx]).trim().toLowerCase() === target) {
      kodeShift = data[r][dateColIdx];
      break;
    }
  }

  if (!kodeShift) return jadwalHariIniEmpty_('Nama tidak ditemukan di Jadwal periode ini');

  const kode = String(kodeShift).trim().toUpperCase();
  if (kode === 'OFF') return jadwalHariIniEmpty_('Hari ini jadwal libur', 'OFF');

  const shiftData = shiftSheet.getDataRange().getValues();
  for (let i = 1; i < shiftData.length; i++) {
    if (String(shiftData[i][0]).trim().toUpperCase() === kode) {
      return {
        namaShift: shiftData[i][1],
        jamMasuk: formatJam_(shiftData[i][2]),
        jamPulang: formatJam_(shiftData[i][3]),
        keterangan: ''
      };
    }
  }
  return jadwalHariIniEmpty_('Kode shift "' + kodeShift + '" tidak ada di SettingJamKerja', kodeShift);
}
