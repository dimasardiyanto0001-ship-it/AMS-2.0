/*************************************************************
 * 33_RelieverData.gs
 * Menu sidebar "Data Reliever": tabel absen reliever per tanggal + download Excel/PDF.
 *
 *   getRelieverData      baris data untuk tabel (rentang tanggal bebas, boleh lintas bulan)
 *   exportRelieverExcel  -> .xlsx (base64) dari data yang sama
 *   exportRelieverPdf    -> .pdf  (base64) dari data yang sama
 *
 * Sumber: sheet LogAbsenOnlineReliever (sama dengan panel Reliever di Dashboard).
 * 1 baris = 1 reliever x 1 tanggal x 1 karyawan yang digantikan, hanya yang punya absen Masuk
 * (aturan yang sama dengan hitungRelieverSummary_ di 24_BackupReliever.gs).
 * Jam diambil dari kolom Timestamp (kolom A) sehingga tidak terpengaruh format sel jam.
 *
 * Izin: menu 'relieverdata' (ada di MENU_LIST, 00_Config.gs; atur lewat Setting > Hak Akses).
 * Export memakai ulang withTempReportSheet_ (29_Export.gs): spreadsheet sementara selalu dibuang.
 * Baris untuk export dihitung ulang di server dari rentang + kata cari, bukan dikirim dari browser.
 *************************************************************/

const RELIEVER_DATA_LOG_BACA_MAKS_ = 20000;
const RELIEVER_DATA_KOLOM_ = ['No', 'Tanggal', 'Nama Reliever', 'Perusahaan', 'Menggantikan', 'Keterangan', 'Masuk', 'Pulang', 'Jarak (m)', 'Status'];

function relieverDataIso_(s) {
  const v = String(s === null || s === undefined ? '' : s).trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(v) ? v : '';
}

/**
 * @param {string} tglMulai 'yyyy-MM-dd' (kosong = dari awal)
 * @param {string} tglAkhir 'yyyy-MM-dd' (kosong = sampai akhir)
 * @param {string} cari     kata cari (nama reliever / perusahaan / karyawan digantikan); kosong = semua
 * @return {Object[]} {tanggal, reliever, perusahaan, menggantikan, keterangan, jamMasuk, jamPulang, jarak, status, foto}
 */
function bacaRelieverDataRows_(tglMulai, tglAkhir, cari) {
  const mulai = relieverDataIso_(tglMulai);
  const akhir = relieverDataIso_(tglAkhir);
  const kata = String(cari || '').trim().toLowerCase();
  const ss = getSpreadsheet_();
  const tz = Session.getScriptTimeZone() || 'GMT+7';
  const str = v => String(v === null || v === undefined ? '' : v).trim();

  const logSheet = ss.getSheetByName(CONFIG.SHEET_LOGABSENRELIEVER);
  if (!logSheet || logSheet.getLastRow() < 2) return [];

  // Nama & perusahaan resmi dari daftar reliever terdaftar (bila ada), seperti di Dashboard.
  const terdaftar = {};
  const dataSheet = ss.getSheetByName(CONFIG.SHEET_DATARELIEVER);
  if (dataSheet && dataSheet.getLastRow() >= 2) {
    dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, 2).getValues().forEach(r => {
      const nama = str(r[0]);
      if (nama) terdaftar[nameKey_(nama)] = { nama: nama, perusahaan: str(r[1]) };
    });
  }

  const last = logSheet.getLastRow();
  const awal = Math.max(2, last - RELIEVER_DATA_LOG_BACA_MAKS_ + 1);
  const pasangan = {};

  logSheet.getRange(awal, 1, last - awal + 1, 13).getValues().forEach(r => {
    const ts = r[0];
    if (!(ts instanceof Date) || isNaN(ts.getTime())) return;
    const tgl = Utilities.formatDate(ts, tz, 'yyyy-MM-dd');
    if (mulai && tgl < mulai) return;
    if (akhir && tgl > akhir) return;

    const nama = str(r[2]);
    const jenis = str(r[6]);
    if (!nama || (jenis !== 'Masuk' && jenis !== 'Pulang')) return;

    const k = tgl + '|' + nameKey_(nama) + '|' + nameKey_(str(r[4]));
    let p = pasangan[k];
    if (!p) {
      const info = terdaftar[nameKey_(nama)];
      p = pasangan[k] = {
        tanggal: tgl,
        reliever: info ? info.nama : nama,
        perusahaan: info ? info.perusahaan : str(r[3]),
        menggantikan: str(r[4]),
        keterangan: str(r[5]),
        jamMasuk: '', jamPulang: '', jarak: '', status: '', foto: '',
        masuk: false
      };
    }

    const jam = Utilities.formatDate(ts, tz, 'HH:mm');
    if (jenis === 'Masuk') {
      if (!p.masuk) {
        p.masuk = true;
        p.jamMasuk = jam;
        const jarak = Number(r[10]);
        p.jarak = (r[10] === '' || r[10] === null || isNaN(jarak)) ? '' : Math.round(jarak);
        p.foto = str(r[12]);
      }
      if (str(r[11]) === 'Di Luar Radius' || !p.status) p.status = str(r[11]);
    } else {
      p.jamPulang = jam;
    }
  });

  return Object.keys(pasangan).map(k => pasangan[k])
    .filter(p => p.masuk)
    .filter(p => !kata || (p.reliever + ' ' + p.perusahaan + ' ' + p.menggantikan).toLowerCase().indexOf(kata) !== -1)
    .sort((a, b) => a.tanggal.localeCompare(b.tanggal) || a.reliever.localeCompare(b.reliever))
    .map(p => ({
      tanggal: p.tanggal, reliever: p.reliever, perusahaan: p.perusahaan, menggantikan: p.menggantikan,
      keterangan: p.keterangan, jamMasuk: p.jamMasuk, jamPulang: p.jamPulang, jarak: p.jarak,
      status: p.status, foto: p.foto
    }));
}

/* ============================ TABEL ============================ */

function getRelieverData(tglMulai, tglAkhir, token) {
  requireMenuAccess_(token, 'relieverdata');
  return { rows: bacaRelieverDataRows_(tglMulai, tglAkhir, '') };
}

/* ============================ EXPORT ============================ */

/** Baris untuk file export. Tanggal ditulis yyyy-MM-dd supaya tidak salah dibaca (dd/MM vs MM/dd) oleh Sheets. */
function relieverDataTabel_(rows) {
  return rows.map((r, i) => [
    i + 1, r.tanggal, r.reliever, r.perusahaan, r.menggantikan, r.keterangan,
    r.jamMasuk || '-', r.jamPulang || '-', (r.jarak === '' || r.jarak === null) ? '-' : r.jarak, r.status
  ]);
}

function relieverDataNamaFile_(tglMulai, tglAkhir, ext) {
  const m = relieverDataIso_(tglMulai) || 'awal';
  const a = relieverDataIso_(tglAkhir) || 'akhir';
  return 'Data_Reliever_' + m + '_sd_' + a + '.' + ext;
}

function ambilRelieverDataUntukExport_(tglMulai, tglAkhir, cari, token) {
  requireMenuAccess_(token, 'relieverdata');
  const rows = bacaRelieverDataRows_(tglMulai, tglAkhir, cari);
  if (!rows.length) throw new Error('Tidak ada data reliever pada rentang tanggal ini.');
  return rows;
}

function exportRelieverExcel(tglMulai, tglAkhir, cari, token) {
  const rows = ambilRelieverDataUntukExport_(tglMulai, tglAkhir, cari, token);

  return withTempReportSheet_('TempReliever_', 'relieverdata', 'Data Reliever', RELIEVER_DATA_KOLOM_, relieverDataTabel_(rows),
    headerRange => headerRange.setBackground('#2563eb').setFontColor('#ffffff'),
    fileId => ({
      base64: Utilities.base64Encode(exportSpreadsheetAsXlsxBlob_(fileId, 'laporan Data Reliever').getBytes()),
      filename: relieverDataNamaFile_(tglMulai, tglAkhir, 'xlsx'),
      mimeType: MIME_XLS_
    })
  );
}

function exportRelieverPdf(tglMulai, tglAkhir, cari, token) {
  const rows = ambilRelieverDataUntukExport_(tglMulai, tglAkhir, cari, token);

  return withTempReportSheet_('TempRelieverPDF_', 'relieverdata', 'Data Reliever', RELIEVER_DATA_KOLOM_, relieverDataTabel_(rows),
    headerRange => headerRange,
    (fileId, sheet) => {
      const url = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export' +
        '?format=pdf&size=A4&portrait=false&fitw=true&gridlines=true' +
        '&printtitle=false&sheetnames=false&pagenumbers=true&gid=' + sheet.getSheetId();
      const response = UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
      });
      if (response.getResponseCode() !== 200) {
        throw new Error('Gagal membuat PDF Data Reliever (kode ' + response.getResponseCode() + ').');
      }
      return {
        base64: Utilities.base64Encode(response.getBlob().getBytes()),
        filename: relieverDataNamaFile_(tglMulai, tglAkhir, 'pdf'),
        mimeType: MIME_PDF_
      };
    }
  );
}
