/*************************************************************
 * 31_MonthlyReport.gs
 * Laporan bulanan otomatis lewat email (trigger tanggal 5 jam 07:00, lihat MONTHLY_REPORT_DAY_):
 * ringkasan kehadiran, analisis tren, 5 terbaik/terendah, daftar kehadiran rendah,
 * plus PDF Dashboard & Rekap.
 *
 * Laporan selalu untuk BULAN LALU. Sumber data dipilih berlapis (resolveMonthlyReportSource_):
 *   1. arsip periode bulan lalu, bila ada (kasus normal: periode sudah ditutup lewat
 *      "Mulai Periode Baru") -> PDF diambil dari sheet arsip;
 *   2. periode aktif, bila periode aktif itu bulan lalu (belum ditutup);
 *   3. selain itu arsip terbaru (atau periode aktif bila belum ada arsip) DAN email memuat
 *      peringatan jelas bahwa data bulan lalu belum diarsipkan.
 * Analisis tren membandingkan dengan arsip tepat sebelum periode yang dilaporkan.
 *
 * Pasang trigger sekali dari editor: jalankan setupMonthlyReportTrigger().
 * Penerima: Script Property REPORT_EMAIL (lihat 00_Config.gs).
 *************************************************************/

/** Jadwal kirim: setiap bulan pada tanggal ini, jam ini. Ubah lalu jalankan ulang setupMonthlyReportTrigger(). */
const MONTHLY_REPORT_DAY_ = 5;
const MONTHLY_REPORT_HOUR_ = 7;

const MONTHLY_REPORT_TOP_N_ = 5;
const TREND_STABLE_POINTS_ = 0.5;

const REPORT_CELL_STYLE_ = 'padding:6px;border:1px solid #e2e8f0;';
const REPORT_TH_STYLE_ = 'padding:6px;border:1px solid #e2e8f0;background:#f1f5f9;text-align:left;';
const REPORT_RIGHT_ = 'text-align:right;';

/* ============================ KIRIM ============================ */

function sendMonthlyReportEmail() {
  let source = resolveMonthlyReportSource_(new Date());

  let data;
  try {
    data = loadMonthlyReportData_(source);
  } catch (e) {
    if (source.kind !== 'arsip') throw e;
    Logger.log('Gagal membaca data arsip untuk laporan bulanan: ' + e.message);
    source = activePeriodReportSource_(source.prevRow,
      'Data arsip periode "' + source.periodText + '" tidak dapat dibaca (' + e.message + '), sehingga laporan ini memakai periode aktif.');
    data = loadMonthlyReportData_(source);
  }

  const prevAvg = source.prevRow ? getArchiveRowAvgAttendance_(source.prevRow) : null;
  const analysis = generateMonthlyAnalysis_(data.dashboard, data.rekap, prevAvg);
  const htmlBody = buildMonthlyReportHtml_(source.periodText, data.dashboard, analysis, source.notice);
  const attachments = collectMonthlyReportAttachments_(source, sanitizePeriodKey_(source.periodText));

  const recipient = getReportEmail_();
  MailApp.sendEmail({
    to: recipient,
    subject: '📊 Laporan Bulanan Kehadiran - ' + source.periodText,
    htmlBody: htmlBody,
    attachments: attachments
  });

  Logger.log('Laporan bulanan berhasil dikirim ke ' + recipient + ' untuk periode ' + source.periodText);
}

/** Pasang trigger bulanan (hapus trigger lama dulu agar tidak dobel). Jalankan sekali dari editor. */
function setupMonthlyReportTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendMonthlyReportEmail') ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('sendMonthlyReportEmail')
    .timeBased()
    .onMonthDay(MONTHLY_REPORT_DAY_)
    .atHour(MONTHLY_REPORT_HOUR_)
    .create();

  Logger.log('Trigger laporan bulanan berhasil dipasang: setiap tanggal ' + MONTHLY_REPORT_DAY_ +
    ' jam ' + ('0' + MONTHLY_REPORT_HOUR_).slice(-2) + ':00.');
}

/* ============================ PILIH PERIODE LAPORAN ============================ */

/** Bulan kalender sebelum `today`. @return {{year:number, monthIndex:number}} */
function previousCalendarMonth_(today) {
  const d = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  return { year: d.getFullYear(), monthIndex: d.getMonth() };
}

function periodMatchesMonth_(periodText, month) {
  const my = parsePeriodMonthYear_(periodText);
  return !!my && my.year === month.year && my.monthIndex === month.monthIndex;
}

/** Semua baris ArsipIndex (tanpa header), dari yang terlama ke terbaru. */
function readArchiveRows_() {
  const idx = getSpreadsheet_().getSheetByName(ARCHIVE_INDEX_SHEET);
  return idx ? idx.getDataRange().getValues().slice(1).filter(r => r[0]) : [];
}

function activePeriodText_() {
  return String(getSheet_(CONFIG.SHEET_JADWAL).getRange(2, 1).getValue() || '').trim() || 'Periode tidak diketahui';
}

function activePeriodReportSource_(prevRow, notice) {
  return { kind: 'aktif', row: null, prevRow: prevRow || null, periodText: activePeriodText_(), notice: notice || '' };
}

/**
 * Tentukan periode yang dilaporkan.
 * @return {{kind:'arsip'|'aktif', row:Array|null, prevRow:Array|null, periodText:string, notice:string}}
 *   row = baris ArsipIndex periode yang dilaporkan (kind 'arsip'); prevRow = arsip tepat sebelumnya.
 */
function resolveMonthlyReportSource_(today) {
  const expected = previousCalendarMonth_(today);
  const archives = readArchiveRows_();

  for (let i = archives.length - 1; i >= 0; i--) {
    if (periodMatchesMonth_(archives[i][0], expected)) {
      return { kind: 'arsip', row: archives[i], prevRow: i > 0 ? archives[i - 1] : null, periodText: String(archives[i][0]), notice: '' };
    }
  }

  const lastArchive = archives.length ? archives[archives.length - 1] : null;
  if (periodMatchesMonth_(activePeriodText_(), expected)) {
    return activePeriodReportSource_(lastArchive, '');
  }

  // Data bulan lalu tidak ditemukan di arsip maupun periode aktif: tetap kirim, tetapi beri peringatan.
  const expectedLabel = INDONESIAN_MONTH_NAMES_UPPER_[expected.monthIndex] + ' ' + expected.year;
  const usedLabel = lastArchive ? String(lastArchive[0]) : activePeriodText_();
  const notice = 'Periode ' + expectedLabel + ' belum diarsipkan dan bukan periode aktif, sehingga laporan ini memakai periode "' +
    usedLabel + '". Jalankan "Mulai Periode Baru" untuk menutup periode bulan lalu.';

  return lastArchive
    ? { kind: 'arsip', row: lastArchive, prevRow: archives.length > 1 ? archives[archives.length - 2] : null, periodText: String(lastArchive[0]), notice: notice }
    : activePeriodReportSource_(null, notice);
}

/** Dashboard + rekap untuk sumber yang dipilih (arsip: dihitung dari sheet arsip). */
function loadMonthlyReportData_(source) {
  if (source.kind === 'arsip') {
    const archived = getArchivedDataInternal_(source.row[1]);
    return { dashboard: archived.dashboard, rekap: archived.rekap };
  }
  return { dashboard: getDashboardSummary_(), rekap: getRekap_() };
}

/** Lampiran PDF Dashboard & Rekap; kegagalan satu lampiran hanya dicatat di log. */
function collectMonthlyReportAttachments_(source, periodKey) {
  const fromArchive = source.kind === 'arsip';
  const jobs = [
    ['Dashboard', 'Dashboard_', () => fromArchive
      ? exportArchivedSheetPdfBlob_(source.row[4], CONFIG.SHEET_DASHBOARD)
      : exportSheetPdfBlob_(CONFIG.SHEET_DASHBOARD)],
    ['Rekap', 'Rekap_', () => fromArchive
      ? exportArchivedSheetPdfBlob_(source.row[5], CONFIG.SHEET_REKAP)
      : exportSheetPdfBlob_(CONFIG.SHEET_REKAP)]
  ];

  const attachments = [];
  jobs.forEach(([label, prefix, makeBlob]) => {
    try {
      attachments.push(makeBlob().setName(prefix + periodKey + '.pdf'));
    } catch (e) {
      Logger.log('Gagal lampirkan PDF ' + label + ': ' + e.message);
    }
  });
  return attachments;
}

/* ============================ ANALISIS ============================ */

/** Baris rekap yang punya nama dan persenKehadiran bertipe angka (pecahan 0..1). */
function validRekapRows_(rekap) {
  return rekap.filter(r => r.nama && String(r.nama).trim() !== '' && typeof r.persenKehadiran === 'number');
}

/** Rata-rata persenKehadiran per karyawan (pecahan 0..1); 0 bila tidak ada baris valid. */
function meanAttendanceFraction_(validRows) {
  return validRows.length
    ? validRows.reduce((sum, r) => sum + r.persenKehadiran, 0) / validRows.length
    : 0;
}

function sumRekapField_(rows, field) {
  return rows.reduce((total, r) => total + (Number(r[field]) || 0), 0);
}

/** prevAvg: rata-rata periode pembanding (pecahan) atau null; bila tidak diberikan, memakai arsip terakhir. */
function generateMonthlyAnalysis_(dashboard, rekap, prevAvg) {
  const validRekap = validRekapRows_(rekap);
  const avgPct = meanAttendanceFraction_(validRekap);

  const sorted = [...validRekap].sort((a, b) => b.persenKehadiran - a.persenKehadiran);
  const lowAttendance = validRekap
    .filter(r => r.persenKehadiran < CONFIG.REPORT_LOW_ATTENDANCE_THRESHOLD)
    .sort((a, b) => a.persenKehadiran - b.persenKehadiran);

  return {
    avgPct,
    topBest: sorted.slice(0, MONTHLY_REPORT_TOP_N_),
    topWorst: [...sorted].reverse().slice(0, MONTHLY_REPORT_TOP_N_),
    lowAttendance,
    totalAlfa: sumRekapField_(validRekap, 'alfa'),
    totalIzin: sumRekapField_(validRekap, 'izin'),
    totalSakit: sumRekapField_(validRekap, 'sakit'),
    totalCuti: sumRekapField_(validRekap, 'cuti'),
    trendText: describeAttendanceTrend_(avgPct, prevAvg === undefined ? getPreviousArchivedAvgAttendance_() : prevAvg)
  };
}

/** Kalimat tren; prevAvg null = belum ada periode arsip untuk dibandingkan. */
function describeAttendanceTrend_(avgPct, prevAvg) {
  if (prevAvg === null) return 'Belum ada data periode sebelumnya yang diarsipkan untuk dibandingkan.';

  const pct = v => (v * 100).toFixed(1);
  const diffPoin = (avgPct - prevAvg) * 100;

  if (Math.abs(diffPoin) < TREND_STABLE_POINTS_) {
    return 'Rata-rata kehadiran relatif STABIL dibanding periode sebelumnya (' + pct(prevAvg) + '%).';
  }
  if (diffPoin > 0) {
    return 'Rata-rata kehadiran NAIK ' + diffPoin.toFixed(1) + ' poin persen dibanding periode sebelumnya (' +
      pct(prevAvg) + '% → ' + pct(avgPct) + '%). 👍';
  }
  return 'Rata-rata kehadiran TURUN ' + Math.abs(diffPoin).toFixed(1) + ' poin persen dibanding periode sebelumnya (' +
    pct(prevAvg) + '% → ' + pct(avgPct) + '%). Perlu ditindaklanjuti. ⚠️';
}

/**
 * Rata-rata kehadiran satu periode arsip (pecahan 0..1), atau null bila sheet arsip tidak
 * lengkap / tidak ada baris valid. Dihitung dengan cara yang SAMA seperti rata-rata periode
 * berjalan (rata-rata persen per karyawan). Versi lama memakai total hadir / total jadwal yang
 * dibulatkan ke persen bulat, sehingga perbandingan tren tidak sebanding.
 */
function getArchiveRowAvgAttendance_(row) {
  const ss = getSpreadsheet_();
  const jadwalSheet = ss.getSheetByName(row[2]);
  const realisasiSheet = ss.getSheetByName(row[3]);
  if (!jadwalSheet || !realisasiSheet) return null;

  const rekap = computeRekapDariGrid_(readGridFromSheetObj_(jadwalSheet), readGridFromSheetObj_(realisasiSheet));
  const valid = validRekapRows_(rekap);
  return valid.length ? meanAttendanceFraction_(valid) : null;
}

/** Rata-rata kehadiran periode yang TERAKHIR diarsipkan, atau null bila belum ada arsip. */
function getPreviousArchivedAvgAttendance_() {
  const rows = readArchiveRows_();
  return rows.length ? getArchiveRowAvgAttendance_(rows[rows.length - 1]) : null;
}

/* ============================ HTML EMAIL ============================ */

function reportPct_(v) {
  return (typeof v === 'number') ? (v * 100).toFixed(1) + '%' : (v == null ? '-' : v);
}

function reportTd_(content, extraStyle) {
  return '<td style="' + REPORT_CELL_STYLE_ + (extraStyle || '') + '">' + content + '</td>';
}

function reportTh_(label, extraStyle) {
  return '<th style="' + REPORT_TH_STYLE_ + (extraStyle || '') + '">' + label + '</th>';
}

function reportSummaryRow_(label, value, color) {
  return '<tr><td style="padding:4px 0;">' + label + '</td><td style="text-align:right;font-weight:bold;' +
    (color ? 'color:' + color + ';' : '') + '">' + value + '</td></tr>';
}

function reportRankTable_(rowsHtml) {
  return '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:16px;">' +
    '<tr>' + reportTh_('Nama') + reportTh_('Jabatan') + reportTh_('% Hadir', REPORT_RIGHT_) + '</tr>' +
    rowsHtml +
    '</table>';
}

/** notice (opsional): peringatan yang ditampilkan di atas ringkasan. */
function buildMonthlyReportHtml_(periodText, dashboard, analysis, notice) {
  const esc = escapeHtmlServer_;
  const empty3 = '<tr><td colspan="3" style="padding:6px;">Belum ada data.</td></tr>';
  const thresholdPct = CONFIG.REPORT_LOW_ATTENDANCE_THRESHOLD * 100;
  const noticeHtml = notice
    ? '<p style="background:#fef2f2;border-left:4px solid #ef4444;padding:10px 14px;font-size:13px;">⚠️ ' + esc(notice) + '</p>'
    : '';

  const rankRow = r => '<tr>' + reportTd_(esc(r.nama)) + reportTd_(esc(r.jabatan)) +
    reportTd_(reportPct_(r.persenKehadiran), REPORT_RIGHT_) + '</tr>';

  const rowsLow = analysis.lowAttendance.length
    ? analysis.lowAttendance.map(r => '<tr>' + reportTd_(esc(r.nama)) + reportTd_(esc(r.jabatan)) +
        reportTd_(reportPct_(r.persenKehadiran), REPORT_RIGHT_) + reportTd_(r.alfa, REPORT_RIGHT_) + '</tr>').join('')
    : '<tr><td colspan="4" style="padding:10px;text-align:center;color:#64748b;">Tidak ada karyawan dengan kehadiran di bawah ' +
      thresholdPct + '%. 🎉</td></tr>';

  return '' +
    '<div style="font-family:Arial,Helvetica,sans-serif;max-width:700px;margin:0 auto;color:#1e293b;">' +
      '<h2 style="color:#1e3a8a;margin-bottom:2px;">📊 Laporan Bulanan Kehadiran</h2>' +
      '<p style="color:#64748b;margin-top:0;">Departemen Housekeeping - PT Swadharma Primautama &mdash; Periode: <strong>' + esc(periodText) + '</strong></p>' +
      noticeHtml +

      '<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:16px;margin:16px 0;">' +
        '<table style="width:100%;border-collapse:collapse;font-size:14px;">' +
          reportSummaryRow_('Total Karyawan', dashboard.totalKaryawan ?? '-') +
          reportSummaryRow_('Rata-rata Kehadiran', reportPct_(analysis.avgPct), '#16a34a') +
          reportSummaryRow_('Total Alfa', analysis.totalAlfa, '#ef4444') +
          reportSummaryRow_('Total Izin', analysis.totalIzin, '#3b82f6') +
          reportSummaryRow_('Total Sakit', analysis.totalSakit, '#f59e0b') +
          reportSummaryRow_('Total Cuti', analysis.totalCuti, '#a855f7') +
        '</table>' +
      '</div>' +

      '<h3 style="color:#1e3a8a;font-size:15px;margin-bottom:6px;">📈 Analisis Tren</h3>' +
      '<p style="background:#fefce8;border-left:4px solid #eab308;padding:10px 14px;font-size:13px;margin-top:0;">' + esc(analysis.trendText) + '</p>' +

      '<h3 style="color:#16a34a;font-size:15px;margin-bottom:6px;">🏆 5 Karyawan Kehadiran Terbaik</h3>' +
      reportRankTable_(analysis.topBest.map(rankRow).join('') || empty3) +

      '<h3 style="color:#ef4444;font-size:15px;margin-bottom:6px;">🔻 5 Karyawan Kehadiran Terendah</h3>' +
      reportRankTable_(analysis.topWorst.map(rankRow).join('') || empty3) +

      '<h3 style="color:#d97706;font-size:15px;margin-bottom:6px;">⚠️ Perlu Perhatian (Kehadiran di bawah ' + thresholdPct + '%)</h3>' +
      '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:16px;">' +
        '<tr>' + reportTh_('Nama') + reportTh_('Jabatan') + reportTh_('% Hadir', REPORT_RIGHT_) + reportTh_('Alfa', REPORT_RIGHT_) + '</tr>' +
        rowsLow +
      '</table>' +

      '<p style="color:#94a3b8;font-size:11px;margin-top:24px;">Laporan ini dibuat otomatis oleh Attendance Monitoring System (AMS) pada ' +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+7', 'dd MMMM yyyy HH:mm') +
      '. Lampiran PDF berisi tampilan lengkap sheet Dashboard dan Rekap.</p>' +
    '</div>';
}
