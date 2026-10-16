/*****************************************************************
 * 15_AIchat.gs
 * Chatbot AI (Google Gemini) untuk tanya-jawab data dan kirim laporan.
 *
 * Alur: chatWithAI() -> AI hanya menerjemahkan maksud (tidak menjalankan aksi).
 * Bila user minta data di-email, server mengembalikan permintaan konfirmasi;
 * pengiriman baru terjadi di confirmSendEmailReport() setelah user klik "Ya",
 * dan hanya untuk Super Admin.
 *
 * Script Properties:
 *   GEMINI_API_KEY      (wajib)
 *   AI_MODEL            (opsional) model utama
 *   AI_FALLBACK_MODELS  (opsional) model cadangan, pisah koma
 *****************************************************************/

const AI_MODEL_DEFAULT_ = 'gemini-3.5-flash';
const AI_FALLBACK_MODELS_DEFAULT_ = ['gemini-2.5-flash', 'gemini-2.5-flash-lite'];
const AI_RETRYABLE_CODES_ = [429, 500, 503, 504]; // error sementara -> retry
const AI_MAX_ATTEMPTS_PER_MODEL_ = 2;
const AI_MAX_MESSAGE_LENGTH_ = 1000;
const AI_EMAIL_RE_ = /^[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+$/;

function getAiModel_() {
  return getSetting_('AI_MODEL', AI_MODEL_DEFAULT_);
}

/** Urutan model: utama dulu, lalu cadangan (Script Property AI_FALLBACK_MODELS, pisah koma). */
function getAiModelChain_() {
  const extra = String(getSetting_('AI_FALLBACK_MODELS', ''))
    .split(',').map(function (s) { return s.trim(); }).filter(Boolean);
  const fallbacks = extra.length ? extra : AI_FALLBACK_MODELS_DEFAULT_;
  return [getAiModel_()].concat(fallbacks)
    .filter(function (m, i, arr) { return arr.indexOf(m) === i; });
}

/** Buang pagar kode ```json ... ``` bila model membungkus JSON-nya. */
function stripCodeFence_(text) {
  const m = String(text).trim().match(/^```(?:json)?\s*([\s\S]*?)\s*```$/i);
  return m ? m[1] : String(text).trim();
}

/** Panggil Gemini API dengan retry + fallback model. systemInstruction = peran AI, promptText = pesan + konteks data. @return {Object} JSON balasan */
function callGeminiAI_(promptText, systemInstruction) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY belum diset di Script Properties.');

  const payload = {
    system_instruction: { parts: [{ text: systemInstruction }] },
    contents: [{ role: 'user', parts: [{ text: promptText }] }],
    generationConfig: { responseMimeType: 'application/json' }
  };
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const models = getAiModelChain_();
  let lastCode = 0;
  let res = null;

  outer:
  for (let m = 0; m < models.length; m++) {
    const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + models[m] + ':generateContent';

    for (let attempt = 0; attempt < AI_MAX_ATTEMPTS_PER_MODEL_; attempt++) {
      res = UrlFetchApp.fetch(url, options);
      lastCode = res.getResponseCode();

      if (lastCode === 200) break outer;

      Logger.log('Gemini ' + models[m] + ' -> ' + lastCode + ': ' + res.getContentText().substring(0, 300));

      if (lastCode === 404) break; // nama model salah/tidak ada -> langsung model berikutnya
      if (AI_RETRYABLE_CODES_.indexOf(lastCode) === -1) { // 400/401/403 dll: bukan masalah sementara
        throw new Error('Gagal memanggil AI (' + lastCode + '). Periksa GEMINI_API_KEY / pengaturan AI.');
      }
      if (attempt < AI_MAX_ATTEMPTS_PER_MODEL_ - 1) Utilities.sleep(1000 * Math.pow(2, attempt)); // 1s, 2s...
    }
  }

  if (!res || lastCode !== 200) {
    throw new Error('Asisten AI sedang sibuk. Coba lagi dalam beberapa saat ya.');
  }

  const json = JSON.parse(res.getContentText());
  const part = json.candidates && json.candidates[0] && json.candidates[0].content &&
    json.candidates[0].content.parts && json.candidates[0].content.parts[0];
  const text = part && part.text;
  if (!text) throw new Error('Respons AI kosong atau tidak terduga.');

  try {
    return JSON.parse(stripCodeFence_(text));
  } catch (e) {
    throw new Error('Balasan AI bukan JSON yang valid. Coba ulangi pertanyaan dengan kalimat lain.');
  }
}

/** Kode Realisasi -> kunci status harian (kode '1' = hadir, '2' = hadir + lembur ditangani terpisah). */
const AI_STATUS_KEY_ = { 'A': 'alfa', 'I': 'izin', 'S': 'sakit', 'C': 'cuti', 'OFF': 'off' };

/**
 * Data harian dari Realisasi (+ Jadwal untuk mendeteksi sel yang belum diisi):
 *  - harian: per tanggal -> daftar nama per status
 *  - perKaryawan: per orang -> daftar tanggal alfa/izin/sakit/cuti
 * Status kosong dibuang supaya konteks ke AI tetap ringkas.
 */
function buildAIDailyContext_(data) {
  const dayCount = Number(data.dayCount) || 0;
  const dayNames = data.dayNames || [];

  const jadwalByName = {};
  (data.jadwal || []).forEach(e => { jadwalByName[nameKey_(e.nama)] = e.days || []; });

  const harian = [];
  for (let i = 0; i < dayCount; i++) {
    harian.push({
      tanggal: i + 1, hari: dayNames[i] || '',
      hadir: [], lembur: [], alfa: [], izin: [], sakit: [], cuti: [], off: [], belumDiisi: [], lainnya: []
    });
  }

  const perKaryawan = [];
  (data.realisasi || []).forEach(emp => {
    const nama = String(emp.nama || '').trim();
    if (!nama) return;
    const days = emp.days || [];
    const jDays = jadwalByName[nameKey_(nama)] || [];
    const tdk = { nama: nama, alfa: [], izin: [], sakit: [], cuti: [] };

    for (let i = 0; i < dayCount; i++) {
      const code = cellCode_(days[i]);
      const d = harian[i];
      if (code === '1') {
        d.hadir.push(nama);
      } else if (code === '2') {
        d.hadir.push(nama);
        d.lembur.push(nama);
      } else if (AI_STATUS_KEY_[code]) {
        const key = AI_STATUS_KEY_[code];
        d[key].push(nama);
        if (tdk[key]) tdk[key].push(i + 1);
      } else if (code === '') {
        const jc = cellCode_(jDays[i]);
        if (jc && jc !== 'OFF') d.belumDiisi.push(nama); // dijadwalkan masuk, realisasi kosong
      } else {
        d.lainnya.push(nama + ' (' + code + ')');
      }
    }
    if (tdk.alfa.length || tdk.izin.length || tdk.sakit.length || tdk.cuti.length) perKaryawan.push(tdk);
  });

  harian.forEach(d => {
    Object.keys(d).forEach(k => { if (Array.isArray(d[k]) && d[k].length === 0) delete d[k]; });
  });
  perKaryawan.forEach(p => {
    ['alfa', 'izin', 'sakit', 'cuti'].forEach(k => { if (p[k].length === 0) delete p[k]; });
  });

  return { harian: harian, perKaryawan: perKaryawan };
}

/** Ringkasan data periode aktif yang dikirim ke AI sebagai konteks (bukan seluruh raw sheet). */
function buildAIDataContext_() {
  const data = getAllDataInternal_();
  const daily = buildAIDailyContext_(data);
  return {
    period: data.period,
    dayCount: data.dayCount,
    tanggalHariIni: Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd'),
    dashboard: data.dashboard,
    harian: daily.harian,
    perKaryawan: daily.perKaryawan,
    rekap: (data.rekap || []).map(r => ({
      nama: r.nama, jabatan: r.jabatan, hadir: r.hadir, alfa: r.alfa,
      izin: r.izin, sakit: r.sakit, cuti: r.cuti, off: r.off,
      persenKehadiran: r.persenKehadiran, totalJamKerja: r.totalJamKerja,
      kekuranganJam: r.kekuranganJam
    }))
  };
}

/** Bulatkan & jepit nomor tanggal ke 1..dayCount; nilai tak terbaca -> fallback. */
function clampDayNumber_(value, fallback, dayCount) {
  const n = parseInt(value, 10);
  const base = isNaN(n) ? fallback : n;
  return Math.min(Math.max(base, 1), dayCount > 0 ? dayCount : 31);
}

/** Alamat email tunggal yang sah, atau null (juga untuk "null"/kosong dari AI). */
function cleanEmailAddress_(value) {
  const s = String(value === null || value === undefined ? '' : value).trim();
  return AI_EMAIL_RE_.test(s) ? s : null;
}

/** Titik masuk chat dari frontend. AI hanya menerjemahkan maksud, TIDAK mengeksekusi aksi apa pun. */
function chatWithAI(pesan, token) {
  const session = verifySession_(token);

  const message = String(pesan === null || pesan === undefined ? '' : pesan).trim();
  if (!message) throw new Error('Pesan tidak boleh kosong.');
  if (message.length > AI_MAX_MESSAGE_LENGTH_) {
    throw new Error('Pesan terlalu panjang (maks ' + AI_MAX_MESSAGE_LENGTH_ + ' karakter).');
  }

  const context = buildAIDataContext_();

  const systemInstruction =
    'Anda asisten AI untuk sistem absensi karyawan. JAWAB HANYA berdasarkan data konteks JSON yang diberikan, ' +
    'jangan mengarang angka yang tidak ada di data. Balas SELALU dalam JSON valid dengan struktur persis: ' +
    '{"action":"jawab","teks":"..."} ATAU ' +
    '{"action":"kirim_email","teks":"ringkasan singkat permintaan","tglMulai":angka,"tglAkhir":angka,"tujuanEmail":"string atau null"}. ' +
    'Gunakan action "kirim_email" HANYA jika user secara eksplisit minta data dikirim/di-email. ' +
    'Kalau user tidak sebutkan tanggal, tglMulai=1 dan tglAkhir=dayCount (jumlah hari periode, ada di data konteks). ' +
    'Kalau user tidak sebutkan alamat email, isi tujuanEmail dengan null. ' +
    'DATA HARIAN: "harian" berisi per tanggal daftar nama untuk status hadir, lembur, alfa, izin, sakit, cuti, off, belumDiisi, lainnya ' +
    '(status yang tidak ada orangnya tidak ditulis). "perKaryawan" berisi daftar tanggal alfa/izin/sakit/cuti per orang. ' +
    '"Tidak masuk" berarti alfa + izin + sakit + cuti; "off" adalah libur terjadwal, sebutkan terpisah dan jangan dihitung tidak masuk. ' +
    '"belumDiisi" = dijadwalkan masuk tetapi realisasi masih kosong; untuk tanggal setelah tanggalHariIni artinya belum terjadi, bukan tidak masuk. ' +
    'Untuk pertanyaan per tanggal atau per orang, sebutkan SEMUA nama/tanggal yang relevan beserta jumlahnya dan jangan dipotong. ' +
    'Jika tanggal yang ditanya tidak ada di data periode ini, katakan terus terang. ' +
    'Isi "teks" dengan teks biasa (tanpa markdown); untuk daftar gunakan baris baru dengan penomoran "1. Nama".';

  const prompt = 'Data konteks (JSON):\n' + JSON.stringify(context) + '\n\nPermintaan user: ' + message;

  const result = callGeminiAI_(prompt, systemInstruction) || {};

  if (result.action === 'kirim_email') {
    const dayCount = Number(context.dayCount) || 0;
    const tglMulai = clampDayNumber_(result.tglMulai, 1, dayCount);
    const tglAkhir = clampDayNumber_(result.tglAkhir, dayCount || 31, dayCount);
    if (tglMulai > tglAkhir) {
      return { action: 'jawab', teks: 'Rentang tanggal tidak valid (tanggal awal ' + tglMulai + ' melebihi tanggal akhir ' + tglAkhir + '). Mohon sebutkan ulang rentang tanggalnya.' };
    }

    const tujuan = cleanEmailAddress_(result.tujuanEmail) || cleanEmailAddress_(session.email);
    if (!tujuan) {
      return { action: 'perlu_email', teks: 'Data periode tanggal ' + tglMulai + '-' + tglAkhir + ' mau dikirim ke alamat email mana?' };
    }
    return {
      action: 'konfirmasi_kirim_email',
      teks: 'Akan kirim data absensi periode "' + context.period + '" tanggal ' + tglMulai + '-' + tglAkhir + ' ke ' + tujuan + '. Kirim sekarang?',
      tglMulai: tglMulai,
      tglAkhir: tglAkhir,
      tujuanEmail: tujuan
    };
  }

  const teks = result.teks === null || result.teks === undefined ? '' : String(result.teks).trim();
  return { action: 'jawab', teks: teks || 'Maaf, saya tidak bisa memberi jawaban untuk pertanyaan itu.' };
}

/**
 * Dipanggil HANYA setelah user klik "Ya, kirim" di popup chat. Super Admin saja
 * (akun biasa tidak boleh memicu pengiriman seluruh data absensi ke alamat bebas).
 * Pengiriman dicatat di Audit Log.
 */
function confirmSendEmailReport(tglMulai, tglAkhir, tujuanEmail, token) {
  const session = requireSuperAdmin_(token, 'Hanya akun dengan role "Super Admin" yang boleh mengirim laporan lewat email.');

  const tujuan = cleanEmailAddress_(tujuanEmail);
  if (!tujuan) throw new Error('Alamat email tujuan tidak valid. Gunakan satu alamat email saja.');

  const excel = generateAttendanceLogExcel('active', 'ALL', tglMulai, tglAkhir, token);
  const blob = Utilities.newBlob(Utilities.base64Decode(excel.base64), excel.mimeType, excel.filename);
  MailApp.sendEmail({
    to: tujuan,
    subject: 'Laporan Absensi - ' + excel.filename,
    body: 'Halo,\n\nTerlampir laporan absensi yang Anda minta lewat chatbot AMS.\n\nSalam,\nAMS Bot',
    attachments: [blob]
  });

  catatAuditLog_(session, 'Kirim Laporan Email', 'Chatbot AI', tujuan, excel.filename);
  return { success: true, tujuanEmail: tujuan };
}
