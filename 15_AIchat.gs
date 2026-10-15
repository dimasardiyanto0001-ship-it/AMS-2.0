/*****************************************************************
 * 15_AIchat.gs
 * Chatbot AI (Google Gemini) untuk tanya-jawab data dan kirim laporan.
 *
 * Alur: chatWithAI() -> AI hanya menerjemahkan maksud (tidak menjalankan aksi).
 * Bila user minta data di-email, server mengembalikan permintaan konfirmasi;
 * pengiriman baru terjadi di confirmSendEmailReport() setelah user klik "Ya",
 * dan hanya untuk Super Admin.
 *
 * Script Properties: GEMINI_API_KEY (wajib), AI_MODEL (opsional).
 *****************************************************************/

const AI_MODEL_DEFAULT_ = 'gemini-3.5-flash';
const AI_MAX_MESSAGE_LENGTH_ = 1000;
const AI_EMAIL_RE_ = /^[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+$/;

function getAiModel_() {
  return getSetting_('AI_MODEL', AI_MODEL_DEFAULT_);
}

/** Buang pagar kode ```json ... ``` bila model membungkus JSON-nya. */
function stripCodeFence_(text) {
  const m = String(text).trim().match(/^```(?:json)?\s*([\s\S]*?)\s*```$/i);
  return m ? m[1] : String(text).trim();
}

/** Panggil Gemini API. systemInstruction = peran AI, promptText = pesan + konteks data. @return {Object} JSON balasan */
function callGeminiAI_(promptText, systemInstruction) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY belum diset di Script Properties.');

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + getAiModel_() + ':generateContent';
  const payload = {
    system_instruction: { parts: [{ text: systemInstruction }] },
    contents: [{ role: 'user', parts: [{ text: promptText }] }],
    generationConfig: { responseMimeType: 'application/json' }
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    throw new Error('Gagal memanggil AI (' + res.getResponseCode() + '): ' + res.getContentText().substring(0, 300));
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

/** Ringkasan data periode aktif yang dikirim ke AI sebagai konteks (bukan seluruh raw sheet). */
function buildAIDataContext_() {
  const data = getAllDataInternal_();
  return {
    period: data.period,
    dayCount: data.dayCount,
    dashboard: data.dashboard,
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
    'Kalau user tidak sebutkan alamat email, isi tujuanEmail dengan null.';

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
