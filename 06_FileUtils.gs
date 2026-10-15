/*************************************************************
 * 06_FileUtils.gs
 * Helper Drive & Excel yang dipakai upload Jadwal/Realisasi,
 * upload finger, dan pembuatan template.
 *
 * Alur upload: file .xls/.xlsx (base64) -> diunggah ke Drive dan
 * dikonversi jadi Google Sheet sementara -> dibaca -> file sementara
 * selalu dibuang (withTempSpreadsheet_).
 *************************************************************/

const MIME_XLSX_ = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const MIME_XLS_ = 'application/vnd.ms-excel';
const MIME_PDF_ = 'application/pdf';
const MIME_GSHEET_ = 'application/vnd.google-apps.spreadsheet';
const DRIVE_UPLOAD_URL_ = 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id';

/** Buang file Drive; kegagalan diabaikan (file sementara tidak boleh menggagalkan proses utama). */
function trashFileQuietly_(fileId) {
  if (!fileId) return;
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) { /* abaikan */ }
}

/**
 * Unggah base64 ke Drive sebagai Google Sheet sementara.
 * @param {string} tempNamePrefix awalan nama file sementara
 * @param {string} sourceMime MIME_XLSX_ atau MIME_XLS_
 * @param {string} errorMessage awal pesan error jika gagal dibaca
 * @return {string} id file sementara
 */
function uploadBase64AsTempSpreadsheet_(base64Data, filename, tempNamePrefix, sourceMime, errorMessage) {
  const boundary = 'sheetupload' + Utilities.getUuid().replace(/-/g, '');
  const metadata = {
    name: tempNamePrefix + filename + '_' + Date.now(),
    mimeType: MIME_GSHEET_
  };

  const body =
    '--' + boundary + '\r\n' +
    'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
    JSON.stringify(metadata) + '\r\n' +
    '--' + boundary + '\r\n' +
    'Content-Type: ' + sourceMime + '\r\n' +
    'Content-Transfer-Encoding: base64\r\n\r\n' +
    base64Data + '\r\n' +
    '--' + boundary + '--';

  const response = UrlFetchApp.fetch(DRIVE_UPLOAD_URL_, {
    method: 'post',
    contentType: 'multipart/related; boundary="' + boundary + '"',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: body,
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const json = JSON.parse(response.getContentText() || '{}');
  if (code !== 200 || !json.id) {
    throw new Error(errorMessage + ' (Detail: ' + response.getContentText() + ')');
  }
  return json.id;
}

/** Jalankan fn(fileId) pada Google Sheet sementara hasil konversi; file selalu dibuang sesudahnya. */
function withTempSpreadsheet_(base64Data, filename, tempNamePrefix, sourceMime, errorMessage, fn) {
  let tempFileId = null;
  try {
    tempFileId = uploadBase64AsTempSpreadsheet_(base64Data, filename, tempNamePrefix, sourceMime, errorMessage);
    return fn(tempFileId);
  } finally {
    trashFileQuietly_(tempFileId);
  }
}

/** Ekspor Google Sheet ke .xlsx. @return {Blob} */
function exportSpreadsheetAsXlsxBlob_(fileId, errorLabel) {
  const response = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=xlsx', {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) {
    throw new Error('Gagal membuat ' + errorLabel + ' (kode ' + response.getResponseCode() + ').');
  }
  return response.getBlob();
}

/** Bentuk respons unduhan yang dipakai frontend. */
function xlsxDownloadResult_(filename, blob) {
  return {
    filename: filename,
    mimeType: MIME_XLSX_,
    base64: Utilities.base64Encode(blob.getBytes())
  };
}

/**
 * Salin satu sheet ke spreadsheet sementara, ubah salinannya (prepare), lalu hasilkan
 * sesuatu darinya (produce). Spreadsheet sementara SELALU dibuang, termasuk saat error.
 * @param {function(Sheet)} prepare ubah salinan (mis. filter baris) sebelum diekspor
 * @param {function(Sheet, string)} produce menerima (salinan, idFileSementara)
 */
function withTempSheetCopy_(srcSheet, tempNamePrefix, copyName, prepare, produce) {
  const tempSs = SpreadsheetApp.create(tempNamePrefix + Date.now());
  const fileId = tempSs.getId();
  try {
    const defaultSheet = tempSs.getSheets()[0];
    const copied = srcSheet.copyTo(tempSs);
    copied.setName(copyName);
    prepare(copied);
    tempSs.deleteSheet(defaultSheet);
    SpreadsheetApp.flush();
    return produce(copied, fileId);
  } finally {
    trashFileQuietly_(fileId);
  }
}
