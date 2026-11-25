const SPREADSHEET_ID = '1USjX1AKZgf-vWP1KY_GQ8izMmsg0MsV8T3fdDdpD2QM';
const SHEET_ID = 0;
const NOTIFICATION_RECIPIENTS = [
  'arsyad07112007@gmail.com',
  'gantengarsyad376@gmail.com'
];
const HEADERS = [
  'Waktu Pendaftaran',
  'Nama Lengkap',
  'NIM / NIS / No. Identitas',
  'Jenis Kelamin',
  'Nomor WhatsApp',
  'Email',
  'Program Studi / Divisi',
  'Status Keanggotaan',
  'Metode Pembayaran',
  'Bukti Transfer'
];
const MAX_PROOF_SIZE = 5 * 1024 * 1024;

function doPost(e) {
  const isFormSubmission = Boolean(e && e.parameter && e.parameter.payload);
  const rawPayload = isFormSubmission
    ? e.parameter.payload
    : e && e.postData && e.postData.contents;
  if (!rawPayload) {
    return registrationResponse({
      result: 'error',
      message: 'Data pendaftaran tidak ditemukan.'
    }, isFormSubmission);
  }

  let payload;
  try {
    payload = JSON.parse(rawPayload);
  } catch (error) {
    return registrationResponse({
      result: 'error',
      message: 'Format data pendaftaran tidak valid.'
    }, isFormSubmission);
  }

  const validation = validateRegistration(payload);
  if (validation) {
    return registrationResponse({
      result: 'error',
      message: validation,
      requestId: payload.requestId
    }, isFormSubmission);
  }

  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetById(SHEET_ID);
  if (!sheet) {
    throw new Error('Sheet dengan ID ' + SHEET_ID + ' tidak ditemukan.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  let proofFile = null;
  try {
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    } else {
      const currentHeaders = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
      let headersChanged = false;
      for (let i = 0; i < HEADERS.length; i++) {
        if (!currentHeaders[i]) {
          currentHeaders[i] = HEADERS[i];
          headersChanged = true;
        }
      }
      if (headersChanged) {
        sheet.getRange(1, 1, 1, HEADERS.length).setValues([currentHeaders]);
        sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      }
    }

    if (payload.metodePembayaran === 'Transfer') {
      proofFile = savePaymentProof(payload.buktiPembayaran);
    }

    sheet.appendRow([
      new Date(),
      safeCellValue(payload.nama),
      safeCellValue(payload.nim),
      safeCellValue(payload.jenisKelamin),
      safeCellValue(payload.whatsapp),
      safeCellValue(payload.email),
      safeCellValue(payload.programStudi),
      safeCellValue(payload.keanggotaan),
      safeCellValue(payload.metodePembayaran),
      proofFile ? safeCellValue(proofFile.getUrl()) : ''
    ]);
    SpreadsheetApp.flush();
  } catch (error) {
    if (proofFile) {
      proofFile.setTrashed(true);
    }
    Logger.log('Pendaftaran gagal disimpan: ' + error.message);
    return registrationResponse({
      result: 'error',
      message: 'Pendaftaran belum tersimpan. Periksa izin Google Drive pada Apps Script, lalu coba lagi.',
      requestId: payload.requestId
    }, isFormSubmission);
  } finally {
    lock.releaseLock();
  }

  let message = 'Pendaftaran berhasil! Data Anda telah tersimpan.';
  try {
    if (proofFile) {
      payload.buktiPembayaranUrl = proofFile.getUrl();
    }
    sendRegistrationNotification(payload);
  } catch (error) {
    Logger.log('Data pendaftaran tersimpan, tetapi email notifikasi gagal dikirim: ' + error.message);
    message += ' Namun, notifikasi email gagal dikirim.';
  }

  return registrationResponse({
    result: 'success',
    message: message,
    requestId: payload.requestId
  }, isFormSubmission);
}

function sendRegistrationNotification(payload) {
  const submittedAt = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'dd MMMM yyyy, HH:mm'
  );
  const subject = 'Pendaftar baru — CLASS LEADERSHIP & SOFT SKILL';
  const body = [
    'Pendaftaran baru telah masuk.',
    '',
    'Nama: ' + payload.nama.trim(),
    'Status: ' + payload.keanggotaan.trim(),
    'Program Studi / Divisi: ' + payload.programStudi.trim(),
    'Metode pembayaran: ' + payload.metodePembayaran.trim(),
    payload.metodePembayaran === 'Transfer' ? 'Bukti transfer: ' + payload.buktiPembayaranUrl : '',
    'Waktu pendaftaran: ' + submittedAt,
    '',
    'Data lengkap tersimpan di Google Spreadsheet pendaftaran.'
  ].filter(Boolean).join('\n');

  MailApp.sendEmail({
    to: NOTIFICATION_RECIPIENTS.join(','),
    subject: subject,
    body: body,
    name: 'Pendaftaran Class Leadership & Soft Skill'
  });
}

function testEmailNotification() {
  MailApp.sendEmail({
    to: NOTIFICATION_RECIPIENTS.join(','),
    subject: 'Tes notifikasi — CLASS LEADERSHIP & SOFT SKILL',
    body: 'Email tes berhasil dikirim dari Google Apps Script.',
    name: 'Pendaftaran Class Leadership & Soft Skill'
  });
  Logger.log('Email tes berhasil dikirim ke: ' + NOTIFICATION_RECIPIENTS.join(', '));
}

function authorizeRegistrationServices() {
  const authorizationFile = DriveApp.createFile(
    'Apps Script authorization check',
    'Temporary file to authorize payment proof uploads.'
  );
  authorizationFile.setTrashed(true);
  MailApp.getRemainingDailyQuota();
  Logger.log('Izin membuat file Google Drive dan menggunakan MailApp berhasil diberikan.');
}

function validateRegistration(payload) {
  if (!payload || typeof payload !== 'object' || Array.isArray(payload)) {
    return 'Data pendaftaran tidak valid.';
  }

  const requiredFields = [
    'nama',
    'nim',
    'jenisKelamin',
    'whatsapp',
    'email',
    'programStudi',
    'keanggotaan'
  ];

  for (let i = 0; i < requiredFields.length; i++) {
    const field = requiredFields[i];
    if (typeof payload[field] !== 'string' || !payload[field].trim()) {
      return 'Semua kolom wajib diisi.';
    }
  }

  if (payload.nama.trim().length < 3) {
    return 'Nama minimal 3 karakter.';
  }

  if (!['Laki-laki', 'Perempuan'].includes(payload.jenisKelamin.trim())) {
    return 'Pilihan jenis kelamin tidak valid.';
  }

  const whatsapp = payload.whatsapp.replace(/\s/g, '');
  if (!/^(\+?62|0)8[0-9]{7,13}$/.test(whatsapp)) {
    return 'Format nomor WhatsApp tidak valid.';
  }

  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(payload.email.trim())) {
    return 'Format email tidak valid.';
  }

  if (!['Pengurus', 'Anggota'].includes(payload.keanggotaan.trim())) {
    return 'Pilihan status keanggotaan tidak valid.';
  }

  if (!['Tunai', 'Transfer'].includes(payload.metodePembayaran)) {
    return 'Pilihan metode pembayaran tidak valid.';
  }

  if (payload.metodePembayaran === 'Transfer') {
    const proofError = validatePaymentProof(payload.buktiPembayaran);
    if (proofError) {
      return proofError;
    }
  }

  return '';
}

function validatePaymentProof(proof) {
  if (!proof || typeof proof !== 'object' || Array.isArray(proof)) {
    return 'Bukti transfer wajib diunggah untuk pembayaran transfer.';
  }
  if (typeof proof.name !== 'string' || !proof.name.trim() || proof.name.length > 150) {
    return 'Nama file bukti transfer tidak valid.';
  }
  if (typeof proof.dataUrl !== 'string') {
    return 'Data bukti transfer tidak valid.';
  }
  if (proof.dataUrl.length > 4 * Math.ceil(MAX_PROOF_SIZE / 3) + 32) {
    return 'Ukuran bukti transfer maksimal 5 MB.';
  }

  const match = proof.dataUrl.match(/^data:(application\/pdf|image\/jpeg|image\/png);base64,([A-Za-z0-9+/]+={0,2})$/);
  if (!match) {
    return 'Format bukti transfer harus JPG, PNG, atau PDF.';
  }

  const extension = proof.name.trim().split('.').pop().toLowerCase();
  const mimeExtensions = {
    'application/pdf': ['pdf'],
    'image/jpeg': ['jpg', 'jpeg'],
    'image/png': ['png']
  };
  if (!mimeExtensions[match[1]].includes(extension)) {
    return 'Ekstensi file tidak sesuai dengan format bukti transfer.';
  }

  const bytes = Utilities.base64Decode(match[2]);
  if (bytes.length === 0 || bytes.length > MAX_PROOF_SIZE) {
    return 'Ukuran bukti transfer harus lebih dari 0 dan maksimal 5 MB.';
  }

  if (!matchesFileSignature(bytes, match[1])) {
    return 'Isi file tidak sesuai dengan format bukti transfer.';
  }

  return '';
}

function matchesFileSignature(bytes, mimeType) {
  if (mimeType === 'application/pdf') {
    return bytes.length >= 5 &&
      bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 &&
      bytes[3] === 0x46 && bytes[4] === 0x2D;
  }
  if (mimeType === 'image/png') {
    return bytes.length >= 8 &&
      (bytes[0] & 0xFF) === 0x89 && bytes[1] === 0x50 &&
      bytes[2] === 0x4E && bytes[3] === 0x47 &&
      bytes[4] === 0x0D && bytes[5] === 0x0A &&
      bytes[6] === 0x1A && bytes[7] === 0x0A;
  }
  return mimeType === 'image/jpeg' &&
    bytes.length >= 3 &&
    (bytes[0] & 0xFF) === 0xFF &&
    (bytes[1] & 0xFF) === 0xD8 &&
    (bytes[2] & 0xFF) === 0xFF;
}

function savePaymentProof(proof) {
  const match = proof.dataUrl.match(/^data:(application\/pdf|image\/jpeg|image\/png);base64,([A-Za-z0-9+/]+={0,2})$/);
  const bytes = Utilities.base64Decode(match[2]);
  const safeName = proof.name.trim().replace(/[\\/:*?"<>|]/g, '_');
  const fileName = 'Bukti-transfer-' + new Date().getTime() + '-' + safeName;
  const blob = Utilities.newBlob(bytes, match[1], fileName);
  return DriveApp.createFile(blob);
}

function safeCellValue(value) {
  const text = value.trim();
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function registrationResponse(data, isFormSubmission) {
  if (!isFormSubmission) {
    return jsonResponse(data);
  }

  const safeData = JSON.stringify(data).replace(/</g, '\\u003c');
  const html = '<!doctype html><html><body><script>' +
    '(function sendRegistrationResult() {' +
    'var result = ' + safeData + ';' +
    'var target = window; var attempts = 0;' +
    'while (target !== target.parent && attempts < 10) {' +
    'target = target.parent; attempts++;' +
    'target.postMessage(result, "*");' +
    '}' +
    'window.postMessage(result, "*");' +
    '})();' +
    '</script></body></html>';
  return HtmlService
    .createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
