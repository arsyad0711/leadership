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
  'Status Keanggotaan'
];

function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return jsonResponse({
      result: 'error',
      message: 'Data pendaftaran tidak ditemukan.'
    });
  }

  let payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (error) {
    return jsonResponse({
      result: 'error',
      message: 'Format data pendaftaran tidak valid.'
    });
  }

  const validation = validateRegistration(payload);
  if (validation) {
    return jsonResponse({
      result: 'error',
      message: validation
    });
  }

  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetById(SHEET_ID);
  if (!sheet) {
    throw new Error('Sheet dengan ID ' + SHEET_ID + ' tidak ditemukan.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      new Date(),
      safeCellValue(payload.nama),
      safeCellValue(payload.nim),
      safeCellValue(payload.jenisKelamin),
      safeCellValue(payload.whatsapp),
      safeCellValue(payload.email),
      safeCellValue(payload.programStudi),
      safeCellValue(payload.keanggotaan)
    ]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  let message = 'Pendaftaran berhasil! Data Anda telah tersimpan.';
  try {
    sendRegistrationNotification(payload);
  } catch (error) {
    Logger.log('Data pendaftaran tersimpan, tetapi email notifikasi gagal dikirim: ' + error.message);
    message += ' Namun, notifikasi email gagal dikirim.';
  }

  return jsonResponse({
    result: 'success',
    message: message
  });
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
    'Waktu pendaftaran: ' + submittedAt,
    '',
    'Data lengkap tersimpan di Google Spreadsheet pendaftaran.'
  ].join('\n');

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

  return '';
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
