const SPREADSHEET_ID = '1SC1HPgM7CyqOXVDX9oFrOWTrsbm5PVaLjfPx_2HSvAQ';
const SHEET_NAME = 'RSVP';
const COMMENTS_SHEET_NAME = 'COMMENTS';

function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action === 'comments') {
      return jsonResponse_({ ok: true, comments: getComments_() });
    }

    return jsonResponse_({ ok: true, message: 'Web App aktif.' });
  } catch (error) {
    console.error(error);
    return jsonResponse_({ ok: false, message: 'Komentar gagal dimuat.' });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse_({ ok: false, message: 'Data RSVP kosong.' });
    }

    const data = JSON.parse(e.postData.contents);
    if (data.action === 'comment') {
      return saveComment_(data);
    }

    const nama = String(data.nama || '').trim();
    const kehadiran = String(data.kehadiran || '').trim();
    const jumlahTamu = Number(data.jumlahTamu || 1);
    const ucapan = String(data.ucapan || '').trim();

    if (!nama || !kehadiran) {
      return jsonResponse_({ ok: false, message: 'Nama dan kehadiran wajib diisi.' });
    }

    if (!Number.isFinite(jumlahTamu) || jumlahTamu < 1 || jumlahTamu > 5) {
      return jsonResponse_({ ok: false, message: 'Jumlah tamu harus antara 1 dan 5.' });
    }

    const sheet = getSheet_();
    sheet.appendRow([
      new Date(),
      nama,
      kehadiran,
      jumlahTamu,
      ucapan,
      String(data.halaman || ''),
      String(data.waktu || '')
    ]);

    return jsonResponse_({ ok: true, message: 'RSVP berhasil disimpan.' });
  } catch (error) {
    console.error(error);
    return jsonResponse_({ ok: false, message: 'Server gagal menyimpan RSVP.' });
  }
}

function saveComment_(data) {
  const nama = String(data.nama || '').trim();
  const ucapan = String(data.ucapan || '').trim();

  if (!nama || !ucapan) {
    return jsonResponse_({ ok: false, message: 'Nama dan ucapan wajib diisi.' });
  }

  if (nama.length > 80 || ucapan.length > 500) {
    return jsonResponse_({ ok: false, message: 'Nama atau ucapan terlalu panjang.' });
  }

  const sheet = getCommentsSheet_();
  sheet.appendRow([
    new Date(),
    nama,
    ucapan,
    String(data.halaman || '')
  ]);

  return jsonResponse_({ ok: true, message: 'Ucapan berhasil disimpan.' });
}

function getComments_() {
  const sheet = getCommentsSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  return sheet.getRange(2, 1, lastRow - 1, 4).getValues()
    .reverse()
    .slice(0, 100)
    .map(row => ({
      waktu: row[0],
      nama: String(row[1] || ''),
      ucapan: String(row[2] || ''),
      halaman: String(row[3] || '')
    }));
}

function setupSheet() {
  const sheet = getSheet_();
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Waktu Server',
      'Nama',
      'Kehadiran',
      'Jumlah Tamu',
      'Ucapan',
      'Halaman',
      'Waktu Klien'
    ]);
    sheet.setFrozenRows(1);
  }

  getCommentsSheet_();
}

function getCommentsSheet_() {
  const spreadsheet = getSpreadsheet_();
  let sheet = spreadsheet.getSheetByName(COMMENTS_SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(COMMENTS_SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Waktu', 'Nama', 'Ucapan', 'Halaman']);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function getSheet_() {
  const spreadsheet = getSpreadsheet_();
  return spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
}

function getSpreadsheet_() {
  if (SPREADSHEET_ID.indexOf('GANTI_DENGAN_') === 0) {
    throw new Error('SPREADSHEET_ID belum dikonfigurasi.');
  }

  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
