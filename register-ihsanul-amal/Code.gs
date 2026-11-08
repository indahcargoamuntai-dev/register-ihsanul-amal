var HEADERS = [
  'waktuDaftar',
  'noRegistrasi',
  'jenjang',
  'programTambahan',
  'statusAsalSekolah',
  'namaAsalSekolah',
  'npsnAsal',
  'namaLengkap',
  'nisn',
  'nik',
  'nomorKK',
  'tempatLahir',
  'tanggalLahir',
  'jenisKelamin',
  'jalanRtRw',
  'provinsi',
  'kabupatenKota',
  'kecamatan',
  'desaKelurahan',
  'nikAyah-namaAyah',
  'PendidikanPekerjaanAyah',
  'nikIbu-namaIbu',
  'PendidikanPekerjaanIbu',
  'noHp',
  'penghasilanOrtu'
];

function doGet(e) {
  var params = e && e.parameter ? e.parameter : {};
  var response = checkRegistrationStatus_(params);
  return sendResponse_(response, params);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  var stage = 'memeriksa data formulir';
  var data = {};

  try {
    data = parseRequestData_(e);
    if (data.action === 'checkStatus') {
      return sendResponse_(checkRegistrationStatus_(data), data);
    }
    validateRegistration_(data);
    stage = 'mengunci proses pendaftaran';
    if (!lock.tryLock(10000)) {
      throw new Error('Server sedang memproses pendaftaran lain. Tunggu sebentar lalu coba lagi.');
    }

    stage = 'memeriksa pengiriman ulang';
    var cache = CacheService.getScriptCache();
    var cachedResponse = null;
    try {
      cachedResponse = cache.get('ppdb:' + data.requestId);
    } catch (_) {}
    if (cachedResponse) {
      var previousResponse = JSON.parse(cachedResponse);
      previousResponse.requestId = previousResponse.requestId || data.requestId;
      previousResponse.nama = previousResponse.nama || data.namaLengkap;
      previousResponse.jenjang = previousResponse.jenjang || data.jenjang;
      return sendResponse_(previousResponse, data);
    }

    stage = 'membuka spreadsheet';
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      throw new Error('Spreadsheet aktif tidak ditemukan. Pastikan Apps Script terikat ke Google Sheets.');
    }

    stage = 'memeriksa tab dan header spreadsheet';
    var sheet = spreadsheet.getActiveSheet();
    ensureHeaders_(sheet);

    var noRegistrasi = generateRegistrationNumber_(sheet, data.jenjang);
    var response = {
      result: 'success',
      requestId: data.requestId,
      nama: data.namaLengkap,
      jenjang: data.jenjang,
      noRegistrasi: noRegistrasi
    };

    stage = 'menyimpan data pendaftaran ke spreadsheet';
    sheet.appendRow([
      new Date(),
      noRegistrasi,
      data.jenjang,
      data.programTambahan || '',
      data.statusAsalSekolah,
      data.namaAsalSekolah,
      data.npsnAsal || '',
      data.namaLengkap,
      data.nisn,
      data.nik,
      data.nomorKK || '',
      data.tempatLahir,
      data.tanggalLahir,
      data.jenisKelamin,
      data.jalanRtRw,
      data.provinsi,
      data.kabupatenKota,
      data.kecamatan,
      data.desaKelurahan,
      combine_(data.nikAyah, data.namaAyah),
      combine_(data.pendidikanAyah, data.pekerjaanAyah),
      combine_(data.nikIbu, data.namaIbu),
      combine_(data.pendidikanIbu, data.pekerjaanIbu),
      data.noHp,
      data.penghasilanOrtu
    ].map(safeCellValue_));
    SpreadsheetApp.flush();

    stage = 'menyimpan tanda pengiriman ulang';
    try {
      cache.put('ppdb:' + data.requestId, JSON.stringify(response), 21600);
    } catch (_) {}
    return sendResponse_(response, data);
  } catch (error) {
    return sendResponse_({
      result: 'error',
      requestId: data && data.requestId ? data.requestId : '',
      code: error && error.code ? error.code : 'REGISTRATION_FAILED',
      error: 'Gagal ' + stage + ': ' + (error && error.message ? error.message : String(error))
    }, data);
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

function parseRequestData_(e) {
  var data = {};
  if (e && e.parameter) {
    Object.keys(e.parameter).forEach(function (key) {
      data[key] = e.parameter[key];
    });
  }
  if (e && e.postData && e.postData.contents) {
    var contentType = String(e.postData.type || '').toLowerCase();
    var raw = e.postData.contents;
    if (contentType.indexOf('application/json') !== -1 || /^\s*\{/.test(raw)) {
      try {
        var parsed = JSON.parse(raw);
        if (parsed && typeof parsed === 'object') {
          Object.keys(parsed).forEach(function (key) {
            data[key] = parsed[key];
          });
          if (!data.responseMode) {
            data.responseMode = 'json';
          }
        }
      } catch (_) {}
    }
  }
  return data;
}

function validateRegistration_(data) {
  if (!data.requestId || !/^[a-zA-Z0-9-]{8,100}$/.test(data.requestId)) {
    throw registrationError_('INVALID_REQUEST_ID', 'ID permintaan tidak valid. Muat ulang halaman lalu coba lagi.');
  }

  var requiredFields = [
    'jenjang', 'statusAsalSekolah', 'namaAsalSekolah', 'namaLengkap',
    'nisn', 'nik', 'tempatLahir', 'tanggalLahir', 'jenisKelamin',
    'jalanRtRw', 'provinsi', 'kabupatenKota', 'kecamatan', 'desaKelurahan',
    'namaAyah', 'pendidikanAyah', 'pekerjaanAyah', 'namaIbu',
    'pendidikanIbu', 'pekerjaanIbu', 'noHp', 'penghasilanOrtu'
  ];

  requiredFields.forEach(function (field) {
    if (!data[field] || !String(data[field]).trim()) {
      throw registrationError_('MISSING_FIELD', 'Field wajib belum diisi: ' + field);
    }
  });

  if (['KB', 'TK', 'SD', 'SMP'].indexOf(data.jenjang) === -1) {
    throw registrationError_('INVALID_SCHOOL_LEVEL', 'Jenjang sekolah tidak valid. Pilih KB, TK, SD, atau SMP.');
  }

  var allowedPrograms = {
    KB: ['', 'TPA'],
    TK: ['', 'TPA'],
    SD: [''],
    SMP: ['', 'Boarding']
  };
  if (allowedPrograms[data.jenjang].indexOf(data.programTambahan || '') === -1) {
    throw registrationError_('INVALID_PROGRAM', 'Program tambahan tidak sesuai dengan jenjang sekolah yang dipilih.');
  }

  if (!/^[0-9]{10}$/.test(data.nisn)) {
    throw registrationError_('INVALID_NISN', 'NISN harus terdiri dari 10 digit angka.');
  }
  if (!/^[0-9]{16}$/.test(data.nik)) {
    throw registrationError_('INVALID_NIK', 'NIK harus terdiri dari 16 digit angka.');
  }

  ['nomorKK', 'nikAyah', 'nikIbu'].forEach(function (field) {
    if (data[field] && !/^[0-9]{16}$/.test(data[field])) {
      throw registrationError_('INVALID_ID_NUMBER', field + ' harus kosong atau terdiri dari 16 digit angka.');
    }
  });

  if (!/^\d{4}-\d{2}-\d{2}$/.test(data.tanggalLahir)) {
    throw registrationError_('INVALID_BIRTH_DATE', 'Format tanggal lahir tidak valid.');
  }
}

function registrationError_(code, message) {
  var error = new Error(message);
  error.code = code;
  return error;
}

function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    return;
  }

  var currentHeaders = sheet.getRange(1, 1, 1, HEADERS.length).getDisplayValues()[0];
  var isBlank = currentHeaders.every(function (header) {
    return !String(header).trim();
  });

  if (isBlank) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    return;
  }

  var matches = HEADERS.every(function (header, index) {
    return String(currentHeaders[index]).trim() === header;
  });
  if (!matches) {
    var mismatchIndex = HEADERS.findIndex(function (header, index) {
      return String(currentHeaders[index]).trim() !== header;
    });
    throw new Error(
      'Header kolom ' + (mismatchIndex + 1) + ' (' + columnName_(mismatchIndex + 1) +
      ') pada tab "' + sheet.getName() + '" tidak cocok. Seharusnya "' +
      HEADERS[mismatchIndex] + '", tetapi ditemukan "' +
      String(currentHeaders[mismatchIndex]).trim() +
      '". Pastikan 25 header berada di baris pertama dan urutannya sesuai.'
    );
  }
}

function columnName_(columnNumber) {
  var name = '';
  while (columnNumber > 0) {
    var remainder = (columnNumber - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return name;
}

function generateRegistrationNumber_(sheet, jenjang) {
  var lastRow = sheet.getLastRow();
  var highestNumber = 0;

  if (lastRow > 1) {
    var registrations = sheet.getRange(2, 2, lastRow - 1, 2).getDisplayValues();
    var pattern = new RegExp('^' + jenjang + '(\\d+)$');

    registrations.forEach(function (row) {
      if (row[1] !== jenjang) return;
      var match = String(row[0]).match(pattern);
      if (match) {
        highestNumber = Math.max(highestNumber, Number(match[1]));
      }
    });
  }

  return jenjang + String(highestNumber + 1).padStart(3, '0');
}

function checkRegistrationStatus_(params) {
  try {
    var rawNoReg = params && (params.noRegistrasi || params.registrationNumber) || '';
    var targetNoReg = String(rawNoReg).trim().toUpperCase();

    if (!targetNoReg) {
      return {
        result: 'error',
        requestId: params && params.requestId ? params.requestId : '',
        code: 'MISSING_REGISTRATION_NUMBER',
        error: 'Nomor pendaftaran wajib diisi.'
      };
    }

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      throw new Error('Spreadsheet aktif tidak ditemukan. Pastikan Apps Script terikat ke Google Sheets.');
    }

    var sheet = spreadsheet.getActiveSheet();
    var lastRow = sheet.getLastRow();

    if (lastRow > 1) {
      var rows = sheet.getRange(2, 1, lastRow - 1, 13).getDisplayValues();
      for (var i = 0; i < rows.length; i++) {
        var row = rows[i];
        if (String(row[1]).trim().toUpperCase() === targetNoReg) {
          return {
            result: 'success',
            requestId: params && params.requestId ? params.requestId : '',
            data: {
              noRegistrasi: row[1],
              namaLengkap: row[7],
              jenjang: row[2],
              namaAsalSekolah: row[5],
              tempatLahir: row[11],
              tanggalLahir: row[12]
            },
            message: 'Data ananda sudah tersimpan. silakan catat nomor pendaftaran anda untuk Tes PPDB di sekolah.'
          };
        }
      }
    }

    return {
      result: 'not_found',
      requestId: params && params.requestId ? params.requestId : '',
      code: 'NOT_FOUND',
      error: 'Data pendaftaran dengan nomor "' + targetNoReg + '" tidak ditemukan. Pastikan nomor pendaftaran yang dimasukkan sudah benar.'
    };
  } catch (error) {
    return {
      result: 'error',
      requestId: params && params.requestId ? params.requestId : '',
      code: 'STATUS_CHECK_FAILED',
      error: 'Gagal memeriksa status pendaftaran: ' + (error && error.message ? error.message : String(error))
    };
  }
}

function combine_(firstValue, secondValue) {
  return [firstValue, secondValue]
    .map(function (value) { return String(value || '').trim(); })
    .filter(function (value) { return value; })
    .join('-');
}

function safeCellValue_(value) {
  if (typeof value !== 'string') return value;
  return /^[=+\-@]/.test(value) ? "'" + value : value;
}

function sendResponse_(payload, data) {
  if (data && data.responseMode === 'iframe') {
    return responsePage_(payload);
  }
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function responsePage_(payload) {
  var serializedPayload = JSON.stringify(payload)
    .replace(/</g, '\\u003c')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');
  var html = '<!doctype html><html><body><script>' +
    '(function(){' +
    'var msg={type:"ppdb-registration-response",payload:' + serializedPayload + '};' +
    'var targets=[];' +
    'try{if(window.top&&window.top!==window)targets.push(window.top);}catch(e){}' +
    'try{var w=window;while(w.parent&&w.parent!==w){w=w.parent;if(targets.indexOf(w)===-1)targets.push(w);}}catch(e){}' +
    'targets.forEach(function(t){try{t.postMessage(msg,"*");}catch(e){}});' +
    '})();' +
    '<\/script></body></html>';

  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
