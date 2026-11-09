var PARENT_DRIVE_FOLDER_ID = '18IroJdtTbrpPnvvGyWp5wVbprGyRdu7L';
var MAX_FILE_SIZE_BYTES = 2 * 1024 * 1024; // 2 MB

var ATTACHMENT_SPECS = [
  { key: 'akta', label: 'Akta Kelahiran', suffix: 'Akta', required: true },
  { key: 'kk', label: 'Kartu Keluarga (KK)', suffix: 'KK', required: true },
  { key: 'kia', label: 'Kartu Identitas Anak (KIA)', suffix: 'KIA', required: false },
  { key: 'ktpAyah', label: 'KTP Ayah', suffix: 'KTP_Ayah', required: true },
  { key: 'ktpIbu', label: 'KTP Ibu', suffix: 'KTP_Ibu', required: true }
];

var HEADERS = [
  'waktuDaftar',
  'noRegistrasi',
  'jenjang',
  'programTambahan',
  'statusAsalSekolah',
  'namaAsalSekolah',
  'namaLengkap',
  'tempatLahir',
  'tanggalLahir',
  'jenisKelamin',
  'jalanRtRw',
  'provinsi',
  'kabupatenKota',
  'kecamatan',
  'desaKelurahan',
  'Data-Ayah',
  'Data-Ibu',
  'Data-Wali',
  'BerkasLampiran'
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
    var sheet = getSheetByJenjang_(spreadsheet, data.jenjang);
    ensureHeaders_(sheet);

    var noRegistrasi = generateRegistrationNumber_(sheet, data.jenjang);

    stage = 'mengunggah berkas lampiran ke Google Drive';
    var folderUrl = saveAttachmentsToDrive_(noRegistrasi, data._parsedAttachments);

    var response = {
      result: 'success',
      requestId: data.requestId,
      nama: data.namaLengkap,
      jenjang: data.jenjang,
      noRegistrasi: noRegistrasi,
      berkasLampiran: folderUrl
    };

    var dataAyahArray = [
      String(data.namaAyah || '').trim(),
      String(data.tempatLahirAyah || '').trim(),
      String(data.tanggalLahirAyah || '').trim(),
      String(data.pendidikanAyah || '').trim(),
      String(data.pekerjaanAyah || '').trim(),
      String(data.noHpAyah || '').trim(),
      String(data.penghasilanAyah || '').trim()
    ];

    var dataIbuArray = [
      String(data.namaIbu || '').trim(),
      String(data.tempatLahirIbu || '').trim(),
      String(data.tanggalLahirIbu || '').trim(),
      String(data.pendidikanIbu || '').trim(),
      String(data.pekerjaanIbu || '').trim(),
      String(data.noHpIbu || '').trim(),
      String(data.penghasilanIbu || '').trim()
    ];

    var namaWali = String(data.namaWali || '').trim();
    var noHpWali = String(data.noHpWali || '').trim();
    var dataWaliArray = (namaWali || noHpWali) ? [namaWali, noHpWali] : [];

    stage = 'menyimpan data pendaftaran ke spreadsheet';
    sheet.appendRow([
      new Date(),
      noRegistrasi,
      data.jenjang,
      data.programTambahan || '',
      data.statusAsalSekolah,
      data.namaAsalSekolah,
      data.namaLengkap,
      data.tempatLahir,
      data.tanggalLahir,
      data.jenisKelamin,
      data.jalanRtRw,
      data.provinsi,
      data.kabupatenKota,
      data.kecamatan,
      data.desaKelurahan,
      JSON.stringify(dataAyahArray),
      JSON.stringify(dataIbuArray),
      dataWaliArray.length ? JSON.stringify(dataWaliArray) : '',
      folderUrl
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
    'tempatLahir', 'tanggalLahir', 'jenisKelamin',
    'jalanRtRw', 'provinsi', 'kabupatenKota', 'kecamatan', 'desaKelurahan',
    'namaAyah', 'tempatLahirAyah', 'tanggalLahirAyah', 'pendidikanAyah', 'pekerjaanAyah', 'noHpAyah', 'penghasilanAyah',
    'namaIbu', 'tempatLahirIbu', 'tanggalLahirIbu', 'pendidikanIbu', 'pekerjaanIbu', 'noHpIbu', 'penghasilanIbu'
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

  ['tanggalLahir', 'tanggalLahirAyah', 'tanggalLahirIbu'].forEach(function (dateField) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(data[dateField])) {
      throw registrationError_('INVALID_BIRTH_DATE', 'Format tanggal lahir tidak valid pada field: ' + dateField);
    }
  });

  data._parsedAttachments = validateAttachments_(data.attachments);
}

function validateAttachments_(rawAttachments) {
  var attachments = rawAttachments;
  if (typeof rawAttachments === 'string') {
    try {
      attachments = JSON.parse(rawAttachments);
    } catch (_) {
      throw registrationError_('INVALID_ATTACHMENTS', 'Data berkas lampiran tidak valid. Silakan pilih ulang berkas Anda.');
    }
  }

  if (!attachments || typeof attachments !== 'object') {
    throw registrationError_('MISSING_ATTACHMENTS', 'Berkas lampiran wajib (Akta, KK, KTP Ayah, KTP Ibu) belum diunggah.');
  }

  var allowedMimeTypes = {
    'application/pdf': 'pdf',
    'image/jpeg': 'jpg',
    'image/jpg': 'jpg',
    'image/png': 'png'
  };

  var validated = {};

  ATTACHMENT_SPECS.forEach(function (spec) {
    var item = attachments[spec.key];
    if (!item || !item.base64) {
      if (spec.required) {
        throw registrationError_('MISSING_FILE_' + spec.key.toUpperCase(), 'Berkas wajib belum diunggah: ' + spec.label + '.');
      }
      return;
    }

    var mimeType = String(item.mimeType || '').toLowerCase().trim();
    var originalName = String(item.name || '').trim();
    var extFromMime = allowedMimeTypes[mimeType];
    var extMatch = originalName.match(/\.(pdf|jpg|jpeg|png)$/i);
    var ext = extMatch ? extMatch[1].toLowerCase() : extFromMime;

    if (ext === 'jpeg') ext = 'jpg';

    if (!ext || ['pdf', 'jpg', 'png'].indexOf(ext) === -1) {
      throw registrationError_(
        'INVALID_FILE_TYPE',
        'Format berkas ' + spec.label + ' tidak didukung. Gunakan file PDF, JPG, atau PNG.'
      );
    }

    if (!mimeType || !allowedMimeTypes[mimeType]) {
      mimeType = ext === 'pdf' ? 'application/pdf' : (ext === 'png' ? 'image/png' : 'image/jpeg');
    }

    var cleanBase64 = String(item.base64).replace(/^data:[^;]+;base64,/, '');
    var bytes;
    try {
      bytes = Utilities.base64Decode(cleanBase64);
    } catch (_) {
      throw registrationError_('CORRUPT_FILE', 'Isi berkas ' + spec.label + ' gagal dibaca. Silakan unggah ulang.');
    }

    if (!bytes || bytes.length === 0) {
      throw registrationError_('EMPTY_FILE', 'Berkas ' + spec.label + ' kosong.');
    }

    if (bytes.length > MAX_FILE_SIZE_BYTES) {
      throw registrationError_(
        'FILE_TOO_LARGE',
        'Ukuran berkas ' + spec.label + ' melebihi batas maksimal 2 MB.'
      );
    }

    validated[spec.key] = {
      bytes: bytes,
      mimeType: mimeType,
      ext: ext,
      suffix: spec.suffix
    };
  });

  return validated;
}

function saveAttachmentsToDrive_(noRegistrasi, validatedAttachments) {
  var parentFolder;
  try {
    parentFolder = DriveApp.getFolderById(PARENT_DRIVE_FOLDER_ID);
  } catch (err) {
    throw new Error('Folder utama Google Drive (' + PARENT_DRIVE_FOLDER_ID + ') tidak ditemukan atau belum diberi izin akses.');
  }

  var existingFolders = parentFolder.getFoldersByName(noRegistrasi);
  var regFolder = existingFolders.hasNext()
    ? existingFolders.next()
    : parentFolder.createFolder(noRegistrasi);

  ATTACHMENT_SPECS.forEach(function (spec) {
    var fileData = validatedAttachments[spec.key];
    if (!fileData) return;
    var fileName = noRegistrasi + '_' + fileData.suffix + '.' + fileData.ext;

    // Hapus file lama dengan nama yang sama jika ada percobaan ulang
    var existingFiles = regFolder.getFilesByName(fileName);
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
    }

    var blob = Utilities.newBlob(fileData.bytes, fileData.mimeType, fileName);
    regFolder.createFile(blob);
  });

  return regFolder.getUrl();
}

function registrationError_(code, message) {
  var error = new Error(message);
  error.code = code;
  return error;
}

function getSheetByJenjang_(spreadsheet, jenjang) {
  var sheetName = String(jenjang || '').trim().toUpperCase();
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  return sheet;
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
      '". Pastikan 19 header berada di baris pertama dan urutannya sesuai.'
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

    var sheetNames = ['KB', 'TK', 'SD', 'SMP'];
    var prefixMatch = targetNoReg.match(/^(KB|TK|SD|SMP)/);
    if (prefixMatch) {
      var preferred = prefixMatch[1];
      sheetNames = [preferred].concat(sheetNames.filter(function (name) {
        return name !== preferred;
      }));
    }

    for (var s = 0; s < sheetNames.length; s++) {
      var sheet = spreadsheet.getSheetByName(sheetNames[s]);
      if (!sheet) continue;
      var lastRow = sheet.getLastRow();
      if (lastRow <= 1) continue;

      var rows = sheet.getRange(2, 1, lastRow - 1, 9).getDisplayValues();
      for (var i = 0; i < rows.length; i++) {
        var row = rows[i];
        if (String(row[1]).trim().toUpperCase() === targetNoReg) {
          return {
            result: 'success',
            requestId: params && params.requestId ? params.requestId : '',
            data: {
              noRegistrasi: row[1],
              namaLengkap: row[6],
              jenjang: row[2],
              namaAsalSekolah: row[5],
              tempatLahir: row[7],
              tanggalLahir: row[8]
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

function authorizeDriveAccess() {
  var folder = DriveApp.getFolderById(PARENT_DRIVE_FOLDER_ID);
  Logger.log('Akses Google Drive berhasil! Nama folder utama: ' + folder.getName() + ' | URL: ' + folder.getUrl());
}
