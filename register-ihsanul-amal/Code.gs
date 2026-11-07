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

function doPost(e) {
  var lock = LockService.getScriptLock();

  try {
    var data = e && e.parameter ? e.parameter : {};
    validateRegistration_(data);
    lock.waitLock(10000);

    var cache = CacheService.getScriptCache();
    var cachedResponse = cache.get('ppdb:' + data.requestId);
    if (cachedResponse) {
      return responsePage_(JSON.parse(cachedResponse));
    }

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      throw new Error('Spreadsheet aktif tidak ditemukan. Pastikan Apps Script terikat ke Google Sheets.');
    }

    var sheet = spreadsheet.getActiveSheet();
    ensureHeaders_(sheet);

    var noRegistrasi = generateRegistrationNumber_(sheet, data.jenjang);
    var response = {
      result: 'success',
      noRegistrasi: noRegistrasi
    };

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

    cache.put('ppdb:' + data.requestId, JSON.stringify(response), 21600);
    return responsePage_(response);
  } catch (error) {
    return responsePage_({
      result: 'error',
      error: error && error.message ? error.message : String(error)
    });
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

function validateRegistration_(data) {
  if (!data.requestId || !/^[a-zA-Z0-9-]{8,100}$/.test(data.requestId)) {
    throw new Error('ID permintaan tidak valid. Muat ulang formulir dan coba lagi.');
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
      throw new Error('Field wajib belum diisi: ' + field);
    }
  });

  if (['KB', 'TK', 'SD', 'SMP'].indexOf(data.jenjang) === -1) {
    throw new Error('Jenjang sekolah tidak valid.');
  }

  var allowedPrograms = {
    KB: ['', 'TPA'],
    TK: ['', 'TPA'],
    SD: [''],
    SMP: ['', 'Boarding']
  };
  if (allowedPrograms[data.jenjang].indexOf(data.programTambahan || '') === -1) {
    throw new Error('Program tambahan tidak sesuai dengan jenjang sekolah.');
  }

  if (!/^[0-9]{10}$/.test(data.nisn)) {
    throw new Error('NISN harus terdiri dari 10 digit angka.');
  }
  if (!/^[0-9]{16}$/.test(data.nik)) {
    throw new Error('NIK harus terdiri dari 16 digit angka.');
  }

  ['nomorKK', 'nikAyah', 'nikIbu'].forEach(function (field) {
    if (data[field] && !/^[0-9]{16}$/.test(data[field])) {
      throw new Error(field + ' harus kosong atau terdiri dari 16 digit angka.');
    }
  });

  if (!/^\d{4}-\d{2}-\d{2}$/.test(data.tanggalLahir)) {
    throw new Error('Format tanggal lahir tidak valid.');
  }
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
    throw new Error('Header sheet tidak cocok. Gunakan 25 header sesuai urutan pada file Code.gs.');
  }
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

function responsePage_(payload) {
  var serializedPayload = JSON.stringify(payload)
    .replace(/</g, '\\u003c')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');
  var html = '<!doctype html><html><body><script>' +
    'window.parent.postMessage({type:"ppdb-registration-response",payload:' +
    serializedPayload + '}, "*");' +
    '<\/script></body></html>';

  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
