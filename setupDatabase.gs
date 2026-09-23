/**
 * ============================================================
 * PMB SETIAWATI
 * DATABASE SETUP V1
 * ============================================================
 *
 * Fungsi utama:
 *   setupDatabase()
 *
 * Database:
 *   Google Spreadsheet berdasarkan DATABASE_SPREADSHEET_ID
 *
 * Fitur:
 *   - Membuat seluruh sheet database
 *   - Membuat header
 *   - Membuat CONFIG
 *   - Membuat ADMIN awal
 *   - Membuat master TINDAKAN awal
 *   - Format database
 *   - Membuat filter
 *   - Freeze header
 *   - Audit log
 *   - Aman dijalankan berulang kali
 *   - Tidak menggunakan SpreadsheetApp.getUi()
 *     sebagai bagian wajib proses
 *
 * ============================================================
 */


/**
 * ============================================================
 * 1. SETUP DATABASE UTAMA
 * ============================================================
 */
function setupDatabase() {

  Logger.log('========================================');
  Logger.log('PMB SETIAWATI - SETUP DATABASE V1');
  Logger.log('========================================');

  try {

    // --------------------------------------------------------
    // Ambil database
    // --------------------------------------------------------
    const ss = getDatabase();

    if (!ss) {
      throw new Error(
        'Google Spreadsheet database tidak ditemukan.'
      );
    }

    Logger.log(
      'Database: ' + ss.getName()
    );

    Logger.log(
      'Spreadsheet ID: ' + ss.getId()
    );


    // --------------------------------------------------------
    // Struktur sheet
    // --------------------------------------------------------
    const sheetDefinitions = getSheetDefinitions();


    // --------------------------------------------------------
    // Buat seluruh sheet
    // --------------------------------------------------------
    Object.keys(sheetDefinitions).forEach(function(sheetName) {

      const definition = sheetDefinitions[sheetName];

      const sheet = getOrCreateSheet_(
        ss,
        sheetName
      );

      setupSheetStructure_(
        sheet,
        definition.headers
      );

    });


    // --------------------------------------------------------
    // CONFIG
    // --------------------------------------------------------
    setupConfigSheet_(ss);


    // --------------------------------------------------------
    // USER ADMIN AWAL
    // --------------------------------------------------------
    setupInitialAdmin_(ss);


    // --------------------------------------------------------
    // MASTER TINDAKAN
    // --------------------------------------------------------
    setupInitialActions_(ss);


    // --------------------------------------------------------
    // Format seluruh database
    // --------------------------------------------------------
    formatAllDatabaseSheets_(
      ss,
      sheetDefinitions
    );


    // --------------------------------------------------------
    // Audit log
    // --------------------------------------------------------
    writeAuditLog_(
      ss,
      'SYSTEM',
      'SETUP_DATABASE',
      'Setup database PMB SETIAWATI berhasil dijalankan.'
    );


    // --------------------------------------------------------
    // Selesai
    // --------------------------------------------------------
    Logger.log('========================================');
    Logger.log('SETUP DATABASE BERHASIL');
    Logger.log('========================================');

    Logger.log(
      'Database: ' + ss.getName()
    );

    Logger.log(
      'Spreadsheet ID: ' + ss.getId()
    );

    Logger.log(
      'Jumlah sheet: ' +
      Object.keys(sheetDefinitions).length
    );

    Logger.log('========================================');


    // --------------------------------------------------------
    // Notifikasi aman
    // --------------------------------------------------------
    showSetupMessage_(
      'Setup Database PMB SETIAWATI berhasil.\n\n' +
      'Database: ' + ss.getName() + '\n' +
      'Jumlah sheet: ' +
      Object.keys(sheetDefinitions).length
    );


  } catch (error) {

    Logger.log('========================================');
    Logger.log('SETUP DATABASE GAGAL');
    Logger.log('========================================');

    Logger.log(
      error &&
      error.stack
        ? error.stack
        : error.message
    );

    Logger.log('========================================');

    // Jangan menggunakan getUi() secara langsung.
    showSetupMessage_(
      'Setup Database gagal.\n\n' +
      'Detail:\n' +
      error.message
    );

    // Tetap lempar error supaya Apps Script menunjukkan
    // lokasi kesalahan secara jelas.
    throw error;
  }
}


/**
 * ============================================================
 * 2. DEFINISI SELURUH SHEET DATABASE
 * ============================================================
 */
function getSheetDefinitions() {

  return {

    CONFIG: {

      headers: [
        'KEY',
        'VALUE',
        'DESCRIPTION',
        'UPDATED_AT'
      ]

    },


    USERS: {

      headers: [
        'USER_ID',
        'USERNAME',
        'PASSWORD_HASH',
        'NAMA',
        'ROLE',
        'STATUS',
        'LAST_LOGIN',
        'CREATED_AT',
        'UPDATED_AT'
      ]

    },


    PASIEN: {

      headers: [
        'PASIEN_ID',
        'NO_RM',
        'NIK',
        'NAMA_LENGKAP',
        'NAMA_PANGGILAN',
        'TEMPAT_LAHIR',
        'TANGGAL_LAHIR',
        'JENIS_KELAMIN',
        'GOL_DARAH',
        'AGAMA',
        'STATUS_PERNIKAHAN',
        'PENDIDIKAN',
        'PEKERJAAN',
        'ALAMAT',
        'RT',
        'RW',
        'KELURAHAN',
        'KECAMATAN',
        'KOTA',
        'PROVINSI',
        'NO_HP',
        'EMAIL',
        'NAMA_SUAMI',
        'NO_HP_SUAMI',
        'NAMA_AYAH',
        'NAMA_IBU',
        'STATUS',
        'CATATAN',
        'CREATED_AT',
        'UPDATED_AT',
        'CREATED_BY',
        'UPDATED_BY'
      ]

    },


    KUNJUNGAN: {

      headers: [
        'KUNJUNGAN_ID',
        'NO_RM',
        'PASIEN_ID',
        'TANGGAL_KUNJUNGAN',
        'JENIS_KUNJUNGAN',
        'POLI',
        'DOKTER_BIDAN',
        'KELUHAN_UTAMA',
        'STATUS_KUNJUNGAN',
        'NO_ANTRIAN',
        'TOTAL_TAGIHAN',
        'STATUS_BAYAR',
        'CATATAN',
        'CREATED_AT',
        'UPDATED_AT',
        'CREATED_BY',
        'UPDATED_BY'
      ]

    },


    ANTRIAN: {

      headers: [
        'ANTRIAN_ID',
        'KUNJUNGAN_ID',
        'NO_RM',
        'NOMOR_ANTRIAN',
        'TANGGAL',
        'JENIS_LAYANAN',
        'STATUS',
        'WAKTU_DAFTAR',
        'WAKTU_PANGGIL',
        'WAKTU_SELESAI',
        'COUNTER',
        'CATATAN'
      ]

    },


    REKAM_MEDIS: {

      headers: [
        'RM_ID',
        'KUNJUNGAN_ID',
        'NO_RM',
        'TANGGAL',
        'TEKANAN_DARAH',
        'NADI',
        'SUHU',
        'RESPIRASI',
        'BERAT_BADAN',
        'TINGGI_BADAN',
        'IMT',
        'KELUHAN',
        'RIWAYAT_PENYAKIT',
        'RIWAYAT_ALERGI',
        'OBAT_DIGUNAKAN',
        'PEMERIKSAAN_FISIK',
        'DIAGNOSIS',
        'TINDAKAN',
        'EDUKASI',
        'RENCANA_TINDAK_LANJUT',
        'CATATAN',
        'PEMERIKSA',
        'CREATED_AT',
        'UPDATED_AT'
      ]

    },


    TINDAKAN: {

      headers: [
        'TINDAKAN_ID',
        'KODE',
        'NAMA_TINDAKAN',
        'KATEGORI',
        'HARGA',
        'SATUAN',
        'STATUS',
        'KETERANGAN',
        'CREATED_AT',
        'UPDATED_AT'
      ]

    },


    OBAT: {

      headers: [
        'OBAT_ID',
        'KODE_OBAT',
        'NAMA_OBAT',
        'KANDUNGAN',
        'SATUAN',
        'HARGA_BELI',
        'HARGA_JUAL',
        'STOK',
        'STOK_MINIMUM',
        'LOKASI',
        'STATUS',
        'KETERANGAN',
        'CREATED_AT',
        'UPDATED_AT'
      ]

    },


    RESEP: {

      headers: [
        'RESEP_ID',
        'KUNJUNGAN_ID',
        'NO_RM',
        'TANGGAL',
        'DOKTER_BIDAN',
        'STATUS',
        'CATATAN',
        'CREATED_AT',
        'UPDATED_AT'
      ]

    },


    RESEP_DETAIL: {

      headers: [
        'RESEP_DETAIL_ID',
        'RESEP_ID',
        'OBAT_ID',
        'KODE_OBAT',
        'NAMA_OBAT',
        'JUMLAH',
        'SATUAN',
        'DOSIS',
        'ATURAN_PAKAI',
        'CATATAN',
        'HARGA',
        'SUBTOTAL'
      ]

    },


    TRANSAKSI: {

      headers: [
        'TRANSAKSI_ID',
        'KUNJUNGAN_ID',
        'NO_RM',
        'TANGGAL',
        'JENIS_TRANSAKSI',
        'TOTAL',
        'DISKON',
        'GRAND_TOTAL',
        'STATUS',
        'KETERANGAN',
        'CREATED_AT',
        'UPDATED_AT'
      ]

    },


    TRANSAKSI_DETAIL: {

      headers: [
        'TRANSAKSI_DETAIL_ID',
        'TRANSAKSI_ID',
        'ITEM_ID',
        'KODE_ITEM',
        'NAMA_ITEM',
        'JENIS_ITEM',
        'QTY',
        'HARGA',
        'DISKON',
        'SUBTOTAL'
      ]

    },


    PEMBAYARAN: {

      headers: [
        'PEMBAYARAN_ID',
        'TRANSAKSI_ID',
        'TANGGAL',
        'METODE_BAYAR',
        'JUMLAH_BAYAR',
        'KEMBALIAN',
        'REFERENSI',
        'STATUS',
        'CATATAN',
        'DIBAYAR_OLEH',
        'CREATED_AT'
      ]

    },


    AUDIT_LOG: {

      headers: [
        'LOG_ID',
        'TIMESTAMP',
        'USER_ID',
        'USERNAME',
        'AKSI',
        'MODUL',
        'REFERENSI_ID',
        'KETERANGAN',
        'IP_ADDRESS'
      ]

    }

  };

}


/**
 * ============================================================
 * 3. GET / CREATE SHEET
 * ============================================================
 */
function getOrCreateSheet_(ss, sheetName) {

  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {

    sheet = ss.insertSheet(sheetName);

    Logger.log(
      'Sheet dibuat: ' + sheetName
    );

  } else {

    Logger.log(
      'Sheet sudah ada: ' + sheetName
    );

  }

  return sheet;
}


/**
 * ============================================================
 * 4. SETUP STRUKTUR SHEET
 * ============================================================
 */
function setupSheetStructure_(sheet, headers) {

  if (!headers || headers.length === 0) {
    return;
  }


  // ----------------------------------------------------------
  // Jika sheet benar-benar kosong
  // ----------------------------------------------------------
  if (
    sheet.getLastRow() === 0 ||
    (
      sheet.getLastRow() === 1 &&
      sheet.getLastColumn() === 1 &&
      String(sheet.getRange(1, 1).getValue()).trim() === ''
    )
  ) {

    sheet
      .getRange(1, 1, 1, headers.length)
      .setValues([headers]);

  } else {

    // --------------------------------------------------------
    // Pastikan header tersedia.
    // Jangan menimpa data yang sudah ada.
    // --------------------------------------------------------
    const existingLastColumn =
      Math.max(sheet.getLastColumn(), 1);

    const existingHeaders =
      sheet
        .getRange(
          1,
          1,
          1,
          existingLastColumn
        )
        .getValues()[0];


    headers.forEach(function(header, index) {

      const current =
        existingHeaders[index];

      if (
        typeof current === 'undefined' ||
        String(current).trim() === ''
      ) {

        sheet
          .getRange(1, index + 1)
          .setValue(header);

      }

    });

  }


  // ----------------------------------------------------------
  // Freeze header
  // ----------------------------------------------------------
  sheet.setFrozenRows(1);


  // ----------------------------------------------------------
  // Filter
  // ----------------------------------------------------------
  try {

    if (sheet.getFilter()) {
      sheet.getFilter().remove();
    }

    const lastColumn =
      Math.max(
        sheet.getLastColumn(),
        headers.length
      );

    const lastRow =
      Math.max(
        sheet.getLastRow(),
        1
      );

    sheet
      .getRange(
        1,
        1,
        lastRow,
        lastColumn
      )
      .createFilter();

  } catch (error) {

    Logger.log(
      'Filter tidak dibuat untuk ' +
      sheet.getName() +
      ': ' +
      error.message
    );

  }

}


/**
 * ============================================================
 * 5. SETUP CONFIG
 * ============================================================
 */
function setupConfigSheet_(ss) {

  const sheet =
    ss.getSheetByName(
      PMB_CONFIG.SHEETS.CONFIG
    );

  if (!sheet) {
    throw new Error(
      'Sheet CONFIG tidak ditemukan.'
    );
  }


  const configData = [

    [
      'APP_NAME',
      PMB_CONFIG.APP_NAME,
      'Nama aplikasi',
      new Date()
    ],

    [
      'APP_VERSION',
      PMB_CONFIG.APP_VERSION,
      'Versi aplikasi',
      new Date()
    ],

    [
      'CLINIC_NAME',
      PMB_CONFIG.CLINIC_NAME,
      'Nama klinik/praktik',
      new Date()
    ],

    [
      'CLINIC_TYPE',
      PMB_CONFIG.CLINIC_TYPE,
      'Jenis fasilitas kesehatan',
      new Date()
    ],

    [
      'ADDRESS',
      PMB_CONFIG.ADDRESS,
      'Alamat praktik',
      new Date()
    ],

    [
      'CITY',
      PMB_CONFIG.CITY,
      'Kota',
      new Date()
    ],

    [
      'TIMEZONE',
      PMB_CONFIG.TIMEZONE,
      'Zona waktu',
      new Date()
    ],

    [
      'RM_PREFIX',
      PMB_CONFIG.RM_PREFIX,
      'Prefix nomor rekam medis',
      new Date()
    ],

    [
      'TRANSACTION_PREFIX',
      PMB_CONFIG.TRANSACTION_PREFIX,
      'Prefix transaksi',
      new Date()
    ],

    [
      'QUEUE_PREFIX',
      PMB_CONFIG.QUEUE_PREFIX,
      'Prefix nomor antrian',
      new Date()
    ]

  ];


  // ----------------------------------------------------------
  // Jika CONFIG masih kosong
  // ----------------------------------------------------------
  if (sheet.getLastRow() <= 1) {

    sheet
      .getRange(
        2,
        1,
        configData.length,
        4
      )
      .setValues(configData);

  } else {

    // --------------------------------------------------------
    // Update berdasarkan KEY tanpa menghapus data lain
    // --------------------------------------------------------
    const lastRow =
      sheet.getLastRow();

    const values =
      sheet
        .getRange(
          2,
          1,
          Math.max(lastRow - 1, 1),
          4
        )
        .getValues();


    const rowMap = {};

    values.forEach(function(row, index) {

      const key =
        String(row[0] || '').trim();

      if (key) {
        rowMap[key] = index + 2;
      }

    });


    configData.forEach(function(item) {

      const key = item[0];

      if (rowMap[key]) {

        sheet
          .getRange(
            rowMap[key],
            2,
            1,
            3
          )
          .setValues([
            [
              item[1],
              item[2],
              new Date()
            ]
          ]);

      } else {

        sheet
          .appendRow(item);

      }

    });

  }


  Logger.log(
    'CONFIG berhasil disiapkan.'
  );
}


/**
 * ============================================================
 * 6. SETUP ADMIN AWAL
 * ============================================================
 */
function setupInitialAdmin_(ss) {

  const sheet =
    ss.getSheetByName(
      PMB_CONFIG.SHEETS.USERS
    );

  if (!sheet) {
    throw new Error(
      'Sheet USERS tidak ditemukan.'
    );
  }


  const lastRow =
    sheet.getLastRow();


  // ----------------------------------------------------------
  // Jika sudah ada user, jangan membuat admin baru
  // ----------------------------------------------------------
  if (lastRow > 1) {

    Logger.log(
      'USERS sudah memiliki data. ' +
      'Admin awal tidak dibuat ulang.'
    );

    return;
  }


  const now =
    new Date();


  const userId =
    generateId_(
      'USR-'
    );


  const username =
    'admin';


  const password =
    'admin123';


  const passwordHash =
    hashPassword_(
      password
    );


  const row = [

    userId,

    username,

    passwordHash,

    'Administrator',

    'SUPERADMIN',

    'AKTIF',

    '',

    now,

    now

  ];


  sheet
    .getRange(
      2,
      1,
      1,
      row.length
    )
    .setValues([row]);


  Logger.log(
    'Admin awal berhasil dibuat.'
  );

  Logger.log(
    'Username: admin'
  );

  Logger.log(
    'Password awal: admin123'
  );

  Logger.log(
    'PENTING: password harus diganti ' +
    'setelah modul login tersedia.'
  );
}


/**
 * ============================================================
 * 7. SETUP MASTER TINDAKAN
 * ============================================================
 */
function setupInitialActions_(ss) {

  const sheet =
    ss.getSheetByName(
      PMB_CONFIG.SHEETS.TINDAKAN
    );

  if (!sheet) {
    throw new Error(
      'Sheet TINDAKAN tidak ditemukan.'
    );
  }


  // ----------------------------------------------------------
  // Jangan membuat data duplikat
  // ----------------------------------------------------------
  const lastRow =
    sheet.getLastRow();


  if (lastRow > 1) {

    Logger.log(
      'TINDAKAN sudah memiliki data. ' +
      'Master awal tidak dibuat ulang.'
    );

    return;
  }


  const now =
    new Date();


  const actions = [

    [
      generateId_('TIN-'),
      'T001',
      'Pemeriksaan Umum',
      'PEMERIKSAAN',
      100000,
      'KALI',
      'AKTIF',
      '',
      now,
      now
    ],

    [
      generateId_('TIN-'),
      'T002',
      'Konsultasi',
      'KONSULTASI',
      75000,
      'KALI',
      'AKTIF',
      '',
      now,
      now
    ],

    [
      generateId_('TIN-'),
      'T003',
      'Pemeriksaan ANC',
      'KEBIDANAN',
      100000,
      'KALI',
      'AKTIF',
      '',
      now,
      now
    ]

  ];


  sheet
    .getRange(
      2,
      1,
      actions.length,
      actions[0].length
    )
    .setValues(actions);


  Logger.log(
    'Master tindakan awal berhasil dibuat.'
  );
}


/**
 * ============================================================
 * 8. FORMAT SELURUH DATABASE
 * ============================================================
 */
function formatAllDatabaseSheets_(
  ss,
  sheetDefinitions
) {

  Object.keys(sheetDefinitions)
    .forEach(function(sheetName) {

      const sheet =
        ss.getSheetByName(sheetName);

      if (!sheet) {
        return;
      }


      const lastColumn =
        Math.max(
          sheet.getLastColumn(),
          1
        );


      const lastRow =
        Math.max(
          sheet.getLastRow(),
          1
        );


      // ------------------------------------------------------
      // Header
      // ------------------------------------------------------
      const headerRange =
        sheet.getRange(
          1,
          1,
          1,
          lastColumn
        );


      headerRange
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');


      // ------------------------------------------------------
      // Tinggi header
      // ------------------------------------------------------
      sheet.setRowHeight(
        1,
        32
      );


      // ------------------------------------------------------
      // Border
      // ------------------------------------------------------
      if (lastRow >= 1) {

        sheet
          .getRange(
            1,
            1,
            lastRow,
            lastColumn
          )
          .setVerticalAlignment('middle');

      }


      // ------------------------------------------------------
      // Format tanggal umum
      // ------------------------------------------------------
      formatDateColumns_(
        sheet,
        sheetDefinitions[sheetName].headers
      );


      // ------------------------------------------------------
      // Auto resize
      // ------------------------------------------------------
      try {

        sheet.autoResizeColumns(
          1,
          lastColumn
        );

      } catch (error) {

        Logger.log(
          'Auto resize gagal pada ' +
          sheetName +
          ': ' +
          error.message
        );

      }


      // ------------------------------------------------------
      // Lebar minimum kolom
      // ------------------------------------------------------
      for (
        let col = 1;
        col <= lastColumn;
        col++
      ) {

        try {

          const width =
            sheet.getColumnWidth(col);

          if (width < 100) {

            sheet.setColumnWidth(
              col,
              120
            );

          }

        } catch (error) {
          // abaikan
        }

      }

    });


  Logger.log(
    'Format seluruh sheet selesai.'
  );
}


/**
 * ============================================================
 * 9. FORMAT KOLOM TANGGAL
 * ============================================================
 */
function formatDateColumns_(
  sheet,
  headers
) {

  const dateKeywords = [

    'TANGGAL',

    'DATE',

    'CREATED_AT',

    'UPDATED_AT',

    'LAST_LOGIN',

    'WAKTU_DAFTAR',

    'WAKTU_PANGGIL',

    'WAKTU_SELESAI',

    'TIMESTAMP'

  ];


  headers.forEach(function(header, index) {

    const normalized =
      String(header)
        .toUpperCase()
        .trim();


    const isDate =
      dateKeywords.some(function(keyword) {

        return normalized.indexOf(
          keyword
        ) !== -1;

      });


    if (isDate) {

      try {

        sheet
          .getRange(
            2,
            index + 1,
            Math.max(
              sheet.getMaxRows() - 1,
              1
            ),
            1
          )
          .setNumberFormat(
            'dd/MM/yyyy HH:mm:ss'
          );

      } catch (error) {

        Logger.log(
          'Format tanggal gagal: ' +
          sheet.getName() +
          ' kolom ' +
          header
        );

      }

    }

  });

}


/**
 * ============================================================
 * 10. AUDIT LOG
 * ============================================================
 */
function writeAuditLog_(
  ss,
  username,
  action,
  description
) {

  try {

    const sheet =
      ss.getSheetByName(
        PMB_CONFIG.SHEETS.AUDIT_LOG
      );


    if (!sheet) {
      return;
    }


    const now =
      new Date();


    const logId =
      generateId_(
        'LOG-'
      );


    sheet.appendRow([

      logId,

      now,

      '',

      username,

      action,

      'SYSTEM',

      '',

      description,

      ''

    ]);


    Logger.log(
      'Audit log berhasil ditulis.'
    );


  } catch (error) {

    Logger.log(
      'Audit log gagal: ' +
      error.message
    );

  }

}


/**
 * ============================================================
 * 11. GENERATE ID
 * ============================================================
 */
function generateId_(prefix) {

  const timestamp =
    new Date()
      .getTime()
      .toString(36)
      .toUpperCase();


  const random =
    Math.random()
      .toString(36)
      .substring(2, 7)
      .toUpperCase();


  return (
    prefix +
    timestamp +
    '-' +
    random
  );
}


/**
 * ============================================================
 * 12. HASH PASSWORD SHA-256
 * ============================================================
 */
function hashPassword_(password) {

  const bytes =
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      String(password),
      Utilities.Charset.UTF_8
    );


  return bytes
    .map(function(byte) {

      const value =
        byte < 0
          ? byte + 256
          : byte;

      return (
        value
          .toString(16)
          .padStart(2, '0')
      );

    })
    .join('');
}


/**
 * ============================================================
 * 13. NOTIFIKASI AMAN
 * ============================================================
 *
 * Fungsi ini sengaja tidak memaksa getUi().
 *
 * Jika script dijalankan dari Spreadsheet UI,
 * alert akan muncul.
 *
 * Jika dijalankan langsung dari Apps Script Editor,
 * pesan masuk ke Logger tanpa menyebabkan error.
 *
 * ============================================================
 */
function showSetupMessage_(message) {

  try {

    const ui =
      SpreadsheetApp.getUi();

    ui.alert(
      'PMB SETIAWATI',
      message,
      ui.ButtonSet.OK
    );

  } catch (error) {

    Logger.log(
      'MESSAGE: ' + message
    );

  }

}
