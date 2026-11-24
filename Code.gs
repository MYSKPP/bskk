/************************************************************
 * PORTAL UNIT KESENIAN BSKK
 * CODE.GS - FASA 1 & FASA 2
 * DATABASE + ADMIN DASHBOARD + COMMENTS
 ************************************************************/

const PORTAL_CONFIG = {
  NAMA: 'PORTAL UNIT KESENIAN BSKK',
  ORGANISASI: 'BAHAGIAN SUKAN, KOKURIKULUM DAN KESENIAN',
  KEMENTERIAN: 'KEMENTERIAN PENDIDIKAN MALAYSIA',
  TIMEZONE: 'Asia/Kuala_Lumpur',
  VERSI: '1.1.0'
};

const SUPER_ADMIN_AWAL = [
  'g-72522241@moe-dl.edu.my',
  'g-81358614@moe-dl.edu.my',
  'sekolah-1428@moe-dl.edu.my'
];

const DATABASE_SCHEMA = {
  POSTS: [
    'ID_POST', 'TAJUK', 'SLUG', 'KANDUNGAN', 'RINGKASAN',
    'ID_KATEGORI', 'IMEJ_UTAMA', 'STATUS', 'PINNED',
    'TARIKH_TERBIT', 'TARIKH_CIPTA', 'TARIKH_KEMASKINI',
    'PENULIS', 'JUMLAH_PAPARAN', 'JUMLAH_LIKE', 'JUMLAH_DISLIKE'
  ],

  CATEGORIES: [
    'ID_KATEGORI', 'NAMA_KATEGORI', 'SLUG',
    'PENERANGAN', 'SUSUNAN', 'STATUS', 'TARIKH_CIPTA'
  ],

  MEDIA: [
    'ID_MEDIA', 'NAMA_FAIL', 'JENIS_FAIL', 'URL_FAIL',
    'URL_THUMBNAIL', 'SAIZ_FAIL', 'FOLDER',
    'DIMUAT_NAIK_OLEH', 'TARIKH_MUAT_NAIK', 'STATUS'
  ],

  ADMINS: [
    'ID_ADMIN', 'EMAIL', 'NAMA', 'PERANAN',
    'STATUS', 'TARIKH_DAFTAR', 'DIDAFTAR_OLEH'
  ],

  SETTINGS: [
    'KUNCI', 'NILAI', 'PENERANGAN', 'TARIKH_KEMASKINI'
  ],

  ACTIVITY_LOG: [
    'ID_LOG', 'TARIKH_MASA', 'EMAIL', 'AKTIVITI', 'BUTIRAN'
  ],

  COMMENTS: [
    'ID_COMMENT',
    'ID_POST',
    'NAMA_PENGUNJUNG',
    'EMAIL',
    'KANDUNGAN',
    'STATUS',
    'TARIKH_CIPTA',
    'TARIKH_KEMASKINI'
  ]
};


/************************************************************
 * FASA 1 - SETUP / KEMAS KINI DATABASE SEDIA ADA
 ************************************************************/

function setupPortalBaru() {

  const properties = PropertiesService.getScriptProperties();

  let spreadsheetId = properties.getProperty('PORTAL_DB_ID');
  let folderId = properties.getProperty('PORTAL_FOLDER_ID');

  let spreadsheet;
  let folder;

  if (!spreadsheetId) {
    spreadsheet = SpreadsheetApp.create(
      'DATABASE PORTAL UNIT KESENIAN BSKK'
    );

    spreadsheetId = spreadsheet.getId();
    properties.setProperty('PORTAL_DB_ID', spreadsheetId);

  } else {
    spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  }

  if (!folderId) {
    folder = DriveApp.createFolder('PORTAL_UNIT_KESENIAN_BSKK');

    folderId = folder.getId();
    properties.setProperty('PORTAL_FOLDER_ID', folderId);

    DriveApp.getFileById(spreadsheetId).moveTo(folder);

  } else {
    folder = DriveApp.getFolderById(folderId);
  }

  spreadsheet.setSpreadsheetTimeZone(PORTAL_CONFIG.TIMEZONE);

  binaSemuaHelaian_(spreadsheet);
  binaFolderMedia_(folder);
  daftarSuperAdminAwal_(spreadsheet);
  masukkanTetapanAsas_(spreadsheet);

  rekodAktiviti_(
    'SISTEM',
    'KEMAS KINI DATABASE',
    'Database disediakan atau dikemas kini.'
  );

  Logger.log('DATABASE: ' + spreadsheet.getUrl());
  Logger.log('FOLDER: ' + folder.getUrl());
  Logger.log('HELAIAN: ' +
    spreadsheet.getSheets().map(s => s.getName()).join(', '));

  return {
    status: 'BERJAYA',
    spreadsheetUrl: spreadsheet.getUrl(),
    folderUrl: folder.getUrl()
  };
}


/************************************************************
 * BINA / SEMAK SEMUA HELAIAN
 * Helaian sedia ada tidak dipadam.
 ************************************************************/

function binaSemuaHelaian_(spreadsheet) {

  Object.keys(DATABASE_SCHEMA).forEach(function(nama) {

    let sheet = spreadsheet.getSheetByName(nama);

    if (!sheet) {

      if (
        nama === 'POSTS' &&
        spreadsheet.getSheets().length === 1 &&
        spreadsheet.getSheets()[0].getLastRow() === 0
      ) {
        sheet = spreadsheet.getSheets()[0];
        sheet.setName(nama);
      } else {
        sheet = spreadsheet.insertSheet(nama);
      }
    }

    const headers = DATABASE_SCHEMA[nama];

    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    formatHelaian_(sheet, headers.length);
  });
}


function formatHelaian_(sheet, jumlahKolum) {

  sheet.setFrozenRows(1);

  sheet.getRange(1, 1, 1, jumlahKolum)
    .setBackground('#172554')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.setRowHeight(1, 38);
  sheet.autoResizeColumns(1, jumlahKolum);
}


/************************************************************
 * FOLDER MEDIA
 ************************************************************/

function binaFolderMedia_(folderUtama) {

  [
    'POST_IMAGES',
    'POST_ATTACHMENTS',
    'DOCUMENTS',
    'VIDEOS',
    'THUMBNAILS'
  ].forEach(function(nama) {

    if (!folderUtama.getFoldersByName(nama).hasNext()) {
      folderUtama.createFolder(nama);
    }

  });
}


/************************************************************
 * DAFTAR SUPER ADMIN AWAL
 ************************************************************/

function daftarSuperAdminAwal_(spreadsheet) {

  const sheet = spreadsheet.getSheetByName('ADMINS');
  const data = sheet.getDataRange().getValues();

  const emails = data.slice(1).map(row =>
    String(row[1] || '').toLowerCase().trim()
  );

  SUPER_ADMIN_AWAL.forEach(function(email) {

    email = email.toLowerCase().trim();

    if (!emails.includes(email)) {
      sheet.appendRow([
        Utilities.getUuid(),
        email,
        '',
        'SUPER ADMIN',
        'AKTIF',
        new Date(),
        'SISTEM'
      ]);
    }

  });
}


/************************************************************
 * TETAPAN PORTAL
 ************************************************************/

function masukkanTetapanAsas_(spreadsheet) {

  const sheet = spreadsheet.getSheetByName('SETTINGS');

  const tetapan = [
    ['NAMA_PORTAL', PORTAL_CONFIG.NAMA, 'Nama rasmi portal'],
    ['ORGANISASI', PORTAL_CONFIG.ORGANISASI, 'Nama organisasi'],
    ['KEMENTERIAN', PORTAL_CONFIG.KEMENTERIAN, 'Nama kementerian'],
    ['TIMEZONE', PORTAL_CONFIG.TIMEZONE, 'Zon masa sistem'],
    ['VERSI_SISTEM', PORTAL_CONFIG.VERSI, 'Versi sistem'],
    ['PAPARAN_PORTAL', 'AKTIF', 'Status laman awam'],
    ['BIL_POST_PER_PAGE', '10', 'Bilangan post setiap halaman'],
    ['WARNA_UTAMA', '#172554', 'Warna utama portal']
  ];

  const data = sheet.getDataRange().getValues();
  const keys = data.slice(1).map(row => String(row[0] || '').trim());

  tetapan.forEach(function(item) {
    if (!keys.includes(item[0])) {
      sheet.appendRow([
        item[0], item[1], item[2], new Date()
      ]);
    }
  });
}


/************************************************************
 * DATABASE HELPER
 ************************************************************/

function getDatabase_() {

  const id = PropertiesService.getScriptProperties()
    .getProperty('PORTAL_DB_ID');

  if (!id) {
    throw new Error(
      'Database tidak dijumpai. Jalankan setupPortalBaru() dahulu.'
    );
  }

  return SpreadsheetApp.openById(id);
}


function getAdminEmail_() {
  return String(Session.getActiveUser().getEmail() || '')
    .toLowerCase().trim();
}


function semakAksesAdmin_() {

  const email = getAdminEmail_();

  if (!email) {
    throw new Error(
      'Email pengguna tidak dapat dikenal pasti. Sila gunakan akaun DELIMa yang berdaftar.'
    );
  }

  const sheet = getDatabase_().getSheetByName('ADMINS');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {

    const emel = String(data[i][1] || '').toLowerCase().trim();
    const status = String(data[i][4] || '').toUpperCase().trim();

    if (emel === email && status === 'AKTIF') {
      return {
        email: email,
        nama: data[i][2] || email,
        peranan: data[i][3] || 'ADMIN'
      };
    }
  }

  throw new Error(
    'Akses ditolak. Akaun anda bukan Admin Aktif.'
  );
}


function semakSuperAdmin_() {

  const admin = semakAksesAdmin_();

  if (String(admin.peranan).toUpperCase() !== 'SUPER ADMIN') {
    throw new Error('Fungsi ini hanya untuk SUPER ADMIN.');
  }

  return admin;
}


/************************************************************
 * LOG AKTIVITI
 ************************************************************/

function rekodAktiviti_(email, aktiviti, butiran) {

  try {
    const db = getDatabase_();
    const sheet = db.getSheetByName('ACTIVITY_LOG');

    sheet.appendRow([
      Utilities.getUuid(),
      new Date(),
      email,
      aktiviti,
      butiran
    ]);

  } catch (err) {
    console.log('Log aktiviti gagal: ' + err.message);
  }
}


/************************************************************
 * DOGET - PAPARAN ADMIN / PUBLIC
 ************************************************************/

function doGet(e) {
  const page = e && e.parameter ? e.parameter.page : '';

  if (page === 'public') {
    return HtmlService
      .createHtmlOutputFromFile('index')
      .setTitle('PORTAL UNIT KESENIAN BSKK')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  return HtmlService
    .createHtmlOutputFromFile('admin')
    .setTitle('Admin - PORTAL UNIT KESENIAN BSKK')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


function getPublicPortalUrl() {
  return ScriptApp.getService().getUrl() + '?page=public';
}


/************************************************************
 * DASHBOARD
 ************************************************************/

function getAdminDashboardData() {

  const pengguna = semakAksesAdmin_();
  const db = getDatabase_();

  const posts = db.getSheetByName('POSTS').getDataRange().getValues();
  const categories = db.getSheetByName('CATEGORIES').getDataRange().getValues();
  const admins = db.getSheetByName('ADMINS').getDataRange().getValues();
  const media = db.getSheetByName('MEDIA').getDataRange().getValues();
  const comments = db.getSheetByName('COMMENTS').getDataRange().getValues();

  const postData = posts.slice(1);
  const commentData = comments.slice(1);

  const countStatus = function(data, index, status) {
    return data.filter(row =>
      String(row[index] || '').toUpperCase() === status
    ).length;
  };

  return {
    pengguna: pengguna,
    portal: PORTAL_CONFIG.NAMA,

    statistik: {
      jumlahPost: postData.filter(row => row[0]).length,
      diterbitkan: countStatus(postData, 7, 'PUBLISHED'),
      draf: countStatus(postData, 7, 'DRAFT'),
      kategori: categories.slice(1).filter(row => row[0]).length,
      media: media.slice(1).filter(row => row[0]).length,
      admin: countStatus(admins.slice(1), 4, 'AKTIF'),
      comments: commentData.filter(row => row[0]).length,
      pending: countStatus(commentData, 5, 'PENDING'),
      approved: countStatus(commentData, 5, 'APPROVED'),
      rejected: countStatus(commentData, 5, 'REJECTED')
    },

    postTerkini: dapatkanPostTerkini_(postData, 5),

    masa: Utilities.formatDate(
      new Date(),
      PORTAL_CONFIG.TIMEZONE,
      'dd/MM/yyyy HH:mm'
    )
  };
}


function dapatkanPostTerkini_(data, had) {

  return data
    .filter(row => row[0])
    .sort((a, b) => {
      const ta = a[10] instanceof Date ? a[10].getTime() : 0;
      const tb = b[10] instanceof Date ? b[10].getTime() : 0;
      return tb - ta;
    })
    .slice(0, had)
    .map(row => ({
      id: row[0],
      tajuk: row[1] || '(Tiada tajuk)',
      status: row[7] || 'DRAFT',
      penulis: row[12] || '-',
      tarikh: formatTarikh_(row[10])
    }));
}


function formatTarikh_(value) {
  return value instanceof Date
    ? Utilities.formatDate(value, PORTAL_CONFIG.TIMEZONE, 'dd/MM/yyyy')
    : '-';
}


/************************************************************
 * SENARAI POST
 ************************************************************/

function getAdminPosts() {

  semakAksesAdmin_();

  const sheet = getDatabase_().getSheetByName('POSTS');
  const data = sheet.getDataRange().getValues();

  return data.slice(1)
    .filter(row => row[0])
    .reverse()
    .map(row => ({
      id: row[0],
      tajuk: row[1] || '(Tiada tajuk)',
      kategori: row[5] || '-',
      status: row[7] || 'DRAFT',
      tarikh: formatTarikh_(row[10]),
      penulis: row[12] || '-'
    }));
}


/************************************************************
 * COMMENTS - SENARAI DAN KIRAAN (ADMIN)
 ************************************************************/

function getAdminComments() {

  semakAksesAdmin_();

  const db = getDatabase_();
  const commentSheet = db.getSheetByName('COMMENTS');
  const postSheet = db.getSheetByName('POSTS');

  const comments = commentSheet.getDataRange().getValues().slice(1);
  const posts = postSheet.getDataRange().getValues().slice(1);

  const postMap = {};

  posts.forEach(function(row) {
    if (row[0]) {
      postMap[String(row[0])] = row[1] || '(Tiada tajuk)';
    }
  });

  const hasil = comments
    .filter(row => row[0])
    .map(row => ({
      id: String(row[0]),
      idPost: String(row[1] || ''),
      tajukPost: postMap[String(row[1])] || '(Post tidak dijumpai)',
      nama: String(row[2] || ''),
      email: String(row[3] || ''),
      kandungan: String(row[4] || ''),
      status: String(row[5] || 'PENDING').toUpperCase(),
      tarikh: row[6] instanceof Date
        ? Utilities.formatDate(row[6], PORTAL_CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm')
        : '-'
    }));

  hasil.sort((a, b) => b.id.localeCompare(a.id));

  return {
    comments: hasil,
    statistik: {
      all: hasil.length,
      pending: hasil.filter(x => x.status === 'PENDING').length,
      approved: hasil.filter(x => x.status === 'APPROVED').length,
      rejected: hasil.filter(x => x.status === 'REJECTED').length
    }
  };
}


/************************************************************
 * COMMENTS - TUKAR STATUS
 ************************************************************/

function tukarStatusComment(id, status) {

  const admin = semakAksesAdmin_();

  id = String(id || '').trim();
  status = String(status || '').toUpperCase().trim();

  if (!id) {
    throw new Error('ID komen tidak sah.');
  }

  if (!['PENDING', 'APPROVED', 'REJECTED'].includes(status)) {
    throw new Error('Status komen tidak sah.');
  }

  const sheet = getDatabase_().getSheetByName('COMMENTS');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {

    if (String(data[i][0]) === id) {

      sheet.getRange(i + 1, 6).setValue(status);
      sheet.getRange(i + 1, 8).setValue(new Date());

      rekodAktiviti_(
        admin.email,
        'KEMAS KINI KOMEN',
        'ID: ' + id + ' | Status: ' + status
      );

      return {
        success: true,
        message: 'Status komen berjaya dikemas kini.'
      };
    }
  }

  throw new Error('Komen tidak dijumpai.');
}


/************************************************************
 * COMMENTS - PADAM
 ************************************************************/

function padamComment(id) {

  const admin = semakAksesAdmin_();

  id = String(id || '').trim();

  if (!id) {
    throw new Error('ID komen tidak sah.');
  }

  const sheet = getDatabase_().getSheetByName('COMMENTS');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {

    if (String(data[i][0]) === id) {

      sheet.deleteRow(i + 1);

      rekodAktiviti_(
        admin.email,
        'PADAM KOMEN',
        'ID komen: ' + id
      );

      return {
        success: true,
        message: 'Komen berjaya dipadam.'
      };
    }
  }

  throw new Error('Komen tidak dijumpai.');
}


/************************************************************
 * COMMENTS - HANTAR KOMEN DARIPADA LAMAN AWAM
 * KOMEN_AUTO_LULUS = true  : komen terus dipaparkan (APPROVED)
 * KOMEN_AUTO_LULUS = false : komen menunggu kelulusan (PENDING)
 * (Versi laju: hanya baca lajur ID_POST dan STATUS)
 ************************************************************/

const KOMEN_AUTO_LULUS = true;

function tambahKomenPublic(data) {

  data = data || {};

  const idPost = String(data.idPost || '').trim();
  let nama = String(data.nama || '').trim();
  const email = String(data.email || '').trim();
  let kandungan = String(data.kandungan || '').trim();

  if (!idPost || !nama || !kandungan) {
    throw new Error('Sila lengkapkan nama dan kandungan komen.');
  }

  if (nama.length > 100 || kandungan.length > 3000 || email.length > 150) {
    throw new Error('Maklumat komen melebihi had yang dibenarkan.');
  }

  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error('Format email tidak sah.');
  }

  // elak formula injection dalam Sheet
  if (/^[=+\-@]/.test(nama)) nama = "'" + nama;
  if (/^[=+\-@]/.test(kandungan)) kandungan = "'" + kandungan;

  const db = getDatabase_();
  const postSheet = db.getSheetByName('POSTS');
  const last = postSheet.getLastRow();

  let postWujud = false;

  if (last >= 2) {
    const ids = postSheet.getRange(2, 1, last - 1, 1).getValues();
    const statuses = postSheet.getRange(2, 8, last - 1, 1).getValues();

    for (let i = 0; i < ids.length; i++) {
      if (
        String(ids[i][0]) === idPost &&
        String(statuses[i][0]).toUpperCase() === 'PUBLISHED'
      ) {
        postWujud = true;
        break;
      }
    }
  }

  if (!postWujud) {
    throw new Error('Post tidak dijumpai atau belum diterbitkan.');
  }

  const id = 'COM-' + Utilities.getUuid().slice(0, 8).toUpperCase();

  db.getSheetByName('COMMENTS').appendRow([
    id,
    idPost,
    nama,
    email,
    kandungan,
    KOMEN_AUTO_LULUS ? 'APPROVED' : 'PENDING',
    new Date(),
    new Date()
  ]);

  return {
    success: true,
    approved: KOMEN_AUTO_LULUS,
    message: KOMEN_AUTO_LULUS
      ? 'Komen anda telah dipaparkan.'
      : 'Komen diterima dan menunggu kelulusan Admin.'
  };
}


/************************************************************
 * COMMENTS - PAPARAN AWAM (HANYA KOMEN APPROVED)
 * Email pengunjung tidak dihantar ke laman awam.
 ************************************************************/

function getPublicComments(idPost) {

  idPost = String(idPost || '').trim();
  if (!idPost) return [];

  const sheet = getDatabase_().getSheetByName('COMMENTS');
  const last = sheet.getLastRow();

  if (last < 2) return [];

  const rows = sheet.getRange(2, 1, last - 1, 7).getValues();

  return rows
    .filter(r =>
      String(r[1]) === idPost &&
      String(r[5]).toUpperCase() === 'APPROVED'
    )
    .sort((a, b) => new Date(b[6]) - new Date(a[6]))
    .map(r => ({
      nama: String(r[2] || ''),
      komen: String(r[4] || ''),
      tarikh: r[6] instanceof Date
        ? Utilities.formatDate(r[6], PORTAL_CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm')
        : ''
    }));
}


/************************************************************
 * LIKE / DISLIKE ARTIKEL (LAMAN AWAM)
 * Disimpan dalam POSTS lajur O (JUMLAH_LIKE) dan P (JUMLAH_DISLIKE).
 * undi / undiSebelum: '' | 'like' | 'dislike'
 ************************************************************/

function undiPost(idPost, undi, undiSebelum) {

  idPost = String(idPost || '').trim();
  undi = String(undi || '').toLowerCase().trim();
  undiSebelum = String(undiSebelum || '').toLowerCase().trim();

  const sah = ['', 'like', 'dislike'];

  if (!idPost) throw new Error('Artikel tidak sah.');
  if (!sah.includes(undi) || !sah.includes(undiSebelum)) {
    throw new Error('Undian tidak sah.');
  }

  const sheet = getDatabase_().getSheetByName('POSTS');
  const last = sheet.getLastRow();

  if (last < 2) throw new Error('Artikel tidak dijumpai.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {

    const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
    const statuses = sheet.getRange(2, 8, last - 1, 1).getValues();

    let rowNumber = -1;

    for (let i = 0; i < ids.length; i++) {
      if (
        String(ids[i][0]) === idPost &&
        String(statuses[i][0]).toUpperCase() === 'PUBLISHED'
      ) {
        rowNumber = i + 2;
        break;
      }
    }

    if (rowNumber === -1) throw new Error('Artikel tidak dijumpai.');

    // pastikan header lajur like/dislike wujud (untuk sheet lama)
    const hdr = sheet.getRange(1, 15, 1, 2);
    const h = hdr.getValues()[0];

    if (!h[0] || !h[1]) {
      sheet.getRange(1, 14).copyTo(
        hdr,
        SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
        false
      );
      hdr.setValues([['JUMLAH_LIKE', 'JUMLAH_DISLIKE']]);
    }

    const cell = sheet.getRange(rowNumber, 15, 1, 2);
    const v = cell.getValues()[0];

    let like = Number(v[0] || 0);
    let dislike = Number(v[1] || 0);

    if (undiSebelum === 'like') like = Math.max(0, like - 1);
    if (undiSebelum === 'dislike') dislike = Math.max(0, dislike - 1);

    if (undi === 'like') like++;
    if (undi === 'dislike') dislike++;

    cell.setValues([[like, dislike]]);

    return { like: like, dislike: dislike };

  } finally {
    lock.releaseLock();
  }
}


/************************************************************
 * PROFIL ADMIN
 ************************************************************/

function getAdminProfile() {
  return semakAksesAdmin_();
}


/************************************************************
 * SEMAK DATABASE
 ************************************************************/

function semakDatabasePortal() {

  const db = getDatabase_();

  Logger.log('NAMA PORTAL: ' + PORTAL_CONFIG.NAMA);
  Logger.log('DATABASE: ' + db.getUrl());
  Logger.log('HELAIAN: ' +
    db.getSheets().map(sheet => sheet.getName()).join(', '));

  return {
    namaPortal: PORTAL_CONFIG.NAMA,
    databaseUrl: db.getUrl(),
    helaian: db.getSheets().map(sheet => sheet.getName())
  };
}


/************************************************************
 * FASA 3 - PENGURUSAN POST
 ************************************************************/

function getPostCategories() {
  semakAksesAdmin_();

  const sheet = getDatabase_().getSheetByName('CATEGORIES');
  const data = sheet.getDataRange().getValues();

  return data.slice(1)
    .filter(row =>
      row[0] &&
      String(row[5] || 'AKTIF').toUpperCase() === 'AKTIF'
    )
    .map(row => ({
      id: String(row[0]),
      nama: String(row[1] || ''),
      status: String(row[5] || 'AKTIF')
    }));
}


/************************************************************
 * DAPATKAN SATU POST UNTUK EDIT
 ************************************************************/

function getAdminPostById(id) {
  semakAksesAdmin_();

  id = String(id || '').trim();

  const sheet = getDatabase_().getSheetByName('POSTS');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) {
      const row = data[i];

      return {
        id: String(row[0]),
        tajuk: String(row[1] || ''),
        slug: String(row[2] || ''),
        kandungan: String(row[3] || ''),
        ringkasan: String(row[4] || ''),
        kategori: String(row[5] || ''),
        imej: String(row[6] || ''),
        status: String(row[7] || 'DRAFT').toUpperCase(),
        pinned: row[8] === true
      };
    }
  }

  throw new Error('Artikel tidak dijumpai.');
}


/************************************************************
 * SIMPAN / KEMAS KINI POST
 ************************************************************/

function simpanPost(data) {
  const admin = semakAksesAdmin_();

  data = data || {};

  const id = String(data.id || '').trim();
  const tajuk = String(data.tajuk || '').trim();
  const kandungan = String(data.kandungan || '').trim();
  const ringkasan = String(data.ringkasan || '').trim();
  const kategori = String(data.kategori || '').trim();
  const status = String(data.status || 'DRAFT').toUpperCase();

  if (!tajuk) {
    throw new Error('Sila masukkan tajuk artikel.');
  }

  if (!kandungan) {
    throw new Error('Sila masukkan kandungan artikel.');
  }

  if (tajuk.length > 250) {
    throw new Error('Tajuk terlalu panjang.');
  }

  if (ringkasan.length > 1000) {
    throw new Error('Ringkasan tidak boleh melebihi 1,000 aksara.');
  }

  if (!['DRAFT', 'PUBLISHED'].includes(status)) {
    throw new Error('Status artikel tidak sah.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const db = getDatabase_();
    const sheet = db.getSheetByName('POSTS');

    let rowNumber = -1;
    let oldRow = null;

    const values = sheet.getDataRange().getValues();

    if (id) {
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][0]) === id) {
          rowNumber = i + 1;
          oldRow = values[i];
          break;
        }
      }

      if (rowNumber === -1) {
        throw new Error('Artikel untuk dikemas kini tidak dijumpai.');
      }
    }

    const postId = id || (
      'POST-' + Utilities.getUuid().slice(0, 8).toUpperCase()
    );

    const slug = janaSlugPost_(tajuk);

    let imej = oldRow ? String(oldRow[6] || '') : '';

    if (data.imejData) {
      imej = muatNaikImejPost_(
        data.imejData,
        data.imejNama,
        data.imejJenis,
        admin.email,
        postId
      );
    }

    const sekarang = new Date();

    let tarikhTerbit = oldRow ? oldRow[9] : '';

    if (status === 'PUBLISHED' && !tarikhTerbit) {
      tarikhTerbit = sekarang;
    }

    if (status === 'DRAFT') {
      tarikhTerbit = '';
    }

    const rekod = [
      postId,
      tajuk,
      slug,
      kandungan,
      ringkasan,
      kategori,
      imej,
      status,
      oldRow ? oldRow[8] : false,
      tarikhTerbit,
      oldRow ? oldRow[10] : sekarang,
      sekarang,
      oldRow ? oldRow[12] : admin.nama,
      oldRow ? Number(oldRow[13] || 0) : 0
    ];

    if (rowNumber > 0) {
      sheet.getRange(rowNumber, 1, 1, rekod.length)
        .setValues([rekod]);
    } else {
      sheet.appendRow(rekod);
    }

    rekodAktiviti_(
      admin.email,
      id ? 'KEMAS KINI POST' : 'CIPTA POST',
      'ID: ' + postId + ' | Tajuk: ' + tajuk +
      ' | Status: ' + status
    );

    return {
      success: true,
      id: postId,
      message: status === 'PUBLISHED'
        ? 'Artikel berjaya diterbitkan.'
        : 'Artikel berjaya disimpan sebagai draf.'
    };

  } finally {
    lock.releaseLock();
  }
}


/************************************************************
 * JANA SLUG
 ************************************************************/

function janaSlugPost_(tajuk) {
  return String(tajuk || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .substring(0, 150);
}


/************************************************************
 * MUAT NAIK IMEJ POST KE GOOGLE DRIVE
 ************************************************************/

function muatNaikImejPost_(dataUrl, namaFail, jenisFail, email, postId) {

  const match = String(dataUrl || '').match(
    /^data:(image\/(?:jpeg|jpg|png|gif|webp));base64,([\s\S]+)$/i
  );

  if (!match) {
    throw new Error('Format imej tidak disokong. Gunakan JPG, PNG, GIF atau WEBP.');
  }

  const mimeType = match[1].toLowerCase();
  const base64 = match[2];

  const bytes = Utilities.base64Decode(base64);

  if (bytes.length > 5 * 1024 * 1024) {
    throw new Error('Saiz imej maksimum ialah 5 MB.');
  }

  const db = getDatabase_();
  const rootId = PropertiesService.getScriptProperties()
    .getProperty('PORTAL_FOLDER_ID');

  if (!rootId) {
    throw new Error('Folder utama portal tidak dijumpai.');
  }

  const root = DriveApp.getFolderById(rootId);
  const folders = root.getFoldersByName('POST_IMAGES');

  if (!folders.hasNext()) {
    throw new Error('Folder POST_IMAGES tidak dijumpai.');
  }

  const folder = folders.next();

  const safeName = String(namaFail || 'imej-post')
    .replace(/[^\w.\- ]/g, '_')
    .substring(0, 100);

  const blob = Utilities.newBlob(
    bytes,
    mimeType,
    Date.now() + '_' + safeName
  );

  const file = folder.createFile(blob);

  try {
    file.setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.VIEW
    );
  } catch (err) {
    file.setTrashed(true);
    throw new Error(
      'Google Drive tidak membenarkan perkongsian imej melalui pautan. ' +
      'Semak tetapan perkongsian organisasi atau akaun Google Workspace.'
    );
  }

  const url = 'https://drive.google.com/uc?export=view&id=' + file.getId();

  db.getSheetByName('MEDIA').appendRow([
    'MED-' + Utilities.getUuid().slice(0, 8).toUpperCase(),
    file.getName(),
    mimeType,
    url,
    url,
    bytes.length,
    'POST_IMAGES',
    email,
    new Date(),
    'AKTIF'
  ]);

  return url;
}


/************************************************************
 * PADAM POST
 ************************************************************/

function padamPost(id) {
  const admin = semakAksesAdmin_();

  id = String(id || '').trim();

  if (!id) {
    throw new Error('ID artikel tidak sah.');
  }

  const sheet = getDatabase_().getSheetByName('POSTS');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) {
      const tajuk = data[i][1] || '';

      sheet.deleteRow(i + 1);

      rekodAktiviti_(
        admin.email,
        'PADAM POST',
        'ID: ' + id + ' | Tajuk: ' + tajuk
      );

      return {
        success: true,
        message: 'Artikel berjaya dipadam.'
      };
    }
  }

  throw new Error('Artikel tidak dijumpai.');
}


/*******************************************************
 * FASA 4 - PAPARAN PORTAL AWAM
 * PORTAL UNIT KESENIAN BSKK
 *******************************************************/

/**
 * TUKAR URL GOOGLE DRIVE KEPADA URL PAPARAN IMEJ
 */
function janaUrlPaparanImej_(url) {
  url = String(url || '').trim();
  if (!url) return '';

  // Kenal pasti ID fail Google Drive
  const match = url.match(/[-\w]{25,}/);

  if (!match) return url;

  const fileId = match[0];

  return 'https://drive.google.com/thumbnail?id=' +
    encodeURIComponent(fileId) + '&sz=w1600';
}


/**
 * FASA 4 - DATA PORTAL AWAM
 */
function getPublicPortalData() {
  const db = getDatabase_();

  const categoryData = db.getSheetByName('CATEGORIES')
    .getDataRange().getValues();

  const postData = db.getSheetByName('POSTS')
    .getDataRange().getValues();

  const categories = categoryData.slice(1)
    .filter(row =>
      row[0] &&
      String(row[5] || 'AKTIF').toUpperCase() === 'AKTIF'
    )
    .map(row => ({
      id: String(row[0]),
      nama: String(row[1] || ''),
      slug: String(row[2] || ''),
      penerangan: String(row[3] || ''),
      susunan: Number(row[4] || 0)
    }))
    .sort((a, b) => a.susunan - b.susunan);

  const categoryMap = {};

  categories.forEach(category => {
    categoryMap[category.id] = category.nama;
  });

  // Kiraan komen APPROVED bagi setiap post (baca lajur ID_POST dan STATUS sahaja)
  const kiraKomen = {};
  const commentSheet = db.getSheetByName('COMMENTS');
  const lastComment = commentSheet.getLastRow();

  if (lastComment >= 2) {
    const cIds = commentSheet.getRange(2, 2, lastComment - 1, 1).getValues();
    const cStatus = commentSheet.getRange(2, 6, lastComment - 1, 1).getValues();

    for (let i = 0; i < cIds.length; i++) {
      if (String(cStatus[i][0]).toUpperCase() === 'APPROVED') {
        const k = String(cIds[i][0]);
        kiraKomen[k] = (kiraKomen[k] || 0) + 1;
      }
    }
  }

  const posts = postData.slice(1)
    .filter(row =>
      row[0] &&
      String(row[7] || '').toUpperCase() === 'PUBLISHED'
    )
    .map(row => ({
      id: String(row[0]),
      tajuk: String(row[1] || ''),
      slug: String(row[2] || ''),
      ringkasan: String(row[4] || ''),
      kategoriId: String(row[5] || ''),
      kategori: categoryMap[String(row[5] || '')] || 'Umum',
      imej: janaUrlPaparanImej_(row[6]),
      pinned: row[8] === true,
      tarikh: formatTarikhAwam_(row[9] || row[10]),
      tarikhSort: row[9] instanceof Date
        ? row[9].getTime()
        : (row[10] instanceof Date ? row[10].getTime() : 0),
      penulis: namaPenulisAwam_(row[12]),
      komen: kiraKomen[String(row[0])] || 0,
      like: Number(row[14] || 0),
      dislike: Number(row[15] || 0)
    }))
    .sort((a, b) => {
      if (a.pinned !== b.pinned) return a.pinned ? -1 : 1;
      return b.tarikhSort - a.tarikhSort;
    });

  return {
    namaPortal: PORTAL_CONFIG.NAMA,
    organisasi: PORTAL_CONFIG.ORGANISASI,
    kementerian: PORTAL_CONFIG.KEMENTERIAN,
    kategori: categories,
    artikel: posts
  };
}


/**
 * FASA 4 - DAPATKAN ARTIKEL MENGIKUT SLUG
 * (Versi laju: cari slug dan status dahulu, kemudian baca satu baris)
 */
function getPublicPostBySlug(slug) {
  slug = String(slug || '').trim();

  if (!slug) throw new Error('Artikel tidak dijumpai.');

  const db = getDatabase_();
  const sheet = db.getSheetByName('POSTS');
  const last = sheet.getLastRow();

  if (last < 2) {
    throw new Error('Artikel tidak dijumpai atau belum diterbitkan.');
  }

  const slugs = sheet.getRange(2, 3, last - 1, 1).getValues();
  const statuses = sheet.getRange(2, 8, last - 1, 1).getValues();

  let rowNumber = -1;

  for (let i = 0; i < slugs.length; i++) {
    if (
      String(slugs[i][0] || '') === slug &&
      String(statuses[i][0] || '').toUpperCase() === 'PUBLISHED'
    ) {
      rowNumber = i + 2;
      break;
    }
  }

  if (rowNumber === -1) {
    throw new Error('Artikel tidak dijumpai atau belum diterbitkan.');
  }

  const row = sheet.getRange(rowNumber, 1, 1, 14).getValues()[0];

  const jumlahPaparan = Number(row[13] || 0) + 1;
  sheet.getRange(rowNumber, 14).setValue(jumlahPaparan);

  let namaKategori = 'Umum';

  const categoryData = db.getSheetByName('CATEGORIES')
    .getDataRange().getValues();

  for (let i = 1; i < categoryData.length; i++) {
    if (
      String(categoryData[i][0]) === String(row[5] || '') &&
      String(categoryData[i][5] || 'AKTIF').toUpperCase() === 'AKTIF'
    ) {
      namaKategori = String(categoryData[i][1] || 'Umum');
      break;
    }
  }

  return {
    id: String(row[0]),
    tajuk: String(row[1] || ''),
    slug: String(row[2] || ''),
    kandungan: String(row[3] || ''),
    ringkasan: String(row[4] || ''),
    kategori: namaKategori,
    imej: janaUrlPaparanImej_(row[6]),
    tarikh: formatTarikhAwam_(row[9] || row[10]),
    penulis: namaPenulisAwam_(row[12])
  };
}


/**
 * Nama penulis untuk laman awam.
 * Jika nilai ialah alamat email, ia tidak dipaparkan (diganti 'Admin').
 */
function namaPenulisAwam_(nilai) {
  const nama = String(nilai || '').trim();
  if (!nama) return '';
  return nama.indexOf('@') > -1 ? 'Admin' : nama;
}


function formatTarikhAwam_(nilai) {
  if (!nilai) return '';

  const tarikh = nilai instanceof Date
    ? nilai
    : new Date(nilai);

  if (isNaN(tarikh.getTime())) return '';

  return Utilities.formatDate(
    tarikh,
    PORTAL_CONFIG.TIMEZONE,
    'dd MMMM yyyy'
  );
}


/************************************************************
 * FASA 5 - PENGURUSAN KATEGORI
 * PORTAL UNIT KESENIAN BSKK
 ************************************************************/

/**
 * SENARAI SEMUA KATEGORI UNTUK ADMIN
 */
function getAdminCategories() {

  semakAksesAdmin_();

  const sheet = getDatabase_().getSheetByName('CATEGORIES');
  const data = sheet.getDataRange().getValues();

  return data.slice(1)
    .filter(row => row[0])
    .map(row => ({
      id: String(row[0]),
      nama: String(row[1] || ''),
      slug: String(row[2] || ''),
      penerangan: String(row[3] || ''),
      susunan: Number(row[4] || 0),
      status: String(row[5] || 'AKTIF').toUpperCase(),
      tarikh: formatTarikh_(row[6])
    }))
    .sort((a, b) => a.susunan - b.susunan);
}


/**
 * TAMBAH / EDIT KATEGORI
 */
function simpanKategori(data) {

  const admin = semakAksesAdmin_();

  data = data || {};

  const id = String(data.id || '').trim();
  const nama = String(data.nama || '').trim();
  const penerangan = String(data.penerangan || '').trim();
  const susunan = Number(data.susunan || 0);
  const status = String(data.status || 'AKTIF')
    .toUpperCase().trim();

  if (!nama) {
    throw new Error('Sila masukkan nama kategori.');
  }

  if (nama.length > 100) {
    throw new Error('Nama kategori terlalu panjang.');
  }

  if (penerangan.length > 500) {
    throw new Error('Penerangan maksimum 500 aksara.');
  }

  if (!Number.isFinite(susunan) || susunan < 0) {
    throw new Error('Susunan kategori tidak sah.');
  }

  if (!['AKTIF', 'TIDAK AKTIF'].includes(status)) {
    throw new Error('Status kategori tidak sah.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {

    const sheet = getDatabase_().getSheetByName('CATEGORIES');
    const dataSheet = sheet.getDataRange().getValues();

    let rowNumber = -1;
    let oldRow = null;

    for (let i = 1; i < dataSheet.length; i++) {

      const namaSediaAda = String(dataSheet[i][1] || '')
        .trim().toLowerCase();

      if (
        namaSediaAda === nama.toLowerCase() &&
        String(dataSheet[i][0]) !== id
      ) {
        throw new Error('Nama kategori ini sudah digunakan.');
      }

      if (id && String(dataSheet[i][0]) === id) {
        rowNumber = i + 1;
        oldRow = dataSheet[i];
      }
    }

    if (id && rowNumber === -1) {
      throw new Error('Kategori yang hendak diedit tidak dijumpai.');
    }

    const kategoriId = id ||
      'KAT-' + Utilities.getUuid().slice(0, 8).toUpperCase();

    const slug = janaSlugPost_(nama);

    const rekod = [
      kategoriId,
      nama,
      slug,
      penerangan,
      susunan,
      status,
      oldRow ? oldRow[6] : new Date()
    ];

    if (rowNumber > 0) {
      sheet.getRange(rowNumber, 1, 1, rekod.length)
        .setValues([rekod]);
    } else {
      sheet.appendRow(rekod);
    }

    rekodAktiviti_(
      admin.email,
      id ? 'KEMAS KINI KATEGORI' : 'TAMBAH KATEGORI',
      'Kategori: ' + nama + ' | Status: ' + status
    );

    return {
      success: true,
      id: kategoriId,
      message: id
        ? 'Kategori berjaya dikemas kini.'
        : 'Kategori baharu berjaya ditambah.'
    };

  } finally {
    lock.releaseLock();
  }
}


/**
 * AKTIFKAN / NYAH AKTIFKAN KATEGORI
 */
function tukarStatusKategori(id, status) {

  const admin = semakAksesAdmin_();

  id = String(id || '').trim();
  status = String(status || '').toUpperCase().trim();

  if (!id) {
    throw new Error('ID kategori tidak sah.');
  }

  if (!['AKTIF', 'TIDAK AKTIF'].includes(status)) {
    throw new Error('Status kategori tidak sah.');
  }

  const sheet = getDatabase_().getSheetByName('CATEGORIES');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {

    if (String(data[i][0]) === id) {

      sheet.getRange(i + 1, 6).setValue(status);

      rekodAktiviti_(
        admin.email,
        'TUKAR STATUS KATEGORI',
        'Kategori: ' + data[i][1] + ' | Status: ' + status
      );

      return {
        success: true,
        message: 'Status kategori berjaya dikemas kini.'
      };
    }
  }

  throw new Error('Kategori tidak dijumpai.');
}


/**
 * PADAM KATEGORI
 * Tidak benarkan padam jika masih digunakan oleh artikel.
 */
function padamKategori(id) {

  const admin = semakAksesAdmin_();

  id = String(id || '').trim();

  if (!id) {
    throw new Error('ID kategori tidak sah.');
  }

  const db = getDatabase_();
  const categorySheet = db.getSheetByName('CATEGORIES');
  const postSheet = db.getSheetByName('POSTS');

  const categories = categorySheet.getDataRange().getValues();
  const posts = postSheet.getDataRange().getValues();

  const digunakan = posts.slice(1).some(row =>
    String(row[5] || '') === id
  );

  if (digunakan) {
    throw new Error(
      'Kategori ini masih digunakan oleh artikel. ' +
      'Sila nyahaktifkan kategori jika tidak mahu ia dipaparkan.'
    );
  }

  for (let i = 1; i < categories.length; i++) {

    if (String(categories[i][0]) === id) {

      const nama = categories[i][1];

      categorySheet.deleteRow(i + 1);

      rekodAktiviti_(
        admin.email,
        'PADAM KATEGORI',
        'Kategori: ' + nama
      );

      return {
        success: true,
        message: 'Kategori berjaya dipadam.'
      };
    }
  }

  throw new Error('Kategori tidak dijumpai.');
}


/***************************************************************
 * FASA 6 - MEDIA LIBRARY
 * PORTAL UNIT KESENIAN BSKK
 ***************************************************************/

/**
 * KONFIGURASI MEDIA
 */
const MEDIA_CONFIG = {
  SHEET_NAME: 'MEDIA',
  MAX_SIZE: 10 * 1024 * 1024,

  FOLDERS: {
    IMAGE: 'POST_IMAGES',
    ATTACHMENT: 'POST_ATTACHMENTS',
    DOCUMENT: 'DOCUMENTS',
    VIDEO: 'VIDEOS'
  }
};


/**
 * DAPATKAN SHEET MEDIA
 */
function getMediaSheet_() {

  const ss = getDatabase_();
  const sheet = ss.getSheetByName(MEDIA_CONFIG.SHEET_NAME);

  if (!sheet) {
    throw new Error('Sheet MEDIA tidak dijumpai.');
  }

  return sheet;
}


/**
 * FORMAT SAIZ FAIL
 */
function formatSaizMedia_(bytes) {

  if (!bytes || bytes <= 0) return '0 KB';

  if (bytes < 1024 * 1024) {
    return (bytes / 1024).toFixed(1) + ' KB';
  }

  return (bytes / (1024 * 1024)).toFixed(2) + ' MB';
}


/**
 * DAPATKAN FOLDER MEDIA
 */
function getMediaFolder_(folderName) {

  const properties = PropertiesService.getScriptProperties();
  const rootId = properties.getProperty('PORTAL_FOLDER_ID');

  if (!rootId) {
    throw new Error('PORTAL_FOLDER_ID tidak dijumpai dalam Script Properties.');
  }

  const root = DriveApp.getFolderById(rootId);
  const folders = root.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  }

  throw new Error('Folder ' + folderName + ' tidak dijumpai dalam Google Drive.');
}


/**
 * DAPATKAN SENARAI MEDIA UNTUK ADMIN
 */
function getAdminMedia() {

  verifyAdmin_();

  const sheet = getMediaSheet_();
  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    return [];
  }

  const headers = values[0].map(function(h) {
    return String(h).trim();
  });

  const result = [];

  for (let i = 1; i < values.length; i++) {

    const row = values[i];

    if (!row[0]) continue;

    const item = {};

    headers.forEach(function(header, index) {
      item[header] = row[index];
    });

    result.push({
      id: String(item.ID_MEDIA || ''),
      nama: String(item.NAMA_FAIL || ''),
      jenis: String(item.JENIS_FAIL || ''),
      url: String(item.URL_FAIL || ''),
      thumbnail: String(item.URL_THUMBNAIL || ''),
      saiz: String(item.SAIZ_FAIL || ''),
      folder: String(item.FOLDER || ''),
      pemuatNaik: String(item.DIMUAT_NAIK_OLEH || ''),
      tarikh: item.TARIKH_MUAT_NAIK
        ? Utilities.formatDate(
            new Date(item.TARIKH_MUAT_NAIK),
            Session.getScriptTimeZone() || 'Asia/Kuala_Lumpur',
            'dd/MM/yyyy HH:mm'
          )
        : '',
      status: String(item.STATUS || '')
    });

  }

  result.reverse();

  return result;
}


/**
 * MUAT NAIK MEDIA KE GOOGLE DRIVE
 *
 * data:
 * {
 *   name: nama fail,
 *   mimeType: jenis MIME,
 *   base64: kandungan base64,
 *   folder: IMAGE / ATTACHMENT / DOCUMENT / VIDEO
 * }
 */
function uploadPortalMedia(data) {

  verifyAdmin_();

  if (!data || !data.base64 || !data.name || !data.mimeType) {
    throw new Error('Maklumat fail tidak lengkap.');
  }

  const namaFail = String(data.name).trim();
  const mimeType = String(data.mimeType).toLowerCase();

  if (!namaFail) {
    throw new Error('Nama fail tidak sah.');
  }

  const folderType = String(data.folder || '').toUpperCase();

  const allowedFolders = [
    'IMAGE',
    'ATTACHMENT',
    'DOCUMENT',
    'VIDEO'
  ];

  if (allowedFolders.indexOf(folderType) === -1) {
    throw new Error('Kategori media tidak sah.');
  }

  const allowedMimeTypes = [
    'image/jpeg',
    'image/png',
    'image/gif',
    'image/webp',
    'image/svg+xml',
    'application/pdf',
    'application/msword',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'application/vnd.ms-excel',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-powerpoint',
    'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    'video/mp4',
    'video/webm'
  ];

  if (allowedMimeTypes.indexOf(mimeType) === -1) {
    throw new Error('Jenis fail ini tidak dibenarkan.');
  }

  const cleanBase64 = String(data.base64).replace(
    /^data:[^;]+;base64,/,
    ''
  );

  const bytes = Utilities.base64Decode(cleanBase64);

  if (bytes.length > MEDIA_CONFIG.MAX_SIZE) {
    throw new Error('Saiz fail melebihi had 10 MB.');
  }

  if (bytes.length === 0) {
    throw new Error('Fail kosong atau tidak sah.');
  }

  const folderName = MEDIA_CONFIG.FOLDERS[folderType];
  const folder = getMediaFolder_(folderName);

  const blob = Utilities.newBlob(
    bytes,
    mimeType,
    namaFail
  );

  const file = folder.createFile(blob);

  file.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );

  const fileId = file.getId();

  const url = file.getUrl();

  const thumbnail = mimeType.indexOf('image/') === 0
    ? 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w600'
    : '';

  const user = verifyAdmin_();

  const sheet = getMediaSheet_();

  const idMedia = 'MED-' + Utilities.getUuid().toUpperCase();

  const tarikh = new Date();

  const jenis = mimeType.indexOf('image/') === 0
    ? 'IMAGE'
    : mimeType.indexOf('video/') === 0
      ? 'VIDEO'
      : mimeType === 'application/pdf'
        ? 'PDF'
        : 'DOCUMENT';

  sheet.appendRow([
    idMedia,
    namaFail,
    jenis,
    url,
    thumbnail,
    formatSaizMedia_(bytes.length),
    folderName,
    user.email,
    tarikh,
    'AKTIF'
  ]);

  return {
    success: true,
    message: 'Fail berjaya dimuat naik.',
    id: idMedia,
    nama: namaFail,
    jenis: jenis,
    url: url,
    thumbnail: thumbnail,
    saiz: formatSaizMedia_(bytes.length),
    folder: folderName
  };

}


/**
 * PADAM MEDIA
 *
 * Rekod database dipadam dan fail Drive
 * dipindahkan ke Trash.
 */
function deletePortalMedia(idMedia) {

  verifyAdmin_();

  if (!idMedia) {
    throw new Error('ID media tidak sah.');
  }

  const sheet = getMediaSheet_();
  const values = sheet.getDataRange().getValues();

  let rowNumber = -1;
  let fileUrl = '';

  for (let i = 1; i < values.length; i++) {

    if (String(values[i][0]) === String(idMedia)) {

      rowNumber = i + 1;
      fileUrl = String(values[i][3] || '');

      break;
    }

  }

  if (rowNumber === -1) {
    throw new Error('Rekod media tidak dijumpai.');
  }

  const match = fileUrl.match(/[-\w]{25,}/);

  if (match && match[0]) {

    try {

      const file = DriveApp.getFileById(match[0]);

      file.setTrashed(true);

    } catch (error) {

      throw new Error(
        'Fail Drive tidak dapat dipadam. Rekod media dikekalkan. ' +
        error.message
      );

    }

  }

  sheet.deleteRow(rowNumber);

  return {
    success: true,
    message: 'Media berjaya dipadam.'
  };

}


/***************************************************************
 * FASA 7 - PENGURUSAN ADMIN / USERS
 * PORTAL UNIT KESENIAN BSKK
 ***************************************************************/

/**
 * KOMPATIBILITI FASA 6
 * Modul Media menggunakan verifyAdmin_().
 * Helper ini menyambungkan pengesahan kepada sistem sedia ada.
 */
function verifyAdmin_() {
  return semakAksesAdmin_();
}


/**
 * SENARAI ADMIN
 * Hanya SUPER ADMIN dibenarkan.
 */
function getAdminUsers() {

  semakSuperAdmin_();

  const sheet = getDatabase_().getSheetByName('ADMINS');
  const data = sheet.getDataRange().getValues();

  return data.slice(1)
    .filter(row => row[0])
    .map(row => ({
      id: String(row[0]),
      email: String(row[1] || ''),
      nama: String(row[2] || ''),
      peranan: String(row[3] || 'ADMIN').toUpperCase(),
      status: String(row[4] || 'TIDAK AKTIF').toUpperCase(),
      tarikh: formatTarikh_(row[5]),
      didaftarOleh: String(row[6] || ''),
      superAdminAwal: SUPER_ADMIN_AWAL.includes(
        String(row[1] || '').toLowerCase().trim()
      )
    }))
    .sort((a, b) => {
      if (a.peranan !== b.peranan) {
        return a.peranan === 'SUPER ADMIN' ? -1 : 1;
      }
      return a.email.localeCompare(b.email);
    });
}


/**
 * TAMBAH / KEMAS KINI ADMIN
 *
 * Admin biasa hanya boleh didaftarkan sebagai ADMIN.
 * Peranan SUPER ADMIN tidak boleh diberikan melalui borang ini.
 */
function simpanAdminUser(data) {

  const pengguna = semakSuperAdmin_();

  data = data || {};

  const id = String(data.id || '').trim();
  const email = String(data.email || '').trim().toLowerCase();
  const nama = String(data.nama || '').trim();

  if (!email) {
    throw new Error('Sila masukkan alamat email admin.');
  }

  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error('Format email tidak sah.');
  }

  if (!email.endsWith('@moe-dl.edu.my')) {
    throw new Error('Hanya akaun DELIMa @moe-dl.edu.my dibenarkan.');
  }

  if (!nama) {
    throw new Error('Sila masukkan nama admin.');
  }

  if (nama.length > 150) {
    throw new Error('Nama terlalu panjang.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {

    const sheet = getDatabase_().getSheetByName('ADMINS');
    const dataSheet = sheet.getDataRange().getValues();

    let rowNumber = -1;
    let oldRow = null;

    for (let i = 1; i < dataSheet.length; i++) {

      const rowEmail = String(dataSheet[i][1] || '')
        .trim().toLowerCase();

      const rowId = String(dataSheet[i][0] || '');

      if (rowEmail === email && rowId !== id) {
        throw new Error('Email ini sudah berdaftar sebagai admin.');
      }

      if (id && rowId === id) {
        rowNumber = i + 1;
        oldRow = dataSheet[i];
      }
    }

    if (id && !oldRow) {
      throw new Error('Rekod admin tidak dijumpai.');
    }

    if (oldRow) {

      const emailLama = String(oldRow[1] || '')
        .trim().toLowerCase();

      if (
        SUPER_ADMIN_AWAL.includes(emailLama) ||
        String(oldRow[3]).toUpperCase() === 'SUPER ADMIN'
      ) {
        throw new Error(
          'Akaun SUPER ADMIN dilindungi dan tidak boleh diubah melalui borang ini.'
        );
      }

      const rekod = [
        oldRow[0],
        email,
        nama,
        'ADMIN',
        oldRow[4],
        oldRow[5],
        oldRow[6]
      ];

      sheet.getRange(rowNumber, 1, 1, 7).setValues([rekod]);

    } else {

      const idAdmin = 'ADM-' +
        Utilities.getUuid().slice(0, 8).toUpperCase();

      sheet.appendRow([
        idAdmin,
        email,
        nama,
        'ADMIN',
        'AKTIF',
        new Date(),
        pengguna.email
      ]);
    }

    rekodAktiviti_(
      pengguna.email,
      id ? 'KEMAS KINI ADMIN' : 'TAMBAH ADMIN',
      'Email: ' + email + ' | Nama: ' + nama
    );

    return {
      success: true,
      message: id
        ? 'Maklumat admin berjaya dikemas kini.'
        : 'Admin baharu berjaya didaftarkan.'
    };

  } finally {
    lock.releaseLock();
  }
}


/**
 * TUKAR STATUS ADMIN
 * Tidak membenarkan SUPER ADMIN awal dinyahaktifkan.
 * Tidak membenarkan pengguna menyahaktifkan akaun sendiri.
 */
function tukarStatusAdminUser(id, status) {

  const pengguna = semakSuperAdmin_();

  id = String(id || '').trim();
  status = String(status || '').toUpperCase().trim();

  if (!id) {
    throw new Error('ID admin tidak sah.');
  }

  if (!['AKTIF', 'TIDAK AKTIF'].includes(status)) {
    throw new Error('Status admin tidak sah.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {

    const sheet = getDatabase_().getSheetByName('ADMINS');
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {

      if (String(data[i][0]) !== id) continue;

      const email = String(data[i][1] || '').toLowerCase().trim();
      const peranan = String(data[i][3] || '').toUpperCase();

      if (
        SUPER_ADMIN_AWAL.includes(email) ||
        peranan === 'SUPER ADMIN'
      ) {
        throw new Error('Status akaun SUPER ADMIN dilindungi.');
      }

      if (email === pengguna.email && status === 'TIDAK AKTIF') {
        throw new Error('Anda tidak boleh menyahaktifkan akaun sendiri.');
      }

      sheet.getRange(i + 1, 5).setValue(status);

      rekodAktiviti_(
        pengguna.email,
        'TUKAR STATUS ADMIN',
        'Email: ' + email + ' | Status: ' + status
      );

      return {
        success: true,
        message: 'Status admin berjaya dikemas kini.'
      };
    }

    throw new Error('Admin tidak dijumpai.');

  } finally {
    lock.releaseLock();
  }
}


/**
 * PADAM ADMIN
 * SUPER ADMIN awal tidak boleh dipadam.
 */
function padamAdminUser(id) {

  const pengguna = semakSuperAdmin_();

  id = String(id || '').trim();

  if (!id) {
    throw new Error('ID admin tidak sah.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {

    const sheet = getDatabase_().getSheetByName('ADMINS');
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {

      if (String(data[i][0]) !== id) continue;

      const email = String(data[i][1] || '').toLowerCase().trim();
      const peranan = String(data[i][3] || '').toUpperCase();

      if (
        SUPER_ADMIN_AWAL.includes(email) ||
        peranan === 'SUPER ADMIN'
      ) {
        throw new Error('Akaun SUPER ADMIN dilindungi dan tidak boleh dipadam.');
      }

      if (email === pengguna.email) {
        throw new Error('Anda tidak boleh memadam akaun sendiri.');
      }

      const nama = String(data[i][2] || '');

      sheet.deleteRow(i + 1);

      rekodAktiviti_(
        pengguna.email,
        'PADAM ADMIN',
        'Email: ' + email + ' | Nama: ' + nama
      );

      return {
        success: true,
        message: 'Admin berjaya dipadam.'
      };
    }

    throw new Error('Admin tidak dijumpai.');

  } finally {
    lock.releaseLock();
  }
}
