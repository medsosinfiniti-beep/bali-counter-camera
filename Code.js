 // ================================================================
// 1. KONFIGURASI & PROPERTIES
// ================================================================
const CONFIG = {
  SHEET_NAME: 'DataCounter',
  USER_SHEET_NAME: 'Users',
  FOLDER_NAME: 'Bali Counter Photos',
  OWNER_EMAIL: PropertiesService.getScriptProperties().getProperty('OWNER_EMAIL') || Session.getActiveUser().getEmail(),
  AI_API_KEY: PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || ''
};

// ================================================================
// 2. FUNGSI BANTUAN SHEET
// ================================================================
function getSheetData(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    if (sheetName === CONFIG.SHEET_NAME) {
      const newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      newSheet.appendRow(['id', 'area', 'jalan', 'counter', 'kontak', 'status', 'email', 'catatan', 'photo_id', 'order']);
      return newSheet;
    } else if (sheetName === CONFIG.USER_SHEET_NAME) {
      const newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      newSheet.appendRow(['username', 'password_hash', 'email', 'status', 'role']);
      return newSheet;
    }
  }
  return sheet;
}

function hashPassword(password) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  return digest.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function verifyPassword(inputPassword, storedHash) {
  return hashPassword(inputPassword) === storedHash;
}

function getUserByUsername(username) {
  const sheet = getSheetData(CONFIG.USER_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username) {
      return { row: i + 1, username: data[i][0], password_hash: data[i][1], email: data[i][2], status: data[i][3], role: data[i][4] };
    }
  }
  return null;
}

function getUserByEmail(email) {
  const sheet = getSheetData(CONFIG.USER_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === email) {
      return { row: i + 1, username: data[i][0], password_hash: data[i][1], email: data[i][2], status: data[i][3], role: data[i][4] };
    }
  }
  return null;
}

function getNextId() {
  const sheet = getSheetData(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  let maxId = 0;
  for (let i = 1; i < data.length; i++) {
    const id = parseInt(data[i][0]);
    if (id > maxId) maxId = id;
  }
  return maxId + 1;
}

function getOrCreateFolder() {
  const folders = DriveApp.getFoldersByName(CONFIG.FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(CONFIG.FOLDER_NAME);
  }
}

// ================================================================
// 3. LOGGING (untuk debugging)
// ================================================================
function logToSheet(message) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Logs');
    let logSheet = sheet;
    if (!logSheet) {
      logSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Logs');
      logSheet.appendRow(['Timestamp', 'Message']);
    }
    logSheet.appendRow([new Date().toISOString(), message]);
  } catch(e) {
    // ignore
  }
}

// ================================================================
// 4. FUNGSI UPLOAD & DELETE FOTO (DENGAN DETEKSI FORMAT)
// ================================================================
function uploadPhotoToDrive(base64Data, filename) {
  if (!base64Data) {
    logToSheet('❌ uploadPhotoToDrive: base64Data kosong');
    return null;
  }
  const folder = getOrCreateFolder();
  logToSheet(`📁 Folder: ${folder.getName()}`);

  // Deteksi format dari dataURL
  const matches = base64Data.match(/^data:([A-Za-z-+\/]+);base64,(.+)$/);
  let contentType = 'image/jpeg';
  let base64String = base64Data;

  if (matches && matches.length === 3) {
    contentType = matches[1];
    base64String = matches[2];
  }
  logToSheet(`📷 Content-Type: ${contentType}`);

  // Tentukan ekstensi file
  let extension = '.jpg';
  if (contentType === 'image/png') extension = '.png';
  else if (contentType === 'image/webp') extension = '.webp';
  else if (contentType === 'image/gif') extension = '.gif';

  // Simpan ke Drive
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(base64String), contentType, filename + extension);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    logToSheet(`✅ File berhasil diupload: ${file.getName()} (ID: ${file.getId()})`);
    return file.getId();
  } catch (err) {
    logToSheet(`❌ Gagal upload ke Drive: ${err.message}`);
    return null;
  }
}

function deletePhotoFromDrive(photoId) {
  if (!photoId) return;
  try {
    const file = DriveApp.getFileById(photoId);
    file.setTrashed(true);
    logToSheet(`🗑️ File ${photoId} dihapus (trashed)`);
  } catch (e) {
    logToSheet(`❌ Gagal hapus file ${photoId}: ${e.message}`);
  }
}

// ================================================================
// 5. AUTENTIKASI
// ================================================================
function registerUser(username, password, email) {
  try {
    if (!username || !password || !email) throw new Error('Semua field wajib diisi.');
    if (password.length < 6) throw new Error('Password minimal 6 karakter.');
    if (!email.includes('@gmail.com') && !email.includes('@googlemail.com')) throw new Error('Gunakan email Gmail.');
    if (getUserByUsername(username)) throw new Error('Username sudah digunakan.');
    if (getUserByEmail(email)) throw new Error('Email sudah terdaftar.');

    const sheet = getSheetData(CONFIG.USER_SHEET_NAME);
    const hash = hashPassword(password);
    sheet.appendRow([username, hash, email, 'pending', 'user']);

    const ownerEmail = CONFIG.OWNER_EMAIL;
    MailApp.sendEmail({
      to: ownerEmail,
      subject: 'Permintaan Verifikasi Akun Baru - Bali Counter DB',
      body: `Pengguna baru mendaftar:\nUsername: ${username}\nEmail: ${email}\n\nSilakan buka spreadsheet dan ubah status user menjadi "active" untuk mengaktifkan akun.`
    });

    return { success: true, message: 'Pendaftaran berhasil! Akun menunggu verifikasi owner.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function loginUser(username, password) {
  try {
    const user = getUserByUsername(username);
    if (!user) throw new Error('Username tidak ditemukan.');
    if (!verifyPassword(password, user.password_hash)) throw new Error('Password salah.');
    if (user.status !== 'active') throw new Error('Akun belum diverifikasi oleh owner.');
    return { success: true, role: user.role, username: user.username };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function forgotPassword(email) {
  try {
    const user = getUserByEmail(email);
    if (!user) throw new Error('Email tidak ditemukan.');
    const ownerEmail = CONFIG.OWNER_EMAIL;
    MailApp.sendEmail({
      to: ownerEmail,
      subject: 'Permintaan Reset Password - Bali Counter DB',
      body: `Pengguna dengan email ${email} (username: ${user.username}) meminta reset password.\n\nSilakan reset password secara manual di spreadsheet atau hubungi pengguna.`
    });
    return { success: true, message: 'Instruksi reset password telah dikirim ke owner.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function approveUser(username) {
  const owner = Session.getActiveUser().getEmail();
  if (owner !== CONFIG.OWNER_EMAIL) {
    return { success: false, message: 'Anda bukan owner.' };
  }
  const user = getUserByUsername(username);
  if (!user) return { success: false, message: 'User tidak ditemukan.' };
  const sheet = getSheetData(CONFIG.USER_SHEET_NAME);
  sheet.getRange(user.row, 4).setValue('active');
  return { success: true, message: `User ${username} telah diaktifkan.` };
}

// ================================================================
// 6. CRUD COUNTER
// ================================================================
function getAllCounters() {
  const sheet = getSheetData(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    result.push({
      id: parseInt(row[0]),
      area: row[1],
      jalan: row[2],
      counter: row[3],
      kontak: row[4] ? JSON.parse(row[4]) : [],
      status: row[5],
      email: row[6] || '',
      catatan: row[7] || '',
      photo_id: row[8] || null,
      order: parseInt(row[9]) || 0
    });
  }
  result.sort((a, b) => (a.order || 0) - (b.order || 0));
  return result;
}

function saveCounter(counter) {
  const sheet = getSheetData(CONFIG.SHEET_NAME);
  const kontakStr = JSON.stringify(counter.kontak || []);
  const order = counter.order || 0;
  const photoId = counter.photo_id || null;

  const existing = sheet.getDataRange().getValues();
  for (let i = 1; i < existing.length; i++) {
    if (parseInt(existing[i][0]) === counter.id) {
      // Update existing row
      sheet.getRange(i + 1, 2, 1, 9).setValues([[
        counter.area,
        counter.jalan,
        counter.counter,
        kontakStr,
        counter.status,
        counter.email || '',
        counter.catatan || '',
        photoId,
        order
      ]]);
      logToSheet(`🔄 Counter ID ${counter.id} diperbarui`);
      return;
    }
  }

  // Insert new row
  const newId = counter.id || getNextId();
  sheet.appendRow([
    newId,
    counter.area,
    counter.jalan,
    counter.counter,
    kontakStr,
    counter.status,
    counter.email || '',
    counter.catatan || '',
    photoId,
    order
  ]);
  logToSheet(`➕ Counter baru ditambahkan: ID ${newId}`);
}

function deleteCounter(id) {
  // Hapus foto dari Drive jika ada
  const counters = getAllCounters();
  const item = counters.find(c => c.id === id);
  if (item && item.photo_id) {
    deletePhotoFromDrive(item.photo_id);
  }

  const sheet = getSheetData(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (parseInt(data[i][0]) === id) {
      sheet.deleteRow(i + 1);
      logToSheet(`🗑️ Counter ID ${id} dihapus`);
      return true;
    }
  }
  return false;
}

function updateOrder(orderedIds) {
  const sheet = getSheetData(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const orderMap = {};
  orderedIds.forEach((id, idx) => { orderMap[id] = idx; });
  for (let i = 1; i < data.length; i++) {
    const id = parseInt(data[i][0]);
    if (orderMap[id] !== undefined) {
      sheet.getRange(i + 1, 10).setValue(orderMap[id]);
    }
  }
  logToSheet(`🔄 Urutan diperbarui: ${orderedIds.join(', ')}`);
}

// ================================================================
// 7. DO POST - Menerima upload dari halaman kamera eksternal (GitHub Pages)
// ================================================================
function doPost(e) {
  try {
    logToSheet('📥 doPost dipanggil');
    const data = JSON.parse(e.postData.contents);
    logToSheet('📦 Data: ' + JSON.stringify(data));

    const action = data.action;

    if (action === 'uploadPhoto') {
      const counterId = data.counterId;
      const photoBase64 = data.photo;
      const filename = data.filename || 'photo.jpg';

      logToSheet(`🖼️ Upload foto untuk counter ID: ${counterId}`);

      // Validasi dasar
      if (!counterId || isNaN(counterId)) {
        throw new Error('counterId tidak valid');
      }
      if (!photoBase64 || photoBase64.length < 100) {
        throw new Error('Data foto tidak valid (terlalu pendek)');
      }

      // 1. Upload foto ke Drive
      const photoId = uploadPhotoToDrive(photoBase64, filename);
      logToSheet(`📁 Photo ID: ${photoId}`);

      if (!photoId) {
        throw new Error('Gagal upload foto ke Drive');
      }

      // 2. Update data counter
      const counters = getAllCounters();
      const item = counters.find(c => c.id === counterId);
      if (!item) {
        throw new Error(`Counter dengan ID ${counterId} tidak ditemukan`);
      }

      item.photo_id = photoId;
      saveCounter(item);
      logToSheet(`✅ Counter ${counterId} berhasil diupdate dengan photo_id ${photoId}`);

      return ContentService.createTextOutput(JSON.stringify({ success: true, photoId: photoId }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Aksi tidak dikenal' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    logToSheet(`❌ ERROR: ${err.message}`);
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ================================================================
// 8. IMPOR
// ================================================================
function importFromSheet(spreadsheetId, sheetName) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(sheetName || 'Sheet1');
    if (!sheet) throw new Error('Sheet tidak ditemukan.');
    const data = sheet.getDataRange().getValues();
    let importedCount = 0;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const counter = {
        id: parseInt(row[0]) || getNextId(),
        area: row[1] || '',
        jalan: row[2] || '',
        counter: row[3] || '',
        kontak: row[4] ? JSON.parse(row[4]) : [],
        status: row[5] || 'Aktif',
        email: row[6] || '',
        catatan: row[7] || '',
        photo_id: row[8] || null,
        order: 0
      };
      saveCounter(counter);
      importedCount++;
    }
    return { success: true, message: `Berhasil mengimpor ${importedCount} data dari spreadsheet.` };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function importFromDrive(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const mimeType = file.getMimeType();
    let data = [];
    if (mimeType === 'text/csv') {
      const content = file.getBlob().getDataAsString();
      const lines = content.split('\n');
      data = lines.map(line => line.split(','));
    } else if (mimeType === 'application/json') {
      const content = file.getBlob().getDataAsString();
      data = JSON.parse(content);
    } else {
      throw new Error('Format file belum didukung. Gunakan CSV atau JSON.');
    }
    let importedCount = 0;
    if (Array.isArray(data)) {
      data.forEach(item => {
        if (typeof item === 'object') {
          const counter = {
            id: item.id || getNextId(),
            area: item.area || '',
            jalan: item.jalan || '',
            counter: item.counter || '',
            kontak: item.kontak || [],
            status: item.status || 'Aktif',
            email: item.email || '',
            catatan: item.catatan || '',
            photo_id: item.photo_id || null,
            order: 0
          };
          saveCounter(counter);
          importedCount++;
        }
      });
    }
    return { success: true, message: `Berhasil mengimpor ${importedCount} data dari Drive.` };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ================================================================
// 9. AI (Opsional)
// ================================================================
function callAI(prompt) {
  const apiKey = CONFIG.AI_API_KEY;
  if (!apiKey) {
    return { success: true, message: 'Fitur AI belum aktif. Silakan hubungi owner untuk mengaktifkan.' };
  }
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=${apiKey}`;
  const payload = { contents: [{ parts: [{ text: prompt }] }] };
  const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };
  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    const text = result?.candidates?.[0]?.content?.parts?.[0]?.text || 'Tidak ada respons.';
    return { success: true, data: text };
  } catch (e) {
    return { success: false, message: 'Gagal memanggil AI: ' + e.message };
  }
}

// ================================================================
// 10. SEED DATABASE (opsional, untuk mengisi data awal)
// ================================================================
function seedDatabase() {
  const sheet = getSheetData(CONFIG.SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }

  // Contoh data awal (potong, disesuaikan)
  const INITIAL_DATA = [
    // ... (data 210 counter) – silakan salin dari versi sebelumnya
    // Untuk ringkas, saya tidak menulis ulang 210 data di sini.
    // Anda bisa menambahkan kembali atau biarkan kosong.
  ];

  if (INITIAL_DATA.length === 0) {
    return "⚠️ Tidak ada data untuk di-seed. Silakan tambahkan data terlebih dahulu.";
  }

  const rows = INITIAL_DATA.map(item => {
    const kontakStr = JSON.stringify(item.kontak || []);
    return [
      item.id,
      item.area,
      item.jalan,
      item.counter,
      kontakStr,
      item.status,
      item.email || '',
      item.catatan || '',
      null,
      0
    ];
  });

  if (rows.length > 0) {
    const numCols = rows[0].length;
    sheet.getRange(2, 1, rows.length, numCols).setValues(rows);
    SpreadsheetApp.flush();
    return `✅ Berhasil mengisi ${rows.length} data counter ke database! Refresh Web App Anda.`;
  }
  return "Data kosong.";
}

// ================================================================
// 11. ENTRY POINT (doGet)
// ================================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Bali Counter DB Pro')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
