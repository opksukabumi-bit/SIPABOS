// ==========================================
// KONFIGURASI NAMA SHEET
// ==========================================
var SHEET_USER     = "User";
var SHEET_LOG      = "Log_Unggah";
var SHEET_SETTINGS = "Settings";
var SHEET_LOCK     = "Kunci_Periode";
var SHEET_PROFILE  = "Profil_Sekolah";
var SHEET_BUDGET   = "Anggaran_Tahunan";

// Waktu tunggu LockService (ms) — 30 detik
var LOCK_WAIT_MS = 30000;

// ==========================================
// HELPER: Ambil lock script-level
// ==========================================
function _getLock() {
  return LockService.getScriptLock();
}

// ==========================================
// 1. AUTENTIKASI
// ==========================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle("SIPABOS - Portal Resmi")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function checkLogin(username, password) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getScriptProperties();
  var admUser = props.getProperty("admin_user") || "admin";
  var admPass = props.getProperty("admin_pass") || "admin123";

  if (username.trim() === admUser && password.trim() === admPass) {
    return { status: "success", role: "admin", name: "Administrator Pusat" };
  }

  var sheetUser = ss.getSheetByName(SHEET_USER);
  if (!sheetUser) return { status: "error", message: "Sheet 'User' tidak ditemukan!" };

  var lastRow = sheetUser.getLastRow();
  if (lastRow < 2) return { status: "error", message: "Username atau Password salah!" };

  var data = sheetUser.getRange(1, 1, lastRow, sheetUser.getLastColumn()).getValues();
  for (var i = 1; i < data.length; i++) {
    var u = data[i][0] ? data[i][0].toString().trim() : "";
    var p = data[i][1] ? data[i][1].toString().trim() : "";
    var n = data[i][2] ? data[i][2].toString().trim() : "";
    if (username.trim() === u && password.trim() === p) {
      return { status: "success", role: "user", name: n };
    }
  }
  return { status: "error", message: "Username atau Password salah!" };
}

function getAppConfig() {
  var props = PropertiesService.getScriptProperties();
  return {
    name: props.getProperty("app_name") || "SIPABOS",
    logo: props.getProperty("app_logo") || "",
    theme: props.getProperty("app_theme") || "dark"
  };
}

// ==========================================
// 1b. PROFIL SEKOLAH
// ==========================================
function _ensureProfileSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_PROFILE);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PROFILE);
    sheet.appendRow(["Nama Sekolah", "NPSN", "Kepala Sekolah", "Alamat", "Telepon", "Email", "Kota/Kabupaten", "Propinsi"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getSchoolProfile(schoolName) {
  var sheet = _ensureProfileSheet();
  var data = sheet.getRange(1, 1, sheet.getLastRow(), 8).getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === schoolName.trim()) {
      return {
        name: data[i][0] ? data[i][0].toString().trim() : "",
        npsn: data[i][1] ? data[i][1].toString().trim() : "",
        kepala: data[i][2] ? data[i][2].toString().trim() : "",
        alamat: data[i][3] ? data[i][3].toString().trim() : "",
        telepon: data[i][4] ? data[i][4].toString().trim() : "",
        email: data[i][5] ? data[i][5].toString().trim() : "",
        kota: data[i][6] ? data[i][6].toString().trim() : "",
        propinsi: data[i][7] ? data[i][7].toString().trim() : ""
      };
    }
  }
  return null;
}

function updateSchoolProfile(schoolName, npsn, kepala, alamat, telepon, email, kota, propinsi) {
  var sheet = _ensureProfileSheet();
  var lock = _getLock();
  try { lock.waitLock(LOCK_WAIT_MS); }
  catch (e) { return { status: 'error', message: 'Sistem sibuk.' }; }
  
  try {
    var data = sheet.getRange(1, 1, sheet.getLastRow(), 8).getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === schoolName.trim()) {
        sheet.getRange(i + 1, 1, 1, 8).setValues([[schoolName, npsn, kepala, alamat, telepon, email, kota, propinsi]]);
        return { status: 'success', message: 'Profil diperbarui.' };
      }
    }
    // Tambah row baru
    sheet.appendRow([schoolName, npsn, kepala, alamat, telepon, email, kota, propinsi]);
    return { status: 'success', message: 'Profil ditambahkan.' };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 1c. ANGGARAN TAHUNAN
// ==========================================
function _ensureBudgetSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_BUDGET);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_BUDGET);
    sheet.appendRow(["Tahun", "Nama Sekolah", "Total Anggaran", "Sumber Dana", "Catatan"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getBudgetByYear(schoolName, year) {
  var sheet = _ensureBudgetSheet();
  var data = sheet.getRange(1, 1, sheet.getLastRow(), 5).getValues();
  var budgets = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][1].toString().trim() === schoolName.trim() && 
        data[i][0] && parseInt(data[i][0]) === year) {
      budgets.push({
        tahun: parseInt(data[i][0]),
        sekolah: data[i][1].toString().trim(),
        jumlah: data[i][2],
        sumber: data[i][3] ? data[i][3].toString().trim() : "",
        catatan: data[i][4] ? data[i][4].toString().trim() : ""
      });
    }
  }
  return budgets;
}

function updateBudget(year, schoolName, totalBudget, sumberDana, catatan) {
  var sheet = _ensureBudgetSheet();
  var lock = _getLock();
  try { lock.waitLock(LOCK_WAIT_MS); }
  catch (e) { return { status: 'error', message: 'Sistem sibuk.' }; }
  
  try {
    var data = sheet.getRange(1, 1, sheet.getLastRow(), 5).getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && parseInt(data[i][0]) === year && 
          data[i][1] && data[i][1].toString().trim() === schoolName.trim()) {
        sheet.getRange(i + 1, 1, 1, 5).setValues([[year, schoolName, totalBudget, sumberDana, catatan]]);
        return { status: 'success', message: 'Anggaran diperbarui.' };
      }
    }
    sheet.appendRow([year, schoolName, totalBudget, sumberDana, catatan]);
    return { status: 'success', message: 'Anggaran ditambahkan.' };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 2. UPLOAD BERKAS USER — Status default: "Pending"
// ==========================================

function uploadFilesBulkToDrive(filesDataArray, schoolName) {
  return _doUploadBulk(filesDataArray, schoolName, "Pending");
}

// ==========================================
// 2b. UPLOAD BERKAS OLEH ADMIN — status langsung "Verified"
// ==========================================

function uploadFilesBulkToDriveAdmin(filesDataArray, schoolName) {
  return _doUploadBulk(filesDataArray, schoolName, "Verified");
}

// ==========================================
// HELPER UPLOAD — [FITUR #4: Ganti file Pending/Revision]
// ==========================================

function _doUploadBulk(filesDataArray, schoolName, initialStatus) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var sheetLog  = ss.getSheetByName(SHEET_LOG);
  var sheetUser = ss.getSheetByName(SHEET_USER);
  if (!sheetLog || !sheetUser) return { status: 'error', message: 'Sheet tidak ditemukan.' };

  var mainFolderId = "";
  var dataUser = sheetUser.getDataRange().getValues();
  for (var i = 1; i < dataUser.length; i++) {
    if (dataUser[i][2] && dataUser[i][2].toString().trim() === schoolName.trim()) {
      mainFolderId = dataUser[i][3] ? dataUser[i][3].toString().trim() : "";
      break;
    }
  }

  var mainFolder = null;
  if (mainFolderId) {
    try { mainFolder = DriveApp.getFolderById(mainFolderId); mainFolder.getName(); }
    catch (e) { mainFolder = null; }
  }
  if (!mainFolder) {
    try {
      var root  = DriveApp.getFoldersByName("SIPABOS_Upload");
      var rootF = root.hasNext() ? root.next() : DriveApp.createFolder("SIPABOS_Upload");
      var schF  = rootF.getFoldersByName(schoolName);
      mainFolder = schF.hasNext() ? schF.next() : rootF.createFolder(schoolName);
    } catch (e2) {
      return { status: 'error', message: 'Akses Drive gagal: ' + e2.toString() };
    }
  }

  var lock = _getLock();
  try { lock.waitLock(LOCK_WAIT_MS); }
  catch (e) { return { status: 'error', message: 'Sistem sedang digunakan proses lain, coba lagi sebentar.' }; }

  try {
    var lastRow = sheetLog.getLastRow();
    var logData = lastRow >= 2
      ? sheetLog.getRange(2, 1, lastRow - 1, 8).getValues()
      : [];

    var suksesCount = 0, skipCount = 0, replaceCount = 0;

    for (var j = 0; j < filesDataArray.length; j++) {
      var fd    = filesDataArray[j];
      var jenis = fd.jenisLaporan.toString().trim();
      var tw    = fd.tw.toString().trim();
      var fname = fd.fileName.toString().trim();

      var replaceRowIdx = -1;
      for (var r = 0; r < logData.length; r++) {
        var rSch    = logData[r][1] ? logData[r][1].toString().trim() : '';
        var rJenis  = logData[r][2] ? logData[r][2].toString().trim() : '';
        var rTw     = logData[r][3] ? logData[r][3].toString().trim() : '';
        var rFname  = logData[r][4] ? logData[r][4].toString().trim() : '';
        var rStatus = logData[r][6] ? logData[r][6].toString().trim() : 'Pending';

        if (rSch.toLowerCase()   !== schoolName.toLowerCase()) continue;
        if (rJenis.toLowerCase() !== jenis.toLowerCase())      continue;
        if (rTw.toLowerCase()    !== tw.toLowerCase())         continue;
        if (rFname.toLowerCase() !== fname.toLowerCase())      continue;

        if (rStatus === 'Rejected' || rStatus === 'Pending' || rStatus === 'Revision Needed') {
          replaceRowIdx = r;
          break;
        }
      }

      var sameFileVerified = logData.some(function(row) {
        return row[1] && row[1].toString().trim().toLowerCase() === schoolName.toLowerCase()
          && row[2] && row[2].toString().trim().toLowerCase() === jenis.toLowerCase()
          && row[3] && row[3].toString().trim().toLowerCase() === tw.toLowerCase()
          && row[4] && row[4].toString().trim().toLowerCase() === fname.toLowerCase()
          && row[6] && row[6].toString().trim() === 'Verified';
      });
      if (sameFileVerified && initialStatus !== 'Verified') {
        skipCount++;
        continue;
      }

      var jenisF, twF;
      var jIt = mainFolder.getFoldersByName(jenis);
      jenisF  = jIt.hasNext() ? jIt.next() : mainFolder.createFolder(jenis);
      var tIt = jenisF.getFoldersByName(tw);
      twF     = tIt.hasNext() ? tIt.next() : jenisF.createFolder(tw);

      var mime   = fd.fileType && fd.fileType !== '' ? fd.fileType : 'application/octet-stream';
      var blob   = Utilities.newBlob(fd.bytes, mime, fname);
      var fileDr = twF.createFile(blob);
      try { fileDr.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(se) {}

      var fileId = fileDr.getId();
      var waktu  = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");

      if (replaceRowIdx !== -1) {
        var sheetRowIdx = replaceRowIdx + 2;
        var oldFileId   = logData[replaceRowIdx][5] ? logData[replaceRowIdx][5].toString().trim() : '';
        if (oldFileId) {
          try { DriveApp.getFileById(oldFileId).setTrashed(true); } catch(de) {}
        }
        sheetLog.getRange(sheetRowIdx, 1, 1, 8).setValues([[
          waktu, schoolName, jenis, tw, fname, fileId, initialStatus, ''
        ]]);
        logData[replaceRowIdx] = [waktu, schoolName, jenis, tw, fname, fileId, initialStatus, ''];
        replaceCount++;
      } else {
        sheetLog.appendRow([waktu, schoolName, jenis, tw, fname, fileId, initialStatus, '']);
        logData.push([waktu, schoolName, jenis, tw, fname, fileId, initialStatus, '']);
        suksesCount++;
      }
    }

    var msg = '';
    if (replaceCount > 0) msg += replaceCount + ' berkas berhasil diganti (menggantikan yang sebelumnya). ';
    if (suksesCount  > 0) msg += suksesCount  + ' berkas baru berhasil disimpan. ';
    if (skipCount    > 0) msg += skipCount     + ' berkas dilewati (sudah terverifikasi).';
    return { status: 'success', message: msg.trim() || 'Selesai.' };

  } catch (e) {
    return { status: 'error', message: 'Gagal: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 3. DASHBOARD & MANAGEMENT DATA
// ==========================================

function getSchoolDashboardData(schName) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheetLog = ss.getSheetByName(SHEET_LOG);
  if (!sheetLog || sheetLog.getLastRow() < 2) return { logs: [], deadlines: getDeadlines() };

  var data = sheetLog.getRange(1, 1, sheetLog.getLastRow(), 8).getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][1].toString().trim() === schName.trim()) {
      var tgl = "";
      try { tgl = Utilities.formatDate(new Date(data[i][0]), "GMT+7", "dd/MM/yyyy HH:mm"); }
      catch(e) { tgl = data[i][0].toString(); }
      result.push([
        tgl,
        data[i][1].toString().trim(),
        data[i][2] ? data[i][2].toString().trim() : "",
        data[i][3] ? data[i][3].toString().trim() : "",
        data[i][4] ? data[i][4].toString().trim() : "",
        data[i][5] ? data[i][5].toString().trim() : "",
        data[i][6] ? data[i][6].toString().trim() : "Pending",
        data[i][7] ? data[i][7].toString().trim() : ""
      ]);
    }
  }
  return { logs: result, deadlines: getDeadlines() };
}

function getMyFiles(schName) {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var sheetLog = ss.getSheetByName(SHEET_LOG);
  var result   = [];

  if (sheetLog && sheetLog.getLastRow() >= 2) {
    var lastRow = sheetLog.getLastRow();
    var data    = sheetLog.getRange(2, 1, lastRow - 1, 8).getValues();
    for (var i = 0; i < data.length; i++) {
      if (!data[i][1]) continue;
      if (data[i][1].toString().trim() !== schName.trim()) continue;
      var tgl = "";
      try { tgl = Utilities.formatDate(new Date(data[i][0]), "GMT+7", "dd/MM/yyyy HH:mm"); }
      catch(e) { tgl = data[i][0].toString(); }
      result.push([
        tgl,
        data[i][1].toString().trim(),
        data[i][2] ? data[i][2].toString().trim() : "",
        data[i][3] ? data[i][3].toString().trim() : "",
        data[i][4] ? data[i][4].toString().trim() : "",
        data[i][5] ? data[i][5].toString().trim() : "",
        data[i][6] ? data[i][6].toString().trim() : "Pending",
        data[i][7] ? data[i][7].toString().trim() : ""
      ]);
    }
  }

  return { logs: result, deadlines: getDeadlines() };
}

function checkUploadedPeriods(schName) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheetLog = ss.getSheetByName(SHEET_LOG);
  if (!sheetLog || sheetLog.getLastRow() < 2) return [];

  var data    = sheetLog.getRange(1, 1, sheetLog.getLastRow(), 7).getValues();
  var uploads = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][1].toString().trim() === schName.trim()) {
      uploads.push({
        jenis:  data[i][2] ? data[i][2].toString().trim() : "",
        tw:     data[i][3] ? data[i][3].toString().trim() : "",
        status: data[i][6] ? data[i][6].toString().trim() : "Pending"
      });
    }
  }
  return uploads;
}

function getAdminDashboardData() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var sheetLog  = ss.getSheetByName(SHEET_LOG);
  var sheetUser = ss.getSheetByName(SHEET_USER);
  var logs = [], users = [];

  if (sheetLog && sheetLog.getLastRow() >= 2) {
    var dLog = sheetLog.getRange(1, 1, sheetLog.getLastRow(), 8).getValues();
    for (var i = 1; i < dLog.length; i++) {
      var tgl = "";
      try { tgl = Utilities.formatDate(new Date(dLog[i][0]), "GMT+7", "dd/MM/yyyy HH:mm"); }
      catch(e) { tgl = dLog[i][0].toString(); }
      logs.push([
        tgl,
        dLog[i][1] ? dLog[i][1].toString().trim() : "",
        dLog[i][2] ? dLog[i][2].toString().trim() : "",
        dLog[i][3] ? dLog[i][3].toString().trim() : "",
        dLog[i][4] ? dLog[i][4].toString().trim() : "",
        dLog[i][5] ? dLog[i][5].toString().trim() : "",
        dLog[i][6] ? dLog[i][6].toString().trim() : "Pending",
        dLog[i][7] ? dLog[i][7].toString().trim() : ""
      ]);
    }
  }

  if (sheetUser && sheetUser.getLastRow() >= 2) {
    var dUser = sheetUser.getRange(1, 1, sheetUser.getLastRow(), 4).getValues();
    for (var j = 1; j < dUser.length; j++) {
      users.push([
        dUser[j][0] ? dUser[j][0].toString().trim() : "",
        dUser[j][1] ? dUser[j][1].toString().trim() : "",
        dUser[j][2] ? dUser[j][2].toString().trim() : "",
        dUser[j][3] ? dUser[j][3].toString().trim() : ""
      ]);
    }
  }
  return { logs: logs, users: users, deadlines: getDeadlines() };
}

// ==========================================
// 3b. SCHOOL DETAIL MODAL DATA
// ==========================================
function getSchoolDetailData(schoolName) {
  var profile = getSchoolProfile(schoolName);
  var dash = getSchoolDashboardData(schoolName);
  return {
    profile: profile,
    logs: dash.logs,
    deadlines: dash.deadlines
  };
}

function getUserData() {
  var sheetUser = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USER);
  if (!sheetUser || sheetUser.getLastRow() < 1) return [];
  return sheetUser.getRange(1, 1, sheetUser.getLastRow(), sheetUser.getLastColumn()).getValues();
}

// ==========================================
// 4. VERIFIKASI & HAPUS
// ==========================================

function actFile(fileId, statusBaru) {
  return actFileWithNote(fileId, statusBaru, '');
}

function actFileWithNote(fileId, statusBaru, catatan) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheetLog = ss.getSheetByName(SHEET_LOG);
  if (!sheetLog) return "Gagal: Sheet Log_Unggah tidak ditemukan.";

  var lock = _getLock();
  try { lock.waitLock(LOCK_WAIT_MS); } catch(e) { return "Gagal: sistem sibuk, coba lagi."; }

  try {
    var lastRow = sheetLog.getLastRow();
    if (lastRow <= 1) return "Gagal: Log kosong.";
    var ids = sheetLog.getRange(1, 6, lastRow, 1).getValues();
    for (var i = 1; i < ids.length; i++) {
      if (ids[i][0] && ids[i][0].toString().trim() === fileId.toString().trim()) {
        sheetLog.getRange(i + 1, 7, 1, 2).setValues([[statusBaru, catatan || '']]);
        try {
          var rowData = sheetLog.getRange(i + 1, 1, 1, 8).getValues()[0];
          var namaSekolah = rowData[1] ? rowData[1].toString().trim() : '';
          var namaFile    = rowData[4] ? rowData[4].toString().trim() : 'Berkas';
          var jenisBerkas = rowData[2] ? rowData[2].toString().trim() : '';
          var twBerkas    = rowData[3] ? rowData[3].toString().trim() : '';
          if (namaSekolah) {
            var pesan = '';
            if (statusBaru === 'Verified')
              pesan = '✅ Berkas ' + jenisBerkas + ' ' + twBerkas + ' Anda telah diverifikasi.';
            else if (statusBaru === 'Rejected')
              pesan = '❌ Berkas ' + jenisBerkas + ' ' + twBerkas + ' Anda ditolak'
                    + (catatan ? ': ' + catatan : '. Silakan perbaiki dan unggah ulang.');
            else if (statusBaru === 'Revision Needed')
              pesan = '✏️ Berkas ' + jenisBerkas + ' ' + twBerkas + ' memerlukan revisi'
                    + (catatan ? ': ' + catatan : '. Silakan periksa dan perbaiki.');
            else if (statusBaru === 'Pending')
              pesan = '⏳ Berkas ' + jenisBerkas + ' ' + twBerkas + ' dikembalikan ke status Pending.';
            if (pesan) _appendNotification(namaSekolah, pesan);
          }
        } catch(ne) { Logger.log('notif error: ' + ne); }
        return "OK";
      }
    }
    return "Gagal: ID tidak ditemukan.";
  } finally {
    lock.releaseLock();
  }
}

function deleteFileAdmin(fileId) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheetLog = ss.getSheetByName(SHEET_LOG);
  if (!sheetLog) return "Gagal: Sheet tidak ditemukan.";

  try { DriveApp.getFileById(fileId).setTrashed(true); } catch(e) {}

  var lock = _getLock();
  try { lock.waitLock(LOCK_WAIT_MS); } catch(e) { return "Gagal: sistem sibuk."; }

  try {
    var lastRow = sheetLog.getLastRow();
    if (lastRow <= 1) return "Log sudah kosong.";
    var ids = sheetLog.getRange(1, 6, lastRow, 1).getValues();
    for (var i = 1; i < ids.length; i++) {
      if (ids[i][0] && ids[i][0].toString().trim() === fileId.toString().trim()) {
        sheetLog.deleteRow(i + 1);
        return "OK";
      }
    }
    return "File Drive terhapus, log tidak ditemukan.";
  } finally {
    lock.releaseLock();
  }
}

function updateCatatan(fileId, catatanBaru) {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var sheetLog = ss.getSheetByName(SHEET_LOG);
  if (!sheetLog) return "Gagal: Sheet tidak ditemukan.";

  var lock = _getLock();
  try { lock.waitLock(LOCK_WAIT_MS); } catch(e) { return "Gagal: sistem sibuk."; }

  try {
    var lastRow = sheetLog.getLastRow();
    if (lastRow <= 1) return "Gagal: Log kosong.";
    var ids = sheetLog.getRange(1, 6, lastRow, 1).getValues();
    for (var i = 1; i < ids.length; i++) {
      if (ids[i][0] && ids[i][0].toString().trim() === fileId.toString().trim()) {
        sheetLog.getRange(i + 1, 8).setValue(catatanBaru || '');
        return "OK";
      }
    }
    return "Gagal: ID tidak ditemukan.";
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// RENAME FILE — ubah nama file di Drive & Log
// ==========================================

function renameFile(fileId, newName) {
  if (!fileId || !newName || !newName.toString().trim()) return "Gagal: Parameter tidak lengkap.";
  newName = newName.toString().trim();

  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var sheetLog = ss.getSheetByName(SHEET_LOG);
  if (!sheetLog) return "Gagal: Sheet Log_Unggah tidak ditemukan.";

  try {
    var driveFile = DriveApp.getFileById(fileId);
    driveFile.setName(newName);
  } catch(e) {
    return "Gagal rename di Drive: " + e.toString();
  }

  var lock = _getLock();
  try { lock.waitLock(LOCK_WAIT_MS); } catch(e) { return "Gagal: sistem sibuk."; }

  try {
    var lastRow = sheetLog.getLastRow();
    if (lastRow <= 1) return "Gagal: Log kosong.";
    var ids = sheetLog.getRange(1, 6, lastRow, 1).getValues();
    for (var i = 1; i < ids.length; i++) {
      if (ids[i][0] && ids[i][0].toString().trim() === fileId.toString().trim()) {
        sheetLog.getRange(i + 1, 5).setValue(newName);
        return "OK";
      }
    }
    return "Gagal: ID tidak ditemukan di log.";
  } finally {
    lock.releaseLock();
  }
}

function updateUserAccount(oldUsername, newUsername, newPassword, newSchoolName, newFolderId) {
  var sheetUsers = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USER);
  if (!sheetUsers) return "Gagal: Sheet User tidak ditemukan.";

  var lock = _getLock();
  try { lock.waitLock(LOCK_WAIT_MS); } catch(e) { return "Gagal: sistem sibuk."; }

  try {
    var lastRow = sheetUsers.getLastRow();
    if (lastRow < 2) return "Gagal: Tidak ada data user.";
    var data = sheetUsers.getRange(1, 1, lastRow, 4).getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === oldUsername.trim()) {
        sheetUsers.getRange(i + 1, 1, 1, 4).setValues([[newUsername, newPassword, newSchoolName, newFolderId]]);
        return "Sukses: Akun " + newSchoolName + " diperbarui.";
      }
    }
    return "Gagal: Akun tidak ditemukan.";
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 5. KUNCI PERIODE
// ==========================================

function getLockPeriodStatus() {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var sheetLock = ss.getSheetByName(SHEET_LOCK);
  if (!sheetLock) {
    sheetLock = ss.insertSheet(SHEET_LOCK);
    sheetLock.appendRow(["Nama Sekolah","TW 1","TW 2","TW 3","TW 4"]);
  }
  var sheetUser = ss.getSheetByName(SHEET_USER);
  if (!sheetUser || sheetUser.getLastRow() < 2) return [];

  var dUser = sheetUser.getRange(1, 1, sheetUser.getLastRow(), 4).getValues();
  var dLock = sheetLock.getDataRange().getValues();
  var lockMap = {};
  for (var i = 1; i < dLock.length; i++) {
    if (dLock[i][0]) lockMap[dLock[i][0].toString().trim()] = [
      dLock[i][1]||"Terbuka", dLock[i][2]||"Terbuka",
      dLock[i][3]||"Terbuka", dLock[i][4]||"Terbuka"
    ];
  }
  var result = [];
  for (var j = 1; j < dUser.length; j++) {
    var n = dUser[j][2] ? dUser[j][2].toString().trim() : "";
    if (n) {
      var s = lockMap[n] || ["Terbuka","Terbuka","Terbuka","Terbuka"];
      result.push({ school:n, tw1:s[0], tw2:s[1], tw3:s[2], tw4:s[3] });
    }
  }
  return result;
}

function togglePeriodLock(schoolName, tw, currentStatus) {
  var sheetLock = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOCK);
  if (!sheetLock) return "Gagal: Sheet kunci tidak ditemukan.";

  var colIdx    = tw==="TW 1"?2:tw==="TW 2"?3:tw==="TW 3"?4:5;
  var newStatus = currentStatus==="Terkunci"?"Terbuka":"Terkunci";

  var lock = _getLock();
  try { lock.waitLock(LOCK_WAIT_MS); } catch(e) { return "Gagal: sistem sibuk."; }

  try {
    var data = sheetLock.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === schoolName.trim()) {
        sheetLock.getRange(i + 1, colIdx).setValue(newStatus);
        return "Status " + tw + " " + schoolName + " → " + newStatus;
      }
    }
    var row = [schoolName,"Terbuka","Terbuka","Terbuka","Terbuka"];
    row[colIdx - 1] = newStatus;
    sheetLock.appendRow(row);
    return "Status " + tw + " " + schoolName + " → " + newStatus;
  } finally {
    lock.releaseLock();
  }
}

function checkIsPeriodLocked(schoolName, tw) {
  var sheetLock = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOCK);
  if (!sheetLock) return false;
  var colIdx = tw==="TW 1"?1:tw==="TW 2"?2:tw==="TW 3"?3:4;
  var data   = sheetLock.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === schoolName.trim()) {
      return data[i][colIdx] === "Terkunci";
    }
  }
  return false;
}

// ==========================================
// 6. PENGATURAN ADMIN
// ==========================================

function updateAdminProfileExtended(newUser, newPass, appName, appLogo, appTheme) {
  var props = PropertiesService.getScriptProperties();
  if (newUser && newUser.trim()) props.setProperty("admin_user", newUser.trim());
  if (newPass && newPass.trim()) props.setProperty("admin_pass", newPass.trim());
  if (appName && appName.trim()) props.setProperty("app_name",  appName.trim());
  if (appLogo)                   props.setProperty("app_logo",  appLogo.trim());
  if (appTheme)                  props.setProperty("app_theme", appTheme.trim());
  return "Sukses";
}

// ==========================================
// 7. TENGGAT WAKTU PER TW
// ==========================================

function getDeadlines() {
  var props = PropertiesService.getScriptProperties();
  return {
    'TW 1': props.getProperty('dl_TW1') || '',
    'TW 2': props.getProperty('dl_TW2') || '',
    'TW 3': props.getProperty('dl_TW3') || '',
    'TW 4': props.getProperty('dl_TW4') || ''
  };
}

function setDeadlines(obj) {
  var props = PropertiesService.getScriptProperties();
  var map = { 'TW 1': 'dl_TW1', 'TW 2': 'dl_TW2', 'TW 3': 'dl_TW3', 'TW 4': 'dl_TW4' };
  for (var tw in map) {
    var val = (obj && obj[tw]) ? obj[tw].toString().trim() : '';
    if (val) props.setProperty(map[tw], val);
    else     props.deleteProperty(map[tw]);
  }
  return { ok: true };
}

// ==========================================
// 8. [FITUR #5] AUTO-LOCK SETELAH TENGGAT
// ==========================================

function autoLockExpiredPeriods() {
  var JENIS_LIST = ['ARKAS Per TW','Bukti Pajak','Buku Pembantu Pajak','LK BOS','Rekap Pajak','Rekening Koran'];
  var TW_LIST    = ['TW 1','TW 2','TW 3','TW 4'];

  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var sheetUser = ss.getSheetByName(SHEET_USER);
  var sheetLog  = ss.getSheetByName(SHEET_LOG);
  var sheetLock = ss.getSheetByName(SHEET_LOCK);

  if (!sheetUser || !sheetLog) {
    Logger.log('autoLockExpiredPeriods: Sheet tidak ditemukan.');
    return;
  }

  if (!sheetLock) {
    sheetLock = ss.insertSheet(SHEET_LOCK);
    sheetLock.appendRow(["Nama Sekolah","TW 1","TW 2","TW 3","TW 4"]);
  }

  var deadlines = getDeadlines();
  var now = new Date();
  now.setHours(0,0,0,0);

  var expiredTWs = [];
  TW_LIST.forEach(function(tw) {
    var dl = deadlines[tw];
    if (!dl) return;
    var dlDate = new Date(dl);
    dlDate.setHours(0,0,0,0);
    if (now > dlDate) expiredTWs.push(tw);
  });

  if (expiredTWs.length === 0) {
    Logger.log('autoLockExpiredPeriods: Tidak ada TW yang melewati deadline.');
    return;
  }

  var logData = [];
  if (sheetLog.getLastRow() >= 2) {
    logData = sheetLog.getRange(2, 1, sheetLog.getLastRow()-1, 7).getValues();
  }

  var users = [];
  if (sheetUser.getLastRow() >= 2) {
    var ud = sheetUser.getRange(2, 1, sheetUser.getLastRow()-1, 3).getValues();
    ud.forEach(function(r){ if(r[2]) users.push(r[2].toString().trim()); });
  }

  var lockData = sheetLock.getDataRange().getValues();
  var lockMap  = {};
  for (var i = 1; i < lockData.length; i++) {
    if (lockData[i][0]) lockMap[lockData[i][0].toString().trim()] = i+1;
  }

  var lockedCount = 0;

  users.forEach(function(schoolName) {
    expiredTWs.forEach(function(tw) {
      var isComplete = JENIS_LIST.every(function(j) {
        return logData.some(function(row) {
          return row[1] && row[1].toString().trim() === schoolName
            && row[2] && row[2].toString().trim() === j
            && row[3] && row[3].toString().trim() === tw
            && row[6] && row[6].toString().trim() === 'Verified';
        });
      });

      if (isComplete) return;

      var colIdx = tw==='TW 1'?2:tw==='TW 2'?3:tw==='TW 3'?4:5;

      if (lockMap[schoolName]) {
        var rowNum = lockMap[schoolName];
        var curVal = lockData[rowNum-1][colIdx-1];
        if (curVal !== 'Terkunci') {
          sheetLock.getRange(rowNum, colIdx).setValue('Terkunci');
          lockedCount++;
        }
      } else {
        var newRow = [schoolName,'Terbuka','Terbuka','Terbuka','Terbuka'];
        newRow[colIdx-1] = 'Terkunci';
        sheetLock.appendRow(newRow);
        lockMap[schoolName] = sheetLock.getLastRow();
        lockedCount++;
      }
    });
  });

  Logger.log('autoLockExpiredPeriods: ' + lockedCount + ' periode dikunci otomatis.');

  if (lockedCount > 0) {
    _sendAutoLockNotifications(users, expiredTWs, logData, JENIS_LIST, deadlines);
  }
}

// ==========================================
// 9. [FITUR #2] NOTIFIKASI EMAIL OTOMATIS
// ==========================================

function _getSchoolEmail(schoolName) {
  try {
    var sheetUser = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USER);
    if (!sheetUser || sheetUser.getLastRow() < 2) return null;
    var data = sheetUser.getRange(2, 1, sheetUser.getLastRow()-1, 5).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][2] && data[i][2].toString().trim() === schoolName.trim()) {
        return data[i][4] ? data[i][4].toString().trim() : null;
      }
    }
  } catch(e) { Logger.log('_getSchoolEmail error: ' + e); }
  return null;
}

function _getAppName() {
  return PropertiesService.getScriptProperties().getProperty('app_name') || 'SIPABOS';
}

function _sendAutoLockNotifications(users, expiredTWs, logData, jenisList, deadlines) {
  var appName = _getAppName();
  users.forEach(function(schoolName) {
    var email = _getSchoolEmail(schoolName);
    if (!email) return;

    var belumList = [];
    expiredTWs.forEach(function(tw) {
      jenisList.forEach(function(j) {
        var ada = logData.some(function(row) {
          return row[1] && row[1].toString().trim() === schoolName
            && row[2] && row[2].toString().trim() === j
            && row[3] && row[3].toString().trim() === tw
            && row[6] && row[6].toString().trim() === 'Verified';
        });
        if (!ada) belumList.push(tw + ' — ' + j);
      });
    });
    if (belumList.length === 0) return;

    try {
      var subject = '[' + appName + '] Periode Dikunci Otomatis — Laporan Belum Lengkap';
      var body =
        'Yth. Operator ' + schoolName + ',\n\n' +
        'Periode upload berikut telah DIKUNCI OTOMATIS karena melewati tenggat waktu dan laporan belum lengkap:\n\n' +
        belumList.map(function(s){ return '  • ' + s; }).join('\n') + '\n\n' +
        'Silakan hubungi Admin Dinas untuk informasi lebih lanjut.\n\n' +
        'Salam,\nSistem ' + appName;
      MailApp.sendEmail(email, subject, body);
    } catch(e) { Logger.log('_sendAutoLockNotifications email error for ' + schoolName + ': ' + e); }
  });
}

function sendDeadlineReminders() {
  var JENIS_LIST = ['ARKAS Per TW','Bukti Pajak','Buku Pembantu Pajak','LK BOS','Rekap Pajak','Rekening Koran'];
  var TW_LIST    = ['TW 1','TW 2','TW 3','TW 4'];

  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var sheetUser = ss.getSheetByName(SHEET_USER);
  var sheetLog  = ss.getSheetByName(SHEET_LOG);
  if (!sheetUser || !sheetLog) return;

  var deadlines = getDeadlines();
  var now = new Date(); now.setHours(0,0,0,0);
  var appName = _getAppName();

  var REMINDER_DAYS = [7, 3];

  var users = [];
  if (sheetUser.getLastRow() >= 2) {
    var ud = sheetUser.getRange(2, 1, sheetUser.getLastRow()-1, 3).getValues();
    ud.forEach(function(r){ if(r[2]) users.push(r[2].toString().trim()); });
  }

  var logData = [];
  if (sheetLog.getLastRow() >= 2) {
    logData = sheetLog.getRange(2, 1, sheetLog.getLastRow()-1, 7).getValues();
  }

  TW_LIST.forEach(function(tw) {
    var dl = deadlines[tw]; if (!dl) return;
    var dlDate = new Date(dl); dlDate.setHours(0,0,0,0);
    var diff = Math.round((dlDate - now) / 86400000);

    if (REMINDER_DAYS.indexOf(diff) === -1) return;

    users.forEach(function(schoolName) {
      var email = _getSchoolEmail(schoolName); if (!email) return;

      var belumList = [];
      JENIS_LIST.forEach(function(j) {
        var ada = logData.some(function(row) {
          return row[1] && row[1].toString().trim() === schoolName
            && row[2] && row[2].toString().trim() === j
            && row[3] && row[3].toString().trim() === tw
            && row[6] && row[6].toString().trim() === 'Verified';
        });
        if (!ada) belumList.push(j);
      });

      if (belumList.length === 0) return;

      try {
        var dlStr = dlDate.getDate() + '/' + (dlDate.getMonth()+1) + '/' + dlDate.getFullYear();
        var subject = '[' + appName + '] Pengingat ' + diff + ' Hari Lagi — Tenggat Upload ' + tw;
        var body =
          'Yth. Operator ' + schoolName + ',\n\n' +
          'Mengingatkan bahwa tenggat upload ' + tw + ' adalah ' + dlStr + ' (' + diff + ' hari lagi).\n\n' +
          'Laporan yang BELUM diunggah:\n' +
          belumList.map(function(s){ return '  • ' + s; }).join('\n') + '\n\n' +
          'Segera lengkapi laporan sebelum tenggat waktu untuk menghindari penguncian otomatis.\n\n' +
          'Salam,\nSistem ' + appName;
        MailApp.sendEmail(email, subject, body);
        Logger.log('Reminder H-' + diff + ' dikirim ke ' + email + ' untuk ' + tw);
      } catch(e) { Logger.log('sendDeadlineReminders error: ' + e); }
    });
  });
}

// ==========================================
// 10. UTILITAS — test otorisasi Drive & email
// ==========================================

function testDriveAccess() {
  try {
    var root = DriveApp.getRootFolder();
    Logger.log('Drive OK: ' + root.getName());
    return 'OK: ' + root.getName();
  } catch(e) {
    Logger.log('Drive ERROR: ' + e.toString());
    return 'ERROR: ' + e.toString();
  }
}

function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'autoLockExpiredPeriods' ||
        t.getHandlerFunction() === 'sendDeadlineReminders') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('autoLockExpiredPeriods')
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .create();

  ScriptApp.newTrigger('sendDeadlineReminders')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();

  Logger.log('Triggers berhasil dipasang.');
  return 'Triggers dipasang: autoLockExpiredPeriods (00:00) + sendDeadlineReminders (07:00)';
}

// ==========================================
// 11. BULK DOWNLOAD — ambil bytes file dari Drive
// ==========================================
function getFilesBytes(fileIds) {
  var results = [];
  for (var i = 0; i < fileIds.length; i++) {
    try {
      var file = DriveApp.getFileById(fileIds[i]);
      var bytes = file.getBlob().getBytes();
      results.push({
        id: fileIds[i],
        name: file.getName(),
        bytes: bytes,
        ok: true
      });
    } catch (e) {
      results.push({ id: fileIds[i], name: '', bytes: [], ok: false });
    }
  }
  return results;
}

// ==========================================
// 12. PENGUMUMAN / BROADCAST ADMIN
// ==========================================

var SHEET_NOTIF = "Notifikasi";

function getAnnouncement() {
  var props = PropertiesService.getScriptProperties();
  return {
    msg:    props.getProperty('announce_msg')    || '',
    active: props.getProperty('announce_active') === 'true',
    ts:     props.getProperty('announce_ts')     || ''
  };
}

function setAnnouncement(msg, active) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('announce_msg',    msg    ? msg.toString().trim() : '');
  props.setProperty('announce_active', active ? 'true' : 'false');
  props.setProperty('announce_ts',
    Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm"));
  return { ok: true };
}

// ==========================================
// 13. NOTIFIKASI IN-APP PER SEKOLAH
// ==========================================

function _ensureNotifSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NOTIF);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NOTIF);
    sheet.appendRow(["Timestamp", "Nama Sekolah", "Pesan", "IsRead"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function _appendNotification(schoolName, message) {
  try {
    var sheet = _ensureNotifSheet();
    var ts    = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
    sheet.appendRow([ts, schoolName, message, false]);
  } catch(e) {
    Logger.log('_appendNotification error: ' + e);
  }
}

function getMyNotifications(schoolName) {
  var sheet = _ensureNotifSheet();
  var last  = sheet.getLastRow();
  if (last < 2) return { notifs: [], unread: 0 };

  var data   = sheet.getRange(2, 1, last - 1, 4).getValues();
  var notifs = [];
  var unread = 0;

  for (var i = data.length - 1; i >= 0; i--) {
    if (!data[i][1] || data[i][1].toString().trim() !== schoolName.trim()) continue;
    var isRead = data[i][3] === true || data[i][3] === 'TRUE';
    notifs.push({
      ts:     data[i][0] ? data[i][0].toString() : '',
      msg:    data[i][2] ? data[i][2].toString() : '',
      isRead: isRead,
      rowIdx: i + 2
    });
    if (!isRead) unread++;
    if (notifs.length >= 30) break;
  }
  return { notifs: notifs, unread: unread };
}

function markNotifsRead(schoolName) {
  var lock = _getLock();
  try { lock.waitLock(LOCK_WAIT_MS); } catch(e) { return false; }
  try {
    var sheet = _ensureNotifSheet();
    var last  = sheet.getLastRow();
    if (last < 2) return true;
    var data = sheet.getRange(2, 1, last - 1, 4).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().trim() === schoolName.trim()
          && data[i][3] !== true && data[i][3] !== 'TRUE') {
        sheet.getRange(i + 2, 4).setValue(true);
      }
    }
    return true;
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 14. BROADCAST TEMPLATE FILE OLEH ADMIN
// ==========================================

var SHEET_TEMPLATE = "Template_Broadcast";
var FOLDER_TEMPLATE = "SIPABOS_Templates";

function _ensureTemplateSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_TEMPLATE);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_TEMPLATE);
    sheet.appendRow(["Timestamp","Nama File","File ID","Kategori","Deskripsi","Diunggah Oleh"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function _getTemplateFolder() {
  var it = DriveApp.getFoldersByName(FOLDER_TEMPLATE);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_TEMPLATE);
}

function uploadTemplateBroadcast(fileData, kategori, deskripsi) {
  if (!fileData || !fileData.bytes || !fileData.bytes.length) {
    return { status: 'error', message: 'Data file kosong.' };
  }

  var lock = _getLock();
  try { lock.waitLock(LOCK_WAIT_MS); }
  catch (e) { return { status: 'error', message: 'Sistem sibuk, coba lagi.' }; }

  try {
    var folder = _getTemplateFolder();
    var mime   = fileData.fileType || 'application/octet-stream';
    var blob   = Utilities.newBlob(fileData.bytes, mime, fileData.fileName);
    var drFile = folder.createFile(blob);
    try { drFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(se) {}

    var sheet = _ensureTemplateSheet();
    var ts    = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
    sheet.appendRow([ts, fileData.fileName, drFile.getId(), kategori || '', deskripsi || '', 'Administrator']);

    return { status: 'success', fileId: drFile.getId(), message: 'Template berhasil diunggah.' };
  } catch (e) {
    return { status: 'error', message: 'Gagal: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function getTemplateBroadcast() {
  var sheet = _ensureTemplateSheet();
  var last  = sheet.getLastRow();
  if (last < 2) return [];

  var data    = sheet.getRange(2, 1, last - 1, 6).getValues();
  var result  = [];
  for (var i = data.length - 1; i >= 0; i--) {
    if (!data[i][2]) continue;
    var ts = '';
    try { ts = Utilities.formatDate(new Date(data[i][0]), "GMT+7", "dd/MM/yyyy HH:mm"); }
    catch(e) { ts = data[i][0].toString(); }
    result.push({
      ts:       ts,
      nama:     data[i][1] ? data[i][1].toString() : '',
      fileId:   data[i][2] ? data[i][2].toString() : '',
      kategori: data[i][3] ? data[i][3].toString() : '',
      desk:     data[i][4] ? data[i][4].toString() : '',
      oleh:     data[i][5] ? data[i][5].toString() : ''
    });
  }
  return result;
}

function deleteTemplate(fileId) {
  if (!fileId) return 'Gagal: ID tidak ada.';
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch(e) {}

  var lock = _getLock();
  try { lock.waitLock(LOCK_WAIT_MS); } catch(e) { return 'Gagal: sistem sibuk.'; }

  try {
    var sheet = _ensureTemplateSheet();
    var last  = sheet.getLastRow();
    if (last < 2) return 'OK';
    var ids = sheet.getRange(2, 3, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (ids[i][0] && ids[i][0].toString().trim() === fileId.trim()) {
        sheet.deleteRow(i + 2);
        return 'OK';
      }
    }
    return 'OK';
  } finally {
    lock.releaseLock();
  }
}
