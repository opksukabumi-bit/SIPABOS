// ==========================================
// ENHANCED FEATURES — SIPABOS V.3.11+
// ==========================================

// ==========================================
// 1. SCHOOL PROFILE MANAGEMENT
// ==========================================

var SHEET_PROFILES = "Profil_Sekolah";

function _ensureProfileSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_PROFILES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PROFILES);
    sheet.appendRow([
      "Nama Sekolah", "NPSN", "Kepala Sekolah", "Alamat", 
      "Telepon", "Email", "Tahun Anggaran", "Status Aktif", "Catatan"
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getSchoolProfile(schoolName) {
  var sheet = _ensureProfileSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === schoolName.trim()) {
      return {
        name: data[i][0] ? data[i][0].toString() : "",
        npsn: data[i][1] ? data[i][1].toString() : "",
        kepala: data[i][2] ? data[i][2].toString() : "",
        alamat: data[i][3] ? data[i][3].toString() : "",
        telepon: data[i][4] ? data[i][4].toString() : "",
        email: data[i][5] ? data[i][5].toString() : "",
        tahunBudget: data[i][6] ? data[i][6].toString() : "",
        aktif: data[i][7] ? data[i][7].toString() : "Ya",
        catatan: data[i][8] ? data[i][8].toString() : ""
      };
    }
  }
  return null;
}

function updateSchoolProfile(schoolName, profileData) {
  var sheet = _ensureProfileSheet();
  var lock = _getLock();
  try { lock.waitLock(LOCK_WAIT_MS); }
  catch (e) { return "Gagal: sistem sibuk."; }

  try {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === schoolName.trim()) {
        sheet.getRange(i + 1, 1, 1, 9).setValues([[
          profileData.name || "",
          profileData.npsn || "",
          profileData.kepala || "",
          profileData.alamat || "",
          profileData.telepon || "",
          profileData.email || "",
          profileData.tahunBudget || "",
          profileData.aktif || "Ya",
          profileData.catatan || ""
        ]]);
        return "Profil berhasil diperbarui.";
      }
    }
    // Jika tidak ditemukan, tambah baris baru
    sheet.appendRow([
      profileData.name || "",
      profileData.npsn || "",
      profileData.kepala || "",
      profileData.alamat || "",
      profileData.telepon || "",
      profileData.email || "",
      profileData.tahunBudget || "",
      profileData.aktif || "Ya",
      profileData.catatan || ""
    ]);
    return "Profil baru berhasil ditambahkan.";
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 2. REVISION STATUS (Status Perlu Revisi)
// ==========================================

function setFileRevisionStatus(fileId, revisionNote) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetLog = ss.getSheetByName(SHEET_LOG);
  if (!sheetLog) return "Gagal: Sheet Log_Unggah tidak ditemukan.";

  var lock = _getLock();
  try { lock.waitLock(LOCK_WAIT_MS); }
  catch (e) { return "Gagal: sistem sibuk."; }

  try {
    var lastRow = sheetLog.getLastRow();
    if (lastRow <= 1) return "Gagal: Log kosong.";
    var ids = sheetLog.getRange(1, 6, lastRow, 1).getValues();
    for (var i = 1; i < ids.length; i++) {
      if (ids[i][0] && ids[i][0].toString().trim() === fileId.toString().trim()) {
        sheetLog.getRange(i + 1, 7, 1, 2).setValues([["Revision", revisionNote || ""]]);
        // Kirim notifikasi
        try {
          var rowData = sheetLog.getRange(i + 1, 1, 1, 8).getValues()[0];
          var namaSekolah = rowData[1] ? rowData[1].toString().trim() : "";
          var jenisBerkas = rowData[2] ? rowData[2].toString().trim() : "";
          var twBerkas = rowData[3] ? rowData[3].toString().trim() : "";
          if (namaSekolah) {
            var pesan = "📝 Berkas ' + jenisBerkas + ' ' + twBerkas + ' perlu direvisi.\n" +
              (revisionNote ? "Instruksi: " + revisionNote : "Silakan perbaiki berkas dan upload ulang.");
            _appendNotification(namaSekolah, pesan);
          }
        } catch (ne) { Logger.log('notif error: ' + ne); }
        return "OK";
      }
    }
    return "Gagal: ID tidak ditemukan.";
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 3. SCHOOL INDIVIDUAL DASHBOARD
// ==========================================

function getSchoolDetailedDashboard(schoolName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetLog = ss.getSheetByName(SHEET_LOG);
  var sheetLock = ss.getSheetByName(SHEET_LOCK);

  if (!sheetLog) return { error: "Sheet Log tidak ditemukan." };

  var profile = getSchoolProfile(schoolName);
  var logs = [];
  var summary = {};

  if (sheetLog.getLastRow() >= 2) {
    var data = sheetLog.getRange(2, 1, sheetLog.getLastRow() - 1, 8).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().trim() === schoolName.trim()) {
        var tgl = "";
        try { tgl = Utilities.formatDate(new Date(data[i][0]), "GMT+7", "dd/MM/yyyy HH:mm"); }
        catch (e) { tgl = data[i][0].toString(); }
        logs.push({
          tgl: tgl,
          jenis: data[i][2] ? data[i][2].toString().trim() : "",
          tw: data[i][3] ? data[i][3].toString().trim() : "",
          fileName: data[i][4] ? data[i][4].toString().trim() : "",
          fileId: data[i][5] ? data[i][5].toString().trim() : "",
          status: data[i][6] ? data[i][6].toString().trim() : "Pending",
          catatan: data[i][7] ? data[i][7].toString().trim() : ""
        });
      }
    }
  }

  // Hitung summary per status
  summary.verified = logs.filter(function(l) { return l.status === "Verified"; }).length;
  summary.pending = logs.filter(function(l) { return l.status === "Pending"; }).length;
  summary.rejected = logs.filter(function(l) { return l.status === "Rejected"; }).length;
  summary.revision = logs.filter(function(l) { return l.status === "Revision"; }).length;

  // Hitung kepatuhan per TW
  var twStatus = {};
  ["TW 1", "TW 2", "TW 3", "TW 4"].forEach(function(tw) {
    var complete = JENIS_LIST.every(function(j) {
      return logs.some(function(l) {
        return l.jenis === j && l.tw === tw && l.status === "Verified";
      });
    });
    var verCount = JENIS_LIST.filter(function(j) {
      return logs.some(function(l) { return l.jenis === j && l.tw === tw && l.status === "Verified"; });
    }).length;
    twStatus[tw] = { complete: complete, verified: verCount, total: JENIS_LIST.length };
  });

  // Baca lock status jika ada
  var locks = {};
  if (sheetLock && sheetLock.getLastRow() >= 2) {
    var lockData = sheetLock.getDataRange().getValues();
    for (var j = 1; j < lockData.length; j++) {
      if (lockData[j][0] && lockData[j][0].toString().trim() === schoolName.trim()) {
        locks = {
          tw1: lockData[j][1] || "Terbuka",
          tw2: lockData[j][2] || "Terbuka",
          tw3: lockData[j][3] || "Terbuka",
          tw4: lockData[j][4] || "Terbuka"
        };
        break;
      }
    }
  }

  return {
    profile: profile,
    logs: logs,
    summary: summary,
    twStatus: twStatus,
    locks: locks,
    deadlines: getDeadlines()
  };
}

// ==========================================
// 4. MULTI-YEAR BUDGET SUPPORT
// ==========================================

function getAllBudgetYears() {
  var sheet = _ensureProfileSheet();
  var years = new Set();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][6]) years.add(data[i][6].toString().trim());
  }
  return Array.from(years).sort().reverse();
}

function getBudgetYearForSchool(schoolName) {
  var profile = getSchoolProfile(schoolName);
  return profile ? profile.tahunBudget : "";
}

// ==========================================
// 5. THEME PREFERENCE (User Preferences)
// ==========================================

var SHEET_PREFS = "User_Preferences";

function _ensurePrefsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_PREFS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PREFS);
    sheet.appendRow(["Username", "Role", "Theme", "ItemsPerPage", "LastLogin"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getUserTheme(username) {
  var sheet = _ensurePrefsSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === username.trim()) {
      return data[i][2] ? data[i][2].toString().trim() : "dark";
    }
  }
  return "dark"; // Default dark theme
}

function setUserTheme(username, theme) {
  var sheet = _ensurePrefsSheet();
  var lock = _getLock();
  try { lock.waitLock(LOCK_WAIT_MS); }
  catch (e) { return false; }

  try {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === username.trim()) {
        sheet.getRange(i + 1, 3).setValue(theme);
        return true;
      }
    }
    return false;
  } finally {
    lock.releaseLock();
  }
}

function getUserItemsPerPage(username) {
  var sheet = _ensurePrefsSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === username.trim()) {
      return data[i][3] ? parseInt(data[i][3]) : 15;
    }
  }
  return 15; // Default
}

function setUserItemsPerPage(username, itemsPerPage) {
  var sheet = _ensurePrefsSheet();
  var lock = _getLock();
  try { lock.waitLock(LOCK_WAIT_MS); }
  catch (e) { return false; }

  try {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === username.trim()) {
        sheet.getRange(i + 1, 4).setValue(itemsPerPage);
        return true;
      }
    }
    // Buat entry baru jika user belum ada
    var sheetUser = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USER);
    var role = "user";
    if (sheetUser && sheetUser.getLastRow() >= 2) {
      var userData = sheetUser.getRange(1, 1, sheetUser.getLastRow(), 3).getValues();
      for (var j = 1; j < userData.length; j++) {
        if (userData[j][0] && userData[j][0].toString().trim() === username.trim()) {
          role = "user";
          break;
        }
      }
    }
    sheet.appendRow([username, role, "dark", itemsPerPage, Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd")]);
    return true;
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 6. ENHANCED STATUS TRACKING
// ==========================================

function getStatusBreakdown(schoolName, tw) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetLog = ss.getSheetByName(SHEET_LOG);
  if (!sheetLog) return { verified: [], pending: [], rejected: [], revision: [] };

  var result = { verified: [], pending: [], rejected: [], revision: [] };
  if (sheetLog.getLastRow() < 2) return result;

  var data = sheetLog.getRange(2, 1, sheetLog.getLastRow() - 1, 8).getValues();
  for (var i = 0; i < data.length; i++) {
    var schName = data[i][1] ? data[i][1].toString().trim() : "";
    var twData = data[i][3] ? data[i][3].toString().trim() : "";
    if (schName === schoolName.trim() && (!tw || twData === tw)) {
      var item = {
        jenis: data[i][2] ? data[i][2].toString().trim() : "",
        fileName: data[i][4] ? data[i][4].toString().trim() : "",
        catatan: data[i][7] ? data[i][7].toString().trim() : ""
      };
      var status = data[i][6] ? data[i][6].toString().trim() : "Pending";
      if (status === "Verified") result.verified.push(item);
      else if (status === "Rejected") result.rejected.push(item);
      else if (status === "Revision") result.revision.push(item);
      else result.pending.push(item);
    }
  }
  return result;
}
