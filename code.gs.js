/* 
  SISTEM ABSENSI GURU V7.3 (WITH CALENDAR FEATURE) - FIXED
  
  Struktur Database:
  1. Sheet 'Users': Data Login.
  2. Sheet 'Database': MASTER Data Semua Guru.
  3. Sheet 'Absensi_[Username]': Data Cepat.
  4. Sheet 'Holidays_Global': Libur Global.
  5. Sheet 'Settings_User_[Username]': Pengaturan Libur Per User.
*/

const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SHEET_USERS_NAME = "Users";
const SHEET_DATABASE_NAME = "Database";
const SHEET_HOLIDAYS_GLOBAL = "Holidays_Global";

// ==========================================
// 1. SETUP DATABASE
// ==========================================
function setupDatabase() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // --- SETUP USERS ---
  let usersSheet = ss.getSheetByName(SHEET_USERS_NAME);
  if (!usersSheet) { usersSheet = ss.insertSheet(SHEET_USERS_NAME); } else { usersSheet.clear(); }

  const usersHeader = ["ID", "Username", "Password", "Nama Lengkap", "Role"];
  usersSheet.getRange(1, 1, 1, usersHeader.length).setValues([usersHeader]);
  usersSheet.getRange(1, 1, 1, usersHeader.length).setFontWeight("bold").setBackground("#4338ca").setFontColor("#fff").setHorizontalAlignment("center");
  usersSheet.setFrozenRows(1);

  const dummyUsers = [
    [1, "admin", "admin", "Administrator", "admin"],
    [2, "guru1", "12345", "Budi Santoso", "guru"],
    [3, "guru2", "rahasia", "Siti Aminah", "guru"]
  ];
  usersSheet.getRange(2, 1, dummyUsers.length, dummyUsers[0].length).setValues(dummyUsers);

  // --- SETUP MASTER DATABASE ---
  let dbSheet = ss.getSheetByName(SHEET_DATABASE_NAME);
  if (!dbSheet) { dbSheet = ss.insertSheet(SHEET_DATABASE_NAME); } else { dbSheet.clear(); }

  const dbHeader = ["ID User", "Username", "Tanggal", "Jam", "Status", "Kategori"];
  dbSheet.getRange(1, 1, 1, dbHeader.length).setValues([dbHeader]);
  dbSheet.getRange(1, 1, 1, dbHeader.length).setFontWeight("bold").setBackground("#059669").setFontColor("#fff").setHorizontalAlignment("center");
  dbSheet.setFrozenRows(1);

  // --- SETUP HOLIDAYS GLOBAL ---
  let holidaysGlobalSheet = ss.getSheetByName(SHEET_HOLIDAYS_GLOBAL);
  if (!holidaysGlobalSheet) { holidaysGlobalSheet = ss.insertSheet(SHEET_HOLIDAYS_GLOBAL); } else { holidaysGlobalSheet.clear(); }

  const holidaysHeader = ["Tanggal", "Nama Libur", "Jenis", "Berulang"];
  holidaysGlobalSheet.getRange(1, 1, 1, holidaysHeader.length).setValues([holidaysHeader]);
  holidaysGlobalSheet.getRange(1, 1, 1, holidaysHeader.length).setFontWeight("bold").setBackground("#dc2626").setFontColor("#fff").setHorizontalAlignment("center");
  holidaysGlobalSheet.setFrozenRows(1);

  // Contoh libur nasional (format: yyyy-MM-dd)
  const sampleHolidays = [
    ["2025-01-01", "Tahun Baru 2025", "Nasional", "Tahunan"],
    ["2025-04-18", "Jumat Agung", "Nasional", "Tahunan"],
    ["2025-05-01", "Hari Buruh", "Nasional", "Tahunan"],
    ["2025-05-29", "Hari Raya Waisak", "Nasional", "Tahunan"],
    ["2025-06-01", "Hari Lahir Pancasila", "Nasional", "Tahunan"],
    ["2025-08-17", "Hari Kemerdekaan RI", "Nasional", "Tahunan"],
    ["2025-12-25", "Hari Natal", "Nasional", "Tahunan"]
  ];
  
  holidaysGlobalSheet.getRange(2, 1, sampleHolidays.length, sampleHolidays[0].length).setValues(sampleHolidays);

  // Hapus Sheet1
  try { const defaultSheet = ss.getSheetByName("Sheet1"); if (defaultSheet) ss.deleteSheet(defaultSheet); } catch (e) { }

  SpreadsheetApp.getUi().alert("Setup V7.3 Selesai!\n\nFitur Kalender Aktif:\n- Libur Global\n- Pengaturan Libur Per User\n- Sistem Periode 26-25\n- Target 20 hari absen");
}

// ==========================================
// 2. CORE LOGIC (doPost) - FIXED
// ==========================================
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    let result = {};
    let usersSheet = ss.getSheetByName(SHEET_USERS_NAME);
    if (!usersSheet) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "DB belum setup." }));
    }

    // --- ACTION: LOGIN ---
    if (action === "login") {
      const user = validateUser(usersSheet, data.username, data.password);
      result = user ? { status: "success", userData: user } : { status: "error", message: "Login Gagal!" };
    }

    // --- ACTION: ABSEN ---
    else if (action === "absen") {
      const user = validateUser(usersSheet, data.username, data.password);
      if (!user) {
        result = { status: "error", message: "Sesi tidak valid (User error)" };
      } else {
        const targetSheetName = "Absensi_" + data.username;
        let userSheet = ss.getSheetByName(targetSheetName);

        if (!userSheet) {
          userSheet = ss.insertSheet(targetSheetName);
          userSheet.appendRow(["Tanggal", "Jam", "Status", "Kategori"]);
          userSheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#4338ca").setFontColor("#fff").setHorizontalAlignment("center");
          userSheet.setFrozenRows(1);
        }

        // Cek Duplicate
        const inputDate = data.date;
        const lastRow = userSheet.getLastRow();
        let isDuplicate = false;

        if (lastRow > 1) {
          const dates = userSheet.getRange(2, 1, lastRow - 1, 1).getValues();
          for (let i = 0; i < dates.length; i++) {
            let sheetDate = dates[i][0] instanceof Date
              ? Utilities.formatDate(dates[i][0], "GMT+7", "yyyy-MM-dd")
              : String(dates[i][0]);
            if (sheetDate === inputDate) { isDuplicate = true; break; }
          }
        }

        if (isDuplicate) {
          result = { status: "error", message: "Sudah absen hari ini!" };
        } else {
          // SIMPAN DATA USER
          userSheet.appendRow([new Date(inputDate), data.time, data.status, data.category]);

          if (!userSheet.isSheetHidden()) {
            userSheet.hideSheet();
          }

          // SIMPAN DATA MASTER
          let dbSheet = ss.getSheetByName(SHEET_DATABASE_NAME);
          if (!dbSheet) {
            dbSheet = ss.insertSheet(SHEET_DATABASE_NAME);
            dbSheet.appendRow(["ID User", "Username", "Tanggal", "Jam", "Status", "Kategori"]);
            dbSheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#059669").setFontColor("#fff");
          }
          dbSheet.appendRow([user.id, data.username, new Date(inputDate), data.time, data.status, data.category]);

          if (userSheet.getLastRow() > 1000) userSheet.deleteRows(2, 100);

          result = { status: "success", message: "Absen Berhasil" };
        }
      }
    }

    // --- ACTION: GET DATA ---
    else if (action === "get_data") {
      const targetSheetName = "Absensi_" + data.username;
      let sheet = ss.getSheetByName(targetSheetName);
      let rawData = [];

      if (sheet && sheet.getLastRow() > 1) {
        rawData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
      }

      const formattedData = rawData.map(row => ({
        date: row[0] instanceof Date ? Utilities.formatDate(row[0], "GMT+7", "yyyy-MM-dd") : row[0],
        time: row[1],
        status: row[2],
        category: row[3]
      }));

      result = { status: "success", data: formattedData };
    }

    // --- ACTION: GET HOLIDAY SETTINGS (FIXED) ---
    else if (action === "get_holiday_settings") {
      // PERBAIKAN: Hanya butuh username, tidak perlu password untuk load settings
      const settings = getHolidaySettings(ss, data.username);
      if (settings) {
        result = { 
          status: "success", 
          settings: {
            additionalHolidays: settings.additionalHolidays || [],
            optionalDays: settings.optionalDays || []
          }
        };
      } else {
        result = { 
          status: "success", 
          settings: {
            additionalHolidays: [],
            optionalDays: []
          }
        };
      }
    }

    // --- ACTION: SAVE HOLIDAY SETTINGS (FIXED) ---
    else if (action === "save_holiday_settings") {
      const user = validateUser(usersSheet, data.username, data.password);
      if (!user) {
        result = { status: "error", message: "Sesi tidak valid" };
      } else {
        try {
          // Pastikan data settings lengkap
          const settingsToSave = {
            additionalHolidays: data.settings.additionalHolidays || [],
            optionalDays: data.settings.optionalDays || []
          };
          
          saveHolidaySettings(ss, user.username, settingsToSave);
          result = { status: "success", message: "Pengaturan libur berhasil disimpan" };
        } catch (error) {
          result = { status: "error", message: "Gagal menyimpan pengaturan: " + error.toString() };
        }
      }
    }

    // --- ACTION: SAVE HOLIDAY (Legacy) ---
    else if (action === "save_holiday") {
      const user = validateUser(usersSheet, data.username, data.password);
      if (!user) {
        result = { status: "error", message: "Sesi tidak valid" };
      } else {
        saveHoliday(ss, user.username, data.holidayData);
        result = { status: "success", message: "Libur berhasil disimpan" };
      }
    }

    // --- ACTION: DELETE HOLIDAY (Legacy) ---
    else if (action === "delete_holiday") {
      const user = validateUser(usersSheet, data.username, data.password);
      if (!user) {
        result = { status: "error", message: "Sesi tidak valid" };
      } else {
        deleteHoliday(ss, user.username, data.date);
        result = { status: "success", message: "Libur berhasil dihapus" };
      }
    }

    else {
      result = { status: "error", message: "Action tidak dikenali: " + action };
    }

    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 4. HOLIDAY SETTINGS FUNCTIONS (FIXED) 
// ==========================================

function getHolidaySettings(ss, username) {
  const settingsSheetName = "Settings_User_" + username;
  let sheet = ss.getSheetByName(settingsSheetName);
  
  // Default settings jika tidak ada
  const defaultSettings = {
    additionalHolidays: [],
    optionalDays: []
  };
  
  if (!sheet || sheet.getLastRow() < 1) {
    return defaultSettings;
  }
  
  try {
    // PERBAIKAN: Ambil data dari sel A1
    const jsonData = sheet.getRange("A1").getValue();
    
    if (!jsonData || jsonData.trim() === '') {
      return defaultSettings;
    }
    
    // Parse JSON dengan error handling
    let settings;
    try {
      settings = JSON.parse(jsonData);
    } catch (parseError) {
      console.error("Error parsing JSON from sheet:", parseError);
      return defaultSettings;
    }
    
    // Pastikan struktur lengkap
    if (!settings || typeof settings !== 'object') {
      return defaultSettings;
    }
    
    return {
      additionalHolidays: Array.isArray(settings.additionalHolidays) ? settings.additionalHolidays : [],
      optionalDays: Array.isArray(settings.optionalDays) ? settings.optionalDays : []
    };
    
  } catch (e) {
    console.error("Error getting holiday settings:", e);
    return defaultSettings;
  }
}

function saveHolidaySettings(ss, username, settings) {
  const settingsSheetName = "Settings_User_" + username;
  let sheet = ss.getSheetByName(settingsSheetName);
  
  try {
    // Validasi settings
    const validatedSettings = {
      additionalHolidays: Array.isArray(settings.additionalHolidays) ? settings.additionalHolidays : [],
      optionalDays: Array.isArray(settings.optionalDays) ? settings.optionalDays : []
    };
    
    // Konversi ke JSON
    const jsonString = JSON.stringify(validatedSettings, null, 2);
    
    if (!sheet) {
      // Buat sheet baru jika belum ada
      sheet = ss.insertSheet(settingsSheetName);
      sheet.getRange("A1").setValue(jsonString);
      sheet.getRange("A1").setFontWeight("bold");
      sheet.getRange("A1").setBackground("#7c3aed");
      sheet.getRange("A1").setFontColor("#fff");
      
      // Sembunyikan sheet
      if (!sheet.isSheetHidden()) {
        sheet.hideSheet();
      }
    } else {
      // Update sheet yang sudah ada
      sheet.getRange("A1").setValue(jsonString);
    }
    
    return true;
    
  } catch (error) {
    console.error("Error saving holiday settings:", error);
    throw error;
  }
}

// ==========================================
// 5. LEGACY HOLIDAY FUNCTIONS (For Compatibility)
// ==========================================

function getHolidaysInRange(ss, sheetName, startDate, endDate) {
  const sheet = ss.getSheetByName(sheetName);
  const holidays = [];
  
  if (!sheet || sheet.getLastRow() < 2) return holidays;
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  
  data.forEach(row => {
    const date = row[0] instanceof Date ? row[0] : new Date(row[0]);
    if (date >= startDate && date <= endDate) {
      holidays.push({
        date: Utilities.formatDate(date, "GMT+7", "yyyy-MM-dd"),
        name: row[1],
        type: row[2],
        recurring: row[3],
        source: sheetName.includes("User_") ? "user" : "global"
      });
    }
  });
  
  return holidays;
}

function saveHoliday(ss, username, holidayData) {
  const userHolidaySheetName = "Holidays_User_" + username;
  let sheet = ss.getSheetByName(userHolidaySheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(userHolidaySheetName);
    sheet.appendRow(["Tanggal", "Nama Libur", "Jenis", "Berulang"]);
    sheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#7c3aed").setFontColor("#fff").setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
    if (!sheet.isSheetHidden()) sheet.hideSheet();
  }
  
  // Cek apakah sudah ada
  const data = sheet.getDataRange().getValues();
  let found = false;
  
  for (let i = 1; i < data.length; i++) {
    const existingDate = data[i][0] instanceof Date ? 
      Utilities.formatDate(data[i][0], "GMT+7", "yyyy-MM-dd") : 
      data[i][0];
    
    if (existingDate === holidayData.date) {
      // Update existing
      sheet.getRange(i + 1, 2, 1, 3).setValues([[holidayData.name, holidayData.type, holidayData.recurring]]);
      found = true;
      break;
    }
  }
  
  if (!found) {
    // Add new
    sheet.appendRow([new Date(holidayData.date), holidayData.name, holidayData.type, holidayData.recurring]);
  }
}

function deleteHoliday(ss, username, date) {
  const userHolidaySheetName = "Holidays_User_" + username;
  const sheet = ss.getSheetByName(userHolidaySheetName);
  
  if (!sheet || sheet.getLastRow() < 2) return;
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const existingDate = data[i][0] instanceof Date ? 
      Utilities.formatDate(data[i][0], "GMT+7", "yyyy-MM-dd") : 
      data[i][0];
    
    if (existingDate === date) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
}

// ==========================================
// 6. HELPERS
// ==========================================
function validateUser(sheet, username, password) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(username) && String(data[i][2]) === String(password)) {
      return { id: data[i][0], username: data[i][1], name: data[i][3], role: data[i][4] };
    }
  }
  return null;
}

// ==========================================
// 7. UTILITY FUNCTIONS
// ==========================================
function hideUserSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheets = ss.getSheets();

  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (name.startsWith("Absensi_") || name.startsWith("Holidays_User_") || name.startsWith("Settings_User_")) {
      if (!sheet.isSheetHidden()) {
        sheet.hideSheet();
      }
    }
  });

  SpreadsheetApp.getUi().alert("Semua Sheet User telah disembunyikan.");
}

function autoHideSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheets = ss.getSheets();

  sheets.forEach(sheet => {
    const name = sheet.getName();
    if ((name.startsWith("Absensi_") || name.startsWith("Holidays_User_") || name.startsWith("Settings_User_")) && !sheet.isSheetHidden()) {
      sheet.hideSheet();
    }
  });
}

function doGet(e) {
  return ContentService.createTextOutput("System Ready V7.3 with Calendar").setMimeType(ContentService.MimeType.TEXT);
}