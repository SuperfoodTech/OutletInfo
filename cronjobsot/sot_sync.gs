/**
 * Google Apps Script: SOT & DBR Sync Engine
 * =========================================
 * Lokasi: /cronjobsot/sot_sync.gs
 * Target Spreadsheet: 15_Xx5ixOcxy0L90U_BrFHfVnAgjIZpiiGBEK4Ce4SyI
 * - Tab 'SOT' (GID: 2135653103): Menerima upsert 37 kolom standar + kolom ke-38 'Terakhir Diperbaharui'
 * - Tab 'DBR' (GID: 0): Menyediakan data master akun owner & kredensial
 *
 * Catatan Deployment:
 * 1. Buka Google Spreadsheet di atas -> Ekstensi -> Apps Script
 * 2. Buat file baru atau ganti isinya dengan kode ini.
 * 3. Deploy -> Deployment Baru / Kelola Deployment -> Jenis 'Aplikasi Web' (Web App)
 *    - Jalankan sebagai: Saya (akun Anda)
 *    - Siapa yang memiliki akses: Siapa saja (Anyone)
 * 4. Masukkan URL deployment ke .env sebagai SOT_APP_SCRIPT_URL (atau APP_SCRIPT_URL).
 */

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // ── Aksi 1: Sinkronisasi Outlet ke Tab SOT (Upsert) ──
    if (data.action === "sync_outlets" || data.action === "sync_sot") {
      var syncRes = handleSyncOutlets(data);
      return ContentService.createTextOutput(JSON.stringify(syncRes)).setMimeType(ContentService.MimeType.JSON);
    }

    // ── Aksi 2: Membaca Data Tab DBR (Master Owner & Kredensial) ──
    if (data.action === "get_dbr" || data.action === "read_dbr") {
      var dbrRes = handleGetDbr(data);
      return ContentService.createTextOutput(JSON.stringify(dbrRes)).setMimeType(ContentService.MimeType.JSON);
    }

    // ── Aksi 3: Upload File Excel ke Google Drive (Opsional jika dijadikan satu deployment) ──
    if (data.fileData || data.fileBase64) {
      return handleDriveUpload(data);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: "Aksi tidak dikenali: " + data.action
    })).setMimeType(ContentService.MimeType.JSON);

  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  var action = e && e.parameter ? e.parameter.action : "";
  if (action === "get_dbr" || action === "read_dbr") {
    try {
      var dbrRes = handleGetDbr({
        spreadsheetId: e.parameter.spreadsheetId || "15_Xx5ixOcxy0L90U_BrFHfVnAgjIZpiiGBEK4Ce4SyI",
        sheetName: e.parameter.sheetName || "DBR",
        gid: e.parameter.gid || "0"
      });
      return ContentService.createTextOutput(JSON.stringify(dbrRes)).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "error",
        message: err.toString()
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: "ready",
    message: "Google Sheets SOT/DBR Sync Web App aktif dan siap menerima request."
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Menangani aksi upsert outlet ke tab 'SOT' pada Google Spreadsheet target.
 * Kolom terakhir 'Terakhir Diperbaharui' akan diperbarui jika outlet sudah ada (ditimpa)
 * atau diisi jika outlet baru ditambahkan.
 */
function handleSyncOutlets(data) {
  var spreadsheetId = data.spreadsheetId || "15_Xx5ixOcxy0L90U_BrFHfVnAgjIZpiiGBEK4Ce4SyI";
  var sheetName = data.sheetName || "SOT";
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    // Fallback pencarian berdasarkan GID jika nama tab berbeda
    var sheets = ss.getSheets();
    for (var s = 0; s < sheets.length; s++) {
      if (String(sheets[s].getSheetId()) === String(data.gid || "2135653103")) {
        sheet = sheets[s];
        break;
      }
    }
  }
  
  if (!sheet) {
    throw new Error("Tab sheet '" + sheetName + "' (GID: " + (data.gid || "2135653103") + ") tidak ditemukan di Spreadsheet.");
  }
  
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  
  var headers = [];
  if (lastRow > 0 && lastCol > 0) {
    headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  }
  
  // Jika sheet kosong, inisialisasi headers dari payload
  var incomingHeaders = data.headers || [];
  if (headers.length === 0 || !headers[0]) {
    headers = incomingHeaders.slice();
    if (headers.indexOf("Terakhir Diperbaharui") === -1) {
      headers.push("Terakhir Diperbaharui");
    }
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    lastRow = 1;
    lastCol = headers.length;
  }
  
  // Pastikan kolom terakhir adalah 'Terakhir Diperbaharui'
  var timestampColIdx = -1;
  for (var h = 0; h < headers.length; h++) {
    if (String(headers[h] || "").trim().toLowerCase() === "terakhir diperbaharui") {
      timestampColIdx = h;
      break;
    }
  }
  if (timestampColIdx === -1) {
    timestampColIdx = headers.length;
    headers.push("Terakhir Diperbaharui");
    sheet.getRange(1, headers.length).setValue("Terakhir Diperbaharui").setFontWeight("bold");
    lastCol = headers.length;
  }
  
  // Indeks kolom penting untuk composite key
  var appColIdx = -1;
  var storeIdColIdx = -1;
  var ownerColIdx = -1;
  var outletColIdx = -1;
  
  for (var h = 0; h < headers.length; h++) {
    var hName = String(headers[h] || "").trim().toLowerCase();
    if (hName === "aplikator" || hName === "aplikasi") appColIdx = h;
    else if (hName === "store id") storeIdColIdx = h;
    else if (hName === "nama pemilik" || hName === "owner") ownerColIdx = h;
    else if (hName === "nama brand" || hName === "nama outlet" || hName === "outlet" || hName === "nama listing") {
      if (outletColIdx === -1) outletColIdx = h;
    }
  }
  
  // Membaca baris existing
  var existingRows = [];
  if (lastRow > 1) {
    existingRows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  }
  
  function getCompositeKey(rowArray) {
    var app = appColIdx !== -1 ? String(rowArray[appColIdx] || "").trim().toLowerCase() : "";
    var sid = storeIdColIdx !== -1 ? String(rowArray[storeIdColIdx] || "").trim() : "";
    if (sid.endsWith(".0")) sid = sid.slice(0, -2);
    
    if (sid && sid !== "-" && sid !== "nan" && sid !== "none") {
      return app + "_" + sid;
    }
    var owner = ownerColIdx !== -1 ? String(rowArray[ownerColIdx] || "").trim().toLowerCase() : "";
    var outlet = outletColIdx !== -1 ? String(rowArray[outletColIdx] || "").trim().toLowerCase() : "";
    return app + "_" + owner + "_" + outlet;
  }
  
  var keyToRowIdx = {};
  for (var r = 0; r < existingRows.length; r++) {
    var k = getCompositeKey(existingRows[r]);
    if (k && k !== "_" && k !== "__") {
      keyToRowIdx[k] = r;
    }
  }
  
  // Waktu WIB sekarang
  var nowStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");
  
  var incomingRows = data.rows || [];
  var updatedCount = 0;
  var insertedCount = 0;
  
  for (var i = 0; i < incomingRows.length; i++) {
    var inRow = incomingRows[i];
    var rowValues = [];
    if (Array.isArray(inRow)) {
      rowValues = inRow.slice(0, headers.length);
      while (rowValues.length < headers.length) {
        rowValues.push("");
      }
    } else {
      for (var c = 0; c < headers.length; c++) {
        var colName = headers[c];
        rowValues.push(inRow[colName] !== undefined ? inRow[colName] : "");
      }
    }
    
    // Perbarui nilai kolom Terakhir Diperbaharui
    rowValues[timestampColIdx] = nowStr;
    
    // Cek apakah data sudah ada (overwrite) atau baru (insert)
    var inKey = getCompositeKey(rowValues);
    if (inKey in keyToRowIdx) {
      var targetIdx = keyToRowIdx[inKey];
      existingRows[targetIdx] = rowValues;
      updatedCount++;
    } else {
      existingRows.push(rowValues);
      keyToRowIdx[inKey] = existingRows.length - 1;
      insertedCount++;
    }
  }
  
  // Batch write ke sheet jika ada baris
  if (existingRows.length > 0) {
    sheet.getRange(2, 1, existingRows.length, headers.length).setValues(existingRows);
  }
  
  return {
    status: "success",
    spreadsheetId: spreadsheetId,
    sheetName: sheet.getName(),
    gid: sheet.getSheetId(),
    updated: updatedCount,
    inserted: insertedCount,
    totalRows: existingRows.length,
    timestamp: nowStr
  };
}

/**
 * Membaca data master kredensial & owner dari tab 'DBR' (GID: 0)
 */
function handleGetDbr(data) {
  var spreadsheetId = data.spreadsheetId || "15_Xx5ixOcxy0L90U_BrFHfVnAgjIZpiiGBEK4Ce4SyI";
  var sheetName = data.sheetName || "DBR";
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    var sheets = ss.getSheets();
    for (var s = 0; s < sheets.length; s++) {
      if (String(sheets[s].getSheetId()) === String(data.gid || "0")) {
        sheet = sheets[s];
        break;
      }
    }
  }
  if (!sheet) {
    sheet = ss.getSheets()[0];
  }
  
  var values = sheet.getDataRange().getValues();
  if (!values || values.length === 0) {
    return { status: "success", headers: [], totalRows: 0, rows: [] };
  }
  
  var headers = values[0];
  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var rowObj = {};
    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c] || "").trim();
      if (h) {
        rowObj[h] = values[r][c];
      }
    }
    rows.push(rowObj);
  }
  
  return {
    status: "success",
    spreadsheetId: spreadsheetId,
    sheetName: sheet.getName(),
    gid: sheet.getSheetId(),
    headers: headers,
    totalRows: rows.length,
    data: rows
  };
}

/**
 * Upload File Excel ke Google Drive (Fallback handler)
 */
function handleDriveUpload(data) {
  var base64Str = data.fileData || data.fileBase64;
  var fileData = Utilities.base64Decode(base64Str);
  var mimeType = data.mimeType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  var blob = Utilities.newBlob(fileData, mimeType, data.fileName);

  var rootFolderId = '19VIrypPcBmNNbjDLGS7kxp_yIdBBwXjB';
  var rootFolder;
  
  try {
    rootFolder = DriveApp.getFolderById(rootFolderId);
  } catch(err) {
    var myDrive = DriveApp.getRootFolder();
    var fIter = myDrive.getFoldersByName("Outlet Info");
    if (fIter.hasNext()) {
      rootFolder = fIter.next();
    } else {
      rootFolder = myDrive.createFolder("Outlet Info");
    }
  }
  
  var targetFolder = rootFolder;
  var ownerName = data.ownerName || data.folderName;
  if (ownerName && ownerName.toString().trim() !== '') {
    var cleanOwner = ownerName.toString().trim();
    var folderIter = rootFolder.getFoldersByName(cleanOwner);
    if (folderIter.hasNext()) {
      targetFolder = folderIter.next();
    } else {
      targetFolder = rootFolder.createFolder(cleanOwner);
    }
  }
  
  var file = targetFolder.createFile(blob);
  
  return ContentService.createTextOutput(JSON.stringify({
    'status': 'success',
    'fileName': data.fileName,
    'folderName': targetFolder.getName(),
    'folderUrl': targetFolder.getUrl(),
    'fileUrl': file.getUrl(),
    'url': targetFolder.getUrl()
  })).setMimeType(ContentService.MimeType.JSON);
}

