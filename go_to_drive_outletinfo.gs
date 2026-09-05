function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var base64Str = data.fileData || data.fileBase64;
    var fileData = Utilities.base64Decode(base64Str);
    var mimeType = data.mimeType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    var blob = Utilities.newBlob(fileData, mimeType, data.fileName);

    // Root Folder ID di Google Drive tujuan
    var rootFolderId = '19VIrypPcBmNNbjDLGS7kxp_yIdBBwXjB';
    var rootFolder;
    
    try {
      rootFolder = DriveApp.getFolderById(rootFolderId);
    } catch(err) {
      // Fallback jika akun Google yang menjalankan script belum diberi akses Editor ke folder tujuan:
      // Simpan ke folder 'Outlet Info' di My Drive akun ini agar file tetap aman tersimpan!
      var myDrive = DriveApp.getRootFolder();
      var fIter = myDrive.getFoldersByName("Outlet Info");
      if (fIter.hasNext()) {
        rootFolder = fIter.next();
      } else {
        rootFolder = myDrive.createFolder("Outlet Info");
      }
    }
    
    // Tentukan folder target (subfolder owner jika disertakan)
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
    
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      'status': 'error',
      'message': error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    'status': 'ready',
    'message': 'Google Drive Upload Web App aktif dan siap menerima data POST.'
  })).setMimeType(ContentService.MimeType.JSON);
}
