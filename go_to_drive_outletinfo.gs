function doPost(e) {
  try {
    // Folder ID sesuai dengan link Drive yang kamu berikan
    var folderId = '1NHTJxWfDLS1dw3VBH_4FnP6Sl8KBtIWg';
    var folder = DriveApp.getFolderById(folderId);
    
    var data = JSON.parse(e.postData.contents);
    var fileData = Utilities.base64Decode(data.fileData);
    var blob = Utilities.newBlob(fileData, data.mimeType, data.fileName);
    
    var file = folder.createFile(blob);
    
    return ContentService.createTextOutput(JSON.stringify({
      'status': 'success',
      'url': file.getUrl()
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      'status': 'error',
      'message': error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
