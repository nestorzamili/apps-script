const CONVERTER_CONFIG = {
  FOLDER_ID: ''
};

function convertXlsxToSheets() {
  const folder = DriveApp.getFolderById(CONVERTER_CONFIG.FOLDER_ID);
  const files = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);
  
  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    
    if (isExcelFile(fileName)) {
      convertFile(file, folder);
    }
  }
}

function isExcelFile(fileName) {
  const lowerName = fileName.toLowerCase();
  return lowerName.endsWith('.xlsx') || lowerName.endsWith('.xls');
}

function convertFile(file, folder) {
  const fileName = file.getName();
  const blob = file.getBlob();
  
  const newFile = {
    title: fileName.replace(/\.(xlsx|xls)$/i, ''),
    parents: [{ id: folder.getId() }],
    mimeType: MimeType.GOOGLE_SHEETS
  };
  
  Drive.Files.insert(newFile, blob, { convert: true });
}
