function findColumnIndex(headerRow, possibleNames) {
  const norm = (s) =>
    String(s || '')
      .toLowerCase()
      .trim()
      .replace(/[\s_\-]/g, '');
  const norms = headerRow.map((h) => norm(h));

  for (let i = 0; i < norms.length; i++) {
    for (const name of possibleNames) {
      if (norms[i] === norm(name)) return i;
    }
  }
  return -1;
}

function parseAmount(value) {
  if (!value) return 0;
  const strValue = String(value).trim();
  const cleaned = strValue.replace(/,/g, '');
  return parseFloat(cleaned) || 0;
}

function extractDate(dateTimeValue) {
  if (!dateTimeValue || dateTimeValue === 'No Data') return null;

  const str = String(dateTimeValue).trim();

  if (str.includes('-')) {
    const datePart = str.split(' ')[0];
    return datePart;
  }

  if (dateTimeValue instanceof Date) {
    const year = dateTimeValue.getFullYear();
    const month = String(dateTimeValue.getMonth() + 1).padStart(2, '0');
    const day = String(dateTimeValue.getDate()).padStart(2, '0');
    return year + '-' + month + '-' + day;
  }

  return null;
}

function extractMonth(dateStr) {
  if (!dateStr) return null;
  
  const parts = dateStr.split('-');
  if (parts.length < 2) return null;
  
  const month = parseInt(parts[1], 10);
  return String(month);
}

function processFolder(folderId, processFn, dataMap) {
  const files = DriveApp.getFolderById(folderId).getFiles();
  let fileCount = 0;
  let rowCount = 0;

  while (files.hasNext()) {
    const file = files.next();
    fileCount++;
    try {
      rowCount += processFn(file, dataMap);
    } catch (e) {
      Logger.log('Error reading file "' + file.getName() + '": ' + e);
    }
  }
  Logger.log('Files processed: ' + fileCount + ', rows: ' + rowCount);
  return { fileCount, rowCount };
}
