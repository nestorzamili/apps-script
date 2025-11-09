function main() {
  const startTime = new Date();
  Logger.log('Process started at: ' + startTime);

  const kiraMap = new Map();
  const pgMap = new Map();
  const bankMap = new Map();

  Logger.log('Reading KIRA folder...');
  processFolder(CONFIG.FOLDER_IDS.KIRA, processKiraFile, kiraMap);

  Logger.log('Reading PG folder...');
  processFolder(CONFIG.FOLDER_IDS.PG, processPGFile, pgMap);

  Logger.log('Reading BANK folder...');
  processFolder(CONFIG.FOLDER_IDS.BANK, processBankFile, bankMap);

  Logger.log('Merging data...');
  const mergeStart = new Date();
  const { output, stats } = mergeData(kiraMap, pgMap, bankMap);
  const mergeDuration = ((new Date() - mergeStart) / 1000).toFixed(2);

  Logger.log('Total merged: ' + output.length + ' (in ' + mergeDuration + 's)');

  let spreadsheet;
  try {
    Logger.log('Opening spreadsheet...');
    spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    Logger.log('Spreadsheet opened');
  } catch (e) {
    Logger.log('Error accessing spreadsheet: ' + e.message);
    return;
  }

  Logger.log('Getting Import Data sheet...');
  const importSheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.IMPORT_DATA);
  if (!importSheet) {
    Logger.log('Error: Import Data sheet not found');
    return;
  }

  Logger.log('Writing Import Data...');
  const writeStart = new Date();
  const newRowsCount = writeToSheet(importSheet, output, CONFIG.BATCH_SIZE);
  const writeDuration = ((new Date() - writeStart) / 1000).toFixed(2);
  Logger.log('Write completed in ' + writeDuration + 's');

  Logger.log('Applying conditional formatting...');
  const formatStart = new Date();
  applyConditionalFormatting(importSheet, output.length);
  const formatDuration = ((new Date() - formatStart) / 1000).toFixed(2);
  Logger.log('Formatting applied in ' + formatDuration + 's');

  const endTime = new Date();
  const duration = ((endTime - startTime) / 1000).toFixed(2);
  
  Logger.log('=== IMPORT DATA COMPLETED ===');
  Logger.log('Duration: ' + duration + 's');
  Logger.log('Import Data rows: ' + output.length);
  Logger.log('New rows added: ' + newRowsCount);
  Logger.log('Mismatched PG: ' + stats.mismatchPG);
  Logger.log('Mismatched Bank: ' + stats.mismatchBank);
  Logger.log('No data PG: ' + stats.noDataPG);
  Logger.log('No data Bank: ' + stats.noDataBank);
  Logger.log('');
  Logger.log('To generate summary, run: runSummaryProcessor()');
}
