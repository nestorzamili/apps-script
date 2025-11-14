function writeToSheet(sheet, newData, batchSize) {
  const headers = [
    'Created On',
    'Merchant',
    'Transaction ID',
    'Merchant Order ID',
    'Payment Method',
    'Transaction Amount',
    'PG Merchant',
    'PG Channel',
    'PG Transaction Date',
    'PG Amount',
    'Bank Merchant',
    'Bank Channel',
    'Bank Transaction Date',
    'Bank Amount',
    'Remarks',
  ];

  Logger.log('Writing data: ' + newData.length + ' rows (batch size: ' + batchSize + ')');
  
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  for (let i = 0; i < newData.length; i += batchSize) {
    const batch = newData.slice(i, Math.min(i + batchSize, newData.length));
    const startRow = i + 2;
    const progress = ((i + batch.length) / newData.length * 100).toFixed(1);
    
    Logger.log('Writing batch: rows ' + (i + 1) + '-' + (i + batch.length) + ' (' + progress + '%)');
    
    writeBatchWithRetry(sheet, startRow, batch, headers.length);
  }
  
  Logger.log('Write completed: ' + newData.length + ' rows');
  return newData.length;
}

function writeBatchWithRetry(sheet, startRow, batch, numCols, maxRetries = 3) {
  let attempt = 0;
  while (attempt < maxRetries) {
    try {
      const currentMaxRows = sheet.getMaxRows();
      const neededRows = startRow + batch.length - 1;
      if (neededRows > currentMaxRows) {
        sheet.insertRowsAfter(currentMaxRows, neededRows - currentMaxRows);
      }
      
      sheet.getRange(startRow, 1, batch.length, numCols).setValues(batch);
      SpreadsheetApp.flush();
      return;
    } catch (e) {
      attempt++;
      if (attempt >= maxRetries) {
        Logger.log('Failed after ' + maxRetries + ' attempts: ' + e.message);
        throw e;
      }
      Logger.log('Retry ' + attempt + '/' + maxRetries + ': ' + e.message);
      Utilities.sleep(2000 * attempt);
    }
  }
}

function writeToSummarySheet(sheet, newData, batchSize) {
  const headers = [
    'PG Merchant',
    'Channel',
    'Transaction Date',
    'Kira Amount',
    'PG Date',
    'Amount PG',
    'Settlement Rule',
    'Settlement Date',
    'Fees',
    'Settlement Amount',
    'PG KIRA Daily Variance',
    'PG KIRA Cumulative Variance',
    'Amount RHB',
    'PG RHB Daily Variance',
    'PG RHB Cumulative Variance',
    'Remarks',
  ];

  Logger.log('Comparing with existing data...');
  
  const existingData = sheet.getDataRange().getValues();
  const hasExistingData = existingData.length > 1;
  
  if (!hasExistingData) {
    Logger.log('No existing data - writing all new data: ' + newData.length + ' rows');
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    const dataToWrite = newData.map(row => {
      const newRow = row.slice();
      newRow[13] = '';
      newRow[14] = '';
      newRow[15] = '';
      newRow[16] = '';
      return newRow;
    });
    
    for (let i = 0; i < dataToWrite.length; i += batchSize) {
      const batch = dataToWrite.slice(i, Math.min(i + batchSize, dataToWrite.length));
      const startRow = i + 2;
      const progress = ((i + batch.length) / dataToWrite.length * 100).toFixed(1);
      
      Logger.log('Writing batch: rows ' + (i + 1) + '-' + (i + batch.length) + ' (' + progress + '%)');
      
      const currentMaxRows = sheet.getMaxRows();
      const neededRows = startRow + batch.length - 1;
      if (neededRows > currentMaxRows) {
        sheet.insertRowsAfter(currentMaxRows, neededRows - currentMaxRows);
      }
      
      sheet.getRange(startRow, 1, batch.length, 13).setValues(batch.map(row => row.slice(0, 13)));
      
      SpreadsheetApp.flush();
    }
    
    Logger.log('Summary write completed: ' + dataToWrite.length + ' rows (columns L-O left empty for manual input)');
    return;
  }
  
  const existingMap = new Map();
  for (let i = 1; i < existingData.length; i++) {
    const row = existingData[i];
    const normalizedDate = extractDate(row[2]);
    const key = row[0] + '|' + row[1] + '|' + normalizedDate;
    existingMap.set(key, {
      rowIndex: i,
      amountRHB: row[11],
      dailyVariance: row[12],
      cumulativeVariance: row[13],
      remarks: row[14]
    });
  }
  
  Logger.log('Found existing data: ' + existingMap.size + ' rows');
  
  const updates = [];
  const appendRows = [];
  
  newData.forEach((row) => {
    const normalizedDate = extractDate(row[2]);
    const key = row[0] + '|' + row[1] + '|' + normalizedDate;
    
    if (existingMap.has(key)) {
      const existing = existingMap.get(key);
      
      row[13] = (existing.amountRHB !== null && existing.amountRHB !== undefined && existing.amountRHB !== '') ? existing.amountRHB : '';
      row[14] = (existing.dailyVariance !== null && existing.dailyVariance !== undefined && existing.dailyVariance !== '') ? existing.dailyVariance : '';
      row[15] = (existing.cumulativeVariance !== null && existing.cumulativeVariance !== undefined && existing.cumulativeVariance !== '') ? existing.cumulativeVariance : '';
      row[16] = (existing.remarks !== null && existing.remarks !== undefined && existing.remarks !== '') ? existing.remarks : '';
      
      updates.push({
        rowIndex: existing.rowIndex + 1,
        data: row
      });
    } else {
      appendRows.push(row);
    }
  });
  
  Logger.log('Updates: ' + updates.length + ' rows, New: ' + appendRows.length + ' rows');
  
  for (let i = 0; i < updates.length; i += batchSize) {
    const batch = updates.slice(i, Math.min(i + batchSize, updates.length));
    const progress = ((i + batch.length) / updates.length * 100).toFixed(1);
    
    Logger.log('Updating batch: ' + (i + 1) + '-' + (i + batch.length) + ' (' + progress + '%)');
    
    batch.forEach((update) => {
      sheet.getRange(update.rowIndex, 1, 1, headers.length).setValues([update.data]);
    });
    
    SpreadsheetApp.flush();
  }
  
  if (appendRows.length > 0) {
    const startRow = existingData.length + 1;
    
    const currentMaxRows = sheet.getMaxRows();
    const neededRows = startRow + appendRows.length - 1;
    if (neededRows > currentMaxRows) {
      sheet.insertRowsAfter(currentMaxRows, neededRows - currentMaxRows);
    }
    
    Logger.log('Appending new rows: ' + appendRows.length + ' (columns L-O left empty for manual input)');
    
    for (let i = 0; i < appendRows.length; i += batchSize) {
      const batch = appendRows.slice(i, Math.min(i + batchSize, appendRows.length));
      const batchStartRow = startRow + i;
      const progress = ((i + batch.length) / appendRows.length * 100).toFixed(1);
      
      Logger.log('Appending batch: rows ' + (i + 1) + '-' + (i + batch.length) + ' (' + progress + '%)');
      
      const currentBatchMaxRows = sheet.getMaxRows();
      const neededBatchRows = batchStartRow + batch.length - 1;
      if (neededBatchRows > currentBatchMaxRows) {
        sheet.insertRowsAfter(currentBatchMaxRows, neededBatchRows - currentBatchMaxRows);
      }
      
      sheet.getRange(batchStartRow, 1, batch.length, 11).setValues(batch.map(row => row.slice(0, 11)));
      
      SpreadsheetApp.flush();
    }
  }
  
  Logger.log('Summary write completed: ' + updates.length + ' updated, ' + appendRows.length + ' appended');
}

function applyConditionalFormatting(sheet, dataLength) {
  if (dataLength === 0) return;

  const dataRange = sheet.getRange(2, 1, dataLength, 15);

  sheet.clearConditionalFormatRules();

  const rules = [];

  const greyRulePG = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(
      '=AND(OR($G2="No Data", $H2="No Data", $I2="No Data", $J2="No Data"), OR($N2="No Data", $F2<>$N2))',
    )
    .setBackground('#d9d9d9')
    .setRanges([dataRange])
    .build();
  rules.push(greyRulePG);

  const greyRuleBank = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(
      '=AND(OR($K2="No Data", $L2="No Data", $M2="No Data", $N2="No Data"), OR($J2="No Data", $F2<>$J2))',
    )
    .setBackground('#e8e8e8')
    .setRanges([dataRange])
    .build();
  rules.push(greyRuleBank);

  const yellowRulePG = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($F2<>"", $J2<>"", $F2<>$J2, $J2<>"No Data", OR($N2="No Data", $F2<>$N2))')
    .setBackground('#ffff00')
    .setRanges([dataRange])
    .build();
  rules.push(yellowRulePG);

  const orangeRuleBank = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($F2<>"", $N2<>"", $F2<>$N2, $N2<>"No Data", OR($J2="No Data", $F2<>$J2))')
    .setBackground('#ffa500')
    .setRanges([dataRange])
    .build();
  rules.push(orangeRuleBank);

  const redRuleBoth = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(
      '=AND($F2<>"", $J2<>"", $N2<>"", $F2<>$J2, $F2<>$N2, $J2<>"No Data", $N2<>"No Data")',
    )
    .setBackground('#ff0000')
    .setFontColor('#ffffff')
    .setRanges([dataRange])
    .build();
  rules.push(redRuleBoth);

  sheet.setConditionalFormatRules(rules);
  Logger.log(
    'Applied 5 conditional format rules to range ' + dataRange.getA1Notation(),
  );
}
