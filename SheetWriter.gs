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
    'Settlement Rule',
    'Settlement Date',
    'Amount PG',
    'Fees',
    'Settlement Amount',
    'Kira Amount',
    'PG KIRA Daily Variance',
    'PG KIRA Cumulative Variance',
    'Amount RHB',
    'PG RHB Daily Variance',
    'PG RHB Cumulative Variance',
    'Remarks',
  ];

  Logger.log('Writing summary: ' + newData.length + ' rows');
  
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  for (let i = 0; i < newData.length; i += batchSize) {
    const batch = newData.slice(i, Math.min(i + batchSize, newData.length));
    const startRow = i + 2;
    const progress = ((i + batch.length) / newData.length * 100).toFixed(1);
    
    Logger.log('Writing summary batch: rows ' + (i + 1) + '-' + (i + batch.length) + ' (' + progress + '%)');
    
    writeBatchWithRetry(sheet, startRow, batch, headers.length);
  }
  
  Logger.log('Summary write completed: ' + newData.length + ' rows');
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
