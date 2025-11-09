function runDepositProcessor() {
  const startTime = new Date();
  Logger.log('=== DEPOSIT PROCESSOR STARTED ===');
  Logger.log('Process started at: ' + startTime);
  
  Logger.log('Opening spreadsheet...');
  const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  Logger.log('Spreadsheet opened');

  Logger.log('Loading parameters...');
  const t1 = new Date();
  const { settlementRuleMap, feeMap } = loadParameterData(spreadsheet);
  const depositFeeMap = loadDepositFeeParameters(spreadsheet);
  const t2 = new Date();
  Logger.log('Parameters loaded: ' + settlementRuleMap.size + ' rules, ' + feeMap.size + ' fees, ' + depositFeeMap.size + ' deposit fees (in ' + ((t2 - t1) / 1000).toFixed(2) + 's)');

  Logger.log('Loading Malaysia holidays...');
  const t3 = new Date();
  const holidays = loadMalaysiaHolidays();
  const t4 = new Date();
  Logger.log('Holidays loaded: ' + holidays.size + ' (in ' + ((t4 - t3) / 1000).toFixed(2) + 's)');

  Logger.log('Reading Import Data sheet...');
  const t5 = new Date();
  const importSheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.IMPORT_DATA);
  if (!importSheet) {
    Logger.log('Error: Import Data sheet not found');
    return;
  }

  const importData = importSheet.getDataRange().getValues();
  if (importData.length <= 1) {
    Logger.log('No data in Import Data sheet');
    return;
  }

  const header = importData[0];
  const rows = importData.slice(1);
  const t6 = new Date();
  Logger.log('Data read: ' + rows.length + ' rows (in ' + ((t6 - t5) / 1000).toFixed(2) + 's)');

  const cols = {};
  for (let i = 0; i < header.length; i++) {
    cols[String(header[i]).trim().toLowerCase()] = i;
  }

  Logger.log('Generating deposit data...');
  const t7 = new Date();
  const depositData = generateDepositData(rows, cols, settlementRuleMap, depositFeeMap, holidays);
  const t8 = new Date();
  Logger.log('Deposit data generated: ' + depositData.length + ' rows (in ' + ((t8 - t7) / 1000).toFixed(2) + 's)');

  Logger.log('Writing deposit data...');
  const t9 = new Date();
  writeToDepositSheet(spreadsheet, depositData);
  const t10 = new Date();
  Logger.log('Deposit data written (in ' + ((t10 - t9) / 1000).toFixed(2) + 's)');

  const endTime = new Date();
  Logger.log('=== DEPOSIT PROCESSOR COMPLETED ===');
  Logger.log('Total duration: ' + ((endTime - startTime) / 1000).toFixed(2) + 's');
  Logger.log('Import Data rows: ' + rows.length);
  Logger.log('Deposit rows: ' + depositData.length);
}

function loadDepositFeeParameters(spreadsheet) {
  const parameterSheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.PARAMETER);
  
  if (!parameterSheet) {
    return new Map();
  }

  const allData = parameterSheet.getDataRange().getValues();
  if (allData.length <= 1) {
    return new Map();
  }

  const header = allData[0];
  
  const cols = {};
  for (let i = 0; i < header.length; i++) {
    const col = String(header[i]).trim().toLowerCase();
    if (!cols[col]) {
      cols[col] = i;
    }
  }

  const depositFeeMap = new Map();

  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    
    const monthValue = String(row[cols['month']] || '').trim();
    const accountValue = String(row[cols['account']] || '').trim();
    
    if (!monthValue || !accountValue) {
      continue;
    }
    
    const monthNum = parseInt(monthValue, 10);
    if (isNaN(monthNum) || monthNum < 1 || monthNum > 12) {
      continue;
    }
    
    const key = monthValue + '|' + accountValue;
    
    depositFeeMap.set(key, {
      FPX: parsePercentage(row[cols['fpx']]),
      ewallet: parsePercentage(row[cols['ewallet']]),
      FPXC: parsePercentage(row[cols['fpxc']]),
    });
  }

  return depositFeeMap;
}

function generateDepositData(rows, cols, settlementRuleMap, depositFeeMap, holidays) {
  const depositMap = new Map();
  
  let processedCount = 0;
  let skippedCount = 0;
  
  rows.forEach((row) => {
    const merchant = String(row[cols['merchant']] || '').trim();
    const pgChannel = String(row[cols['pg channel']] || '').trim();
    const pgMerchant = String(row[cols['pg merchant']] || '').trim();
    const createdOn = row[cols['created on']];
    const pgAmount = parseAmount(row[cols['pg amount']]);
    const bankAmount = parseAmount(row[cols['bank amount']]);
    
    if (!merchant || !pgChannel || pgChannel === 'No Data' || !pgMerchant || pgMerchant === 'No Data' || !createdOn) {
      skippedCount++;
      return;
    }
    
    const kiraAmount = (pgAmount !== 0) ? pgAmount : bankAmount;
    
    if (kiraAmount === 0) {
      skippedCount++;
      return;
    }
    
    processedCount++;
    
    const transactionDate = extractDate(createdOn);
    if (!transactionDate) {
      skippedCount++;
      return;
    }
    
    let channel;
    if (pgChannel === 'FPX') {
      channel = 'FPX';
    } else if (pgChannel === 'FPXC') {
      channel = 'FPXC';
    } else {
      channel = 'ewallet';
    }
    
    const key = merchant + '|' + channel + '|' + pgMerchant + '|' + transactionDate;
    
    if (!depositMap.has(key)) {
      depositMap.set(key, {
        merchant: merchant,
        channel: channel,
        pgMerchant: pgMerchant,
        transactionDate: transactionDate,
        kiraAmount: 0
      });
    }
    
    const entry = depositMap.get(key);
    entry.kiraAmount += kiraAmount;
  });
  
  Logger.log('Processed: ' + processedCount + ' rows, skipped: ' + skippedCount + ' rows, unique deposits: ' + depositMap.size);
  
  const depositData = [];
  
  depositMap.forEach((entry) => {
    if (entry.merchant === 'Test Merchant Sdn Bhd') {
      return;
    }
    
    const settlementRule = getSettlementRule(entry.merchant, entry.channel, settlementRuleMap);
    
    const settlementDate = calculateSettlementDate(entry.transactionDate, settlementRule, holidays);
    
    const fees = calculateDepositFees(entry.transactionDate, entry.merchant, entry.channel, entry.kiraAmount, depositFeeMap);
    
    const grossAmount = entry.kiraAmount - fees;
    
    depositData.push([
      entry.merchant,
      entry.channel,
      entry.pgMerchant,
      entry.transactionDate,
      settlementRule,
      settlementDate,
      entry.kiraAmount,
      fees,
      grossAmount,
      ''
    ]);
  });
  
  depositData.sort((a, b) => {
    const dateCompare = a[3].localeCompare(b[3]);
    if (dateCompare !== 0) return dateCompare;
    
    const merchantCompare = a[0].localeCompare(b[0]);
    if (merchantCompare !== 0) return merchantCompare;
    
    const channelCompare = a[1].localeCompare(b[1]);
    if (channelCompare !== 0) return channelCompare;
    
    return a[2].localeCompare(b[2]);
  });
  
  return depositData;
}

function calculateDepositFees(transactionDate, merchant, channel, amount, depositFeeMap) {
  if (!transactionDate || !merchant || !channel || amount === 0) {
    return 0;
  }
  
  const month = extractMonth(transactionDate);
  if (!month) {
    return 0;
  }
  
  const key = month + '|' + merchant;
  const fees = depositFeeMap.get(key);
  
  if (!fees) {
    return 0;
  }
  
  let feeRate = 0;
  
  if (channel === 'FPX') {
    feeRate = fees.FPX || 0;
  } else if (channel === 'FPXC') {
    feeRate = fees.FPXC || 0;
  } else {
    feeRate = fees.ewallet || 0;
  }
  
  return amount * feeRate;
}

function writeToDepositSheet(spreadsheet, depositData) {
  let depositSheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.DEPOSIT);
  
  if (!depositSheet) {
    depositSheet = spreadsheet.insertSheet(CONFIG.SHEET_NAMES.DEPOSIT);
  } else {
    depositSheet.clear();
  }
  
  const headers = [
    'Merchant',
    'Channel',
    'PG Merchant',
    'Transaction Date',
    'Settlement Rule',
    'Settlement Date',
    'Kira Amount',
    'Fees',
    'Gross Amount (Deposit)',
    'Remarks'
  ];
  
  depositSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  if (depositData.length > 0) {
    Logger.log('Writing deposit: ' + depositData.length + ' rows');
    
    for (let i = 0; i < depositData.length; i += CONFIG.BATCH_SIZE) {
      const batch = depositData.slice(i, Math.min(i + CONFIG.BATCH_SIZE, depositData.length));
      const startRow = i + 2;
      const progress = ((i + batch.length) / depositData.length * 100).toFixed(1);
      
      Logger.log('Writing deposit batch: rows ' + (i + 1) + '-' + (i + batch.length) + ' (' + progress + '%)');
      
      writeBatchWithRetry(depositSheet, startRow, batch, headers.length);
    }
    
    Logger.log('Deposit write completed: ' + depositData.length + ' rows');
    
    depositSheet.getRange(2, 7, depositData.length, 1).setNumberFormat('#,##0.##');
    depositSheet.getRange(2, 8, depositData.length, 1).setNumberFormat('#,##0.##');
    depositSheet.getRange(2, 9, depositData.length, 1).setNumberFormat('#,##0.##');
  }
}
