function runSummaryProcessor() {
  const startTime = new Date();
  Logger.log('=== SUMMARY PROCESSOR STARTED ===');
  Logger.log('Process started at: ' + startTime);

  let spreadsheet;
  try {
    Logger.log('Opening spreadsheet...');
    spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    Logger.log('Spreadsheet opened');
  } catch (e) {
    Logger.log('Error accessing spreadsheet: ' + e.message);
    return;
  }

  Logger.log('Loading parameters...');
  const paramStart = new Date();
  const { settlementRuleMap, feeMap } = loadParameterData(spreadsheet);
  const paramDuration = ((new Date() - paramStart) / 1000).toFixed(2);
  Logger.log('Parameters loaded: ' + settlementRuleMap.size + ' rules, ' + feeMap.size + ' fees (in ' + paramDuration + 's)');

  Logger.log('Loading Malaysia holidays...');
  const holidayStart = new Date();
  const holidaySet = loadMalaysiaHolidays();
  const holidayDuration = ((new Date() - holidayStart) / 1000).toFixed(2);
  Logger.log('Holidays loaded: ' + holidaySet.size + ' (in ' + holidayDuration + 's)');

  Logger.log('Reading Import Data sheet...');
  const importSheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.IMPORT_DATA);
  if (!importSheet) {
    Logger.log('Error: Import Data sheet not found');
    return;
  }

  const readStart = new Date();
  const dataRange = importSheet.getDataRange();
  const allData = dataRange.getValues();
  
  if (allData.length <= 1) {
    Logger.log('Error: No data found in Import Data sheet');
    return;
  }

  const mergedData = allData.slice(1);
  const readDuration = ((new Date() - readStart) / 1000).toFixed(2);
  Logger.log('Data read: ' + mergedData.length + ' rows (in ' + readDuration + 's)');

  Logger.log('Generating summary...');
  const summaryStart = new Date();
  const summaryData = generateSummary(mergedData, settlementRuleMap, holidaySet, feeMap);
  const summaryDuration = ((new Date() - summaryStart) / 1000).toFixed(2);
  Logger.log('Summary generated: ' + summaryData.length + ' rows (in ' + summaryDuration + 's)');

  const summarySheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.SUMMARY);
  if (!summarySheet) {
    Logger.log('Error: Summary sheet not found');
    return;
  }

  Logger.log('Writing summary...');
  const writeStart = new Date();
  writeToSummarySheet(summarySheet, summaryData, CONFIG.BATCH_SIZE);
  const writeDuration = ((new Date() - writeStart) / 1000).toFixed(2);
  Logger.log('Summary written (in ' + writeDuration + 's)');

  const endTime = new Date();
  const duration = ((endTime - startTime) / 1000).toFixed(2);
  
  Logger.log('=== SUMMARY PROCESSOR COMPLETED ===');
  Logger.log('Total duration: ' + duration + 's');
  Logger.log('Import Data rows: ' + mergedData.length);
  Logger.log('Summary rows: ' + summaryData.length);
}

function generateSummary(mergedData, settlementRuleMap, holidaySet, feeMap) {
  const kiraMap = new Map();
  const pgMap = new Map();

  mergedData.forEach((row) => {
    const createdOn = row[0];
    const merchant = row[1];
    const kiraAmount = row[5];
    const pgMerchant = row[6];
    const pgChannel = row[7];

    if (pgMerchant === 'No Data' || pgChannel === 'No Data') {
      return;
    }

    const transactionDate = extractDate(createdOn);
    if (!transactionDate) return;

    const kiraKey = pgMerchant + '|' + pgChannel + '|' + transactionDate;
    
    if (!kiraMap.has(kiraKey)) {
      kiraMap.set(kiraKey, 0);
    }

    if (typeof kiraAmount === 'number' && kiraAmount > 0) {
      kiraMap.set(kiraKey, kiraMap.get(kiraKey) + kiraAmount);
    }
  });

  mergedData.forEach((row) => {
    const createdOn = row[0];
    const merchant = row[1];
    const pgMerchant = row[6];
    const pgChannel = row[7];
    const pgTransactionDate = row[8];
    const pgAmount = row[9];

    if (pgMerchant === 'No Data' || pgChannel === 'No Data') {
      return;
    }

    const transactionDate = extractDate(createdOn);
    if (!transactionDate) return;

    const pgDate = pgTransactionDate === 'No Data' ? 'No Data' : extractDate(pgTransactionDate) || 'No Data';
    const pgKey = pgMerchant + '|' + pgChannel + '|' + pgDate;

    if (!pgMap.has(pgKey)) {
      const settlementRule = getSettlementRule(merchant, pgChannel, settlementRuleMap);
      const settlementDate = calculateSettlementDate(pgDate === 'No Data' ? transactionDate : pgDate, settlementRule, holidaySet);
      
      pgMap.set(pgKey, {
        pgMerchant: pgMerchant,
        channel: pgChannel,
        pgDate: pgDate,
        settlementRule: settlementRule,
        settlementDate: settlementDate,
        amountPG: 0,
        transactionDates: new Set(),
      });
    }

    const summary = pgMap.get(pgKey);
    summary.transactionDates.add(transactionDate);

    if (typeof pgAmount === 'number' && pgAmount > 0) {
      summary.amountPG += pgAmount;
    }
  });

  const summaryMap = new Map();
  
  pgMap.forEach((pgSummary, pgKey) => {
    const parts = pgKey.split('|');
    const pgMerchant = parts[0];
    const pgChannel = parts[1];
    const pgDate = parts[2];
    
    pgSummary.transactionDates.forEach((transactionDate) => {
      const kiraKey = pgMerchant + '|' + pgChannel + '|' + transactionDate;
      const kiraAmount = kiraMap.get(kiraKey) || 0;
      
      const summaryKey = pgMerchant + '|' + pgChannel + '|' + transactionDate + '|' + pgDate;
      summaryMap.set(summaryKey, {
        pgMerchant: pgMerchant,
        channel: pgChannel,
        transactionDate: transactionDate,
        pgDate: pgDate,
        settlementRule: pgSummary.settlementRule,
        settlementDate: pgSummary.settlementDate,
        amountPG: pgSummary.amountPG,
        kiraAmount: kiraAmount,
      });
    });
  });
  
  kiraMap.forEach((kiraAmount, kiraKey) => {
    const parts = kiraKey.split('|');
    const pgMerchant = parts[0];
    const pgChannel = parts[1];
    const transactionDate = parts[2];
    
    let hasAnyPGData = false;
    for (const key of summaryMap.keys()) {
      if (key.startsWith(pgMerchant + '|' + pgChannel + '|' + transactionDate + '|')) {
        hasAnyPGData = true;
        break;
      }
    }
    
    if (!hasAnyPGData) {
      let merchant = '';
      for (const row of mergedData) {
        if (row[6] === pgMerchant && row[7] === pgChannel) {
          merchant = row[1];
          break;
        }
      }
      
      const settlementRule = getSettlementRule(merchant, pgChannel, settlementRuleMap);
      const settlementDate = calculateSettlementDate(transactionDate, settlementRule, holidaySet);
      
      const noDataKey = pgMerchant + '|' + pgChannel + '|' + transactionDate + '|No Data';
      summaryMap.set(noDataKey, {
        pgMerchant: pgMerchant,
        channel: pgChannel,
        transactionDate: transactionDate,
        pgDate: 'No Data',
        settlementRule: settlementRule,
        settlementDate: settlementDate,
        amountPG: 0,
        kiraAmount: kiraAmount,
      });
    }
  });

  const summaryArray = Array.from(summaryMap.values());

  summaryArray.sort((a, b) => {
    const dateCompare = a.transactionDate.localeCompare(b.transactionDate);
    if (dateCompare !== 0) {
      return dateCompare;
    }
    const merchantCompare = a.pgMerchant.localeCompare(b.pgMerchant);
    if (merchantCompare !== 0) {
      return merchantCompare;
    }
    return a.channel.localeCompare(b.channel);
  });

  const output = summaryArray.map((s, index) => {
    const month = extractMonth(s.pgDate === 'No Data' ? s.transactionDate : s.pgDate);
    const feeRate = getFeeRate(month, s.pgMerchant, s.channel, feeMap);
    const fees = s.amountPG * feeRate;
    const settlementAmount = s.amountPG - fees;

    const pgKiraDailyVariance = s.amountPG - s.kiraAmount;

    let pgKiraCumulativeVariance = pgKiraDailyVariance;

    if (index > 0) {
      const prev = summaryArray[index - 1];
      pgKiraCumulativeVariance = (prev.pgKiraCumulativeVariance || 0) + pgKiraDailyVariance;
    }

    summaryArray[index].pgKiraCumulativeVariance = pgKiraCumulativeVariance;

    return [
      s.pgMerchant,
      s.channel,
      s.transactionDate,
      s.kiraAmount || 0,
      s.pgDate,
      s.amountPG === 0 ? 'No Data' : s.amountPG,
      s.settlementRule,
      s.settlementDate,
      fees,
      settlementAmount,
      pgKiraDailyVariance,
      pgKiraCumulativeVariance,
      '',
      '',
      '',
      '',
    ];
  });

  return output;
}
