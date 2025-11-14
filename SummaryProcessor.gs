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
  const summaryMap = new Map();

  mergedData.forEach((row) => {
    const createdOn = row[0];
    const merchant = row[1];
    const kiraAmount = row[5];
    const pgMerchant = row[6];
    const pgChannel = row[7];
    const pgTransactionDate = row[8];
    const pgAmount = row[9];
    const bankAmount = row[13];

    if (pgMerchant === 'No Data' || pgChannel === 'No Data') {
      return;
    }

    const transactionDate = extractDate(createdOn);
    if (!transactionDate) return;

    const key = pgMerchant + '|' + pgChannel + '|' + transactionDate;

    if (!summaryMap.has(key)) {
      const settlementRule = getSettlementRule(merchant, pgChannel, settlementRuleMap);
      const settlementDate = calculateSettlementDate(transactionDate, settlementRule, holidaySet);
      
      summaryMap.set(key, {
        pgMerchant: pgMerchant,
        channel: pgChannel,
        transactionDate: transactionDate,
        settlementRule: settlementRule,
        settlementDate: settlementDate,
        pgTransactionDate: pgTransactionDate === 'No Data' ? 'No Data' : extractDate(pgTransactionDate) || 'No Data',
        amountPG: 0,
        kiraAmount: 0,
      });
    }

    const summary = summaryMap.get(key);

    if (typeof pgAmount === 'number' && pgAmount > 0) {
      summary.amountPG += pgAmount;
    }

    if (typeof kiraAmount === 'number' && kiraAmount > 0) {
      summary.kiraAmount += kiraAmount;
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
    const month = extractMonth(s.transactionDate);
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
      s.pgTransactionDate || 'No Data',
      s.amountPG === 0 ? 'No Data' : (s.amountPG || 0),
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
