const EDIT_COLUMNS = {
  MERCHANT_LEDGER: {
    MANUAL_INPUT: [13, 14, 15],
    REMARKS: 19
  },
  AGENT_LEDGER: [2, 4, 10],
  SUMMARY_RHB: 13
};

const SHEET_ROW_START = 5;

function onEdit(e) {
  if (!e?.range || !e.source) return;

  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const { row, col } = { row: e.range.getRow(), col: e.range.getColumn() };

  const handlers = {
    [CONFIG.SHEET_NAMES.SUMMARY]: () => handleSummaryEdit(sheet, row, col),
    'Merchants Balance & Settlement Ledger': () => handleMerchantLedgerEdit(sheet, e.range, row, col),
    'Agents Balance & Settlement Ledger': () => handleAgentLedgerEdit(sheet, e.range, row, col)
  };

  handlers[sheetName]?.();
}

function handleSummaryEdit(sheet, row, col) {
  if (col === EDIT_COLUMNS.SUMMARY_RHB && row >= 2) {
    updateRHBVariance(sheet);
  }
}

function handleMerchantLedgerEdit(sheet, range, row, col) {
  const notation = range.getA1Notation();
  
  if (notation === 'B1' || notation === 'B2') {
    importmerchant();
    return;
  }

  if (row < SHEET_ROW_START) return;

  if (EDIT_COLUMNS.MERCHANT_LEDGER.MANUAL_INPUT.includes(col)) {
    saveManualInput(sheet, row);
    recalculateMerchantClosingBalance(sheet, row);
  } else if (col === EDIT_COLUMNS.MERCHANT_LEDGER.REMARKS) {
    saveManualInput(sheet, row);
  }
}

function handleAgentLedgerEdit(sheet, range, row, col) {
  const notation = range.getA1Notation();
  
  if (notation === 'B1' || notation === 'B2') {
    importAgent();
    return;
  }

  if (row >= SHEET_ROW_START && EDIT_COLUMNS.AGENT_LEDGER.includes(col)) {
    saveAgentManualInput(sheet, row);
    updateSingleRow(sheet, row);
  }
}

function updateRHBVariance(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  let cumulativeVariance = 0;
  const updates = data.slice(1).map(row => {
    const amountPG = row[5];
    const amountRHB = row[12];

    if (typeof amountRHB === 'number') {
      const dailyVariance = amountPG - amountRHB;
      cumulativeVariance += dailyVariance;
      return [dailyVariance, cumulativeVariance];
    }
    
    return ['', ''];
  });

  sheet.getRange(2, 14, updates.length, 2).setValues(updates);
}

function recalculateMerchantClosingBalance(sheet, startRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const merchant = sheet.getRange('B1').getValue();
  const monthName = sheet.getRange('B2').getValue();
  
  if (!merchant || !monthName) return;

  const monthNum = getMonthNumber(monthName);
  const year = 2025;
  const withdrawRate = getWithdrawalRate(
    ss.getSheetByName('Parameter').getDataRange().getValues(),
    merchant,
    monthNum
  );

  const lastRow = sheet.getLastRow();
  if (lastRow < SHEET_ROW_START) return;

  const data = sheet.getRange(SHEET_ROW_START, 1, lastRow - 4, 19).getValues();
  
  let initialBalance = 0;
  if (startRow === SHEET_ROW_START) {
    initialBalance = getPreviousMonthClosingBalance(sheet, merchant, monthNum, year);
  }

  const updates = [];
  for (let i = startRow - SHEET_ROW_START; i < data.length; i++) {
    if (!data[i][0]) break;

    let prevBalance;
    if (i === 0) {
      prevBalance = initialBalance;
    } else if (i === startRow - SHEET_ROW_START && startRow > SHEET_ROW_START) {
      prevBalance = parseNumber(sheet.getRange(startRow - 1, 17).getValue());
    } else {
      prevBalance = parseNumber(data[i - 1][16]);
    }

    const availTotal = parseNumber(data[i][11]);
    const fund = parseNumber(data[i][12]);
    const charges = parseNumber(data[i][13]);
    const withdraw = parseNumber(data[i][14]);
    const withdrawCharges = withdraw * withdrawRate;
    const closing = prevBalance + availTotal - (fund + charges + withdraw + withdrawCharges);

    updates.push([withdrawCharges, closing, formatDate(new Date(), DATETIME_FORMAT)]);
  }

  if (updates.length > 0) {
    sheet.getRange(startRow, 16, updates.length, 3).setValues(updates);
  }
}
