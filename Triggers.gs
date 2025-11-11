function onEdit(e) {
  if (!e || !e.range || !e.source) return;
  
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  
  if (sheetName === CONFIG.SHEET_NAMES.SUMMARY) {
    handleSummaryEdit(sheet, row, col);
  } else if (sheetName === "Merchants Balance & Settlement Ledger") {
    handleMerchantLedgerEdit(sheet, range, row, col);
  } else if (sheetName === "Agents Balance & Settlement Ledger") {
    handleAgentLedgerEdit(sheet, range, row, col);
  }
}

function handleSummaryEdit(sheet, row, col) {
  if (col !== 12 || row < 2) return;
  updateRHBVariance(sheet);
}

function handleMerchantLedgerEdit(sheet, range, row, col) {
  if (range.getA1Notation() === "B1" || range.getA1Notation() === "B2") {
    importmerchant();
    return;
  }
  
  if (row >= 5 && [13, 14, 15, 19].includes(col)) {
    saveManualInput(sheet, row);
  }
}

function handleAgentLedgerEdit(sheet, range, row, col) {
  if (range.getA1Notation() === "B1" || range.getA1Notation() === "B2") {
    importagent();
    return;
  }
  
  if (row >= 5 && [2, 4, 10].includes(col)) {
    saveAgentManualInput(sheet, row);
    updateSingleRow(sheet, row);
  }
}

function updateRHBVariance(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;
  
  const updates = [];
  let cumulativeVariance = 0;
  
  for (let i = 1; i < data.length; i++) {
    const amountPG = data[i][5];
    const amountRHB = data[i][11];
    
    let dailyVariance = '';
    let cumVariance = '';
    
    if (amountRHB !== '' && amountRHB !== null && amountRHB !== undefined && typeof amountRHB === 'number') {
      dailyVariance = amountPG - amountRHB;
      cumulativeVariance += dailyVariance;
      cumVariance = cumulativeVariance;
    }
    
    updates.push([dailyVariance, cumVariance]);
  }
  
  if (updates.length > 0) {
    sheet.getRange(2, 13, updates.length, 2).setValues(updates);
  }
}
