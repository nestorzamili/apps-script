function loadParameterData(spreadsheet) {
  const parameterSheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.PARAMETER);
  
  if (!parameterSheet) {
    Logger.log('Warning: Parameter sheet not found');
    return { settlementRuleMap: new Map(), feeMap: new Map() };
  }

  const dataRange = parameterSheet.getDataRange();
  const allData = dataRange.getValues();

  if (allData.length <= 1) {
    Logger.log('Warning: No data found in Parameter sheet');
    return { settlementRuleMap: new Map(), feeMap: new Map() };
  }

  const header = allData[0];
  
  const cols = {};
  const colsSecond = {};
  const colsThird = {};
  
  for (let i = 0; i < header.length; i++) {
    const col = String(header[i]).trim().toLowerCase();
    if (!cols[col]) {
      cols[col] = i;
    } else if (!colsSecond[col]) {
      colsSecond[col] = i;
    } else if (!colsThird[col]) {
      colsThird[col] = i;
    }
  }

  const settlementRuleMap = new Map();
  const feeMap = new Map();

  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    
    const hasSettlementRule = cols['settlement rule'] !== undefined && 
                              String(row[cols['settlement rule']] || '').trim() !== '';
    const hasFeeParams = cols['month'] !== undefined && 
                        cols['pg'] !== undefined &&
                        String(row[cols['month']] || '').trim() !== '' &&
                        String(row[cols['pg']] || '').trim() !== '';
    
    if (hasSettlementRule) {
      const merchant = String(row[cols['settlement rule']] || '').trim();
      const rules = {
        FPX: String(row[colsThird['fpx']] || '').trim(),
        FPXC: String(row[colsThird['fpxc']] || '').trim(),
        ewallet: String(row[colsThird['ewallet']] || '').trim(),
      };
      settlementRuleMap.set(merchant, rules);
    }
    
    if (hasFeeParams) {
      const month = String(row[cols['month']] || '').trim();
      const pg = String(row[cols['pg']] || '').trim();
      const key = month + '|' + pg;
      
      const fees = {
        FPX: parsePercentage(row[colsSecond['fpx']]),
        ewallet: parsePercentage(row[colsSecond['ewallet']]),
        FPXC: parsePercentage(row[colsSecond['fpxc']]),
        TNG: parsePercentage(row[cols['tng']]),
        Shopee: parsePercentage(row[cols['shopee']]),
        BOOST: parsePercentage(row[cols['boost']]),
      };
      feeMap.set(key, fees);
    }
  }

  return { settlementRuleMap, feeMap };
}

function loadSettlementRules(spreadsheet) {
  return loadParameterData(spreadsheet).settlementRuleMap;
}

function loadFeeParameters(spreadsheet) {
  return loadParameterData(spreadsheet).feeMap;
}

function parsePercentage(value) {
  if (!value) return 0;
  
  const str = String(value).trim();
  
  if (str === '' || str === '0' || str === '0%') return 0;
  
  if (str.includes('%')) {
    const num = parseFloat(str.replace('%', ''));
    if (isNaN(num)) return 0;
    return num / 100;
  }
  
  const num = parseFloat(str);
  if (isNaN(num)) return 0;
  
  if (num > 0 && num < 1) {
    return num;
  }
  
  if (num >= 1) {
    return num / 100;
  }
  
  return 0;
}

function extractPGName(pgMerchant) {
  if (!pgMerchant || pgMerchant === 'No Data') {
    return '';
  }

  const parts = String(pgMerchant).trim().split(/\s+/);
  
  if (parts.length >= 2) {
    return parts[1];
  }
  
  return pgMerchant;
}

function getFeeRate(month, pgMerchant, channel, feeMap) {
  if (!month || !pgMerchant || !channel) {
    return 0;
  }

  const pgName = extractPGName(pgMerchant);
  const key = String(month) + '|' + pgName;
  const fees = feeMap.get(key);
  
  if (!fees) {
    return 0;
  }

  let feeKey = null;
  
  const channelNorm = String(channel).trim().toUpperCase();
  
  if (channelNorm === 'FPX') {
    feeKey = 'FPX';
  } else if (channelNorm === 'FPXC') {
    feeKey = 'FPXC';
  } else if (channelNorm === 'TNG' || channelNorm === 'TOUCH N GO' || channelNorm === 'TOUCHNGO') {
    feeKey = 'TNG';
  } else if (channelNorm === 'SHOPEE' || channelNorm === 'SHOPEEPAY') {
    feeKey = 'Shopee';
  } else if (channelNorm === 'BOOST') {
    feeKey = 'BOOST';
  } else {
    feeKey = 'ewallet';
  }
  
  if (fees[feeKey] !== undefined) {
    return fees[feeKey];
  }

  return 0;
}

function getSettlementRule(merchant, channel, settlementRuleMap) {
  if (!merchant || !channel || merchant === 'No Data' || channel === 'No Data') {
    return '';
  }

  const rules = settlementRuleMap.get(merchant);
  if (!rules) {
    return '';
  }

  if (channel === 'FPX') {
    return rules.FPX || '';
  } else if (channel === 'FPXC') {
    return rules.FPXC || '';
  } else {
    return rules.ewallet || '';
  }
}
