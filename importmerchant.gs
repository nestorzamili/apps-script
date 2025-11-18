const MONTH_NAMES = [
  'Jan',
  'Feb',
  'Mar',
  'Apr',
  'May',
  'Jun',
  'Jul',
  'Aug',
  'Sep',
  'Oct',
  'Nov',
  'Dec',
];
const TIMEZONE = 'GMT+8';
const DATE_FORMAT = 'yyyy-MM-dd';
const DATETIME_FORMAT = 'yyyy-MM-dd HH:mm:ss';

function getMonthNumber(monthName) {
  return MONTH_NAMES.indexOf(monthName) + 1;
}

function formatDate(date, format = DATE_FORMAT) {
  return Utilities.formatDate(new Date(date), TIMEZONE, format);
}

function parseNumber(value) {
  return typeof value === 'number'
    ? value
    : parseFloat(String(value).replace(/[^\d.-]/g, '')) || 0;
}

function getWithdrawalRate(paramData, merchant, month) {
  for (let i = 1; i < paramData.length; i++) {
    if (
      paramData[i][0] === month &&
      String(paramData[i][1]).trim() === String(merchant).trim()
    ) {
      return parseFloat(String(paramData[i][5]).replace('%', '')) / 100;
    }
  }
  return 0;
}

function buildDepositMaps(depData, merchant, monthNum) {
  const header = depData[0].map((h) => String(h).trim());
  const idx = {};
  header.forEach((h, i) => (idx[h] = i));

  const trxMap = new Map();
  const settleMap = new Map();

  for (let i = 1; i < depData.length; i++) {
    const row = depData[i];
    if (String(row[idx['Merchant']]).trim() !== String(merchant).trim())
      continue;

    const trxDate = new Date(row[idx['Transaction Date']]);
    const settleDate = new Date(row[idx['Settlement Date']]);
    const channel = String(row[idx['Channel']] || '').toLowerCase();
    const channelType = channel.includes('fpx') ? 'fpx' : 'ewallet';

    const kira = parseNumber(row[idx['Kira Amount']]);
    const fee = parseNumber(row[idx['Fees']]);
    const gross = parseNumber(row[idx['Gross Amount (Deposit)']]);

    if (trxDate.getMonth() + 1 === monthNum) {
      const key = formatDate(trxDate);
      if (!trxMap.has(key)) {
        trxMap.set(key, {
          fpx: { kira: 0, fee: 0, gross: 0 },
          ewallet: { kira: 0, fee: 0, gross: 0 },
        });
      }
      const map = trxMap.get(key)[channelType];
      map.kira += kira;
      map.fee += fee;
      map.gross += gross;
    }

    if (settleDate.getMonth() + 1 === monthNum) {
      const key = formatDate(settleDate);
      if (!settleMap.has(key)) settleMap.set(key, { fpx: 0, ewallet: 0 });
      settleMap.get(key)[channelType] += gross;
    }
  }

  return { trxMap, settleMap };
}

function buildManualInputMap(storeData) {
  const storeMap = new Map();
  for (let i = 1; i < storeData.length; i++) {
    storeMap.set(storeData[i][0], {
      fund: storeData[i][1],
      charges: storeData[i][2],
      withdraw: storeData[i][3],
      remarks: storeData[i][4],
    });
  }
  return storeMap;
}

function saveManualInput(sheet, row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const store = ss.getSheetByName('ManualInputStorage');
  const merchant = sheet.getRange('B1').getValue();
  const monthName = sheet.getRange('B2').getValue();
  const date = sheet.getRange(row, 1).getValue();

  if (!date || !merchant || !monthName) return;

  const monthNum = getMonthNumber(monthName);
  const dateObj = new Date(date);
  const dateMonth = dateObj.getMonth() + 1;

  if (dateMonth !== monthNum) {
    Logger.log(
      `Warning: Date ${date} does not match selected month ${monthName}`,
    );
    return;
  }

  const key = `${merchant}@${monthNum}@${formatDate(dateObj)}`;
  const values = [
    sheet.getRange(row, 13).getValue(),
    sheet.getRange(row, 14).getValue(),
    sheet.getRange(row, 15).getValue(),
    sheet.getRange(row, 19).getValue(),
  ];

  const data = store.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) {
      store.getRange(i + 1, 2, 1, 4).setValues([values]);
      return;
    }
  }

  store.appendRow([key, ...values]);
}

function importmerchant() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledger = ss.getSheetByName('Merchants Balance & Settlement Ledger');
  const deposit = ss.getSheetByName('Deposit');
  const param = ss.getSheetByName('Parameter');
  const store = ss.getSheetByName('ManualInputStorage');

  const merchant = ledger.getRange('B1').getValue();
  const monthName = ledger.getRange('B2').getValue();

  if (!merchant || !monthName) return;

  const monthNum = getMonthNumber(monthName);
  const year = 2025;
  const lastDay = new Date(year, monthNum, 0).getDate();

  const withdrawRate = getWithdrawalRate(
    param.getDataRange().getValues(),
    merchant,
    monthNum,
  );
  const { trxMap, settleMap } = buildDepositMaps(
    deposit.getDataRange().getValues(),
    merchant,
    monthNum,
  );
  const storeMap = buildManualInputMap(store.getDataRange().getValues());
  const prevMonthClosingBalance = getPreviousMonthClosingBalance(
    ledger,
    merchant,
    monthNum,
    year,
  );

  const lastRowWithData = ledger.getLastRow();
  if (lastRowWithData >= 5) {
    ledger.getRange(5, 1, lastRowWithData - 4, 19).clearContent();
  }

  const results = [];
  let prevBalance = prevMonthClosingBalance;

  for (let d = 1; d <= lastDay; d++) {
    const dateStr = formatDate(new Date(year, monthNum - 1, d));
    const t = trxMap.get(dateStr) || {
      fpx: { kira: 0, fee: 0, gross: 0 },
      ewallet: { kira: 0, fee: 0, gross: 0 },
    };
    const s = settleMap.get(dateStr) || { fpx: 0, ewallet: 0 };

    const totalGross = t.fpx.gross + t.ewallet.gross;
    const totalFee = t.fpx.fee + t.ewallet.fee;
    const availTotal = s.fpx + s.ewallet;

    const key = `${merchant}@${monthNum}@${dateStr}`;
    const manual = storeMap.get(key) || {};

    const fund = parseNumber(manual.fund);
    const charges = parseNumber(manual.charges);
    const withdraw = parseNumber(manual.withdraw);
    const withdrawCharges = withdraw * withdrawRate;

    const closing =
      prevBalance + availTotal - (fund + charges + withdraw + withdrawCharges);

    prevBalance = closing;

    results.push([
      dateStr,
      t.fpx.kira,
      t.fpx.fee,
      t.fpx.gross,
      t.ewallet.kira,
      t.ewallet.fee,
      t.ewallet.gross,
      totalGross,
      totalFee,
      s.fpx,
      s.ewallet,
      availTotal,
      fund,
      charges,
      withdraw,
      withdrawCharges,
      closing,
      formatDate(new Date(), DATETIME_FORMAT),
      manual.remarks || '',
    ]);
  }

  ledger.getRange(5, 1, results.length, results[0].length).setValues(results);
  Logger.log(`Imported ${results.length} days for ${merchant} (${monthName})`);
}

function getPreviousMonthClosingBalance(ledger, merchant, currentMonth, year) {
  const prevMonthNum = currentMonth === 1 ? 12 : currentMonth - 1;
  const prevYear = currentMonth === 1 ? year - 1 : year;
  return calculateMonthClosingBalance(merchant, prevMonthNum, prevYear);
}

function calculateMonthClosingBalance(merchant, month, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const deposit = ss.getSheetByName('Deposit');
  const param = ss.getSheetByName('Parameter');
  const store = ss.getSheetByName('ManualInputStorage');

  if (!deposit || !param || !store) return 0;

  const withdrawRate = getWithdrawalRate(
    param.getDataRange().getValues(),
    merchant,
    month,
  );
  const lastDay = new Date(year, month, 0).getDate();

  const depData = deposit.getDataRange().getValues();
  const header = depData[0].map((h) => String(h).trim());
  const idx = {};
  header.forEach((h, i) => (idx[h] = i));

  const settleMap = new Map();
  for (let i = 1; i < depData.length; i++) {
    const row = depData[i];
    if (String(row[idx['Merchant']]).trim() !== String(merchant).trim())
      continue;

    const settleDate = new Date(row[idx['Settlement Date']]);
    const channel = String(row[idx['Channel']] || '').toLowerCase();
    const gross = parseNumber(row[idx['Gross Amount (Deposit)']]);

    if (
      settleDate.getMonth() + 1 === month &&
      settleDate.getFullYear() === year
    ) {
      const key = formatDate(settleDate);
      const channelType = channel.includes('fpx') ? 'fpx' : 'ewallet';
      if (!settleMap.has(key)) settleMap.set(key, { fpx: 0, ewallet: 0 });
      settleMap.get(key)[channelType] += gross;
    }
  }

  const storeMap = buildManualInputMap(store.getDataRange().getValues());

  let cumulativeBalance = 0;
  if (month > 1 || year > 2025) {
    cumulativeBalance = getPreviousMonthClosingBalance(
      null,
      merchant,
      month,
      year,
    );
  }

  for (let d = 1; d <= lastDay; d++) {
    const dateStr = formatDate(new Date(year, month - 1, d));
    const s = settleMap.get(dateStr) || { fpx: 0, ewallet: 0 };
    const availTotal = s.fpx + s.ewallet;

    const key = `${merchant}@${month}@${dateStr}`;
    const manual = storeMap.get(key) || {};

    const withdrawCharges = parseNumber(manual.withdraw) * withdrawRate;
    cumulativeBalance +=
      availTotal -
      (parseNumber(manual.fund) +
        parseNumber(manual.charges) +
        parseNumber(manual.withdraw) +
        withdrawCharges);
  }

  return cumulativeBalance;
}

function cleanupManualInputStorage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const store = ss.getSheetByName('ManualInputStorage');
  if (!store) return;

  const data = store.getDataRange().getValues();
  const validRows = [data[0]];
  let removedCount = 0;

  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0]).trim();
    if (!key) continue;

    const parts = key.split('@');
    if (parts.length !== 3) {
      Logger.log(`Invalid key format: ${key}`);
      removedCount++;
      continue;
    }

    const month = parseInt(parts[1]);
    const dateMonth = new Date(parts[2]).getMonth() + 1;

    if (month !== dateMonth) {
      Logger.log(
        `Month mismatch in key: ${key} (month=${month}, date month=${dateMonth})`,
      );
      removedCount++;
      continue;
    }

    validRows.push(data[i]);
  }

  store.clear();
  if (validRows.length > 0) {
    store
      .getRange(1, 1, validRows.length, validRows[0].length)
      .setValues(validRows);
  }

  const message = `Cleanup complete!\nRemoved: ${removedCount} invalid entries\nKept: ${
    validRows.length - 1
  } valid entries`;
  Logger.log(message);
  SpreadsheetApp.getUi().alert(message);
}
