function saveAgentManualInput(sheet, row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const storage = ss.getSheetByName("AgentStorage");
  if (!storage) return;

  const merchant = sheet.getRange("B1").getValue();
  const monthName = sheet.getRange("B2").getValue();
  const date = sheet.getRange(row, 1).getValue();
  if (!merchant || !monthName || !date) return;

  const monthNum = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"].indexOf(monthName)+1;
  const key = `${merchant}@${monthNum}@${Utilities.formatDate(new Date(date),"GMT+8","yyyy-MM-dd")}`;

  const rateFPX = sheet.getRange(row,2).getValue();
  const rateEwallet = sheet.getRange(row,4).getValue();
  const withdrawal = sheet.getRange(row,10).getValue();

  // Fetch existing data starting from row 2
  const lastRow = storage.getLastRow();
  const data = lastRow >= 2 ? storage.getRange(2,1,lastRow-1,4).getValues() : [];
  let found = false;

  for (let i=0;i<data.length;i++) {
    if (String(data[i][0]).trim() === key) {
      storage.getRange(i+2,2,1,3).setValues([[rateFPX, rateEwallet, withdrawal]]);
      found = true;
      break;
    }
  }

  if (!found) {
    storage.appendRow([key, rateFPX, rateEwallet, withdrawal]);
  }
}

// Update only the manually edited row
function updateSingleRow(sheet, row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const storage = ss.getSheetByName("AgentStorage");
  const deposit = ss.getSheetByName("Deposit");
  if (!storage || !deposit) return;

  const merchant = sheet.getRange("B1").getValue();
  const monthName = sheet.getRange("B2").getValue();
  if (!merchant || !monthName) return;

  const monthNum = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"].indexOf(monthName)+1;
  const date = Utilities.formatDate(new Date(sheet.getRange(row,1).getValue()),"GMT+8","yyyy-MM-dd");
  const key = `${merchant}@${monthNum}@${date}`;

  // Get manual input
  const storeData = storage.getDataRange().getValues();
  const storeRowIndex = storeData.slice(1).findIndex(r => r[0] === key);
  const manual = storeRowIndex >= 0 ? storeData[storeRowIndex+1] : [key,0,0,0];
  const rateFPX = manual[1] || 0;
  const rateEwallet = manual[2] || 0;
  const withdrawal = manual[3] || 0;

  // Calculate FPX/Ewallet/Gross/Available for this date
  const depData = deposit.getDataRange().getValues();
  const depHeader = depData[0].map(h => String(h).trim());
  const idx = {};
  depHeader.forEach((h,i)=> idx[h] = i);

  let fpXAmount=0, ewAmount=0, availFPX=0, availEwallet=0;
  for (let i=1;i<depData.length;i++) {
    const rowD = depData[i];
    if(String(rowD[idx["Merchant"]]).trim() !== merchant) continue;

    const trxDate = Utilities.formatDate(new Date(rowD[idx["Transaction Date"]]),"GMT+8","yyyy-MM-dd");
    const settleDate = Utilities.formatDate(new Date(rowD[idx["Settlement Date"]]),"GMT+8","yyyy-MM-dd");
    const channel = String(rowD[idx["Channel"]]||"").toLowerCase();
    const kira = parseFloat(rowD[idx["Kira Amount"]]) || 0;

    if(trxDate === date){
      if(channel.includes("fpx")) fpXAmount += kira * rateFPX;
      else ewAmount += kira * rateEwallet;
    }

    if(settleDate === date){
      if(channel.includes("fpx")) availFPX += kira * rateFPX;
      else availEwallet += kira * rateEwallet;
    }
  }

  const gross = fpXAmount + ewAmount;
  const availTotal = availFPX + availEwallet;

  // Calculate cumulative balance for this row
  const previousBalance = row > 5 ? sheet.getRange(row-1,11).getValue() : 0;
  const balance = previousBalance + availTotal - withdrawal;

  const timestamp = Utilities.formatDate(new Date(),"GMT+8","yyyy-MM-dd HH:mm:ss");

  // Update the row in ledger
  sheet.getRange(row,2,1,11).setValues([[ 
    rateFPX, fpXAmount, rateEwallet, ewAmount, gross,
    availFPX, availEwallet, availTotal,
    withdrawal, balance, timestamp
  ]]);

  // Make Withdrawal Amount font black
  sheet.getRange(row,10).setFontColor("black");

  // Sync Withdrawal in AgentStorage
  if(storeRowIndex >= 0){
    storage.getRange(storeRowIndex+2,4).setValue(withdrawal);
  } else {
    storage.appendRow([key, rateFPX, rateEwallet, withdrawal]);
  }
}

// Full import of the month
function importAgent() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledger = ss.getSheetByName("Agents Balance & Settlement Ledger");
  const deposit = ss.getSheetByName("Deposit");
  const storage = ss.getSheetByName("AgentStorage");
  if (!ledger || !deposit || !storage) return;

  const merchant = ledger.getRange("B1").getValue();
  const monthName = ledger.getRange("B2").getValue();
  if (!merchant || !monthName) return;

  const monthNum = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"].indexOf(monthName)+1;
  const year = 2025;
  const lastDay = new Date(year, monthNum, 0).getDate();

  const depData = deposit.getDataRange().getValues();
  const depHeader = depData[0].map(h=>String(h).trim());
  const idx = {};
  depHeader.forEach((h,i)=>idx[h]=i);

  const parseNum = v=>typeof v==="number"?v:parseFloat(String(v).replace(/[^\d.-]/g,""))||0;

  // Build maps
  const trxMap = new Map();
  const settleMap = new Map();
  for(let i=1;i<depData.length;i++){
    const row = depData[i];
    if(String(row[idx["Merchant"]]).trim() !== String(merchant).trim()) continue;

    const trxDate = new Date(row[idx["Transaction Date"]]);
    const settleDate = new Date(row[idx["Settlement Date"]]);
    const channel = String(row[idx["Channel"]]||"").toLowerCase();
    const kira = parseNum(row[idx["Kira Amount"]]);
    const map = channel.includes("fpx") ? "fpx":"ewallet";

    if(trxDate.getMonth()+1===monthNum){
      const key = Utilities.formatDate(trxDate,"GMT+8","yyyy-MM-dd");
      if(!trxMap.has(key)) trxMap.set(key,{fpx:0,ewallet:0});
      trxMap.get(key)[map] += kira;
    }

    if(settleDate.getMonth()+1===monthNum){
      const key = Utilities.formatDate(settleDate,"GMT+8","yyyy-MM-dd");
      if(!settleMap.has(key)) settleMap.set(key,{fpx:0,ewallet:0});
      settleMap.get(key)[map] += kira;
    }
  }

  // Manual input map
  const storeData = storage.getDataRange().getValues();
  const storeMap = new Map();
  for(let i=1;i<storeData.length;i++){
    const key = String(storeData[i][0]).trim();
    storeMap.set(key,{
      rateFPX: parseNum(storeData[i][1]),
      rateEwallet: parseNum(storeData[i][2]),
      withdrawal: parseNum(storeData[i][3])
    });
  }

  // Generate ledger for the month with cumulative balance
  let cumulativeBalance = 0;
  const results = [];
  for(let d=1;d<=lastDay;d++){
    const dateStr = Utilities.formatDate(new Date(year,monthNum-1,d),"GMT+8","yyyy-MM-dd");
    const trx = trxMap.get(dateStr) || {fpx:0,ewallet:0};
    const settle = settleMap.get(dateStr) || {fpx:0,ewallet:0};
    const key = `${merchant}@${monthNum}@${dateStr}`;
    const manual = storeMap.get(key) || {};

    const fpXAmount = trx.fpx * (manual.rateFPX||0);
    const ewAmount = trx.ewallet * (manual.rateEwallet||0);
    const gross = fpXAmount + ewAmount;

    const availFPX = settle.fpx * (manual.rateFPX||0);
    const availEwallet = settle.ewallet * (manual.rateEwallet||0);
    const availTotal = availFPX + availEwallet;

    const withdrawal = manual.withdrawal||0;

    cumulativeBalance = cumulativeBalance + availTotal - withdrawal;

    results.push([
      dateStr,
      manual.rateFPX||0,
      fpXAmount,
      manual.rateEwallet||0,
      ewAmount,
      gross,
      availFPX,
      availEwallet,
      availTotal,
      withdrawal,
      cumulativeBalance,
      Utilities.formatDate(new Date(),"GMT+8","yyyy-MM-dd HH:mm:ss")
    ]);
  }

  ledger.getRange(5,1,results.length,results[0].length).setValues(results);
  ledger.getRange(5,10,results.length,1).setFontColor("black");
  Logger.log(`Agents Ledger updated for ${merchant} (${monthName})`);
}
