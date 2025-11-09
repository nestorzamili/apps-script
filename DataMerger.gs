function mergeData(kiraMap, pgMap, bankMap) {
  const allTids = new Set([
    ...kiraMap.keys(),
    ...pgMap.keys(),
    ...bankMap.keys(),
  ]);
  const output = [];
  let mismatchPG = 0;
  let mismatchBank = 0;
  let noDataPG = 0;
  let noDataBank = 0;

  allTids.forEach((tid) => {
    const k = kiraMap.get(tid) || {};
    const pg = pgMap.get(tid) || {};
    const bank = bankMap.get(tid) || {};
    
    const kAmt = k.transactionAmount || 0;
    const pgAmt = pg.pgAmount || 0;
    const bankAmt = bank.bankAmount || 0;
    
    const hasKiraData = kiraMap.has(tid);
    const hasPGData = pgMap.has(tid);
    const hasBankData = bankMap.has(tid);

    let remarks = '';
    if (hasKiraData) {
      const pgMatch = hasPGData && kAmt === pgAmt;
      const bankMatch = hasBankData && kAmt === bankAmt;
      
      if (hasPGData && hasBankData) {
        if (pgMatch && bankMatch) {
          remarks = 'Match';
        } else if (!pgMatch && !bankMatch) {
          remarks = 'Not Match (PG & Bank)';
        } else if (!pgMatch) {
          remarks = 'Not Match (PG)';
        } else {
          remarks = 'Not Match (Bank)';
        }
      } else if (hasPGData) {
        remarks = pgMatch ? 'Match (PG only)' : 'Not Match (PG)';
      } else if (hasBankData) {
        remarks = bankMatch ? 'Match (Bank only)' : 'Not Match (Bank)';
      } else {
        remarks = 'No Data (PG & Bank)';
      }
    } else {
      if (hasPGData && hasBankData) {
        remarks = 'No Kira Data';
      } else if (hasPGData) {
        remarks = 'No Kira Data (PG only)';
      } else if (hasBankData) {
        remarks = 'No Kira Data (Bank only)';
      }
    }

    const row = [
      k.createdOn || (hasKiraData ? '' : 'No Data'),
      k.merchant || (hasKiraData ? '' : 'No Data'),
      tid || '',
      k.merchantOrderId || (hasKiraData ? '' : 'No Data'),
      k.paymentMethod || (hasKiraData ? '' : 'No Data'),
      kAmt || (hasKiraData ? 0 : 'No Data'),
      pg.pgMerchant || (hasPGData ? '' : 'No Data'),
      pg.pgChannel || (hasPGData ? '' : 'No Data'),
      pg.pgTransactionDate || (hasPGData ? '' : 'No Data'),
      pgAmt || (hasPGData ? 0 : 'No Data'),
      bank.bankMerchant || (hasBankData ? '' : 'No Data'),
      bank.bankChannel || (hasBankData ? '' : 'No Data'),
      bank.bankTransactionDate || (hasBankData ? '' : 'No Data'),
      bankAmt || (hasBankData ? 0 : 'No Data'),
      remarks || '',
    ];
    
    output.push(row);

    if (hasKiraData && !hasPGData) {
      noDataPG++;
    } else if (kAmt && pgAmt && kAmt !== pgAmt) {
      mismatchPG++;
    }

    if (hasKiraData && !hasBankData) {
      noDataBank++;
    } else if (kAmt && bankAmt && kAmt !== bankAmt) {
      mismatchBank++;
    }
  });

  return {
    output,
    stats: {
      mismatchPG,
      mismatchBank,
      noDataPG,
      noDataBank,
    },
  };
}

