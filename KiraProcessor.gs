function processKiraFile(file, kiraMap) {
  const ss = SpreadsheetApp.openById(file.getId());
  const sh = ss.getSheets()[0];
  const all = sh.getDataRange().getValues();

  if (all.length <= 1) return 0;

  const header = all[0];
  const data = all.slice(1);

  const indices = {
    createdOn: findColumnIndex(header, ['Created On']),
    merchant: findColumnIndex(header, ['Merchant']),
    txnId: findColumnIndex(header, ['Transaction ID', 'transactionid']),
    merchantOrder: findColumnIndex(header, [
      'Merchant Order ID',
      'merchantOrderNo',
    ]),
    paymentMethod: findColumnIndex(header, ['Payment Method', 'paymentmethod']),
    txnAmount: findColumnIndex(header, [
      'Transaction Amount',
      'transactionamount',
    ]),
  };

  let rowCount = 0;
  data.forEach((r) => {
    const tid = String(r[indices.txnId] || '').trim();
    if (!tid || kiraMap.has(tid)) return;

    const rawPaymentMethod = indices.paymentMethod >= 0 ? r[indices.paymentMethod] : '';
    const normalizedPaymentMethod = normalizePaymentMethod(rawPaymentMethod);

    kiraMap.set(tid, {
      createdOn: indices.createdOn >= 0 ? r[indices.createdOn] : '',
      merchant: indices.merchant >= 0 ? r[indices.merchant] : '',
      merchantOrderId:
        indices.merchantOrder >= 0 ? r[indices.merchantOrder] : '',
      paymentMethod: normalizedPaymentMethod,
      transactionAmount: parseFloat(r[indices.txnAmount]) || 0,
    });
    rowCount++;
  });

  return rowCount;
}

function normalizePaymentMethod(value) {
  if (!value) return 'ewallet';
  const str = String(value).toUpperCase().trim();
  if (str === 'FPX' || str.includes('FPX B2C') || str.includes('CASA')) return 'FPX';
  if (str === 'FPXC' || str.includes('FPX B2B')) return 'FPXC';
  if (str === 'TNG' || str.includes('TOUCH') || str.includes('TOUCHNGO')) return 'TNG';
  if (str === 'BOOST' || str.includes('BOOST')) return 'BOOST';
  if (str === 'SHOPEE' || str.includes('SHOPEE')) return 'Shopee';
  return 'ewallet';
}
