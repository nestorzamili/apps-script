function processBankFile(file, bankMap) {
  const fileName = file.getName();
  const meta = parseBankFileName(fileName);

  const ss = SpreadsheetApp.openById(file.getId());
  const sh = ss.getSheets()[0];
  const all = sh.getDataRange().getValues();

  if (all.length <= 1) return 0;

  const header = all[0];
  const data = all.slice(1);

  const colPaymentMode = findColumnIndex(header, [
    'Payment Mode',
    'paymentmode',
    'Payment Method',
    'paymentmethod',
  ]);

  const colOrderId = findColumnIndex(header, ['Order ID', 'orderid']);
  const colMerchantOrderNo = findColumnIndex(header, [
    'merchantOrderNo',
    'merchantorderno',
  ]);
  const colTxnId = findColumnIndex(header, [
    'transactionID',
    'transactionid',
    'order number',
    'ordernumber',
    'order_no',
    'order no',
  ]);

  const colAmount = findColumnIndex(header, [
    'Amount',
    'transactionAmount',
    'Payment Amount',
    'paymentamount',
    'Amount (RM)',
    'amount(rm)',
  ]);

  const colDate = findColumnIndex(header, [
    'Payment Time',
    'createdDate',
    'Date',
    'Transaction Date',
    'Created Date',
  ]);

  let colFinalTxn = -1;
  if (colOrderId >= 0) {
    colFinalTxn = colOrderId;
  } else if (colMerchantOrderNo >= 0) {
    colFinalTxn = colMerchantOrderNo;
  } else {
    colFinalTxn = colTxnId;
  }

  if (colFinalTxn === -1 || colAmount === -1) {
    Logger.log(
      'Skipping Bank file "' + fileName + '" due to missing key columns',
    );
    return 0;
  }

  let rowCount = 0;
  data.forEach((r) => {
    const tid = String(r[colFinalTxn] || '').trim();
    if (!tid || bankMap.has(tid)) return;

    let channel = meta.channel;
    if (colPaymentMode >= 0) {
      const paymentMode = String(r[colPaymentMode] || '').toLowerCase();
      if (paymentMode.includes('fpx b2c') || paymentMode.includes('fpx casa')) {
        channel = 'FPX';
      } else if (paymentMode.includes('fpx b2b') || paymentMode.includes('fpxc')) {
        channel = 'FPXC';
      } else if (paymentMode.includes('tng') || paymentMode.includes('touch')) {
        channel = 'TNG';
      } else if (paymentMode.includes('boost')) {
        channel = 'BOOST';
      } else if (
        paymentMode.includes('shopeepay') ||
        paymentMode.includes('shopee')
      ) {
        channel = 'Shopee';
      } else if (paymentMode) {
        channel = paymentMode;
      }
    }

    const normalizedChannel = normalizeBankChannel(channel);

    bankMap.set(tid, {
      bankMerchant: meta.bankMerchant,
      bankChannel: normalizedChannel,
      bankTransactionDate: colDate >= 0 ? r[colDate] : '',
      bankAmount: parseAmount(r[colAmount]),
    });
    rowCount++;
  });

  return rowCount;
}

function normalizeBankChannel(value) {
  if (!value) return 'ewallet';
  const str = String(value).toUpperCase().trim();
  if (str === 'FPX' || str.includes('FPX B2C') || str.includes('CASA')) return 'FPX';
  if (str === 'FPXC' || str.includes('FPX B2B')) return 'FPXC';
  if (str === 'TNG' || str.includes('TOUCH') || str.includes('TOUCHNGO')) return 'TNG';
  if (str === 'BOOST' || str.includes('BOOST')) return 'BOOST';
  if (str === 'SHOPEE' || str.includes('SHOPEE')) return 'Shopee';
  return 'ewallet';
}

function parseBankFileName(name) {
  const base = name.replace(/\.[^/.]+$/, '');
  const parts = base.split('_');
  const lower = parts.map((p) => p.toLowerCase());

  let bankMerchant = 'Unknown';
  let channel = 'Unknown';

  const channelLower = lower.some((p) => p.includes('fpx'))
    ? 'FPX'
    : 'ewallet';
  const channelIndex = lower.findIndex(
    (p) => p.includes('fpx') || p.includes('ewallet'),
  );
  
  let rawMerchant =
    channelIndex > 0
      ? parts.slice(0, channelIndex).join('_')
      : parts.slice(0, 2).join('_') || parts[0] || 'Unknown';
  
  bankMerchant = rawMerchant.replace(/_Axaipay$/i, '').replace(/_axaipay$/i, '');
  
  channel = channelLower;

  return { bankMerchant, channel };
}
