function processPGFile(file, pgMap) {
  const fileName = file.getName();
  const meta = parsePGFileName(fileName);

  const ss = SpreadsheetApp.openById(file.getId());
  const sh = ss.getSheets()[0];
  const all = sh.getDataRange().getValues();

  if (all.length <= 1) return 0;

  const header = all[0];
  const data = all.slice(1);

  const colPaymentMode = findColumnIndex(header, [
    'Payment Mode',
    'paymentmode',
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

  let colAmount = -1;
  if (meta.isRHB) {
    colAmount = findColumnIndex(header, ['Amount (RM)', 'amount(rm)']);
    if (colAmount === -1) {
      colAmount = findColumnIndex(header, [
        'Amount',
        'transactionAmount',
        'Payment Amount',
        'paymentamount',
      ]);
    }
  } else {
    colAmount = findColumnIndex(header, [
      'Amount',
      'transactionAmount',
      'Payment Amount',
      'paymentamount',
      'Amount (RM)',
    ]);
  }

  const colDate = findColumnIndex(header, [
    'Payment Time',
    'createdDate',
    'Date',
    'Transaction Date',
  ]);

  let colFinalTxn = -1;
  if (meta.isRHB && colOrderId >= 0) {
    colFinalTxn = colOrderId;
  } else if (colMerchantOrderNo >= 0) {
    colFinalTxn = colMerchantOrderNo;
  } else {
    colFinalTxn = colTxnId;
  }

  if (colFinalTxn === -1 || colAmount === -1) {
    Logger.log(
      'Skipping PG file "' + fileName + '" due to missing key columns',
    );
    return 0;
  }

  let rowCount = 0;
  data.forEach((r) => {
    const tid = String(r[colFinalTxn] || '').trim();
    if (!tid || pgMap.has(tid)) return;

    let channel = meta.channel;
    
    if (colPaymentMode >= 0) {
      const paymentMode = String(r[colPaymentMode] || '').toLowerCase().trim();
      if (paymentMode) {
        if (paymentMode.includes('fpx b2c') || paymentMode.includes('fpx casa') || paymentMode.includes('casa')) {
          channel = 'FPX';
        } else if (paymentMode.includes('fpx b2b') || paymentMode.includes('fpxc')) {
          channel = 'FPXC';
        } else if (paymentMode.includes('tng') || paymentMode.includes('touch')) {
          channel = 'TNG';
        } else if (paymentMode.includes('boost')) {
          channel = 'BOOST';
        } else if (paymentMode.includes('shopeepay') || paymentMode.includes('shopee')) {
          channel = 'Shopee';
        } else if (!paymentMode.includes('ewallet')) {
          channel = paymentMode;
        }
      }
    }

    const normalizedChannel = normalizePGChannel(channel);

    pgMap.set(tid, {
      pgMerchant: meta.pgMerchant,
      pgChannel: normalizedChannel,
      pgTransactionDate: colDate >= 0 ? normalizePGDate(r[colDate]) : '',
      pgAmount: parseAmount(r[colAmount]),
    });
    rowCount++;
  });

  return rowCount;
}

function normalizePGDate(value) {
  if (!value) return '';
  const str = String(value).trim();
  const pattern = /^(\d{2}):(\d{2})\s+(\d{4}-\d{2}-\d{2})$/;
  const match = str.match(pattern);
  if (match) {
    const hour = match[1];
    const minute = match[2];
    const date = match[3];
    return date + ' ' + hour + ':' + minute + ':00';
  }
  return str;
}

function normalizePGChannel(value) {
  if (!value) return 'ewallet';
  const str = String(value).toUpperCase().trim();
  if (str === 'FPX' || str.includes('FPX B2C') || str.includes('CASA')) return 'FPX';
  if (str === 'FPXC' || str.includes('FPX B2B')) return 'FPXC';
  if (str === 'TNG' || str.includes('TOUCH') || str.includes('TOUCHNGO')) return 'TNG';
  if (str === 'BOOST' || str.includes('BOOST')) return 'BOOST';
  if (str === 'SHOPEE' || str.includes('SHOPEE')) return 'Shopee';
  return 'ewallet';
}

function parsePGFileName(name) {
  const base = name.replace(/\.[^/.]+$/, '');
  const parts = base.split('_');
  const lower = parts.map((p) => p.toLowerCase());

  const isRHB =
    lower.some((p) => p.includes('rhb')) &&
    lower.some((p) => p.includes('axaipay'));

  let pgMerchant = 'Unknown';
  let channel = 'Unknown';

  if (isRHB) {
    const firstPart = parts[0];
    const cleanedPart = firstPart.replace(/\s*RHB\s*/gi, '').trim();
    pgMerchant = cleanedPart ? cleanedPart + ' RHB' : firstPart + ' RHB';
    channel = 'Dynamic';
  } else if (lower.some((p) => p.includes('ragnarok'))) {
    const merchantPart = parts[0];
    const channelPart = parts[2] || '';
    
    const merchantLower = merchantPart.toLowerCase();

    if (merchantLower.includes('infinetix')) {
      pgMerchant = merchantLower.includes(' axai') ? merchantPart : 'Infinetix Axai';
    } else if (merchantLower.includes('ms')) {
      pgMerchant = merchantLower.includes(' axai') ? merchantPart : 'MS Axai';
    } else if (merchantLower.includes('vbina')) {
      pgMerchant = merchantLower.includes(' axai') ? merchantPart : 'Vbina Axai';
    } else {
      pgMerchant = merchantLower.includes(' axai') ? merchantPart : merchantPart + ' Axai';
    }

    const filenameLower = base.toLowerCase();
    
    if (channelPart.toLowerCase() === 'fpx') {
      channel = 'FPX';
    } else if (filenameLower.includes('shopeepay')) {
      channel = 'Shopee';
    } else if (filenameLower.includes('tng')) {
      channel = 'TNG';
    } else if (filenameLower.includes('boost')) {
      channel = 'BOOST';
    } else {
      channel = 'ewallet';
    }
  } else if (lower.some((p) => p.includes('m1pay'))) {
    const merchantPart = parts[0];
    const channelPart = parts[2] || '';
    const typePart = parts[3] || '';

    const merchantLower = merchantPart.toLowerCase();
    
    if (merchantLower.includes('infinetix')) {
      pgMerchant = merchantLower.includes(' m1') ? merchantPart : 'Infinetix M1';
    } else if (merchantLower.includes('ms')) {
      pgMerchant = merchantLower.includes(' m1') ? merchantPart : 'MS M1';
    } else if (merchantLower.includes('vbina')) {
      pgMerchant = merchantLower.includes(' m1') ? merchantPart : 'Vbina M1';
    } else {
      pgMerchant = merchantLower.includes(' m1') ? merchantPart : merchantPart + ' M1';
    }

    const filenameLower = base.toLowerCase();
    
    if (
      channelPart.toLowerCase() === 'fpx' ||
      typePart.toLowerCase().includes('fpx') ||
      filenameLower.includes('transactionallfpx')
    ) {
      channel = 'FPX';
    } else if (filenameLower.includes('shopeepay')) {
      channel = 'Shopee';
    } else if (filenameLower.includes('tng')) {
      channel = 'TNG';
    } else if (filenameLower.includes('boost')) {
      channel = 'BOOST';
    } else {
      channel = 'ewallet';
    }
  } else {
    const channelLower = lower.some((p) => p.includes('fpx'))
      ? 'FPX'
      : 'ewallet';
    const channelIndex = lower.findIndex(
      (p) => p.includes('fpx') || p.includes('ewallet'),
    );
    pgMerchant =
      channelIndex > 0
        ? parts.slice(0, channelIndex).join('_')
        : parts.slice(0, 2).join('_') || parts[0] || 'Unknown';
    channel = channelLower;
  }

  return { pgMerchant, channel, isRHB };
}
