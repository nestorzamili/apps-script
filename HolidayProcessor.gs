function loadMalaysiaHolidays() {
  try {
    const res = UrlFetchApp.fetch(CONFIG.MALAYSIA_HOLIDAYS_URL);
    const lines = res.getContentText().split(/\r?\n/);

    const holidays = new Set();
    let date = "", name = "";

    lines.forEach(l => {
      if (l.startsWith("DTSTART")) {
        const m = l.match(/(\d{4})(\d{2})(\d{2})/);
        if (m) date = `${m[1]}-${m[2]}-${m[3]}`;
      }
      if (l.startsWith("SUMMARY:")) {
        name = l.replace("SUMMARY:", "").trim();
        if (date && name) {
          holidays.add(date);
          date = "";
          name = "";
        }
      }
    });

    Logger.log('Loaded ' + holidays.size + ' Malaysia holidays');
    return holidays;
  } catch (e) {
    Logger.log('Error loading Malaysia holidays: ' + e.message);
    return new Set();
  }
}

function isWeekend(date) {
  const day = date.getDay();
  return day === 0 || day === 6;
}

function isHoliday(dateStr, holidaySet) {
  return holidaySet.has(dateStr);
}

function formatDateString(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return year + '-' + month + '-' + day;
}

function calculateSettlementDate(transactionDateStr, settlementRule, holidaySet) {
  if (!transactionDateStr || !settlementRule) {
    return '';
  }

  const match = settlementRule.match(/T\+(\d+)/i);
  if (!match) {
    return '';
  }

  const daysToAdd = parseInt(match[1], 10);
  
  const dateParts = transactionDateStr.split('-');
  if (dateParts.length !== 3) {
    return '';
  }

  let currentDate = new Date(
    parseInt(dateParts[0], 10),
    parseInt(dateParts[1], 10) - 1,
    parseInt(dateParts[2], 10)
  );

  let businessDaysAdded = 0;

  while (businessDaysAdded < daysToAdd) {
    currentDate.setDate(currentDate.getDate() + 1);
    
    const dateStr = formatDateString(currentDate);
    
    if (!isWeekend(currentDate) && !isHoliday(dateStr, holidaySet)) {
      businessDaysAdded++;
    }
  }

  while (isWeekend(currentDate) || isHoliday(formatDateString(currentDate), holidaySet)) {
    currentDate.setDate(currentDate.getDate() + 1);
  }

  return formatDateString(currentDate);
}
