// ═══════════════════════════════════════════════════════════
//  BALISTO — Apps Script Connector
//  Attach to: For Calculator spreadsheet
//  Deploy as Web App → Execute as Me → Anyone can access
// ═══════════════════════════════════════════════════════════

const SHEET_ID        = '126KyTuj9PBK8A_kYRVigMrehpF7FlOkUrl6TXqkUsss';
const DATE_WEEK_TAB   = 'Date and Week';
const LOG_TAB         = 'Quotation Log';

// ── ROUTER ──────────────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action || '';
  let result;

  try {
    if (action === 'getBasePrice') {
      result = getBasePriceForDate(e.parameter.date);
    } else if (action === 'logQuotation') {
      result = logQuotation(e.parameter);
    } else {
      result = { status: 'Balisto connector live ✅' };
    }
  } catch(err) {
    result = { error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET BASE PRICE FOR DATE ──────────────────────────────────
function getBasePriceForDate(dateStr) {
  if (!dateStr) throw new Error('No date provided');

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(DATE_WEEK_TAB);
  if (!sheet) throw new Error('Tab "' + DATE_WEEK_TAB + '" not found');

  const data   = sheet.getDataRange().getValues();
  const target = new Date(dateStr);
  target.setHours(0, 0, 0, 0);

  // Row 0 = headers: Date | Week | Single Battery | Double Battery
  for (let i = 1; i < data.length; i++) {
    const rowDate = new Date(data[i][0]);
    rowDate.setHours(0, 0, 0, 0);
    if (rowDate.getTime() === target.getTime()) {
      return {
        found:      true,
        week:       data[i][1],
        singleBase: data[i][2],
        doubleBase: data[i][3]
      };
    }
  }

  // Date beyond sheet range — use last row as fallback
  const last = data[data.length - 1];
  return {
    found:      false,
    fallback:   true,
    week:       last[1],
    singleBase: last[2],
    doubleBase: last[3]
  };
}

// ── LOG QUOTATION TO SHEET ───────────────────────────────────
function logQuotation(p) {
  const ss  = SpreadsheetApp.openById(SHEET_ID);
  let log   = ss.getSheetByName(LOG_TAB);

  if (!log) {
    log = ss.insertSheet(LOG_TAB);
    log.appendRow([
      'Quotation No', 'Date', 'Customer', 'Address', 'Phone',
      'Single Qty', 'Single Price (₹)',
      'Double Qty', 'Double Price (₹)',
      'Grand Total (₹)'
    ]);
    log.getRange('1:1')
       .setFontWeight('bold')
       .setBackground('#1A1A2E')
       .setFontColor('#FFFFFF');
    log.setFrozenRows(1);
    log.setColumnWidth(1, 180);
    log.setColumnWidth(3, 160);
  }

  log.appendRow([
    p.quotNo       || '',
    p.date         || '',
    p.customer     || '',
    p.address      || '',
    p.phone        || '',
    parseInt(p.singleQty)    || 0,
    parseFloat(p.singlePrice)|| 0,
    parseInt(p.doubleQty)    || 0,
    parseFloat(p.doublePrice)|| 0,
    parseFloat(p.grand)      || 0
  ]);

  return { success: true, quotNo: p.quotNo };
}
