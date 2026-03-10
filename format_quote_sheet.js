/****************************************************
 * FilePro Quote Sheet Formatter - Version 1.1.0
 * - Deployed as Google Apps Script web app
 * - Called by filepro_sync.py after each sheet sync
 * - Styling consistent with EAR tools suite
 * - Conservative syntax (var, no optional chaining)
 * - All setBorder() calls use 8 parameters
 ****************************************************/

var APP_VERSION   = '1.1.0';
var SHEET_VERSION = '2.3';

/****************************************************
 * WEB APP ENTRY POINTS
 ****************************************************/
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheetId = data.sheet_id || extractSheetId(data.sheet_url);
    var quoteNumber = data.quote_number || 'Unknown';

    if (!sheetId) {
      return jsonResponse({status: 'error', message: 'No sheet_id or sheet_url provided'});
    }

    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheets()[0];
    formatSheet(sheet, quoteNumber);

    return jsonResponse({status: 'ok', quote_number: quoteNumber, formatted: true});

  } catch (error) {
    return jsonResponse({status: 'error', message: error.toString()});
  }
}

// Also support GET for testing
function doGet(e) {
  var sheetId = e.parameter.sheet_id;
  if (sheetId) {
    try {
      var ss = SpreadsheetApp.openById(sheetId);
      var sheet = ss.getSheets()[0];
      formatSheet(sheet, '');
      return jsonResponse({status: 'ok', formatted: true});
    } catch (error) {
      return jsonResponse({status: 'error', message: error.toString()});
    }
  }
  return jsonResponse({
    status: 'ready',
    message: 'FilePro Quote Formatter v' + APP_VERSION + '. POST with sheet_url or sheet_id.'
  });
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function extractSheetId(url) {
  if (!url) return null;
  var match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

// Manual test — update sheet ID before running in Apps Script editor
function testFormat() {
  var ss = SpreadsheetApp.openById('1NqwkGRrJ-R3gZq3A0ibupJmxPHclQgLxrwAWQ1R8BKc');
  var sheet = ss.getSheets()[0];
  formatSheet(sheet, '');
}

/****************************************************
 * MAIN FORMATTING FUNCTION
 ****************************************************/
function formatSheet(sheet, quoteNumber) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 2) return;

  // Colors — consistent with EAR tools suite
  var earBlue       = '#1a73e8';
  var lightBlue     = '#d4e6f1';
  var white         = '#FFFFFF';
  var rowAlt        = '#f8f9fa';
  var totalRowGreen = '#4CAF50';
  var borderColor   = '#CCCCCC';

  var dataStartRow = findDataStartRow(sheet);
  var totalsRow    = findTotalsStartRow(sheet);
  var dataEndRow   = (totalsRow > 0) ? totalsRow - 1 : lastRow;

  formatHeaderSection(sheet, lastCol, earBlue, white);
  formatColumnHeaders(sheet, dataStartRow, lastCol, earBlue, white, borderColor);
  formatLineItems(sheet, dataStartRow, dataEndRow, lastCol, white, rowAlt, borderColor);
  formatTotalsSection(sheet, totalsRow, lastRow, lastCol, lightBlue, totalRowGreen, white, borderColor);
  applyFinalFormatting(sheet, dataStartRow, lastCol);
}

/****************************************************
 * HEADER SECTION (rows 1–2: QUOTATION + date)
 ****************************************************/
function formatHeaderSection(sheet, lastCol, earBlue, white) {
  var headerRange = sheet.getRange(1, 1, 2, lastCol);
  headerRange.setBackground(earBlue);
  headerRange.setFontColor(white);
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(12);

  try { sheet.getRange(1, 1, 1, lastCol).merge(); } catch (e) {}
  sheet.getRange(1, 1).setFontSize(16).setHorizontalAlignment('left');
}

/****************************************************
 * COLUMN HEADER ROW
 ****************************************************/
function formatColumnHeaders(sheet, dataStartRow, lastCol, earBlue, white, borderColor) {
  if (dataStartRow <= 0) return;

  sheet.getRange(dataStartRow, 1, 1, lastCol)
    .setBackground(earBlue)
    .setFontColor(white)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, borderColor, SpreadsheetApp.BorderStyle.SOLID);
}

/****************************************************
 * LINE ITEM ROWS — alternating colors
 ****************************************************/
function formatLineItems(sheet, dataStartRow, dataEndRow, lastCol, white, rowAlt, borderColor) {
  if (dataStartRow <= 0 || dataEndRow <= dataStartRow) return;

  var i;
  for (i = dataStartRow + 1; i <= dataEndRow; i++) {
    var bg = ((i - dataStartRow) % 2 === 0) ? rowAlt : white;
    sheet.getRange(i, 1, 1, lastCol)
      .setBackground(bg)
      .setBorder(true, true, true, true, false, false, borderColor, SpreadsheetApp.BorderStyle.SOLID);
  }

  // Currency format on price columns (4 & 5)
  if (lastCol >= 5) {
    sheet.getRange(dataStartRow + 1, 4, dataEndRow - dataStartRow, 2)
      .setNumberFormat('$#,##0.00')
      .setHorizontalAlignment('right');
  }

  // Qty column — center
  sheet.getRange(dataStartRow + 1, 1, dataEndRow - dataStartRow, 1)
    .setHorizontalAlignment('center');
}

/****************************************************
 * TOTALS SECTION (Sub Total / Tax / Shipping / TOTAL)
 ****************************************************/
function formatTotalsSection(sheet, totalsRow, lastRow, lastCol, lightBlue, totalRowGreen, white, borderColor) {
  if (totalsRow <= 0) return;

  // Sub Total, Tax, Shipping rows — light blue
  sheet.getRange(totalsRow, 1, lastRow - totalsRow, 2)
    .setBackground(lightBlue)
    .setFontWeight('bold')
    .setBorder(true, true, true, true, true, true, borderColor, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(totalsRow, 2, lastRow - totalsRow, 1)
    .setNumberFormat('$#,##0.00')
    .setHorizontalAlignment('right');

  // TOTAL row — solid green highlight
  var j;
  for (j = totalsRow; j <= lastRow; j++) {
    var label = sheet.getRange(j, 1).getValue().toString().toUpperCase();
    if (label === 'TOTAL' || label.indexOf('TOTAL') === 0) {
      sheet.getRange(j, 1, 1, 2)
        .setBackground(totalRowGreen)
        .setFontColor(white)
        .setFontWeight('bold')
        .setFontSize(12)
        .setBorder(true, true, true, true, true, true, borderColor, SpreadsheetApp.BorderStyle.SOLID);
    }
  }
}

/****************************************************
 * FINAL FORMATTING — font, gridlines, freeze, resize
 ****************************************************/
function applyFinalFormatting(sheet, dataStartRow, lastCol) {
  // Roboto font across all content
  var lr = sheet.getLastRow();
  var lc = sheet.getLastColumn();
  if (lr && lc) {
    sheet.getRange(1, 1, lr, lc).setFontFamily('Roboto');
  }

  // Hide gridlines
  sheet.setHiddenGridlines(true);

  // Auto-resize all columns
  var c;
  for (c = 1; c <= lastCol; c++) {
    sheet.autoResizeColumn(c);
  }

  // Freeze through column header row
  sheet.setFrozenRows(dataStartRow > 0 ? dataStartRow : 4);
}

/****************************************************
 * ROW FINDERS
 ****************************************************/
function findDataStartRow(sheet) {
  var data = sheet.getDataRange().getValues();
  var i;
  for (i = 0; i < data.length; i++) {
    var row = data[i].join(' ').toLowerCase();
    if (row.indexOf('qty') >= 0 || row.indexOf('part') >= 0 || row.indexOf('description') >= 0) {
      return i + 1;
    }
  }
  return 8;
}

function findTotalsStartRow(sheet) {
  var data = sheet.getDataRange().getValues();
  var i;
  for (i = data.length - 1; i >= 0; i--) {
    var cellA = data[i][0].toString().toLowerCase();
    if (cellA.indexOf('sub total') >= 0 || cellA.indexOf('subtotal') >= 0) {
      return i + 1;
    }
  }
  return 0;
}
