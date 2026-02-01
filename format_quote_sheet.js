/**
 * FilePro Quote Sheet Formatter - Web App
 * Deployed as web app, called by filepro_sync.py after each sync
 *
 * Deploy:
 *   1. Deploy → New deployment → Web app
 *   2. Execute as: Me
 *   3. Who has access: Anyone
 *   4. Copy the deployment URL to filepro_sync.py CONFIG['webhook_url']
 */

// Web app entry point - receives POST from Python
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
    formatSheet(sheet);

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
      formatSheet(sheet);
      return jsonResponse({status: 'ok', formatted: true});
    } catch (error) {
      return jsonResponse({status: 'error', message: error.toString()});
    }
  }
  return jsonResponse({status: 'ready', message: 'FilePro Quote Formatter. POST with sheet_url or sheet_id.'});
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

// Manual test function
function testFormat() {
  var ss = SpreadsheetApp.openById('1NqwkGRrJ-R3gZq3A0ibupJmxPHclQgLxrwAWQ1R8BKc');
  var sheet = ss.getSheets()[0];
  formatSheet(sheet);
}

/**
 * Main formatting function
 */
function formatSheet(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 2) return;

  // Colors
  var headerBlue = '#1F4F76';
  var lightBlue = '#D6E9F4';
  var white = '#FFFFFF';
  var lightGray = '#F5F5F5';
  var borderColor = '#CCCCCC';
  var totalsGreen = '#E8F5E9';

  // Header section (Rows 1-2)
  var headerRange = sheet.getRange('A1:F2');
  headerRange.setBackground(headerBlue);
  headerRange.setFontColor(white);
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(12);

  sheet.getRange('A1:B1').merge();
  sheet.getRange('A1').setFontSize(16);

  // Customer info section
  var customerRange = sheet.getRange('A3:B6');
  customerRange.setFontSize(10);
  sheet.getRange('A3:A6').setFontWeight('bold');

  // Find data rows
  var dataStartRow = findDataStartRow(sheet);
  var totalsRow = findTotalsStartRow(sheet);
  var dataEndRow = (totalsRow > 0) ? totalsRow - 1 : lastRow;

  // Column headers
  if (dataStartRow > 0) {
    var headerRow = sheet.getRange(dataStartRow, 1, 1, lastCol);
    headerRow.setBackground(lightBlue);
    headerRow.setFontWeight('bold');
    headerRow.setHorizontalAlignment('center');
    headerRow.setBorder(true, true, true, true, true, true, borderColor, SpreadsheetApp.BorderStyle.SOLID);
  }

  // Line items with alternating colors
  if (dataStartRow > 0 && dataEndRow > dataStartRow) {
    var i;
    for (i = dataStartRow + 1; i <= dataEndRow; i++) {
      var rowRange = sheet.getRange(i, 1, 1, lastCol);
      if ((i - dataStartRow) % 2 === 0) {
        rowRange.setBackground(lightGray);
      } else {
        rowRange.setBackground(white);
      }
      rowRange.setBorder(null, true, null, true, false, false, borderColor, SpreadsheetApp.BorderStyle.SOLID);
    }

    // Format price columns as currency
    if (lastCol >= 5) {
      var priceRange = sheet.getRange(dataStartRow + 1, 4, dataEndRow - dataStartRow, 2);
      priceRange.setNumberFormat('$#,##0.00');
      priceRange.setHorizontalAlignment('right');
    }

    // Format qty column
    var qtyRange = sheet.getRange(dataStartRow + 1, 1, dataEndRow - dataStartRow, 1);
    qtyRange.setHorizontalAlignment('center');
  }

  // Totals section
  if (totalsRow > 0) {
    var totalsRange = sheet.getRange(totalsRow, 1, lastRow - totalsRow + 1, 2);
    totalsRange.setBackground(totalsGreen);
    totalsRange.setFontWeight('bold');
    totalsRange.setBorder(true, true, true, true, true, true, borderColor, SpreadsheetApp.BorderStyle.SOLID);

    var totalsAmounts = sheet.getRange(totalsRow, 2, lastRow - totalsRow + 1, 1);
    totalsAmounts.setNumberFormat('$#,##0.00');
    totalsAmounts.setHorizontalAlignment('right');

    // Highlight TOTAL row
    var j;
    for (j = totalsRow; j <= lastRow; j++) {
      var label = sheet.getRange(j, 1).getValue().toString().toUpperCase();
      if (label === 'TOTAL' || label.indexOf('TOTAL') === 0) {
        var totalRowRange = sheet.getRange(j, 1, 1, 2);
        totalRowRange.setBackground('#4CAF50');
        totalRowRange.setFontColor(white);
        totalRowRange.setFontSize(12);
      }
    }
  }

  // Auto-fit columns
  var c;
  for (c = 1; c <= lastCol; c++) {
    sheet.autoResizeColumn(c);
  }

  // Freeze header rows
  sheet.setFrozenRows(dataStartRow > 0 ? dataStartRow : 4);
}

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
