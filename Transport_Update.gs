/**
 * TANTU Transport Updater — Standalone Script
 * 
 * This is a SEPARATE project that ONLY handles transport column updates.
 * It opens the spreadsheet by ID (not bound to any sheet).
 * 
 * Deploy this as: Web App → Execute as "Me" → Access "Anyone"
 */

const SPREADSHEET_ID = '1V6loraO_LvVCZDdJ8ky24aHURZuceDmgicdESns6SbY';

function doPost(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const data = JSON.parse(e.postData.contents);

    // Only handle updateTransport action
    if (data.action !== 'updateTransport') {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'error',
        message: 'Unknown action: ' + (data.action || 'none')
      })).setMimeType(ContentService.MimeType.JSON);
    }

    const monthTab = data.monthTab;
    const invNumber = String(data.invNumber || '').trim();

    if (!monthTab || !invNumber) {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'error',
        message: 'monthTab and invNumber are required'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    const sheet = ss.getSheetByName(monthTab);
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'error',
        message: 'Sheet tab not found: ' + monthTab
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Find the invoice row by matching column B
    const colB = sheet.getRange('B5:B500').getValues();
    let targetRow = -1;
    for (let i = 0; i < colB.length; i++) {
      if (String(colB[i][0]).trim() === invNumber) {
        targetRow = 5 + i;
        break;
      }
    }

    if (targetRow === -1) {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'error',
        message: 'Invoice not found: ' + invNumber + ' in tab ' + monthTab
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Write transport data to Columns O (15), P (16), Q (17)
    const freight = Number(data.transportFreight) || 0;
    const tCgst = Number(data.transportCgst) || 0;
    const tSgst = Number(data.transportSgst) || 0;
    sheet.getRange(targetRow, 15, 1, 3).setValues([[freight, tCgst, tSgst]]);

    // Ensure Column R has transport tax formula
    sheet.getRange(targetRow, 18).setFormula(
      '=IF(AND(P' + targetRow + '="",Q' + targetRow + '=""),"",SUM(P' + targetRow + ':Q' + targetRow + '))'
    );
    // Ensure Column S has combined tax formula
    sheet.getRange(targetRow, 19).setFormula(
      '=IF(AND(M' + targetRow + '="",R' + targetRow + '=""),"",SUM(M' + targetRow + ',R' + targetRow + '))'
    );

    // Ensure Column S header exists in row 4
    var s4 = sheet.getRange(4, 19);
    if (!s4.getValue()) {
      s4.setValue('Total Tax To Pay ₹ (Sales + Transport)')
        .setFontWeight('bold')
        .setBackground('#fee2e2')
        .setFontColor('#991b1b');
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      action: 'updateTransport',
      monthTab: monthTab,
      invNumber: invNumber,
      row: targetRow,
      transportFreight: freight,
      transportCgst: tCgst,
      transportSgst: tSgst
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    message: 'Transport Updater is running. Use POST to update transport data.'
  })).setMimeType(ContentService.MimeType.JSON);
}

