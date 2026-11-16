const MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

function setupCurrentMonthTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const tabName = getMonthTabName(todayStr);
  getOrCreateMonthSheet(ss, tabName);
}

function getMonthTabName(invDateStr) {
  let d = new Date();
  if (invDateStr && /^\d{4}-\d{2}-\d{2}$/.test(invDateStr)) {
    const parts = invDateStr.split('-');
    d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  }
  return MONTH_NAMES[d.getMonth()] + " " + d.getFullYear(); // e.g., "Oct 2026"
}

function getOrCreateMonthSheet(ss, tabName) {
  let sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);

    // Row 1: Monthly Summary Headers
    const summaryHeaders = [
      "Month",
      "Total Invoices",
      "Sales Subtotal (₹)",
      "Sales CGST (₹)",
      "Sales SGST (₹)",
      "Sales IGST (₹)",
      "Total Sales Tax (₹)",
      "Total Transport Freight (₹)",
      "Transport CGST 2.5% (₹)",
      "Transport SGST 2.5% (₹)",
      "Total Transport Tax (₹)",
      "GRAND TOTAL TAX TO PAY (₹)",
      "Sales Grand Total (₹)"
    ];
    sheet.getRange(1, 1, 1, summaryHeaders.length).setValues([summaryHeaders])
      .setFontWeight("bold")
      .setBackground("#1e293b")
      .setFontColor("#ffffff");

    // Row 2: Live Formulas for Auto Calculation
    const summaryFormulas = [
      tabName,
      "=COUNTA(B5:B)",
      "=SUM(I5:I)",
      "=SUM(J5:J)",
      "=SUM(K5:K)",
      "=SUM(L5:L)",
      "=SUM(M5:M)",
      "=SUM(O5:O)",
      "=SUM(P5:P)",
      "=SUM(Q5:Q)",
      "=SUM(R5:R)",
      "=SUM(M5:M)+SUM(R5:R)",
      "=SUM(N5:N)"
    ];
    sheet.getRange(2, 1, 1, summaryFormulas.length).setValues([summaryFormulas])
      .setFontWeight("bold")
      .setBackground("#f8fafc");

    // Highlight summary cells
    sheet.getRange("K1").setBackground("#d97706").setFontColor("#ffffff");
    sheet.getRange("K2").setBackground("#fffbeb").setFontColor("#b45309");
    sheet.getRange("L1").setBackground("#dc2626").setFontColor("#ffffff");
    sheet.getRange("L2").setBackground("#fef2f2").setFontColor("#dc2626").setFontSize(11);

    // Row 4: Table Column Headers including Column S
    const tableHeaders = [
      "Timestamp", "Invoice Number", "Invoice Date", "Customer Name",
      "Customer GSTIN", "Phone", "Address", "Items Summary",
      "Subtotal (₹)", "Sales CGST (₹)", "Sales SGST (₹)", "Sales IGST (₹)",
      "Total Sales Tax (₹)", "Invoice Grand Total (₹)",
      "Transport Bill Amount ₹ (Manual)",
      "Transport CGST 2.5% ₹ (Manual)", "Transport SGST 2.5% ₹ (Manual)",
      "Total Transport Tax ₹ (Auto/Manual)",
      "Total Tax To Pay ₹ (Sales + Transport)"
    ];
    sheet.getRange(4, 1, 1, tableHeaders.length).setValues([tableHeaders])
      .setFontWeight("bold");

    sheet.getRange(4, 1, 1, 14).setBackground("#dbeafe").setFontColor("#1e3a8a");
    sheet.getRange(4, 15, 1, 4).setBackground("#fef08a").setFontColor("#854d0e");
    sheet.getRange(4, 19).setBackground("#fee2e2").setFontColor("#991b1b");

    // Pre-fill Transport Tax formula down Column R and Combined Tax formula down Column S for rows 5 to 200
    const transportFormulas = [];
    const combinedTaxFormulas = [];
    for (let r = 5; r <= 200; r++) {
      transportFormulas.push([`=IF(AND(P${r}="",Q${r}=""),"",SUM(P${r}:Q${r}))`]);
      combinedTaxFormulas.push([`=IF(AND(M${r}="",R${r}=""),"",SUM(M${r},R${r}))`]);
    }
    sheet.getRange(5, 18, transportFormulas.length, 1).setFormulas(transportFormulas);
    sheet.getRange(5, 19, combinedTaxFormulas.length, 1).setFormulas(combinedTaxFormulas);

    sheet.setFrozenRows(4);
  }
  return sheet;
}

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = JSON.parse(e.postData.contents);

    // ---- ACTION: Fetch all invoices (same logic as doGet, but via POST) ----
    if (data.action === 'fetchInvoices') {
      return fetchAllInvoicesResponse_(ss);
    }

    // ---- ACTION: Update Transport Columns (O, P, Q) for an existing invoice ----
    if (data.action === 'updateTransport') {
      const monthTab = data.monthTab;
      const invNumber = String(data.invNumber || '').trim();
      if (!monthTab || !invNumber) throw new Error('monthTab and invNumber are required');

      const sheet = ss.getSheetByName(monthTab);
      if (!sheet) throw new Error('Sheet tab not found: ' + monthTab);

      const colB = sheet.getRange('B5:B500').getValues();
      let targetRow = -1;
      for (let i = 0; i < colB.length; i++) {
        if (String(colB[i][0]).trim() === invNumber) {
          targetRow = 5 + i;
          break;
        }
      }
      if (targetRow === -1) throw new Error('Invoice not found: ' + invNumber);

      // Write transport data to Columns O (15), P (16), Q (17)
      const freight = Number(data.transportFreight) || 0;
      const tCgst = Number(data.transportCgst) || 0;
      const tSgst = Number(data.transportSgst) || 0;
      sheet.getRange(targetRow, 15, 1, 3).setValues([[ freight, tCgst, tSgst ]]);

      // Ensure Column R has the transport tax formula
      sheet.getRange(targetRow, 18).setFormula(
        '=IF(AND(P' + targetRow + '="",Q' + targetRow + '=""),"",SUM(P' + targetRow + ':Q' + targetRow + '))'
      );
      // Ensure Column S has the combined tax formula
      sheet.getRange(targetRow, 19).setFormula(
        '=IF(AND(M' + targetRow + '="",R' + targetRow + '=""),"",SUM(M' + targetRow + ',R' + targetRow + '))'
      );

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
    }

    // ---- DEFAULT ACTION: Save new invoice (Columns A-N) ----
    const subtotal = Number(data.subtotal) || 0;
    const cgst = Number(data.cgst) || 0;
    const sgst = Number(data.sgst) || 0;
    const igst = Number(data.igst) || 0;
    const totalTax = Number(data.totalTax) || (cgst + sgst + igst);
    const grandTotal = Number(data.grandTotal) || 0;

    const monthTabName = getMonthTabName(data.invDate);
    const monthSheet = getOrCreateMonthSheet(ss, monthTabName);

    const colB = monthSheet.getRange("B5:B500").getValues();
    let targetRow = 5;
    for (let i = 0; i < colB.length; i++) {
      if (colB[i][0] === "" || colB[i][0] === null) {
        targetRow = 5 + i;
        break;
      }
    }

    // Write website invoice data to Columns A to N
    monthSheet.getRange(targetRow, 1, 1, 14).setValues([[
      new Date().toISOString(),
      data.invNumber || '',
      data.invDate || '',
      data.cName || '',
      data.cGst || 'URP',
      data.cPhone || '',
      data.cAddress || '',
      data.itemsSummary || '',
      subtotal,
      cgst,
      sgst,
      igst,
      totalTax,
      grandTotal
    ]]);

    // Ensure Column S has the combined tax formula (=M + R) for this row
    monthSheet.getRange(targetRow, 19).setFormula(`=IF(AND(M${targetRow}="",R${targetRow}=""),"",SUM(M${targetRow},R${targetRow}))`);

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      monthTab: monthTabName,
      row: targetRow
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Convert a cell value to the canonical "Mon YYYY" month name.
 * Handles: Date objects, date-like strings, and plain "Mon YYYY" strings.
 * Always falls back to the sheet tab name when the cell is blank or unparseable.
 */
function resolveMonthName_(cellValue, sheetName) {
  if (cellValue instanceof Date && !isNaN(cellValue.getTime())) {
    return MONTH_NAMES[cellValue.getMonth()] + ' ' + cellValue.getFullYear();
  }
  var s = String(cellValue === null || cellValue === undefined ? '' : cellValue).trim();
  if (!s) return sheetName;
  // If it already looks like "Oct 2026", use it directly.
  if (/^[A-Za-z]{3}\s+\d{4}$/.test(s)) return s;
  // Try to parse as a date string (e.g. "Mon Oct 01 2026 ..." from toString()).
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return MONTH_NAMES[d.getMonth()] + ' ' + d.getFullYear();
  }
  return sheetName;
}

/**
 * Shared helper: reads all invoice data from all monthly tabs and returns
 * a ContentService response (JSON).  Used by both doGet and doPost(fetchInvoices).
 */
function fetchAllInvoicesResponse_(ss) {
    SpreadsheetApp.flush();

    var invoices = [];
    var summaries = [];

    ss.getSheets().forEach(function(sh) {
      var sheetName = sh.getName();
      if (sheetName === 'Sheet1') return;

      var lastRow = sh.getLastRow();
      if (lastRow < 2) return;

      var maxCol = Math.max(sh.getLastColumn(), 13);
      var headers = sh.getRange(1, 1, 1, maxCol)
        .getDisplayValues()[0].map(normalizeHeaderForApi_);

      var isSummaryLayout = headers[1] === 'total invoices' &&
        headers.some(function(h) { return h.indexOf('grand total tax to pay') !== -1; });

      if (isSummaryLayout) {
        var row = sh.getRange(2, 1, 1, 13).getValues()[0];
        var monthName = resolveMonthName_(row[0], sheetName);
        var count = apiNumber_(row[1]);
        var salesTax = apiNumber_(row[6]);
        var transportTax = apiNumber_(row[10]);
        var combinedTax = apiNumber_(row[11]);
        var revenue = apiNumber_(row[12]);
        summaries.push({
          month: monthName,
          totalInvoices: count,
          invoiceCount: count,
          count: count,
          subtotal: apiNumber_(row[2]),
          salesSubtotal: apiNumber_(row[2]),
          salesCgst: apiNumber_(row[3]),
          salesSgst: apiNumber_(row[4]),
          salesIgst: apiNumber_(row[5]),
          salesTax: salesTax,
          totalSalesTax: salesTax,
          totalTransportFreight: apiNumber_(row[7]),
          transportCgst: apiNumber_(row[8]),
          transportSgst: apiNumber_(row[9]),
          transportTax: transportTax,
          totalTransportTax: transportTax,
          combinedTax: combinedTax,
          totalCombinedTax: combinedTax,
          grandTotalTaxToPay: combinedTax,
          revenue: revenue,
          salesGrandTotal: revenue,
          grandTotalRevenue: revenue
        });
      }

      if (lastRow >= 5) {
        var rows = sh.getRange(5, 1, lastRow - 4, Math.max(sh.getLastColumn(), 19)).getValues();
        rows.forEach(function(r) {
          var invoiceNumber = String(r[1] === null || r[1] === undefined ? '' : r[1]).trim();
          if (!invoiceNumber) return;

          var customerName = String(r[3] || '').trim();
          var itemsSummary = String(r[7] || '').trim();
          var subtotal = apiNumber_(r[8]);
          var sTax = apiNumber_(r[12]);
          var tTax = apiNumber_(r[17]);
          var cTax = isApiBlank_(r[18]) ? sTax + tTax : apiNumber_(r[18]);
          var total = !isApiBlank_(r[13]) ? apiNumber_(r[13])
            : (!isApiBlank_(r[19]) ? apiNumber_(r[19]) : subtotal + sTax);

          var normalizedNumber = normalizeHeaderForApi_(invoiceNumber);
          if (normalizedNumber === 'invoice number' || normalizedNumber === 'invoice no' ||
              normalizedNumber === 'total invoices') return;
          if (!customerName && !itemsSummary && subtotal === 0 && !r[2]) return;

          var invDate = '';
          if (r[2]) {
            invDate = r[2] instanceof Date
              ? Utilities.formatDate(r[2], Session.getScriptTimeZone(), 'yyyy-MM-dd')
              : String(r[2]);
          }

          invoices.push({
            month: sheetName,
            monthTab: sheetName,
            timestamp: r[0],
            invNumber: invoiceNumber,
            invDate: invDate,
            cName: customerName || 'Customer',
            cGst: String(r[4] || 'URP'),
            cPhone: String(r[5] || ''),
            cAddress: String(r[6] || ''),
            itemsSummary: itemsSummary,
            subtotal: subtotal,
            cgst: apiNumber_(r[9]),
            sgst: apiNumber_(r[10]),
            igst: apiNumber_(r[11]),
            salesTax: sTax,
            totalSalesTax: sTax,
            grandTotal: total,
            transportFreight: apiNumber_(r[14]),
            transportCgst: apiNumber_(r[15]),
            transportSgst: apiNumber_(r[16]),
            transportTax: tTax,
            totalTransportTax: tTax,
            totalCombinedTax: cTax,
            hasTransportData: !isApiBlank_(r[14])
          });
        });
      }
    });

    invoices.reverse();
    var response = {
      status: 'success',
      summaries: summaries,
      monthlySummaries: summaries,
      invoices: invoices
    };
    return ContentService.createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var result = fetchAllInvoicesResponse_(ss);
    // JSONP support
    if (e && e.parameter && e.parameter.callback) {
      var jsonStr = result.getContent();
      return ContentService.createTextOutput(e.parameter.callback + '(' + jsonStr + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return result;
  } catch (err) {
    var errorResponse = {
      status: 'error',
      message: String(err && err.message ? err.message : err),
      summaries: [],
      monthlySummaries: [],
      invoices: []
    };
    var errorStr = JSON.stringify(errorResponse);
    if (e && e.parameter && e.parameter.callback) {
      return ContentService.createTextOutput(e.parameter.callback + '(' + errorStr + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(errorStr)
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function normalizeHeaderForApi_(value) {
  return String(value === null || value === undefined ? '' : value)
    .toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}

function isApiBlank_(value) {
  return value === null || value === undefined || String(value).trim() === '';
}

function apiNumber_(value) {
  if (isApiBlank_(value)) return 0;
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  const parsed = Number(String(value).replace(/[^0-9.-]/g, ''));
  return isFinite(parsed) ? parsed : 0;
}
