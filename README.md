# Tantu Bamboo Crafts — GST Invoice Tracker & Tax Calculator

Smartphone-first GST Invoice Generator and Google Sheets Monthly Tax Tracker for **Tantu Bamboo Crafts**.

## Features

- **Create Bill** — Generate GST invoices with auto-calculated CGST/SGST/IGST
- **Transport Entry** — Add transport bill & tax for saved invoices
- **PDF Export** — Download crisp A4 PDF invoices
- **Google Sheets Sync** — Auto-save invoices to monthly tabs in Google Sheets
- **Mobile Optimized** — Designed for smartphone screens (360px–430px)

## Files

| File | Description |
|------|-------------|
| `Tantu_Tax_Tracker_fixed.html` | Main HTML app (open in browser) |
| `Code_gs_fixed.gs` | Google Apps Script for invoice saving (bound to Google Sheet) |
| `Transport_Update.gs` | Standalone Apps Script for transport data updates |
| `START_SERVER.bat` | Optional local HTTP server for testing |

## Setup

1. Open `Tantu_Tax_Tracker_fixed.html` in any browser
2. Deploy `Code_gs_fixed.gs` as a Web App in Google Apps Script (bound to your Google Sheet)
3. Deploy `Transport_Update.gs` as a separate standalone Web App
4. Update the URLs in the HTML file (`SHEETS_WEB_APP_URL` and `TRANSPORT_WEB_APP_URL`)

## Google Sheet Structure

- **Row 1**: Monthly summary headers
- **Row 2**: Auto-calculated summary formulas
- **Row 4**: Column headers (A–S)
- **Row 5+**: Invoice data
- **Columns A–N**: Invoice data (auto-saved)
- **Columns O–Q**: Transport bill & tax (manual entry via Transport tab)
- **Column R**: Total Transport Tax (auto formula)
- **Column S**: Total Tax To Pay = Sales Tax + Transport Tax (auto formula)

## Tech Stack

- HTML5 / Tailwind CSS / Vanilla JavaScript
- Google Apps Script (server-side)
- Google Sheets (database)
- html2pdf.js (PDF generation)

