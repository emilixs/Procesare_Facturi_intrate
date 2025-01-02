/**
 * Creates the custom menu in Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Invoice Tools')
    .addItem('To Numbers', 'processInvoiceData')
    .addToUi();
} 