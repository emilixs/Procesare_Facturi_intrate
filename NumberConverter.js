/**
 * Converts specified numeric columns in the invoice data to standardized number format
 * @param {Array[][]} data The raw data from the spreadsheet
 * @param {number} headerRow The row number where the header is located (0-based)
 * @return {Array[][]} The processed data with numeric columns converted
 */
function convertHeaderColumnsToNumbers(data, headerRow = 0) {
  // Define the specific columns to be converted
  const numericColumns = [
    'Suma TVA',    // Column E
    'Suma',        // Column F
    'Suma ramasa'  // Column G
  ];
  
  // Get header row and find indices of numeric columns
  const headers = data[headerRow];
  const numericColumnIndices = numericColumns.map(colName => headers.indexOf(colName))
    .filter(index => index !== -1);
  
  // Process all rows except header
  for (let i = headerRow + 1; i < data.length; i++) {
    for (const colIndex of numericColumnIndices) {
      const value = data[i][colIndex];
      if (value !== null && value !== undefined && value !== '') {
        // Remove currency symbols, spaces, and standardize decimal separator
        const cleanValue = String(value)
          .replace(/[^\d,.-]/g, '')  // Remove any character that's not a digit, comma, dot, or minus
          .replace(/\s/g, '')        // Remove spaces
          .replace(/,/g, '.');       // Replace commas with dots for standardization
        
        // Convert to number, default to 0 if conversion fails
        const numericValue = Number(cleanValue);
        data[i][colIndex] = isNaN(numericValue) ? 0 : numericValue;
      } else {
        // Set empty/null values to 0
        data[i][colIndex] = 0;
      }
    }
  }
  
  return data;
}

/**
 * Main function to process invoice data
 * Triggered from the menu
 */
function processInvoiceData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const range = sheet.getDataRange();
    const data = range.getValues();
    
    // Process the data
    const processedData = convertHeaderColumnsToNumbers(data);
    
    // Write back to sheet
    range.setValues(processedData);
    
    // Notify user
    SpreadsheetApp.getUi().alert(
      'Success', 
      'Numeric columns (Suma TVA, Suma, Suma ramasa) have been processed successfully!', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    // Log error for debugging
    console.error('Error in processInvoiceData:', error);
    
    // Notify user of error
    SpreadsheetApp.getUi().alert(
      'Error', 
      'An error occurred while processing numeric columns: ' + error.message, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Adds a menu item to trigger the invoice processing
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Invoice Processing')
    .addItem('Convert Numeric Columns', 'processInvoiceData')
    .addToUi();
} 