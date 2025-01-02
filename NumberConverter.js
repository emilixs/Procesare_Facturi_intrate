/**
 * Converts specified columns in the invoice header to number format
 * @param {Array[][]} data The raw data from the spreadsheet
 * @param {number} headerRow The row number where the header is located (0-based)
 * @return {Array[][]} The processed data with numeric columns converted
 */
function convertHeaderColumnsToNumbers(data, headerRow = 0) {
  // Columns to be converted to numbers
  const numericColumns = [
    'Suma incasata',
    'Incasata prin',
    'Suma ramasa de incasat',
    'Valoare',
    'TVA',
    'Total',
    'CursValutar'
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
        // Remove any currency symbols, spaces, and convert commas to dots
        const cleanValue = String(value)
          .replace(/[^\d,.-]/g, '')  // Remove any character that's not a digit, comma, dot, or minus
          .replace(/,/g, '.');       // Replace commas with dots
        
        // Convert to number
        data[i][colIndex] = Number(cleanValue) || 0;
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
    const data = sheet.getDataRange().getValues();
    
    // Process the data
    const processedData = convertHeaderColumnsToNumbers(data);
    
    // Write back to sheet
    sheet.getDataRange().setValues(processedData);
    
    // Notify user
    SpreadsheetApp.getUi().alert('Success', 'Numeric columns have been processed successfully!', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'An error occurred: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
} 