/**
 * Main function to process invoice data
 * Triggered from the menu
 */
function processInvoiceData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    
    // Define the columns to process
    const columnsToProcess = {
      'E': 'Suma TVA',
      'F': 'Suma',
      'G': 'Suma ramasa'
    };

    // Get the data for processing
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    // Process each column separately
    Object.entries(columnsToProcess).forEach(([col, name]) => {
      const colIndex = headers.findIndex(header => header.toString().trim() === name);
      if (colIndex !== -1) {
        // Get only the column data
        const columnRange = sheet.getRange(2, colIndex + 1, lastRow - 1, 1);
        const columnData = columnRange.getValues();
        
        // Process the numbers
        const processedData = columnData.map(([cell]) => [convertToNumber(cell)]);
        
        // Update only this column
        columnRange.setValues(processedData);
      }
    });
    
    // Notify user
    SpreadsheetApp.getUi().alert(
      'Success', 
      'Numeric columns (Suma TVA, Suma, Suma ramasa) have been processed successfully! All other columns were preserved.', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
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
 * Converts a value to a standardized number format
 * @private
 */
function convertToNumber(value) {
  if (typeof value === 'number') return value;
  if (!value) return 0;

  // Convert to string and clean up
  let strValue = value.toString()
    .replace(/[^\d.,\-]/g, '') // Remove all except digits, dots, commas and minus
    .trim();

  if (!strValue) return 0;

  // Handle European format (1.234,56)
  if (strValue.includes(',') && strValue.includes('.')) {
    strValue = strValue.replace(/\./g, '').replace(',', '.');
  }
  // Handle simple comma as decimal separator
  else if (strValue.includes(',')) {
    strValue = strValue.replace(',', '.');
  }

  const number = parseFloat(strValue);
  return isNaN(number) ? 0 : number;
} 