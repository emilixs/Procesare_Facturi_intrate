/**
 * Converts header columns containing numbers to standardized number format
 * @param {Array} data - The spreadsheet data
 * @param {number} headerRow - The row containing headers (default: 0)
 * @returns {Array} - The processed data
 */
function convertHeaderColumnsToNumbers(data, headerRow = 0) {
  // Define columns to process (only up to column N)
  const columnsToProcess = {
    'E': 'Suma TVA',
    'F': 'Suma',
    'G': 'Suma ramasa'
  };

  // Get column indices
  const columnIndices = {};
  data[headerRow].forEach((header, index) => {
    if (index <= 13) { // Only process up to column N (index 13)
      const headerText = header.toString().trim();
      Object.entries(columnsToProcess).forEach(([col, name]) => {
        if (headerText === name) {
          columnIndices[col] = index;
        }
      });
    }
  });

  // Process each row
  return data.map((row, rowIndex) => {
    if (rowIndex === headerRow) return row;

    // Create a new row array
    return row.map((cell, colIndex) => {
      // Only process if column is in our list and is before column O
      if (colIndex <= 13 && Object.values(columnIndices).includes(colIndex)) {
        return convertToNumber(cell);
      }
      return cell; // Return unchanged for other columns
    });
  });
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
      'Numeric columns have been processed successfully! Formulas after column N were preserved.', 
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