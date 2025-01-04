/**
 * Creates a PLReconciliation service instance
 * @param {string} spreadsheetUrl - URL of the target spreadsheet
 * @param {string} month - Month to process (e.g., "October")
 * @return {Object} PLReconciliation service methods
 */
function createPLReconciliationService(spreadsheetUrl, month) {
  // Private variables
  const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const targetSpreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
  const expensesSheet = targetSpreadsheet.getSheetByName('Expenses');
  const staffingSheet = targetSpreadsheet.getSheetByName('Staffing');
  const monthColumn = `${month} real`;

  /**
   * Creates a matching query for the LLM
   * @param {string} supplier - Supplier name to match
   * @param {Array} targetData - Array of potential matches
   * @return {string} Formatted query
   */
  function createMatchingQuery(supplier, targetData) {
    return `Compare "${supplier}" with the following potential matches and determine the best match if any exists:\n${targetData.join('\n')}`;
  }

  /**
   * Process LLM response for matching decision
   * @param {string} response - LLM response
   * @return {Object} Matching decision
   */
  function processLLMResponse(response) {
    // Implementation will depend on Claude API integration
    return {
      isMatch: false,
      confidence: 0,
      matchedEntry: null
    };
  }

  /**
   * Updates the matched status in the source file
   * @param {number} row - Row number
   * @param {Object} matchResult - Match result details
   */
  function updateMatchedStatus(row, matchResult) {
    const sheet = sourceSpreadsheet.getActiveSheet();
    const matchedCell = sheet.getRange(row, 16); // Column P

    if (matchResult.isMatch) {
      matchedCell.setValue(matchResult.reference);
      matchedCell.setBackground('#b7e1cd'); // Green
    } else {
      matchedCell.setValue('No match');
      matchedCell.setBackground('#cccccc'); // Gray
    }
  }

  /**
   * Matches and updates a single entry
   * @param {Object} entry - Entry data
   * @return {Object} Match results
   */
  function matchAndUpdateEntry(entry) {
    // Check Expenses sheet
    const expensesMatch = checkSheetForMatch(entry, expensesSheet, 'C', 'Expenses');
    if (expensesMatch.isMatch) {
      updateAmount(expensesMatch, entry.amount);
      return expensesMatch;
    }

    // Check Staffing sheet
    const staffingMatch = checkSheetForMatch(entry, staffingSheet, 'D', 'Staffing');
    if (staffingMatch.isMatch) {
      updateAmount(staffingMatch, entry.amount);
      return staffingMatch;
    }

    return { isMatch: false };
  }

  /**
   * Main reconciliation process
   */
  function processReconciliation() {
    const sheet = sourceSpreadsheet.getActiveSheet();
    const data = sheet.getDataRange().getValues();
    const headerRow = data[0];

    // Process each row starting from row 2
    for (let i = 1; i < data.length; i++) {
      const entry = {
        supplier: data[i][1], // Column B
        amount: data[i][5],   // Column F
        isMatched: data[i][15] // Column P
      };

      // Skip if already matched
      if (entry.isMatched && entry.isMatched !== '') continue;

      const matchResult = matchAndUpdateEntry(entry);
      updateMatchedStatus(i + 1, matchResult);
    }
  }

  // Return public methods
  return {
    processReconciliation,
    matchAndUpdateEntry
  };
} 