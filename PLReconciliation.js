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
  const startTime = new Date();

  // Validate sheets exist
  const expensesSheet = targetSpreadsheet.getSheetByName('Expenses');
  if (!expensesSheet) {
    throw new Error('Expenses sheet not found in target spreadsheet');
  }

  const staffingSheet = targetSpreadsheet.getSheetByName('Staffing');
  if (!staffingSheet) {
    throw new Error('Staffing sheet not found in target spreadsheet');
  }

  // Validate sheet structures
  function validateSheetStructure(sheet, sheetName, requiredColumn) {
    const headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    const columnIndex = requiredColumn.toUpperCase().charCodeAt(0) - 65;
    
    if (columnIndex < 0 || columnIndex >= headers.length) {
      throw new Error(`Required column ${requiredColumn} not found in ${sheetName} sheet`);
    }
  }

  // Validate both sheets
  validateSheetStructure(expensesSheet, 'Expenses', 'C');  // Validate Furnizor column
  validateSheetStructure(staffingSheet, 'Staffing', 'D');  // Validate Partener column

  const monthColumn = `${month} real`;

  /**
   * Creates a matching query for the LLM
   * @private
   */
  function createMatchingQuery(supplier, targetData) {
    const prompt = `Act as an experienced accountant who understands company names and their variations.

Task: Find if this company "${supplier}" matches any company in the list below.

When trying to recognize act as a normal human being, and user your brain to understand the contex. 
The files are used for different purposes so extra information might be added to the company name. 
Do not consider the company specific denominators as SRL, S.R.L., SA, S.A. as they might or might not be present. 

List of potential matches:
${targetData.join('\n')}

Reply in JSON format:
{
  "matched": boolean,
  "lineNumber": number or null,
  "confidence": number (0.0-1.0),
  "explanation": "Detailed explanation of why this is or isn't a match, including what patterns or variations were considered"
}

Remember: It's better to find a correct match with medium confidence than miss a valid match due to strict matching.`;

    // Log only the prompt
    console.log("\n=== LLM REQUEST ===");
    console.log(prompt);
    console.log("=== END REQUEST ===\n");

    return prompt;
  }

  /**
   * Process LLM response for matching decision
   * @private
   */
  function processLLMResponse(response) {
    try {
      // Log only the response
      console.log("\n=== LLM RESPONSE ===");
      console.log(response);
      console.log("=== END RESPONSE ===\n");

      return JSON.parse(response);
    } catch (error) {
      throw error;
    }
  }

  /**
   * Updates the matched status in the source file
   * @private
   */
  function updateMatchedStatus(row, matchResult) {
    try {
      const sheet = sourceSpreadsheet.getActiveSheet();
      const matchedCell = sheet.getRange(row, 16); // Column P

      if (matchResult.isMatch) {
        matchedCell.setValue(matchResult.reference);
        matchedCell.setBackground('#b7e1cd'); // Green
      } else {
        matchedCell.setValue('No match');
        matchedCell.setBackground('#cccccc'); // Gray
      }
    } catch (error) {
      throw error;
    }
  }

  /**
   * Matches and updates a single entry
   * @private
   */
  function matchAndUpdateEntry(entry) {
    try {
      // Get data from both sheets
      const expensesData = expensesSheet.getDataRange().getValues();
      const staffingData = staffingSheet.getDataRange().getValues();

      // Create potential matches from both sheets with direct cell references
      const expensesMatches = expensesData.slice(1)
        .map((row, index) => ({
          name: row[2].toString(), // Column C
          reference: `Expenses!C${index + 2}`
        }))
        .filter(match => match.name.trim() !== '');

      const staffingMatches = staffingData.slice(1)
        .map((row, index) => ({
          name: row[3].toString(), // Column D
          reference: `Staffing!D${index + 2}`
        }))
        .filter(match => match.name.trim() !== '');

      // Combine matches from both sheets
      const allPotentialMatches = [...expensesMatches, ...staffingMatches];

      // Create single prompt with all matches
      const prompt = createMatchingQuery(entry.supplier, 
        allPotentialMatches.map(m => `${m.reference}: ${m.name}`));

      const claude = getClaudeService();
      // Pass the matches with name and reference
      const matchResult = claude.matchClient(entry.supplier, allPotentialMatches);

      if (matchResult.matched && matchResult.confidence > 0.5) {
        // Parse the reference to get sheet and cell
        const [sheetName, cellRef] = matchResult.reference.split('!');
        const targetSheet = sheetName === 'Expenses' ? expensesSheet : staffingSheet;
        
        // Get the row number from the cell reference (e.g., C128 -> 128)
        const rowNumber = parseInt(cellRef.match(/\d+/)[0]);
        
        // Find the month column in the target sheet
        const headers = targetSheet.getRange(2, 1, 1, targetSheet.getLastColumn()).getValues()[0]
          .map(header => header.toString().trim());
        
        const monthColumnIndex = headers.findIndex(header => {
          const headerStr = header.toString().trim();
          const monthStr = monthColumn.toString().trim();
          return headerStr === monthStr;
        });
        
        if (monthColumnIndex === -1) {
          const nonEmptyHeaders = headers.filter(h => h !== '');
          throw new Error(`Column "${monthColumn}" not found in ${sheetName} sheet. Available non-empty columns: ${nonEmptyHeaders.join(', ')}`);
        }

        // Update the amount in the target sheet using the exact row number
        const targetCell = targetSheet.getRange(rowNumber, monthColumnIndex + 1);
        const currentValue = targetCell.getValue() || 0;
        const newValue = currentValue + entry.amount;
        targetCell.setValue(newValue);

        return {
          isMatch: true,
          reference: matchResult.reference,
          confidence: matchResult.confidence,
          explanation: matchResult.explanation,
          sheet: sheetName
        };
      }

      return { isMatch: false };
    } catch (error) {
      throw error;
    }
  }

  /**
   * Converts a column number to letter reference (e.g., 1 -> A, 27 -> AA)
   * @private
   */
  function columnToLetter(column) {
    let temp, letter = '';
    while (column > 0) {
      temp = (column - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      column = (column - temp - 1) / 26;
    }
    return letter;
  }

  /**
   * Main reconciliation process
   */
  function processReconciliation(testMode = true) {
    try {
      const sheet = sourceSpreadsheet.getActiveSheet();
      const data = sheet.getDataRange().getValues();
      
      let processedCount = 0;
      let matchedCount = 0;
      
      // Calculate how many rows to process
      const maxRows = testMode ? Math.min(11, data.length) : data.length;

      // Process each row starting from row 2
      for (let i = 1; i < maxRows; i++) {
        const entry = {
          supplier: data[i][1],    // Column B (Furnizor)
          amount: data[i][14],     // Column O (Suma in EUR)
          isMatched: data[i][15]   // Column P (Matched P&L)
        };

        // Skip if there's a valid match reference
        if (entry.isMatched && 
            entry.isMatched !== '' && 
            entry.isMatched !== 'No match' && 
            entry.isMatched.includes('!')) {
          continue;
        }

        processedCount++;
        const matchResult = matchAndUpdateEntry(entry);
        if (matchResult.isMatch) {
          matchedCount++;
        }
        updateMatchedStatus(i + 1, matchResult);
      }

      return {
        processedCount,
        matchedCount,
        successRate: (matchedCount / processedCount * 100).toFixed(2) + '%',
        mode: testMode ? 'test' : 'full',
        rowsProcessed: maxRows - 1
      };
    } catch (error) {
      throw error;
    }
  }

  /**
   * Checks a sheet for supplier match
   * @private
   * @param {Object} entry - The entry to match
   * @param {Sheet} sheet - The sheet to check
   * @param {string} matchColumn - The column letter to match against
   * @param {string} sheetName - Name of the sheet for logging
   * @returns {Object} Match result
   */
  function checkSheetForMatch(entry, sheet, matchColumn, sheetName) {
    try {
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const matchColumnIndex = matchColumn.toUpperCase().charCodeAt(0) - 65;

      if (matchColumnIndex < 0 || matchColumnIndex >= headers.length) {
        throw new Error(`Invalid column ${matchColumn} in ${sheetName}`);
      }

      // Create potential matches
      const potentialMatches = data.slice(1)
        .map((row, index) => ({
          name: row[matchColumnIndex].toString(),
          reference: `${sheetName}!${matchColumn}${index + 2}`
        }))
        .filter(match => match.name.trim() !== '');

      const prompt = createMatchingQuery(entry.supplier, 
        potentialMatches.map(m => `${m.reference}: ${m.name}`));

      const claude = getClaudeService();
      const matchResult = claude.matchClient(entry.supplier, potentialMatches);

      if (matchResult.matched && matchResult.confidence > 0.5) {
        return {
          isMatch: true,
          reference: matchResult.reference,
          confidence: matchResult.confidence,
          explanation: matchResult.explanation
        };
      }

      return { isMatch: false };

    } catch (error) {
      throw error;
    }
  }

  // Return public methods
  return {
    processReconciliation,
    matchAndUpdateEntry,
    checkSheetForMatch,
    processTestReconciliation: () => processReconciliation(true),
    processFullReconciliation: () => processReconciliation(false)
  };
} 