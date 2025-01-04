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

  // Initialize logging context first
  const loggingContext = {
    sessionId: Utilities.getUuid(),
    startTimestamp: startTime.toISOString(),
    sourceSpreadsheetId: sourceSpreadsheet.getId(),
    targetSpreadsheetId: targetSpreadsheet.getId(),
    month: month
  };

  /**
   * Log structured data with context
   * @private
   */
  function logEvent(eventType, data) {
    const logEntry = {
      ...loggingContext,
      timestamp: new Date().toISOString(),
      eventType,
      processingTime: new Date() - startTime,
      ...data
    };
    console.log(JSON.stringify(logEntry));
  }
  
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
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const columnIndex = requiredColumn.toUpperCase().charCodeAt(0) - 65;
    
    if (columnIndex < 0 || columnIndex >= headers.length) {
      throw new Error(`Required column ${requiredColumn} not found in ${sheetName} sheet`);
    }

    // Log sheet validation
    logEvent('sheet_validation', {
      sheet: sheetName,
      headers: headers,
      requiredColumn: requiredColumn,
      columnIndex: columnIndex,
      headerFound: headers[columnIndex]
    });
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
  "confidence": number (0-1),
  "explanation": "Detailed explanation of why this is or isn't a match, including what patterns or variations were considered"
}

Remember: It's better to find a correct match with medium confidence than miss a valid match due to strict matching.`;
    
    // Pre-request logging with full prompt
    logEvent('llm_request_preparation', {
      supplier,
      targetDataSize: targetData.length,
      context: 'supplier_matching',
      prompt: prompt,
      targetData: targetData,
      timestamp_sent: new Date().toISOString()
    });

    return prompt;
  }

  /**
   * Process LLM response for matching decision
   * @private
   */
  function processLLMResponse(response) {
    try {
      // Log raw response immediately upon receiving
      logEvent('llm_response_received', {
        responseSize: response.length,
        status: 'success',
        rawResponse: response,
        timestamp_received: new Date().toISOString()
      });

      const result = JSON.parse(response);
      
      // Log processed result with full context
      logEvent('llm_response_processed', {
        matchFound: result.matched,
        confidence: result.confidence,
        parsedResponse: result,
        processingStatus: 'success'
      });

      return result;
    } catch (error) {
      logEvent('llm_response_error', {
        error: error.message,
        stack: error.stack,
        rawResponse: response,
        processingStatus: 'failed',
        errorType: 'parsing_error',
        timestamp_error: new Date().toISOString()
      });
      throw error;
    }
  }

  /**
   * Updates the matched status in the source file
   * @private
   */
  function updateMatchedStatus(row, matchResult) {
    logEvent('status_update_start', {
      row,
      matchResult: {
        isMatch: matchResult.isMatch,
        reference: matchResult.reference
      }
    });

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

      logEvent('status_update_complete', {
        row,
        success: true
      });
    } catch (error) {
      logEvent('status_update_error', {
        row,
        error: error.message,
        stack: error.stack
      });
      throw error;
    }
  }

  /**
   * Matches and updates a single entry
   * @private
   */
  function matchAndUpdateEntry(entry) {
    logEvent('entry_processing_start', {
      supplier: entry.supplier,
      amount: entry.amount,
      timestamp_start: new Date().toISOString()
    });

    try {
      // Get data from both sheets
      const expensesData = expensesSheet.getDataRange().getValues();
      const staffingData = staffingSheet.getDataRange().getValues();

      // Create potential matches from both sheets
      const expensesMatches = expensesData.slice(1)
        .map((row, index) => ({
          text: row[2].toString(), // Column C
          rowIndex: index + 2,
          reference: `Expenses!C${index + 2}`
        }))
        .filter(match => match.text.trim() !== '');

      const staffingMatches = staffingData.slice(1)
        .map((row, index) => ({
          text: row[3].toString(), // Column D
          rowIndex: index + 2,
          reference: `Staffing!D${index + 2}`
        }))
        .filter(match => match.text.trim() !== '');

      // Combine matches from both sheets
      const allPotentialMatches = [...expensesMatches, ...staffingMatches];

      // Log combined matches info
      logEvent('combined_matches_preparation', {
        supplier: entry.supplier,
        expensesMatchCount: expensesMatches.length,
        staffingMatchCount: staffingMatches.length,
        totalMatches: allPotentialMatches.length
      });

      // Create single prompt with all matches
      const prompt = createMatchingQuery(entry.supplier, 
        allPotentialMatches.map(m => `${m.reference}: ${m.text}`));

      const claude = getClaudeService();
      const matchResult = claude.matchClient(entry.supplier, allPotentialMatches);

      // Log match result
      logEvent('match_attempt', {
        supplier: entry.supplier,
        matched: matchResult.matched,
        confidence: matchResult.confidence,
        lineNumber: matchResult.lineNumber,
        matchedText: matchResult.lineNumber ? allPotentialMatches[matchResult.lineNumber - 2].text : null,
        matchedIn: matchResult.lineNumber ? allPotentialMatches[matchResult.lineNumber - 2].reference.split('!')[0] : null
      });

      if (matchResult.matched && matchResult.confidence > 0.5) {
        const matchedEntry = allPotentialMatches[matchResult.lineNumber - 2];
        return {
          isMatch: true,
          reference: matchedEntry.reference,
          rowIndex: matchedEntry.rowIndex,
          confidence: matchResult.confidence,
          explanation: matchResult.explanation,
          sheet: matchedEntry.reference.split('!')[0]
        };
      }

      logEvent('no_match_found', {
        supplier: entry.supplier,
        timestamp_nomatch: new Date().toISOString()
      });

      return { isMatch: false };
    } catch (error) {
      logEvent('entry_processing_error', {
        supplier: entry.supplier,
        error: error.message,
        stack: error.stack
      });
      throw error;
    }
  }

  /**
   * Main reconciliation process
   */
  function processReconciliation(testMode = true) {
    logEvent('reconciliation_start', {
      totalRows: sourceSpreadsheet.getActiveSheet().getLastRow() - 1,
      mode: testMode ? 'test' : 'full'
    });

    try {
      const sheet = sourceSpreadsheet.getActiveSheet();
      const data = sheet.getDataRange().getValues();
      const headerRow = data[0];

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

        // Log the entry data being processed
        logEvent('processing_entry', {
          row: i + 1,
          supplier: entry.supplier,
          amountEUR: entry.amount,
          columnUsed: 'O',
          mode: testMode ? 'test' : 'full',
          currentMatchStatus: entry.isMatched
        });

        // Only skip if there's a valid match reference
        // Process if empty, "No match", or invalid value
        if (entry.isMatched && 
            entry.isMatched !== '' && 
            entry.isMatched !== 'No match' && 
            entry.isMatched.includes('!')) {  // Valid match references contain '!' for cell reference
          logEvent('skip_matched_entry', {
            row: i + 1,
            supplier: entry.supplier,
            existingMatch: entry.isMatched
          });
          continue;
        }

        // If we're reprocessing a previous "No match", log it
        if (entry.isMatched === 'No match') {
          logEvent('reprocessing_no_match', {
            row: i + 1,
            supplier: entry.supplier,
            previousStatus: entry.isMatched
          });
        }

        processedCount++;
        const matchResult = matchAndUpdateEntry(entry);
        if (matchResult.isMatch) {
          matchedCount++;
          // Log if we successfully matched a previously unmatched entry
          if (entry.isMatched === 'No match') {
            logEvent('no_match_converted', {
              row: i + 1,
              supplier: entry.supplier,
              newMatch: matchResult.reference
            });
          }
        }
        updateMatchedStatus(i + 1, matchResult);
      }

      logEvent('reconciliation_complete', {
        processedCount,
        matchedCount,
        successRate: (matchedCount / processedCount * 100).toFixed(2) + '%',
        mode: testMode ? 'test' : 'full',
        rowsProcessed: maxRows - 1
      });
    } catch (error) {
      logEvent('reconciliation_error', {
        error: error.message,
        stack: error.stack,
        systemState: {
          month: month,
          activeSheet: sourceSpreadsheet.getActiveSheet().getName(),
          mode: testMode ? 'test' : 'full'
        }
      });
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
    logEvent('sheet_check_start', {
      supplier: entry.supplier,
      sheet: sheetName,
      matchColumn
    });

    try {
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const matchColumnIndex = matchColumn.toUpperCase().charCodeAt(0) - 65;

      // Log just essential info
      logEvent('column_mapping', {
        sheet: sheetName,
        columnLetter: matchColumn,
        columnIndex: matchColumnIndex
      });

      if (matchColumnIndex < 0 || matchColumnIndex >= headers.length) {
        throw new Error(`Invalid column ${matchColumn} in ${sheetName}`);
      }

      // Create more concise potential matches
      const potentialMatches = data.slice(1)
        .map((row, index) => ({
          text: row[matchColumnIndex].toString(),
          rowIndex: index + 2,
          reference: `${sheetName}!${matchColumn}${index + 2}`
        }))
        .filter(match => match.text.trim() !== ''); // Only include non-empty matches

      // Log potential matches count
      logEvent('potential_matches', {
        sheet: sheetName,
        supplier: entry.supplier,
        matchCount: potentialMatches.length
      });

      const prompt = createMatchingQuery(entry.supplier, 
        potentialMatches.map(m => `${m.reference}: ${m.text}`));

      const claude = getClaudeService();
      const matchResult = claude.matchClient(entry.supplier, potentialMatches);

      // Log match result with essential info
      logEvent('match_attempt', {
        supplier: entry.supplier,
        sheet: sheetName,
        matched: matchResult.matched,
        confidence: matchResult.confidence,
        lineNumber: matchResult.lineNumber,
        matchedText: matchResult.lineNumber ? potentialMatches[matchResult.lineNumber - 2].text : null
      });

      if (matchResult.matched && matchResult.confidence > 0.5) {
        const matchedEntry = potentialMatches[matchResult.lineNumber - 2];
        return {
          isMatch: true,
          reference: matchedEntry.reference,
          rowIndex: matchedEntry.rowIndex,
          confidence: matchResult.confidence,
          explanation: matchResult.explanation
        };
      }

      return { isMatch: false };

    } catch (error) {
      logEvent('sheet_check_error', {
        supplier: entry.supplier,
        sheet: sheetName,
        error: error.message
      });
      throw error;
    }
  }

  // Return public methods with test mode option
  return {
    processReconciliation,
    matchAndUpdateEntry,
    checkSheetForMatch,
    processTestReconciliation: () => processReconciliation(true),
    processFullReconciliation: () => processReconciliation(false)
  };
} 