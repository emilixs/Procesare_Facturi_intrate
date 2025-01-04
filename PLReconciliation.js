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

Consider these matching scenarios:
1. Different legal entity formats:
   - "Company SRL" = "Company S.R.L." = "COMPANY S.R.L"
   - "Company SA" = "Company S.A." = "COMPANY SA"

2. Common abbreviations and variations:
   - "International" = "Intl" = "Int."
   - "Technology" = "Tech" = "Technologies"
   - "Solutions" = "Sol." = "Sols"

3. Brand names vs legal names:
   - A company might be listed by its brand name in one place and legal name in another
   - Example: "Google Cloud" might match "GOOGLE *CLOUD" or "Google LLC"

4. Special characters and formatting:
   - Ignore spaces, dots, dashes, underscores
   - Example: "Tech-Soft" = "TechSoft" = "Tech Soft"

5. Context clues:
   - Consider the business context
   - Look for unique identifying parts of names
   - Consider if the entry might be a subdivision or department of a larger company

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
      // Check Expenses sheet
      const expensesMatch = checkSheetForMatch(entry, expensesSheet, 'C', 'Expenses');
      if (expensesMatch.isMatch) {
        logEvent('match_found', {
          sheet: 'Expenses',
          supplier: entry.supplier,
          matchDetails: expensesMatch,
          llmPromptUsed: expensesMatch.promptUsed, // Log the prompt that led to this match
          llmResponse: expensesMatch.rawResponse,   // Log the raw response that led to this match
          timestamp_match: new Date().toISOString()
        });
        updateAmount(expensesMatch, entry.amount);
        return expensesMatch;
      }

      // Check Staffing sheet
      const staffingMatch = checkSheetForMatch(entry, staffingSheet, 'D', 'Staffing');
      if (staffingMatch.isMatch) {
        logEvent('match_found', {
          sheet: 'Staffing',
          supplier: entry.supplier,
          matchDetails: staffingMatch,
          llmPromptUsed: staffingMatch.promptUsed,  // Log the prompt that led to this match
          llmResponse: staffingMatch.rawResponse,    // Log the raw response that led to this match
          timestamp_match: new Date().toISOString()
        });
        updateAmount(staffingMatch, entry.amount);
        return staffingMatch;
      }

      logEvent('no_match_found', {
        supplier: entry.supplier,
        lastPromptTried: entry.lastPromptTried,     // Log the last prompt that was tried
        lastResponseReceived: entry.lastResponse,    // Log the last response that led to no match
        timestamp_nomatch: new Date().toISOString()
      });
      return { isMatch: false };
    } catch (error) {
      logEvent('entry_processing_error', {
        supplier: entry.supplier,
        error: error.message,
        stack: error.stack,
        lastPromptTried: entry.lastPromptTried,     // Log the prompt that caused the error
        lastResponseReceived: entry.lastResponse,    // Log the response that caused the error
        timestamp_error: new Date().toISOString()
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
      // Get all data from the sheet
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      
      // Convert column letter to index (0-based)
      const matchColumnIndex = matchColumn.toUpperCase().charCodeAt(0) - 65; // 'A'=0, 'B'=1, 'C'=2, etc.
      
      logEvent('column_mapping', {
        columnLetter: matchColumn,
        columnIndex: matchColumnIndex,
        headerFound: headers[matchColumnIndex]
      });

      if (matchColumnIndex < 0 || matchColumnIndex >= headers.length) {
        throw new Error(`Invalid column ${matchColumn} in ${sheetName}`);
      }

      // Create array of potential matches for LLM
      const potentialMatches = data.slice(1) // Skip header row
        .map((row, index) => ({
          text: row[matchColumnIndex].toString(),
          rowIndex: index + 2, // +2 because we skipped header and array is 0-based
          reference: `${sheetName}!${matchColumn}${index + 2}`
        }));

      // Create LLM query
      const prompt = createMatchingQuery(entry.supplier, 
        potentialMatches.map(m => `${m.reference}: ${m.text}`));

      // Get Claude service
      const claude = getClaudeService();
      
      // Get match decision from LLM
      const matchResult = claude.matchClient(entry.supplier, potentialMatches);

      // Log the match attempt
      logEvent('match_attempt', {
        supplier: entry.supplier,
        sheet: sheetName,
        matchResult,
        promptUsed: prompt
      });

      if (matchResult.matched && matchResult.confidence > 0.5) {
        const matchedEntry = potentialMatches[matchResult.lineNumber - 2];
        logEvent('match_found', {
          supplier: entry.supplier,
          matchedWith: matchedEntry.text,
          confidence: matchResult.confidence,
          explanation: matchResult.explanation,
          matchingPatterns: {
            exactMatch: entry.supplier.toLowerCase() === matchedEntry.text.toLowerCase(),
            partialMatch: entry.supplier.toLowerCase().includes(matchedEntry.text.toLowerCase()) || 
                         matchedEntry.text.toLowerCase().includes(entry.supplier.toLowerCase()),
            withoutLegalEntity: entry.supplier.toLowerCase().replace(/s\.?r\.?l\.?|s\.?a\.?/gi, '').trim() ===
                               matchedEntry.text.toLowerCase().replace(/s\.?r\.?l\.?|s\.?a\.?/gi, '').trim()
          }
        });
        return {
          isMatch: true,
          reference: matchedEntry.reference,
          rowIndex: matchedEntry.rowIndex,
          confidence: matchResult.confidence,
          explanation: matchResult.explanation,
          promptUsed: prompt,
          rawResponse: matchResult
        };
      }

      return {
        isMatch: false,
        promptUsed: prompt,
        rawResponse: matchResult
      };

    } catch (error) {
      logEvent('sheet_check_error', {
        supplier: entry.supplier,
        sheet: sheetName,
        error: error.message,
        stack: error.stack
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