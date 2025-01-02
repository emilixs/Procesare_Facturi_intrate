/**
 * Shows a progress dialog
 */
function showProgressDialog() {
  const html = HtmlService.createTemplate(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; }
          .progress-bar {
            width: 100%;
            height: 4px;
            background-color: #e8eaed;
            margin: 20px 0;
            border-radius: 2px;
            overflow: hidden;
          }
          .progress-fill {
            height: 100%;
            background-color: #1a73e8;
            width: 0%;
            transition: width 0.3s ease;
          }
          .status { color: #5f6368; margin-top: 10px; }
          .stats { margin-top: 20px; }
        </style>
      </head>
      <body>
        <div class="progress-bar">
          <div class="progress-fill" id="progressBar"></div>
        </div>
        <div class="status" id="status">Initializing...</div>
        <div class="stats" id="stats"></div>
        
        <script>
          window.onmessage = function(e) {
            const data = e.data;
            if (data.progress) {
              document.getElementById('progressBar').style.width = data.progress + '%';
            }
            if (data.status) {
              document.getElementById('status').textContent = data.status;
            }
            if (data.stats) {
              document.getElementById('stats').textContent = 
                'Processed: ' + data.processed + 
                ' | Matches: ' + data.matches +
                ' | Time: ' + data.time + 's';
            }
          };
        </script>
      </body>
    </html>
  `);
  
  const userInterface = html.evaluate()
    .setWidth(400)
    .setHeight(150)
    .setTitle('P&L Reconciliation Progress');
    
  return userInterface;
}

/**
 * Starts the P&L reconciliation process
 * @param {string} month Reference month (e.g., "January")
 * @param {string} plUrl URL of the P&L spreadsheet
 */
function startPLReconciliation(month, plUrl) {
  // Show progress dialog
  const progressDialog = showProgressDialog();
  const ui = SpreadsheetApp.getUi();
  const htmlOutput = ui.showModelessDialog(progressDialog, 'Processing...');
  
  const startTime = new Date();
  const requestId = Utilities.getUuid();
  
  try {
    // Initialize Claude service once
    const claude = getClaudeService();
    
    // 1. Get source (invoice) data
    const sourceSheet = SpreadsheetApp.getActiveSheet();
    const sourceData = sourceSheet.getDataRange().getValues();
    
    // 2. Ensure Matched P&L column exists
    const matchedColumnIndex = ensureMatchedColumn(sourceSheet);
    
    // 3. Open P&L spreadsheet
    const plFileId = plUrl.match(/[-\w]{25,}/)[0];
    const plSpreadsheet = SpreadsheetApp.openById(plFileId);
    const revenuesSheet = plSpreadsheet.getSheetByName('revenues');
    
    if (!revenuesSheet) {
      throw new Error('Could not find "revenues" sheet in P&L file');
    }
    
    // 4. Find target column in P&L (e.g., "January real")
    const plHeaders = revenuesSheet.getRange(2, 1, 1, revenuesSheet.getLastColumn()).getValues()[0];
    const targetColumnName = `${month.toLowerCase()} real`;
    const targetColumnIndex = plHeaders.findIndex(header => 
      header.toString().toLowerCase() === targetColumnName) + 1;
    
    if (targetColumnIndex === 0) {
      throw new Error(`Could not find column "${month} real" in P&L`);
    }
    
    // 5. Get P&L client list from column D
    const plClients = revenuesSheet.getRange('D:D')
      .getValues()
      .map((row, index) => ({
        name: row[0],
        line: index + 1
      }))
      .filter(client => client.name); // Remove empty rows
    
    // 6. Create or get log sheet early with updated headers
    let logSheet = sourceSheet.getParent().getSheetByName('Reconciliation Log');
    if (!logSheet) {
      logSheet = sourceSheet.getParent().insertSheet('Reconciliation Log');
      logSheet.appendRow([
        'Timestamp',
        'Request ID',
        'Operation',
        'Invoice Client',
        'P&L Client',
        'Value Added',
        'Previous Value',
        'New Value',
        'Match Confidence',
        'Month',
        'Processing Time (ms)'
      ]);
    }
    
    // 7. Process each invoice row
    const log = [];
    let updatedCount = 0;
    
    // Find source data columns
    const sourceHeaders = sourceData[0];
    const clientColumnIndex = sourceHeaders.findIndex(header => 
      header.toString().toLowerCase().includes('client') || 
      header.toString().toLowerCase().includes('nume'));
    const valueColumnIndex = sourceHeaders.indexOf('Suma in EUR');
    
    if (clientColumnIndex === -1 || valueColumnIndex === -1) {
      throw new Error('Required columns not found in invoice sheet. Need "Client"/"Nume Client" and "Suma in EUR"');
    }
    
    // Process each row
    let skippedCount = 0;
    for (let i = 1; i < sourceData.length; i++) {
      const progress = Math.round((i / (sourceData.length - 1)) * 100);
      const timeElapsed = Math.round((new Date() - startTime) / 1000);
      
      // Update progress dialog
      htmlOutput.getDialogWindow().postMessage({
        progress: progress,
        status: `Processing ${sourceData[i][clientColumnIndex]}...`,
        stats: {
          processed: i,
          matches: updatedCount,
          time: timeElapsed
        }
      }, '*');
      
      const invoiceClient = sourceData[i][clientColumnIndex];
      const invoiceValue = sourceData[i][valueColumnIndex];
      
      if (!invoiceClient || !invoiceValue) continue;
      
      // Check if already matched
      const matchedCell = sourceSheet.getRange(i + 1, matchedColumnIndex);
      const existingMatch = matchedCell.getValue();
      
      if (existingMatch) {
        console.log(`[${requestId}] Skipping already matched row ${i + 1}: ${invoiceClient}`);
        skippedCount++;
        continue;
      }
      
      // Pre-request logging
      console.log(`[${requestId}] Processing client match request:`, {
        invoiceClient,
        plClientsCount: plClients.length,
        timestamp: new Date().toISOString()
      });
      
      const matchStartTime = new Date();
      // Find matching client in P&L
      const match = claude.matchClient(invoiceClient, plClients);
      const matchProcessingTime = new Date() - matchStartTime;
      
      // Post-response logging
      console.log(`[${requestId}] Client match response:`, {
        matched: match.matched,
        confidence: match.confidence,
        processingTime: matchProcessingTime,
        timestamp: new Date().toISOString(),
        requestStatus: match.matched ? 'success' : 'no_match'
      });
      
      if (match.matched && match.confidence > 0.8) {
        const currentValue = revenuesSheet.getRange(match.lineNumber, targetColumnIndex).getValue() || 0;
        const newValue = currentValue + Number(invoiceValue);
        
        // Update P&L
        revenuesSheet.getRange(match.lineNumber, targetColumnIndex).setValue(newValue);
        
        // Update source sheet with match reference
        const cellRef = `${columnToLetter(targetColumnIndex)}${match.lineNumber}`;
        matchedCell.setValue(cellRef);
        matchedCell.setBackground(COLORS.MATCHED);
        
        // Create structured log entry
        log.push({
          timestamp: new Date(),
          requestId,
          invoiceClient,
          plClient: plClients[match.lineNumber - 1].name,
          value: invoiceValue,
          oldValue: currentValue,
          newValue: newValue,
          confidence: match.confidence,
          month,
          processingTime: matchProcessingTime,
          plReference: cellRef
        });
        
        updatedCount++;
      } else {
        // Mark as unmatched
        matchedCell.setBackground(COLORS.UNMATCHED);
      }
      
      // Add a small delay between requests to avoid rate limiting
      if (i < sourceData.length - 1) {
        Utilities.sleep(200);
      }
    }
    
    // Add log entries
    log.forEach(entry => {
      logSheet.appendRow(createLogEntry(entry));
    });
    
    const totalProcessingTime = new Date() - startTime;
    
    // Final execution log
    console.log(`[${requestId}] Reconciliation completed:`, {
      processedEntries: sourceData.length - 1,
      updatedEntries: updatedCount,
      totalProcessingTime,
      timestamp: new Date().toISOString()
    });
    
    // Show summary to user
    const message = `Reconciliation completed:\n` +
                   `- Processed ${sourceData.length - 1} invoice entries\n` +
                   `- Skipped ${skippedCount} already matched entries\n` +
                   `- Updated ${updatedCount} P&L entries\n` +
                   `- Total processing time: ${totalProcessingTime}ms\n` +
                   `- Request ID: ${requestId}\n` +
                   `- Check 'Reconciliation Log' sheet for details`;
                   
    SpreadsheetApp.getUi().alert('Success', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    // Error logging
    console.error(`[${requestId}] Reconciliation error:`, {
      error: error.message,
      stack: error.stack,
      timestamp: new Date().toISOString(),
      processingTime: new Date() - startTime
    });
    
    SpreadsheetApp.getUi().alert('Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Creates a structured log entry
 * @param {Object} params Log parameters
 * @returns {Array} Formatted log row
 */
function createLogEntry({
  timestamp,
  invoiceClient,
  plClient,
  value,
  oldValue,
  newValue,
  confidence,
  month,
  requestId,
  processingTime
}) {
  return [
    timestamp,
    requestId,
    'Client Match',
    invoiceClient,
    plClient,
    value,
    oldValue,
    newValue,
    confidence,
    month,
    processingTime
  ];
}

/**
 * Updates the progress dialog
 * @param {Object} data Progress data
 */
function updateProgress(data) {
  return data; // This function exists just to be called from the client side
}

/**
 * Convert column number to letter reference (e.g., 1 -> A, 27 -> AA)
 * @param {number} column Column number (1-based)
 * @returns {string} Column letter reference
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
 * Ensures the Matched P&L column exists and returns its index
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Source sheet
 * @returns {number} Column index (1-based)
 */
function ensureMatchedColumn(sheet) {
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const matchedColumnIndex = headers.indexOf('Matched P&L');
  
  if (matchedColumnIndex === -1) {
    // Add new column if it doesn't exist
    const newColumnIndex = lastColumn + 1;
    sheet.getRange(1, newColumnIndex).setValue('Matched P&L');
    return newColumnIndex;
  }
  
  return matchedColumnIndex + 1; // Convert to 1-based index
}

// Define color constants
const COLORS = {
  MATCHED: '#b7e1cd',    // Light green
  UNMATCHED: '#eaecef'   // Light gray
}; 