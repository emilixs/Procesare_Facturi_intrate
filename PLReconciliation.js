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
          body { 
            font-family: Arial, sans-serif; 
            padding: 20px;
            background-color: #f8f9fa;
          }
          .container {
            background-color: white;
            padding: 25px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
          }
          .title {
            color: #1a73e8;
            margin-bottom: 20px;
            font-size: 18px;
            font-weight: 500;
            text-align: center;
          }
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
            animation: pulse 2s infinite;
          }
          .loading-bar {
            width: 100%;
            height: 36px;
            background: #f1f3f4;
            border-radius: 4px;
            margin: 15px 0;
            position: relative;
            overflow: hidden;
          }
          .loading-bar::after {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            height: 100%;
            width: 30%;
            background: linear-gradient(
              90deg,
              transparent,
              rgba(26, 115, 232, 0.2),
              transparent
            );
            animation: loading 1.5s infinite;
          }
          @keyframes loading {
            0% { transform: translateX(-100%); }
            100% { transform: translateX(400%); }
          }
          @keyframes pulse {
            0% { opacity: 1; }
            50% { opacity: 0.5; }
            100% { opacity: 1; }
          }
          .status { 
            color: #5f6368; 
            margin-top: 10px;
            font-size: 14px;
            text-align: center;
          }
          .company-name {
            color: #1a73e8;
            font-weight: 500;
          }
          .stats { 
            margin-top: 20px;
            padding: 12px;
            background-color: #f8f9fa;
            border-radius: 4px;
            font-size: 13px;
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 10px;
            text-align: center;
          }
          .stat-item {
            padding: 8px;
            background: white;
            border-radius: 4px;
            box-shadow: 0 1px 2px rgba(0,0,0,0.05);
          }
          .stat-label {
            color: #5f6368;
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
          }
          .stat-value {
            color: #1a73e8;
            font-weight: 500;
            font-size: 14px;
            margin-top: 4px;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="title">P&L Reconciliation</div>
          
          <div class="progress-bar">
            <div class="progress-fill" id="progressBar"></div>
          </div>
          
          <div class="loading-bar"></div>
          
          <div class="status">
            Processing: <span class="company-name" id="currentCompany">Initializing...</span>
          </div>
          
          <div class="stats">
            <div class="stat-item">
              <div class="stat-label">Processed</div>
              <div class="stat-value" id="processed">0</div>
            </div>
            <div class="stat-item">
              <div class="stat-label">Matches</div>
              <div class="stat-value" id="matches">0</div>
            </div>
            <div class="stat-item">
              <div class="stat-label">Time</div>
              <div class="stat-value" id="time">0s</div>
            </div>
          </div>
        </div>
        
        <script>
          let width = 0;
          
          function updateProgress(data) {
            if (data.progress) {
              document.getElementById('progressBar').style.width = data.progress + '%';
            }
            if (data.status) {
              document.getElementById('currentCompany').textContent = data.status;
            }
            if (data.stats) {
              document.getElementById('processed').textContent = data.stats.processed;
              document.getElementById('matches').textContent = data.stats.matches;
              document.getElementById('time').textContent = data.stats.time + 's';
            }
          }
          
          // Auto-increment progress bar for visual feedback
          setInterval(() => {
            if (width < 90) {
              width += 0.5;
              document.getElementById('progressBar').style.width = width + '%';
            }
          }, 500);
        </script>
      </body>
    </html>
  `);
  
  return html.evaluate()
    .setWidth(450)
    .setHeight(300)
    .setTitle('P&L Reconciliation Progress');
}

/**
 * Updates the progress information
 */
function updateProgressInfo(data) {
  PropertiesService.getScriptProperties().setProperty('progress_data', JSON.stringify(data));
}

/**
 * Process a batch of invoice entries
 */
function processBatch(batch, context) {
  const {
    claude,
    plClients,
    sourceSheet,
    revenuesSheet,
    targetColumnIndex,
    matchedColumnIndex,
    requestId,
    month
  } = context;
  
  console.log(`[${requestId}] Processing batch of ${batch.length} entries`);
  const batchStartTime = new Date();
  
  const results = [];
  
  // Process each entry in the batch
  for (const entry of batch) {
    const { row, invoiceClient, invoiceValue } = entry;
    
    // Check if already matched
    const matchedCell = sourceSheet.getRange(row + 1, matchedColumnIndex);
    const existingMatch = matchedCell.getValue();
    
    if (existingMatch) {
      console.log(`[${requestId}] Skipping already matched row ${row + 1}: ${invoiceClient}`);
      results.push({ skipped: true, row });
      continue;
    }
    
    // Pre-request logging
    console.log(`[${requestId}] Processing client match request:`, {
      invoiceClient,
      plClientsCount: plClients.length,
      timestamp: new Date().toISOString()
    });
    
    const matchStartTime = new Date();
    const match = claude.matchClient(invoiceClient, plClients);
    const matchProcessingTime = new Date() - matchStartTime;
    
    if (match.matched && match.confidence > 0.8) {
      const currentValue = revenuesSheet.getRange(match.lineNumber, targetColumnIndex).getValue() || 0;
      const newValue = currentValue + Number(invoiceValue);
      
      // Update P&L
      revenuesSheet.getRange(match.lineNumber, targetColumnIndex).setValue(newValue);
      
      // Update source sheet
      const cellRef = `${columnToLetter(targetColumnIndex)}${match.lineNumber}`;
      matchedCell.setValue(cellRef);
      matchedCell.setBackground(COLORS.MATCHED);
      
      results.push({
        matched: true,
        row,
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
    } else {
      matchedCell.setBackground(COLORS.UNMATCHED);
      results.push({ matched: false, row });
    }
  }
  
  console.log(`[${requestId}] Batch completed:`, {
    processed: results.length,
    matches: results.filter(r => r.matched).length,
    skipped: results.filter(r => r.skipped).length,
    processingTime: new Date() - batchStartTime
  });
  
  return results;
}

/**
 * Starts the P&L reconciliation process
 */
function startPLReconciliation(month, plUrl) {
  // Show progress dialog
  const progressDialog = showProgressDialog();
  const ui = SpreadsheetApp.getUi();
  ui.showModelessDialog(progressDialog, 'Processing...');
  
  const startTime = new Date();
  const requestId = Utilities.getUuid();
  
  try {
    // Initialize Claude service once
    const claude = getClaudeService();
    
    // 1. Get source (invoice) data
    const sourceSheet = SpreadsheetApp.getActiveSheet();
    const sourceData = sourceSheet.getDataRange().getValues();
    
    // Find required columns
    const sourceHeaders = sourceData[0];
    const clientColumnIndex = sourceHeaders.findIndex(header => 
      String(header).toLowerCase().includes('client') || 
      String(header).toLowerCase().includes('nume'));
    const valueColumnIndex = sourceHeaders.findIndex(header =>
      String(header).toLowerCase() === 'suma in eur');
    
    if (clientColumnIndex === -1) {
      throw new Error('Could not find client column (should contain "Client" or "Nume")');
    }
    if (valueColumnIndex === -1) {
      throw new Error('Could not find "Suma in EUR" column');
    }
    
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
    
    // Prepare entries for batch processing
    const entries = [];
    for (let i = 1; i < sourceData.length; i++) {
      const invoiceClient = sourceData[i][clientColumnIndex];
      const invoiceValue = sourceData[i][valueColumnIndex];
      
      if (!invoiceClient || !invoiceValue) continue;
      
      entries.push({
        row: i,
        invoiceClient,
        invoiceValue
      });
    }
    
    // Process in batches of 10
    const BATCH_SIZE = 10;
    const batches = [];
    for (let i = 0; i < entries.length; i += BATCH_SIZE) {
      batches.push(entries.slice(i, i + BATCH_SIZE));
    }
    
    const context = {
      claude,
      plClients,
      sourceSheet,
      revenuesSheet,
      targetColumnIndex,
      matchedColumnIndex,
      requestId,
      month
    };
    
    let processedCount = 0;
    let matchedCount = 0;
    let skippedCount = 0;
    const log = [];
    
    // Process each batch
    for (let batchIndex = 0; batchIndex < batches.length; batchIndex++) {
      const batch = batches[batchIndex];
      const batchStartTime = new Date();
      
      // Update progress
      const progress = Math.round((batchIndex * BATCH_SIZE / entries.length) * 100);
      const timeElapsed = Math.round((new Date() - startTime) / 1000);
      
      updateProgressInfo({
        progress,
        status: `Processing batch ${batchIndex + 1}/${batches.length}...`,
        stats: {
          processed: processedCount,
          matches: matchedCount,
          time: timeElapsed
        },
        currentBatch: {
          index: batchIndex + 1,
          total: batches.length,
          size: batch.length
        }
      });
      
      // Process batch
      const results = processBatch(batch, context);
      
      // Update counts and logs
      results.forEach(result => {
        processedCount++;
        if (result.skipped) {
          skippedCount++;
        } else if (result.matched) {
          matchedCount++;
          log.push(result);
        }
      });
      
      // Delay between batches
      if (batchIndex < batches.length - 1) {
        Utilities.sleep(500); // Longer delay between batches
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
      updatedEntries: matchedCount,
      totalProcessingTime,
      timestamp: new Date().toISOString()
    });
    
    // Show summary to user
    const message = `Reconciliation completed:\n` +
                   `- Processed ${sourceData.length - 1} invoice entries\n` +
                   `- Skipped ${skippedCount} already matched entries\n` +
                   `- Updated ${matchedCount} P&L entries\n` +
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
  const MATCHED_COLUMN_INDEX = 19; // Column S is the 19th column
  const HEADER_NAME = 'Matched P&L';
  
  // Get the current header in column S
  const currentHeader = sheet.getRange(1, MATCHED_COLUMN_INDEX).getValue();
  console.log('Current header in column S:', currentHeader);
  
  // If column S doesn't have "Matched P&L", set it
  if (currentHeader !== HEADER_NAME) {
    console.log('Setting header in column S to:', HEADER_NAME);
    sheet.getRange(1, MATCHED_COLUMN_INDEX).setValue(HEADER_NAME);
  }
  
  return MATCHED_COLUMN_INDEX; // Always return column S index
}

// Define color constants
const COLORS = {
  MATCHED: '#b7e1cd',    // Light green
  UNMATCHED: '#eaecef'   // Light gray
}; 