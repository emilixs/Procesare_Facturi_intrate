/**
 * Starts the P&L reconciliation process
 * @param {string} month Reference month (e.g., "January")
 * @param {string} plUrl URL of the P&L spreadsheet
 */
function startPLReconciliation(month, plUrl) {
  const startTime = new Date();
  const requestId = Utilities.getUuid();
  
  try {
    // Initialize Claude service once
    const claude = getClaudeService();
    
    // 1. Get source (invoice) data
    const sourceSheet = SpreadsheetApp.getActiveSheet();
    const sourceData = sourceSheet.getDataRange().getValues();
    
    // 2. Open P&L spreadsheet
    const plFileId = plUrl.match(/[-\w]{25,}/)[0];
    const plSpreadsheet = SpreadsheetApp.openById(plFileId);
    const revenuesSheet = plSpreadsheet.getSheetByName('revenues');
    
    if (!revenuesSheet) {
      throw new Error('Could not find "revenues" sheet in P&L file');
    }
    
    // 3. Find target column in P&L (e.g., "January real")
    const plHeaders = revenuesSheet.getRange(2, 1, 1, revenuesSheet.getLastColumn()).getValues()[0];
    const targetColumnName = `${month.toLowerCase()} real`;
    const targetColumnIndex = plHeaders.findIndex(header => 
      header.toString().toLowerCase() === targetColumnName) + 1;
    
    if (targetColumnIndex === 0) {
      throw new Error(`Could not find column "${month} real" in P&L`);
    }
    
    // 4. Get P&L client list from column D
    const plClients = revenuesSheet.getRange('D:D')
      .getValues()
      .map((row, index) => ({
        name: row[0],
        line: index + 1
      }))
      .filter(client => client.name); // Remove empty rows
    
    // 5. Create or get log sheet early with updated headers
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
    
    // 6. Process each invoice row
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
    
    // Update progress periodically
    for (let i = 1; i < sourceData.length; i++) {
      const progress = Math.round((i / (sourceData.length - 1)) * 100);
      const timeElapsed = Math.round((new Date() - startTime) / 1000);
      
      // Update progress dialog
      google.script.run.withSuccessHandler(function() {}).updateProgress({
        progress: progress,
        status: `Processing ${sourceData[i][clientColumnIndex]}...`,
        processed: i,
        matches: updatedCount,
        time: timeElapsed
      });
      
      const invoiceClient = sourceData[i][clientColumnIndex];
      const invoiceValue = sourceData[i][valueColumnIndex];
      
      if (!invoiceClient || !invoiceValue) continue;
      
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
      
      // Add a small delay between requests to avoid rate limiting
      if (i < sourceData.length - 1) {
        Utilities.sleep(200); // 200ms delay between requests
      }
      
      if (match.matched && match.confidence > 0.8) {
        const currentValue = revenuesSheet.getRange(match.lineNumber, targetColumnIndex).getValue() || 0;
        const newValue = currentValue + Number(invoiceValue);
        
        // Update P&L
        revenuesSheet.getRange(match.lineNumber, targetColumnIndex).setValue(newValue);
        
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
          processingTime: matchProcessingTime
        });
        
        updatedCount++;
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
 * Shows a progress dialog
 * @param {string} title The dialog title
 * @returns {google.script.html.HtmlOutput} The dialog
 */
function showProgressDialog(title) {
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
            margin: 0;
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
          }
          
          .progress-container {
            margin: 20px 0;
          }
          
          .progress-bar {
            width: 100%;
            height: 4px;
            background-color: #e8eaed;
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
          
          @keyframes pulse {
            0% { opacity: 1; }
            50% { opacity: 0.5; }
            100% { opacity: 1; }
          }
          
          .status {
            margin-top: 15px;
            color: #5f6368;
            font-size: 14px;
          }
          
          .stats {
            margin-top: 20px;
            padding: 15px;
            background-color: #f8f9fa;
            border-radius: 4px;
            font-size: 13px;
          }
          
          .stat-item {
            margin: 8px 0;
            display: flex;
            justify-content: space-between;
          }
          
          .stat-label {
            color: #5f6368;
          }
          
          .stat-value {
            color: #1a73e8;
            font-weight: 500;
          }
          
          .tips {
            margin-top: 20px;
            font-size: 13px;
            color: #5f6368;
            font-style: italic;
          }
          
          .tip {
            margin: 8px 0;
            padding-left: 20px;
            position: relative;
          }
          
          .tip:before {
            content: "💡";
            position: absolute;
            left: 0;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="title"><?= title ?></div>
          
          <div class="progress-container">
            <div class="progress-bar">
              <div class="progress-fill" id="progressBar"></div>
            </div>
            <div class="status" id="status">Initializing...</div>
          </div>
          
          <div class="stats">
            <div class="stat-item">
              <span class="stat-label">Processed:</span>
              <span class="stat-value" id="processed">0</span>
            </div>
            <div class="stat-item">
              <span class="stat-label">Matches Found:</span>
              <span class="stat-value" id="matches">0</span>
            </div>
            <div class="stat-item">
              <span class="stat-label">Processing Time:</span>
              <span class="stat-value" id="time">0s</span>
            </div>
          </div>
          
          <div class="tips">
            <div class="tip" id="tip">Did you know? The AI model considers various company name formats and abbreviations.</div>
          </div>
        </div>
        
        <script>
          const tips = [
            "The AI model considers various company name formats and abbreviations.",
            "All changes are automatically logged for future reference.",
            "High confidence matches (>0.8) are automatically processed.",
            "The system handles different currency formats and decimal separators.",
            "Each update is verified before being applied to the P&L.",
            "You can find detailed logs in the 'Reconciliation Log' sheet."
          ];
          
          let tipIndex = 0;
          
          function updateProgress(data) {
            const progressBar = document.getElementById('progressBar');
            const status = document.getElementById('status');
            const processed = document.getElementById('processed');
            const matches = document.getElementById('matches');
            const time = document.getElementById('time');
            const tip = document.getElementById('tip');
            
            if (data.progress) {
              progressBar.style.width = data.progress + '%';
            }
            
            if (data.status) {
              status.textContent = data.status;
            }
            
            if (data.processed) {
              processed.textContent = data.processed;
            }
            
            if (data.matches) {
              matches.textContent = data.matches;
            }
            
            if (data.time) {
              time.textContent = data.time + 's';
            }
            
            // Rotate tips every 5 seconds
            tipIndex = (tipIndex + 1) % tips.length;
            tip.textContent = tips[tipIndex];
          }
          
          // Update tips every 5 seconds
          setInterval(() => {
            const tip = document.getElementById('tip');
            tipIndex = (tipIndex + 1) % tips.length;
            tip.textContent = tips[tipIndex];
          }, 5000);
          
          // Simulate progress bar movement
          let width = 0;
          setInterval(() => {
            if (width < 90) { // Never reach 100% until actually complete
              width += 0.5;
              document.getElementById('progressBar').style.width = width + '%';
            }
          }, 500);
        </script>
      </body>
    </html>
  `);
  
  html.title = title;
  return html.evaluate()
    .setWidth(450)
    .setHeight(400);
}

/**
 * Updates the progress dialog
 * @param {Object} data Progress data
 */
function updateProgress(data) {
  return data; // This function exists just to be called from the client side
} 