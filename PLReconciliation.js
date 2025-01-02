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
    
    // Process each row
    for (let i = 1; i < sourceData.length; i++) {
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
        timestamp: new Date().toISOString()
      });
      
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