/**
 * Creates the custom menu in Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Invoice Processing')
    .addItem('Convert Numeric Columns', 'processInvoiceData')
    .addItem('P&L Reconciliation', 'showPLReconciliationDialog')
    .addToUi();
}

/**
 * Shows the P&L reconciliation dialog
 */
function showPLReconciliationDialog() {
  const html = HtmlService.createTemplate(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            padding: 20px;
            background-color: #f5f5f5;
            margin: 0;
          }
          
          .container {
            background-color: white;
            padding: 25px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
          }
          
          .form-group {
            margin-bottom: 20px;
          }
          
          .title {
            color: #1a73e8;
            margin-bottom: 25px;
            font-size: 20px;
            font-weight: 500;
            text-align: center;
          }
          
          label {
            display: block;
            margin-bottom: 8px;
            color: #202124;
            font-weight: 500;
            font-size: 14px;
          }
          
          input {
            width: 100%;
            padding: 8px 12px;
            border: 1px solid #dadce0;
            border-radius: 4px;
            font-size: 14px;
            box-sizing: border-box;
            transition: border-color 0.2s;
          }
          
          input:focus {
            outline: none;
            border-color: #1a73e8;
          }
          
          .error {
            color: #d93025;
            font-size: 12px;
            margin-top: 4px;
            display: none;
          }
          
          button {
            background-color: #1a73e8;
            color: white;
            padding: 10px 24px;
            border: none;
            border-radius: 4px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            width: 100%;
            transition: background-color 0.2s;
          }
          
          button:hover {
            background-color: #1557b0;
          }
          
          button:disabled {
            background-color: #dadce0;
            cursor: not-allowed;
          }
          
          .loading {
            display: none;
            text-align: center;
            margin-top: 10px;
            color: #5f6368;
            font-size: 13px;
          }
          
          .spinner {
            display: inline-block;
            width: 16px;
            height: 16px;
            border: 2px solid #dadce0;
            border-top: 2px solid #1a73e8;
            border-radius: 50%;
            animation: spin 1s linear infinite;
            margin-right: 8px;
            vertical-align: middle;
          }
          
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          
          .info-text {
            color: #5f6368;
            font-size: 12px;
            margin-top: 4px;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="title">P&L Reconciliation</div>
          
          <div class="form-group">
            <label for="month">Reference Month</label>
            <input type="text" id="month" name="month" required 
                   placeholder="e.g., January" autocomplete="off">
            <div id="monthError" class="error">Please enter a valid month</div>
            <div class="info-text">Enter the full month name (e.g., January, February)</div>
          </div>
          
          <div class="form-group">
            <label for="plUrl">P&L File URL</label>
            <input type="text" id="plUrl" name="plUrl" required 
                   placeholder="https://docs.google.com/spreadsheets/d/..." autocomplete="off">
            <div id="urlError" class="error">Please enter a valid Google Sheets URL</div>
            <div class="info-text">Paste the full URL of your P&L Google Sheet</div>
          </div>
          
          <button onclick="submitForm()" id="submitBtn">Start Reconciliation</button>
          
          <div id="loading" class="loading">
            <div class="spinner"></div>
            Processing reconciliation...
          </div>
        </div>
        
        <script>
          function submitForm() {
            const month = document.getElementById('month').value;
            const plUrl = document.getElementById('plUrl').value;
            const submitBtn = document.getElementById('submitBtn');
            const loading = document.getElementById('loading');
            
            if (validateMonth(month) && validateUrl(plUrl)) {
              // Show loading state
              submitBtn.disabled = true;
              loading.style.display = 'block';
              
              // Add logging to check if this is being called
              console.log('Starting reconciliation with:', {month, plUrl});
              
              google.script.run
                .withSuccessHandler(onSuccess)
                .withFailureHandler(onFailure)
                .startPLReconciliation(month.trim(), plUrl.trim());
            }
          }
          
          function onSuccess(result) {
            console.log('Reconciliation completed:', result);
            google.script.host.close();
          }
          
          function onFailure(error) {
            console.error('Reconciliation failed:', error);
            const submitBtn = document.getElementById('submitBtn');
            const loading = document.getElementById('loading');
            
            submitBtn.disabled = false;
            loading.style.display = 'none';
            
            alert('Error: ' + (error.message || 'An unexpected error occurred'));
          }
          
          function validateMonth(month) {
            const months = ['january', 'february', 'march', 'april', 'may', 'june', 
                          'july', 'august', 'september', 'october', 'november', 'december'];
            return months.includes(month.toLowerCase().trim());
          }
          
          function validateUrl(url) {
            return url.trim().includes('docs.google.com/spreadsheets');
          }
        </script>
      </body>
    </html>
  `);
  
  const userInterface = html.evaluate()
    .setWidth(450)
    .setHeight(500)
    .setTitle('P&L Reconciliation');
    
  SpreadsheetApp.getUi().showModalDialog(userInterface, 'P&L Reconciliation');
}

/**
 * Entry point for P&L reconciliation
 */
function startPLReconciliation(month, plUrl) {
  // Add initial logging
  console.log('Starting P&L reconciliation:', {month, plUrl});
  
  try {
    const service = createPLReconciliationService(plUrl, month);
    return service.processReconciliation(true); // true for test mode
  } catch (error) {
    console.error('Error in startPLReconciliation:', error);
    throw error;
  }
} 