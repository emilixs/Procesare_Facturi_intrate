/**
 * Creates the custom menu in Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Invoice Tools')
    .addItem('To Numbers', 'processInvoiceData')
    .addItem('P&L Reconciliation', 'showPLReconciliationDialog')
    .addToUi();
}

/**
 * Shows the P&L reconciliation dialog
 */
function showPLReconciliationDialog() {
  const html = HtmlService.createHtmlTemplate(`
    <style>
      .form-group { margin-bottom: 15px; }
      label { display: block; margin-bottom: 5px; }
      input { width: 100%; padding: 5px; }
      .error { color: red; display: none; }
      button { padding: 8px 15px; }
    </style>
    
    <div class="form-group">
      <label for="month">Reference Month:</label>
      <input type="text" id="month" name="month" required placeholder="e.g., January">
      <div id="monthError" class="error">Please enter a valid month</div>
    </div>
    
    <div class="form-group">
      <label for="plUrl">P&L File URL:</label>
      <input type="text" id="plUrl" name="plUrl" required placeholder="Paste Google Sheets URL here">
      <div id="urlError" class="error">Please enter a valid Google Sheets URL</div>
    </div>
    
    <div class="form-group">
      <button onclick="submitForm()">Start Reconciliation</button>
    </div>
    
    <script>
      function validateMonth(month) {
        const months = ['january', 'february', 'march', 'april', 'may', 'june', 
                       'july', 'august', 'september', 'october', 'november', 'december'];
        return months.includes(month.toLowerCase());
      }
      
      function validateUrl(url) {
        return url.includes('docs.google.com/spreadsheets');
      }
      
      function submitForm() {
        const month = document.getElementById('month').value;
        const plUrl = document.getElementById('plUrl').value;
        
        document.getElementById('monthError').style.display = 'none';
        document.getElementById('urlError').style.display = 'none';
        
        let isValid = true;
        
        if (!validateMonth(month)) {
          document.getElementById('monthError').style.display = 'block';
          isValid = false;
        }
        
        if (!validateUrl(plUrl)) {
          document.getElementById('urlError').style.display = 'block';
          isValid = false;
        }
        
        if (isValid) {
          google.script.run
            .withSuccessHandler(onSuccess)
            .withFailureHandler(onFailure)
            .startPLReconciliation(month, plUrl);
        }
      }
      
      function onSuccess(result) {
        google.script.host.close();
      }
      
      function onFailure(error) {
        alert('Error: ' + error.message);
      }
    </script>
  `);
  
  const userInterface = html.evaluate()
    .setWidth(400)
    .setHeight(300)
    .setTitle('P&L Reconciliation');
    
  SpreadsheetApp.getUi().showModalDialog(userInterface, 'P&L Reconciliation');
} 