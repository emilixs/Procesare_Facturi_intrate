/**
 * Simple Claude Integration Service for Procesare_Facturi
 * Handles client name matching using Claude AI
 */
function createClaudeService() {
  // Initialize or get the log sheet
  /*
  function getLogSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('LLM_Logs');
    if (!sheet) {
      sheet = ss.insertSheet('LLM_Logs');
      sheet.getRange('A1:C1').setValues([['Timestamp', 'Request', 'Response']]);
      sheet.setFrozenRows(1);
    }
    return sheet;
  }

  // Log to sheet function
  function logToSheet(request, response) {
    const logSheet = getLogSheet();
    const timestamp = new Date().toISOString();
    logSheet.appendRow([timestamp, request, response]);
  }
  */

  return {
    apiKey: PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY'),
    endpoint: 'https://api.anthropic.com/v1/messages',
    
    /**
     * Compare a client name from invoice with P&L client list
     * @param {string} invoiceClient - Client name from invoice
     * @param {Array<{name: string, reference: string}>} plClients - Array of P&L clients with cell references
     * @returns {Object} Match result with cell reference and confidence
     */
    matchClient: function(invoiceClient, plClients) {
      if (!this.apiKey) {
        throw new Error('Anthropic API key not found in Script Properties');
      }

      // Validate the plClients array
      if (!Array.isArray(plClients) || plClients.some(c => !c.name || !c.reference)) {
        throw new Error('Invalid P&L clients data structure');
      }

      const prompt = `
Compare this invoice client name: "${invoiceClient}"
with these P&L client names:
${plClients.map(c => `${c.reference}: ${c.name}`).join('\n')}

Rules:
1. Ignore case, spaces, and special characters
2. Consider company type variations (SRL, S.R.L., LLC, etc.)
3. Look for the closest match

Reply only with a JSON object in this format:
{
  "matched": true/false,
  "reference": string (the exact cell reference provided, e.g. "Expenses!C128"),
  "confidence": 0.0-1.0
}`;

      try {
        const response = this.callClaude(prompt);
        return JSON.parse(response);
      } catch (error) {
        // Implement retry logic
        try {
          Utilities.sleep(1000); // Wait 1 second before retry
          const retryResponse = this.callClaude(prompt);
          return JSON.parse(retryResponse);
        } catch (retryError) {
          return {
            matched: false,
            reference: null,
            confidence: 0
          };
        }
      }
    },

    /**
     * Make API call to Claude
     * @private
     */
    callClaude: function(prompt) {
      // Log the request
      console.log("\n=== LLM REQUEST ===");
      console.log(prompt);
      console.log("=== END REQUEST ===\n");

      const options = {
        method: 'POST',
        headers: {
          'x-api-key': this.apiKey,
          'Content-Type': 'application/json',
          'anthropic-version': '2023-06-01'
        },
        muteHttpExceptions: true,
        payload: JSON.stringify({
          model: "claude-3-5-sonnet-latest",
          max_tokens: 4000,
          temperature: 0,
          system: "You are a helpful assistant that matches company names. You only respond with JSON. When you find a match, return the exact cell reference that was provided in the input.",
          messages: [
            {
              role: "user",
              content: prompt
            }
          ]
        })
      };

      try {
        const response = UrlFetchApp.fetch(this.endpoint, options);
        const responseCode = response.getResponseCode();
        const responseBody = response.getContentText();
        
        if (responseCode !== 200) {
          const error = new Error('Claude API request failed');
          error.details = {
            statusCode: responseCode,
            response: responseBody,
            headers: response.getHeaders()
          };
          throw error;
        }

        const parsedResponse = JSON.parse(responseBody);
        if (!parsedResponse.content || !parsedResponse.content[0] || !parsedResponse.content[0].text) {
          throw new Error('Invalid response structure');
        }

        const llmResponse = parsedResponse.content[0].text;
        
        // Log the actual LLM response text
        console.log("\n=== LLM RESPONSE ===");
        console.log(llmResponse);
        console.log("=== END RESPONSE ===\n");

        // Comment out sheet logging
        // logToSheet(prompt, llmResponse);

        return llmResponse;
      } catch (parseError) {
        throw parseError;
      }
    }
  };
}

/**
 * Get instance of Claude Service
 */
function getClaudeService() {
  return createClaudeService();
} 