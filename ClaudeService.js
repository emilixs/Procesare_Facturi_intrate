/**
 * Simple Claude Integration Service for Procesare_Facturi
 * Handles client name matching using Claude AI
 */
function createClaudeService() {
  return {
    apiKey: PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY'),
    endpoint: 'https://api.anthropic.com/v1/messages',
    
    /**
     * Compare a client name from invoice with P&L client list
     * @param {string} invoiceClient - Client name from invoice
     * @param {Array<{name: string, line: number}>} plClients - Array of P&L clients with line numbers
     * @returns {Object} Match result with line number and confidence
     */
    matchClient: function(invoiceClient, plClients) {
      if (!this.apiKey) {
        throw new Error('Anthropic API key not found in Script Properties');
      }

      // Validate the plClients array
      if (!Array.isArray(plClients) || plClients.some(c => !c.name || !c.line)) {
        throw new Error('Invalid P&L clients data structure');
      }

      const prompt = `
Compare this invoice client name: "${invoiceClient}"
with these P&L client names:
${plClients.map(c => `Line ${c.line}: ${c.name}`).join('\n')}

Rules:
1. Ignore case, spaces, and special characters
2. Consider company type variations (SRL, S.R.L., LLC, etc.)
3. Look for the closest match

Reply only with a JSON object in this format:
{
  "matched": true/false,
  "lineNumber": number or null,
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
            lineNumber: null,
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
          model: "claude-3-sonnet-20240229",
          max_tokens: 4000,
          temperature: 0,
          system: "You are a helpful assistant that matches company names. You only respond with JSON.",
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

        // Log the actual LLM response text
        console.log("\n=== LLM RESPONSE ===");
        console.log(parsedResponse.content[0].text);
        console.log("=== END RESPONSE ===\n");

        return parsedResponse.content[0].text;
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