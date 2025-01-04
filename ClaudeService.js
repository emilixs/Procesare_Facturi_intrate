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
        console.error('Invalid plClients format:', JSON.stringify(plClients));
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
        // Log the plClients data for debugging
        console.log('P&L Clients:', JSON.stringify(plClients));
        
        const response = this.callClaude(prompt);
        return JSON.parse(response);
      } catch (error) {
        console.error('Client matching error:', {
          error: error.message,
          stack: error.stack,
          invoiceClient,
          plClientsCount: plClients.length,
          timestamp: new Date().toISOString(),
          details: error.details || 'No additional details'
        });
        
        // Implement retry logic
        try {
          console.log('Retrying API call...');
          Utilities.sleep(1000); // Wait 1 second before retry
          const retryResponse = this.callClaude(prompt);
          return JSON.parse(retryResponse);
        } catch (retryError) {
          console.error('Retry failed:', {
            error: retryError.message,
            invoiceClient,
            timestamp: new Date().toISOString()
          });
          
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
      // Log the prompt being sent
      console.log('Prompt sent to Claude:', prompt);

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
        
        // Log the raw response
        console.log('Claude Response:', JSON.stringify(responseBody));
        
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
        return parsedResponse.content[0].text;
      } catch (parseError) {
        console.error('Response parsing error:', {
          error: parseError.message,
          responseBody: responseBody
        });
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