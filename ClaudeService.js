/**
 * Simple Claude Integration Service for Procesare_Facturi
 * Handles client name matching using Claude AI
 */
class ClaudeService {
  constructor() {
    this.apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
    if (!this.apiKey) {
      throw new Error('Claude API key not found in Script Properties');
    }
    
    this.endpoint = 'https://api.anthropic.com/v1/complete';
  }

  /**
   * Compare a client name from invoice with P&L client list
   * @param {string} invoiceClient - Client name from invoice
   * @param {Array<{name: string, line: number}>} plClients - Array of P&L clients with line numbers
   * @returns {Object} Match result with line number and confidence
   */
  matchClient(invoiceClient, plClients) {
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
      console.error('Client matching error:', error);
      return {
        matched: false,
        lineNumber: null,
        confidence: 0
      };
    }
  }

  /**
   * Make API call to Claude
   * @private
   */
  callClaude(prompt) {
    const options = {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${this.apiKey}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true,
      payload: JSON.stringify({
        prompt: prompt,
        max_tokens: 150,
        temperature: 0,
        model: 'claude-2'
      })
    };

    const response = UrlFetchApp.fetch(this.endpoint, options);
    if (response.getResponseCode() !== 200) {
      throw new Error('Claude API request failed');
    }

    return JSON.parse(response.getContentText()).completion;
  }
}

/**
 * Get instance of Claude Service
 */
function getClaudeService() {
  return new ClaudeService();
} 