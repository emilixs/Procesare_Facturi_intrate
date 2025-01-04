/**
 * Claude AI Integration Library
 * Handles communication with Claude AI for text analysis and matching
 */

/**
 * Configuration object for Claude API
 * @typedef {Object} ClaudeConfig
 * @property {string} apiKey - The Claude API key
 * @property {string} apiEndpoint - The API endpoint URL
 * @property {number} maxTokens - Maximum tokens for response (default: 150)
 * @property {number} temperature - Response temperature (default: 0)
 */

class ClaudeService {
  /**
   * Initialize the Claude service
   * @param {ClaudeConfig} config - Configuration object
   */
  constructor(config) {
    this.validateConfig(config);
    this.apiKey = config.apiKey;
    this.apiEndpoint = config.apiEndpoint;
    this.maxTokens = config.maxTokens || 4000;
    this.temperature = config.temperature || 0;
  }

  /**
   * Validate the configuration object
   * @param {ClaudeConfig} config 
   * @private
   */
  validateConfig(config) {
    if (!config.apiKey) throw new Error('API key is required');
    if (!config.apiEndpoint) throw new Error('API endpoint is required');
  }

  /**
   * Compare client names and find matches
   * @param {string} sourceClient - Client name from invoice
   * @param {string[]} targetClients - Array of client names from P&L
   * @returns {Promise<{matched: boolean, lineNumber: number|null, confidence: number}>}
   */
  async findClientMatch(sourceClient, targetClients) {
    const prompt = this.buildMatchingPrompt(sourceClient, targetClients);
    
    try {
      const response = await this.callClaude(prompt);
      return this.parseMatchingResponse(response);
    } catch (error) {
      console.error('Error in client matching:', error);
      throw new Error(`Client matching failed: ${error.message}`);
    }
  }

  /**
   * Build the prompt for client matching
   * @param {string} sourceClient 
   * @param {string[]} targetClients 
   * @returns {string}
   * @private
   */
  buildMatchingPrompt(sourceClient, targetClients) {
    return `
Task: Find if and where the source client name matches any of the target client names.
Source client: "${sourceClient}"
Target clients (with line numbers):
${targetClients.map((client, index) => `${index + 1}. ${client}`).join('\n')}

Rules:
1. Consider variations in spelling, abbreviations, and company suffixes
2. Account for different languages (e.g., English vs Romanian company types)
3. Ignore case and special characters
4. Consider partial matches if they uniquely identify the company

Response format (JSON):
{
  "matched": boolean,
  "lineNumber": number or null,
  "confidence": number (0-1),
  "explanation": "brief explanation of the match or non-match"
}
`;
  }

  /**
   * Parse Claude's response for client matching
   * @param {string} response 
   * @returns {Object}
   * @private
   */
  parseMatchingResponse(response) {
    try {
      const result = JSON.parse(response);
      return {
        matched: result.matched,
        lineNumber: result.lineNumber,
        confidence: result.confidence
      };
    } catch (error) {
      throw new Error('Failed to parse Claude response');
    }
  }

  /**
   * Make the API call to Claude
   * @param {string} prompt 
   * @returns {Promise<string>}
   * @private
   */
  async callClaude(prompt) {
    const options = {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${this.apiKey}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true,
      payload: JSON.stringify({
        prompt: prompt,
        max_tokens: this.maxTokens,
        temperature: this.temperature
      })
    };

    try {
      const response = UrlFetchApp.fetch(this.apiEndpoint, options);
      const responseCode = response.getResponseCode();
      
      if (responseCode !== 200) {
        throw new Error(`API request failed with status ${responseCode}`);
      }
      
      const responseData = JSON.parse(response.getContentText());
      return responseData.completion;
    } catch (error) {
      throw new Error(`API call failed: ${error.message}`);
    }
  }
}

/**
 * Create a new instance of the Claude service
 * @param {ClaudeConfig} config 
 * @returns {ClaudeService}
 */
function createClaudeService(config) {
  return new ClaudeService(config);
} 