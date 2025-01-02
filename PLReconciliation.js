/**
 * Match invoice client with P&L clients
 * @param {string} invoiceClient 
 * @param {Array<{name: string, line: number}>} plClients 
 */
function findMatchingClient(invoiceClient, plClients) {
  const claude = getClaudeService();
  const result = claude.matchClient(invoiceClient, plClients);
  
  if (result.matched && result.confidence > 0.8) {
    return result.lineNumber;
  }
  return null;
} 