# Invoice Processing System Documentation

## 1. System Overview
A Google Apps Script-based system for processing invoices and reconciling them with P&L entries using LLM-powered matching.

## 2. File Structure and Functionalities

### 2.1 NumberConverter.js
#### Purpose
Processes and standardizes numeric values in invoice headers, converting various formats to consistent number representations.

#### Core Functions
##### convertHeaderColumnsToNumbers(data, headerRow)
- Converts specified numeric columns in invoice data
- Parameters:
  - data: Array[][] (spreadsheet data)
  - headerRow: number (default: 0)
- Returns: Processed data array with standardized numbers

#### Numeric Columns Processed
- Suma TVA (Column E)
- Suma (Column F)
- Suma ramasa (Column G)

#### Source File Structure
1. Factura / Bon
2. Furnizor
3. Numar
4. Data emitere
5. Suma TVA
6. Suma
7. Suma ramasa
8. Moneda
9. Scadenta
10. Inregistrat
11. Data upload
12. Uploadat de
13. Centru de cost
14. Status
15. Suma in EUR
16. Matched P&L
17. EUR/RON

#### Processing Features
- Removes currency symbols
- Standardizes decimal separators
- Handles European and US number formats
- Validates numeric conversions

#### Error Handling
- Invalid number format detection
- Error notification via UI
- Zero fallback for invalid conversions

### 2.2 PLReconciliation.js
#### Purpose
Handles the reconciliation of invoice data with P&L entries across Expenses and Staffing sheets.

#### Core Functions
##### createPLReconciliationService(spreadsheetUrl, month)
- Creates service instance for P&L reconciliation
- Parameters:
  - spreadsheetUrl: Target spreadsheet URL
  - month: Processing month (e.g., "October")
- Returns: Service object with reconciliation methods

##### processReconciliation(testMode)
- Processes invoice entries against P&L sheets
- Parameters:
  - testMode: boolean (default: true) - When true, processes only first 10 entries
- Returns: void

##### processTestReconciliation()
- Convenience method to run reconciliation in test mode (10 entries)
- Returns: void

##### processFullReconciliation()
- Convenience method to run full reconciliation
- Returns: void

##### matchAndUpdateEntry(entry)
- Matches single entry against both sheets
- Updates amounts when match found
- Returns match results

#### Data Processing
- Source File Column Usage:
  - Column B (Furnizor): Supplier name for matching
  - Column O (Suma in EUR): Amount to be added for reconciliation
  - Column P (Matched P&L): Match status tracking

- Target Sheets:
  1. Expenses Sheet:
     - Match Column: C (Furnizor)
     - Update Column: "{month} real"
     - Amount Added: EUR value from source Column O
  
  2. Staffing Sheet:
     - Match Column: D (Partener)
     - Update Column: "{month} real"
     - Amount Added: EUR value from source Column O

### 2.3 ClaudeService.js
#### Purpose
Manages all LLM (Claude) interactions for supplier matching.

#### Core Functions
##### createMatchingQuery(supplier, targetData)
- Formats supplier matching queries
- Returns structured query for LLM

##### processLLMResponse(response)
- Processes LLM matching decisions
- Returns structured match results

### 2.4 UI Components
#### Status Tracking
- Column P Format: "{SheetName}!{CellReference}"
- Color Coding:
  - Green: Successfully matched
  - Gray: No match found

## 3. Data Structures

### 3.1 Source File Headers
1. Factura / Bon
2. Furnizor
3. Numar
4. Data emitere
5. Suma TVA
6. Suma
7. Suma ramasa
8. Moneda
9. Scadenta
10. Inregistrat
11. Data upload
12. Uploadat de
13. Centru de cost
14. Status
15. Suma in EUR (used as primary amount for P&L reconciliation)
16. Matched P&L
17. EUR/RON

### 3.2 Target Files Structure
#### Expenses Sheet
- Column C: Furnizor (matching column)
- Month Columns: "{month} real"

#### Staffing Sheet
- Column D: Partener (matching column)
- Month Columns: "{month} real"

## 4. Processing Logic

### 4.1 Reconciliation Flow
1. Load source and target spreadsheets
2. Determine processing mode (test/full)
3. In test mode:
   - Process only first 10 entries
   - Provide quick feedback for iteration
4. In full mode:
   - Process all entries
5. For each entry:
   - Check Expenses sheet
   - Check Staffing sheet
   - Use LLM for matching
   - Update amounts if matched
   - Update status tracking

### 4.2 Amount Handling
- Uses EUR amounts from Column O (Suma in EUR)
- Adds to existing amounts (not replace)
- Validates numeric values
- Updates in corresponding month column
- All reconciliation is done in EUR currency

### 4.3 Match Processing
- LLM-powered supplier name matching
- Confidence threshold validation
- Match reference tracking

## 5. Error Handling
- Invalid spreadsheet URLs
- Missing sheets/columns
- LLM service failures
- Amount validation errors
- Access permission issues

## 6. Reprocessing Features
- Skip matched entries
- Clear previous matches
- Selective reprocessing

## 7. Security & Performance
- Spreadsheet permission validation
- LLM rate limiting
- Data validation before updates
- Error logging and monitoring 

## 3. Data Processing Workflows

### 3.1 Number Conversion Workflow
1. User triggers processInvoiceData()
2. System processes specific columns (Suma TVA, Suma, Suma ramasa)
3. For each row:
   - Cleans numeric values (removes symbols, spaces)
   - Standardizes decimal separators
   - Converts to number format
4. Updates spreadsheet with processed data
5. Provides success/error feedback

### 3.2 P&L Reconciliation Workflow
[Previous reconciliation workflow documentation remains the same...]

## 4. Input/Output Specifications

### 4.1 Number Converter Input Formats
- European format: "1.234,56"
- US format: "1,234.56"
- Mixed formats: "1 234,56"
- Currency symbols: "€", "RON", "LEI"
- Formula cells (preserved)

### 4.2 Number Converter Output Format
- Standardized numbers
- Preserved formulas (columns R onwards)
- Zero for invalid conversions

### 4.3 P&L Reconciliation Formats
[Previous P&L format documentation remains the same...] 

### 6. Testing Features
#### Quick Iteration Mode
- Process only first 10 transactions
- Enabled by default for faster testing
- Comprehensive logging in test mode
- Easy switching between test/full mode
- Separate methods for test and full processing 