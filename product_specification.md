# Invoice Data Processing Tools Documentation

## Overview
This collection of tools is designed to process and manipulate invoicing data for P&L (Profit & Loss) integration. The tools handle various data transformations and calculations related to invoice processing.

## User Interface

### Custom Menu: Invoice Tools
Located in the Google Sheets menu bar, provides access to all invoice processing tools.

**Menu Items:**
1. **To Numbers**
   - Function: Converts specified numeric columns from text to number format
   - Trigger: Calls `processInvoiceData()`
   - Usage: Select this option when you need to convert text-formatted numbers to actual number format

## Functions

### convertHeaderColumnsToNumbers
**Purpose:** Transforms specific numeric columns in the invoice data from text/string format to numbers.
**Location:** NumberConverter.js

**Columns Processed:**
- Suma incasata (Amount Received)
- Incasata prin (Received Through)
- Suma ramasa de incasat (Amount Remaining to be Received)
- Valoare (Value)
- TVA (VAT)
- Total
- CursValutar (Exchange Rate)

**Limitations:**
- Only processes columns up to column Q to preserve formulas in later columns

**Input:** Raw invoice data with text/string formatted numbers
**Output:** Processed data with properly formatted numbers in specified columns

**Processing Details:**
- Removes currency symbols and special characters
- Converts comma decimal separators to dots
- Handles empty cells and invalid numbers (converts to 0)
- Preserves negative numbers
- Processes only non-header rows
- Preserves formulas in columns after Q (column R onwards)

**Error Handling:**
- Invalid numbers are converted to 0
- User is notified of success or failure via UI alert

### onOpen
**Purpose:** Creates the custom menu in Google Sheets when the spreadsheet is opened.
**Location:** Menu.js
**Trigger:** Automatically runs when the spreadsheet is opened

### P&L Reconciliation Process
**Purpose:** Automatically reconciles invoice data with P&L entries by matching clients and updating corresponding revenue values.
**Location:** PLReconciliation.js

**User Interface:**
- Dialog prompt requesting:
  - Reference month (e.g., "January")
  - URL of the P&L spreadsheet

**Process Flow:**
1. **P&L File Navigation**
   - Locates "revenues" sheet in P&L file
   - Identifies target column by finding "{month} real" in row 2
   - Maps client names from column D for matching

2. **Client Matching Process**
   - For each invoice line in source file:
     - Extracts client information
     - Queries Claude AI for client match against P&L list
     - Receives matching line number from P&L if found
   - Logs each matching attempt (success/failure)

3. **Value Update Process**
   - For successful matches:
     - Retrieves "Suma in EUR" value from column R of source file
     - Adds value to corresponding cell in P&L
     - Logs each update with before/after values

**Dependencies:**
- Claude AI Integration Library (separate project)
  - Location: ClaudeIntegration.js
  - Purpose: Handles all AI communication
  - Configuration: API key and endpoint management

**Data Requirements:**
- Source Invoice File:
  - Client information in standard format
  - "Suma in EUR" values in column R
- P&L File:
  - Sheet named "revenues"
  - Month headers in row 2 format: "{month} real"
  - Client list in column D

**Error Handling:**
- Invalid P&L URL handling
- Missing sheets/columns detection
- Client match failures logging
- Value update verification
- Network/API failure management

**Logging:**
- Client match attempts
- Successful matches with line numbers
- Value updates with before/after states
- Error conditions and failure points

**Technical Requirements:**
- Google Apps Script
- Claude AI API access
- Cross-spreadsheet permissions
- Logging infrastructure

**Security Considerations:**
- P&L file access permissions
- API key management
- Data validation before updates

**Limitations:**
- Requires stable internet connection
- Dependent on Claude AI availability
- Processing time may vary with data volume

## Technical Requirements
- Platform: Google Apps Script
- Input Format: Google Sheets data
- Data Processing: In-memory transformation 