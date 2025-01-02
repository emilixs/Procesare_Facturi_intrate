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

## Technical Requirements
- Platform: Google Apps Script
- Input Format: Google Sheets data
- Data Processing: In-memory transformation 