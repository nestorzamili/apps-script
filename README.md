# Payment Gateway Reconciliation System

Automated data reconciliation system for payment gateway transactions using Google Apps Script.

## Overview

This system processes and reconciles transaction data from three sources (KIRA, PG, and BANK) and generates summary reports for financial analysis.

## Features

- **Automated Import**: Reads files from Google Drive folders (KIRA, PG, BANK)
- **Data Reconciliation**: Merges transaction data from multiple sources
- **Summary Generation**: Creates consolidated Kira-PG-Bank Tally Summary
- **Deposit Tracking**: Generates deposit reports with fees calculation
- **Settlement Calculation**: Calculates settlement dates based on Malaysia holidays
- **Fee Processing**: Applies merchant-specific and PG-specific fees

## File Structure

### Core Modules
- `Config.gs` - Configuration and constants
- `Main.gs` - Main import and merge process
- `Menu.gs` - Custom menu for Google Sheets

### Processors
- `SummaryProcessor.gs` - Generates Kira-PG-Bank Tally Summary
- `DepositProcessor.gs` - Generates deposit reports
- `ParameterProcessor.gs` - Loads fee and settlement parameters
- `HolidayProcessor.gs` - Malaysia holiday calendar integration

### Data Handlers
- `KiraProcessor.gs` - Processes KIRA transaction files
- `PGProcessor.gs` - Processes payment gateway files
- `BankProcessor.gs` - Processes bank settlement files
- `DataMerger.gs` - Merges data from all sources

### Utilities
- `Utils.gs` - Common utility functions
- `SheetWriter.gs` - Google Sheets write operations

## Data Flow

```
KIRA Folder → KiraProcessor
PG Folder   → PGProcessor     → DataMerger → Import Data Sheet
BANK Folder → BankProcessor

Import Data → SummaryProcessor → Kira-PG-Bank Tally Summary
Import Data → DepositProcessor → Deposit Sheet
```

## Configuration

Edit `Config.gs` to update:
- Spreadsheet ID
- Folder IDs (KIRA, PG, BANK)
- Sheet names
- Batch size

## Usage

### From Custom Menu (⚙️ Import Tools)

1. **Import Data** - Import and merge all transaction files
2. **Update Kira-PG-Bank Tally Summary** - Generate summary report
3. **Update Deposit** - Generate deposit report

### From Apps Script Editor

```javascript
main()                    // Import data
runSummaryProcessor()     // Generate summary
runDepositProcessor()     // Generate deposit report
```
