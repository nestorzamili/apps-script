#  Payment Gateway Financial Sheets Automation

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
