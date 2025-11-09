function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Import Tools')
    .addItem('📥 Import Data', 'main')
    .addSeparator()
    .addItem('📊 Update Kira-PG-Bank Tally Summary', 'runSummaryProcessor')
    .addItem('💰 Update Deposit', 'runDepositProcessor')
    .addSeparator()
    .addItem('Take KIRA', 'takeKIRA')
    .addItem('Take PG', 'takePG')
    .addItem('Import Deposit', 'importdeposit')
    .addItem('Import Merchant', 'importmerchant')
    .addItem('Import Agent', 'importagent')
    .addToUi();
}
