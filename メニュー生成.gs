function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('突合')
    // .addSubMenu(ui.createMenu('ナビ面談')
    //     .addItem('マスタ⇒加工データ', 'performNaviReconciliation_MasterToProcessed')
    //     .addItem('加工データ⇒マスタ', 'performNaviReconciliation_ProcessedToMaster'))
    // .addSubMenu(ui.createMenu('AGT面談')
    //     .addItem('マスタ⇒加工データ', 'performAGTReconciliation_MasterToProcessed')
    //     .addItem('加工データ⇒マスタ', 'performAGTReconciliation_ProcessedToMaster'))
    .addSubMenu(ui.createMenu('翌日面談連絡')
        .addItem('翌日面談連絡メール送付', 'generateCompanySpecificEmailsWithConfirmation'))
    .addToUi();
}
