function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('突合')
    // .addSubMenu(ui.createMenu('ナビ面談')
    //     .addItem('マスタ⇒加工データ', 'performNaviReconciliation_MasterToProcessed')
    //     .addItem('加工データ⇒マスタ', 'performNaviReconciliation_ProcessedToMaster'))
    .addSubMenu(ui.createMenu('開発中：翌日面談連絡【ナビ】')
        .addItem('開発中', 'manual_sendNaviRemindEmails'))
    .addSubMenu(ui.createMenu('翌日面談連絡')
        .addItem('翌日面談連絡メール送付', 'manual_generateAgtRemindEmails'))
    .addToUi();
  ui.createMenu("📄 PDF生成")
    .addItem("選択行のPDFを手動生成", "generatePdfManual")
    .addItem("使い方", "showHelp")
    .addToUi();
}
