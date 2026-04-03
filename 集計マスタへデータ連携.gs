// --- 設定項目 ---
const CONFIG_shukeiSync = {
  // 共通の転記先スプレッドシートID
  TARGET_SS_ID: '17fN5a50OYPlBNquYUI2by77UowcY09MZzTp-XjaRAps',
  
  // バッチサイズ（1回で処理する行数。27,000行なら3,000〜5,000くらいが安全）
  BATCH_SIZE: 3000,

  // タスク1の設定
  TASK1: {
    srcName: '求職者管理',
    srcRange: 'A:AX', // A〜AX
    dstName: '参照用(新求職者管理)',
    dstStartCol: 1    // A列から
  },

  // タスク2の設定
  TASK2: {
    srcName: '送客管理',
    srcRange: 'A:BM', // A〜BM
    dstName: '参照用(新送客済み管理)',
    dstStartCol: 2    // B列から
  }
};

/**
 * 同期を開始する（最初の一回だけ手動で実行）
 * 以前の進捗をリセットして、トリガーを設定します
 */
function startFullSync() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('LAST_ROW_TASK1', '0');
  props.setProperty('LAST_ROW_TASK2', '0');
  props.setProperty('SYNC_STATUS', 'RUNNING');
  
  // 既存のトリガーがあれば削除（重複防止）
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'executeBatchSync') ScriptApp.deleteTrigger(t);
  });
  
  // 5分おきに実行するトリガーを作成
  ScriptApp.newTrigger('executeBatchSync')
    .timeBased()
    .everyMinutes(5)
    .create();
  
  console.log("同期を開始しました。5分おきにバッチ処理が走ります。");
  executeBatchSync(); // 初回実行
}

/**
 * 分割実行されるメイン処理
 */
function executeBatchSync() {
  const props = PropertiesService.getScriptProperties();
  const status = props.getProperty('SYNC_STATUS');
  if (status !== 'RUNNING') return;

  const targetSs = SpreadsheetApp.openById(CONFIG_shukeiSync.TARGET_SS_ID);
  
  // タスク1の処理
  const lastRow1 = parseInt(props.getProperty('LAST_ROW_TASK1') || '0');
  const finished1 = processTask(CONFIG_shukeiSync.TASK1, lastRow1, targetSs, 'LAST_ROW_TASK1');

  // タスク1が終わっていればタスク2を処理
  if (finished1) {
    const lastRow2 = parseInt(props.getProperty('LAST_ROW_TASK2') || '0');
    const finished2 = processTask(CONFIG_shukeiSync.TASK2, lastRow2, targetSs, 'LAST_ROW_TASK2');
    
    if (finished2) {
      props.setProperty('SYNC_STATUS', 'COMPLETED');
      // トリガーを削除
      const triggers = ScriptApp.getProjectTriggers();
      triggers.forEach(t => ScriptApp.deleteTrigger(t));
      console.log("すべての同期が完了しました。トリガーを停止しました。");
    }
  }
}

/**
 * 各シートの転記処理
 */
function processTask(task, startRow, targetSs, propKey) {
  const srcSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(task.srcName);
  const dstSheet = targetSs.getSheetByName(task.dstName);
  
  // 初回実行時のみ転記先をクリア
  if (startRow === 0) {
    dstSheet.clearContents();
  }

  const allValues = srcSheet.getRange(task.srcRange).getValues();
  const totalRows = allValues.length;
  
  // ヘッダー行を含まないデータがあるか確認
  if (startRow >= totalRows) return true;

  // 今回処理する範囲を切り出し
  const endRow = Math.min(startRow + CONFIG_shukeiSync.BATCH_SIZE, totalRows);
  const chunk = allValues.slice(startRow, endRow);
  
  // 転記先へ書き込み
  dstSheet.getRange(startRow + 1, task.dstStartCol, chunk.length, chunk[0].length).setValues(chunk);
  
  // 進捗を保存
  PropertiesService.getScriptProperties().setProperty(propKey, endRow.toString());
  console.log(`${task.srcName}: ${startRow}行目から${endRow}行目まで転記完了`);

  return endRow >= totalRows;
}