/**データ移行時のみ使っていました
 * J列をスクリーニングし、日付以外のデータをB列に移動します。
 * 実行する前に、下の「▼▼ ユーザー設定 ▼▼」セクションを編集してください。
 */
function cleanUpJColumn() {
  
  // ▼▼ ユーザー設定 ▼▼
  // 1. 処理対象のシート名
  const SHEET_NAME = '送客管理'; 
  
  // 2. データが始まる行番号（見出しの次の行）
  // (前回の画像から11行目、または Set2 スクリプトの 5行目と仮定)
  const START_ROW = 5; 
  // ▲▲ ユーザー設定 ▲▲

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error(`シートが見つかりません: ${SHEET_NAME}`);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < START_ROW) {
      Logger.log('処理するデータがありません。');
      return;
    }

    // B列(2)からJ列(10)までの範囲を取得 (列数は 10 - 2 + 1 = 9列)
    const numRows = lastRow - START_ROW + 1;
    const numCols = 9; // B列からJ列までの列数
    const range = sheet.getRange(START_ROW, 2, numRows, numCols);
    const values = range.getValues();

    let processedCount = 0;

    // データをループ処理
    // values[i][0] が B列
    // values[i][8] が J列
    for (let i = 0; i < values.length; i++) {
      const jValue = values[i][8]; // J列の値

      // J列に値があり、かつそれが日付オブジェクトではない場合
      if (jValue !== '' && !(jValue instanceof Date)) {
        
        // J列の値をB列にコピー
        // (B列にすでにある値を上書きします。追記する場合は values[i][0] + jValue に変更)
        values[i][0] = jValue; 
        
        // J列の値を空白にする
        values[i][8] = ''; 
        
        processedCount++;
      }
    }

    // 変更したデータを一括書き戻し
    if (processedCount > 0) {
      range.setValues(values);
      Logger.log(`${processedCount} 件のデータを処理しました。`);
      SpreadsheetApp.getUi().alert(`${processedCount} 件のデータをB列に移動し、J列をクリアしました。`);
    } else {
      Logger.log('処理対象のデータはありませんでした。');
      SpreadsheetApp.getUi().alert('処理対象のデータはありませんでした。');
    }

  } catch (e) {
    Logger.log(e);
    SpreadsheetApp.getUi().alert(`エラーが発生しました: ${e.message}`);
  }
}