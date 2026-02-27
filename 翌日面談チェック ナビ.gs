/**
 * 2. 自動実行用：トリガーで実行する場合（ログ出力強化版）
 */
function auto_sendNaviRemindEmails() {
  console.info("--- [自動送信ジョブ開始] 明日の面談リマインド ---");
  
  let successCount = 0;
  let skipCount = 0;
  let errorCount = 0;

  const summaryReport = processNaviRemind((naviName, recipient, body, itemCount) => {
    try {
      // 実際の送信
      sendEmail(recipient, body);
      
      // 成功ログ（詳細）
      console.info(`【送信成功】担当者: ${naviName} / 宛先: ${recipient} / 面談件数: ${itemCount}件`);
      successCount++;
      return null; 
    } catch (e) {
      // エラーログ
      console.error(`【送信失敗】担当者: ${naviName} / 原因: ${e.message}`);
      errorCount++;
      return null;
    }
  });

  // 予定が一件もなかった場合の処理
  if (summaryReport === "明日の予定はありません。") {
    console.warn("--- [終了] 明日の面談予定がスプレッドシートに見つかりませんでした ---");
    return;
  }

  // スキップされた（アドレス不明など）のログを解析して出力
  if (summaryReport) {
    const lines = summaryReport.split("\n");
    lines.forEach(line => {
      if (line.includes("【不可】")) {
        console.warn(line);
        skipCount++;
      }
    });
  }

  console.info(`--- [自動送信ジョブ終了] 成功: ${successCount}件 / スキップ: ${skipCount}件 / エラー: ${errorCount}件 ---`);
}

/**
 * 3. 共通ロジック（コールバック引数にitemCountを追加）
 */
function processNaviRemind(sendLogic) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_NAVI.SHEET_NAME);
  if (!sheet) {
    console.error("シート「" + CONFIG_NAVI.SHEET_NAME + "」が見つかりません。");
    return "シートが見つかりません。";
  }

  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowStr = Utilities.formatDate(tomorrow, 'JST', 'M/d');
  const tomorrowCompare = Utilities.formatDate(tomorrow, 'JST', 'yyyy/MM/dd');

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG_NAVI.START_ROW) return "データがありません。";
  
  const data = sheet.getRange(CONFIG_NAVI.START_ROW, 1, lastRow - CONFIG_NAVI.START_ROW + 1, 21).getValues();
  const naviDataMap = groupDataByNavi(data, tomorrowCompare);

  const naviKeys = Object.keys(naviDataMap);
  if (naviKeys.length === 0) return "明日の予定はありません。";

  const results = [];
  for (const naviName in naviDataMap) {
    // 担当者名の名寄せ
    const recipient = NAVI_EMAIL_MAP[naviName] || NAVI_EMAIL_MAP[naviName + "さん"];
    const items = naviDataMap[naviName];

    if (!recipient) {
      results.push(`【不可】${naviName} さんのアドレスがMAPに登録されていません。`);
      continue;
    }

    const body = generateMailBody(tomorrowStr, items);
    
    // コールバック実行（第4引数に件数を追加）
    const res = sendLogic(naviName, recipient, body, items.length);
    if (res) results.push(res);
  }

  return results.length > 0 ? results.join("\n") : "";
}