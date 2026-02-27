/**
 * 設定情報
 */
const NAVI_EMAIL_MAP = {
  '菅谷さん': 'naomisugaya721@gmail.com',
  '小木曽さん': 'yuki9517@gmail.com',
  '佐藤さん': 'nishi.shima217@gmail.com',
  '浅富さん': 'naomi600143@gmail.com',
  '中村さん': 'hitomi10321107@gmail.com',
  '千田さん': 'shuntaro8710@gmail.com',
  '瀬戸さん': 'k-seto@circus-group.jp'
};

const CONFIG_NAVI = {
  SHEET_NAME: '📅面談予定',
  START_ROW: 76,
  COLUMN_INDICES: {
    NAVI_NAME: 12,      // M列
    INTERVIEW_TIME: 13, // N列
    CANDIDATE_ID: 14,   // O列
    CANDIDATE_NAME: 15, // P列
    PHONE_NUMBER: 17,   // R列
    EMAIL: 18,          // S列
    INTERVIEW_TYPE: 19  // T列
  }
};

const CC_ADDRESS = "agentnavi@circus-group.jp";

// ==========================================================
// 1. エントリーポイント（ここを実行・トリガー設定する）
// ==========================================================

/**
 * 【手動用】ボタンやメニューから実行
 * 送信前に確認ダイアログを出します
 */
function manual_sendNaviRemindEmails() {
  const ui = SpreadsheetApp.getUi();
  
  const report = processNaviRemind((naviName, recipient, body, itemCount) => {
    // 送信確認ダイアログ
    const confirm = ui.alert(
      `${naviName}へ送信確認 (${itemCount}件)`,
      `宛先: ${recipient}\n\n${body}`,
      ui.ButtonSet.YES_NO
    );

    if (confirm === ui.Button.YES) {
      sendEmail(recipient, body);
      return `【完了】${naviName}に送信しました。`;
    } else {
      return `【スキップ】${naviName}の送信をキャンセルしました。`;
    }
  });

  if (report) ui.alert("実行結果:\n\n" + report);
}

/**
 * 【自動用】時間トリガーで実行
 * 確認なしで即送信し、詳細なログを残します
 */
function auto_sendNaviRemindEmails() {
  console.info("--- [自動リマインド開始] ---");
  
  let successCount = 0;
  let skipCount = 0;
  let errorCount = 0;

  const summary = processNaviRemind((naviName, recipient, body, itemCount) => {
    try {
      sendEmail(recipient, body);
      console.info(`✅ 送信成功: ${naviName} (${recipient}) / 面談数: ${itemCount}件`);
      successCount++;
      return null; // summaryには含めない
    } catch (e) {
      console.error(`❌ 送信失敗: ${naviName} / エラー: ${e.message}`);
      errorCount++;
      return `【エラー】${naviName}: ${e.message}`;
    }
  });

  // スキップ（アドレス未登録など）の情報をログに吐き出す
  if (summary && summary !== "明日の予定はありません。") {
    summary.split("\n").forEach(line => {
      if (line.includes("【不可】")) {
        console.warn(`⚠️ ${line}`);
        skipCount++;
      }
    });
  } else if (summary === "明日の予定はありません。") {
    console.info("明日の面談予定は0件でした。処理を終了します。");
    return;
  }

  console.info(`--- [自動リマインド終了] 成功: ${successCount}件 / スキップ: ${skipCount}件 / エラー: ${errorCount}件 ---`);
}


// ==========================================================
// 2. 共通ロジック
// ==========================================================

/**
 * データの抽出からループ処理までのメインロジック
 */
function processNaviRemind(sendLogic) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_NAVI.SHEET_NAME);
  if (!sheet) return "エラー: シートが見つかりません。";

  // 日付準備
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowStr = Utilities.formatDate(tomorrow, 'JST', 'M/d');
  const tomorrowCompare = Utilities.formatDate(tomorrow, 'JST', 'yyyy/MM/dd');

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG_NAVI.START_ROW) return "データがありません。";
  
  const data = sheet.getRange(CONFIG_NAVI.START_ROW, 1, lastRow - CONFIG_NAVI.START_ROW + 1, 21).getValues();
  
  // 担当者ごとにデータをまとめる
  const naviDataMap = groupDataByNavi(data, tomorrowCompare);
  const naviKeys = Object.keys(naviDataMap);

  if (naviKeys.length === 0) return "明日の予定はありません。";

  const results = [];
  for (const naviName of naviKeys) {
    // 名寄せ（「佐藤し」→「佐藤」、「佐藤」→「佐藤さん」など）
    const recipient = NAVI_EMAIL_MAP[naviName] || NAVI_EMAIL_MAP[naviName + "さん"];
    const items = naviDataMap[naviName];

    if (!recipient) {
      results.push(`【不可】${naviName} のアドレスが未登録です。`);
      continue;
    }

    const body = generateMailBody(tomorrowStr, items);
    
    // 外部から渡された送信処理を実行
    const res = sendLogic(naviName, recipient, body, items.length);
    if (res) results.push(res);
  }

  return results.length > 0 ? results.join("\n") : "";
}

/**
 * データを担当者ごとにグルーピング
 */
function groupDataByNavi(data, dateStr) {
  const map = {};
  data.forEach(row => {
    const timeValue = row[CONFIG_NAVI.COLUMN_INDICES.INTERVIEW_TIME];
    if (!(timeValue instanceof Date)) return;
    
    // 日付チェック
    if (Utilities.formatDate(timeValue, 'JST', 'yyyy/MM/dd') !== dateStr) return;

    // 名前整形
    let name = row[CONFIG_NAVI.COLUMN_INDICES.NAVI_NAME].toString()
                .replace(/[【】]/g, '') // 【】を消す
                .replace(/し$/, '')     // 佐藤し 対策
                .trim();
    
    if (!name) return;
    if (!map[name]) map[name] = [];

    map[name].push({
      time: Utilities.formatDate(timeValue, 'JST', 'HH:mm'),
      candidateName: row[CONFIG_NAVI.COLUMN_INDICES.CANDIDATE_NAME],
      candidateId: row[CONFIG_NAVI.COLUMN_INDICES.CANDIDATE_ID],
      phone: row[CONFIG_NAVI.COLUMN_INDICES.PHONE_NUMBER],
      email: row[CONFIG_NAVI.COLUMN_INDICES.EMAIL],
      type: row[CONFIG_NAVI.COLUMN_INDICES.INTERVIEW_TYPE]
    });
  });
  return map;
}

/**
 * メール本文の生成
 */
function generateMailBody(dateStr, items) {
  const listText = items.map(item => 
    `${item.time}〜\n${item.candidateId} ${item.candidateName} 様\n${item.phone}  ${item.email}  ${item.type}`
  ).join("\n\n");

  return `お疲れ様です。
現時点での明日(${dateStr})のナビ・個別面談のリマインドを送信いたします。

${listText}

変更・追加があればLINEにてご連絡いたします。
宜しくお願いします！`;
}

/**
 * 実際のメール送信
 */
function sendEmail(to, body) {
  // GmailApp.sendEmail(to, "明日の面談リマインドのご連絡", body, {
  //   name: "転職エージェントナビ事務局",
  //   cc: CC_ADDRESS
  // });
}