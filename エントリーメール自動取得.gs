/**
 * ==========================================
 * 設定エリア
 * ==========================================
 */
const GET_MAIL_CONFIG = {
  // ★ true: テストモード（既読にしない・アーカイブしない）
  // ★ false: 本番モード（処理後に既読化・アーカイブする）
  IS_TEST_MODE: false, 

  SHEET_NAME: "テスト中メール自動取得", 
  
  SEARCH: {
    AGENT_NAVI: 'is:unread from:agentnavi@circus-group.jp subject:"[自動送信 | 転職エージェントナビ ]" AND "自動送信" newer_than:2d',
    FB_LEAD: 'is:unread subject:FBリードからお問い合わせがありました newer_than:2d'
  }
};

/**
 * メイン実行関数
 */
function main_getEntryMail() {
  const modeName = GET_MAIL_CONFIG.IS_TEST_MODE ? "【テストモード】" : "【本番モード】";
  console.log(`=== 処理開始 ${modeName} ===`);
  if (GET_MAIL_CONFIG.IS_TEST_MODE) console.log("※Gmailの既読化・アーカイブは行いません。");

  console.log("--- AgentNavi ---");
  processEmailGroup(GET_MAIL_CONFIG.SEARCH.AGENT_NAVI, extractAgentNaviData);

  console.log("--- FB Lead ---");
  processEmailGroup(GET_MAIL_CONFIG.SEARCH.FB_LEAD, extractFbLeadData);
}

/**
 * メール検索〜書き込みまでの統括処理
 */
function processEmailGroup(query, extractionCallback) {
  console.log(`検索クエリ: [${query}]`);
  const threads = GmailApp.search(query, 0, 500);
  console.log(`ヒット件数: ${threads.length}件`);
  
  if (threads.length === 0) return;

  const newData = [];

  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      if (message.isUnread()) {
        const extractedData = extractionCallback(message);
        if (extractedData) {
          newData.push(extractedData);
        }
        if (!GET_MAIL_CONFIG.IS_TEST_MODE) message.markRead(); 
      }
    });
    if (!GET_MAIL_CONFIG.IS_TEST_MODE) thread.moveToArchive();
  });

  if (newData.length > 0) {
    saveDataToSheet7(newData);
  } else {
    console.log("抽出対象のデータはありませんでした。");
  }
}

/**
 * 「シート7」に書き込む関数
 * ★修正点: A,B,C,E列を上書きしないよう、書き込みを2分割しました
 */
function saveDataToSheet7(data) {
  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GET_MAIL_CONFIG.SHEET_NAME);
  
  if (!targetSheet) {
    console.error(`エラー: シート「${GET_MAIL_CONFIG.SHEET_NAME}」が見つかりません。`);
    return;
  }

  // D列(4列目)を見て最終行を判定
  const lastRowWithData = getLastRowInColumn(targetSheet, 4); 
  const startRow = lastRowWithData === 0 ? 1 : lastRowWithData + 1;

  // 1. 【D列】のデータを作成（日付＆時間）
  const dataForColD = data.map(row => [`${row[0]} ${row[1]}`]);

  // 2. 【F列以降】のデータを作成
  //    行番号が必要なため、map内で計算
  const dataForColF_Onwards = data.map((row, index) => {
    return createRowFromF(row, startRow + index);
  });

  // 書き込み実行
  // (1) D列 (4列目) に書き込み
  targetSheet.getRange(startRow, 4, dataForColD.length, 1).setValues(dataForColD);
  
  // (2) F列 (6列目) 以降に書き込み
  targetSheet.getRange(startRow, 6, dataForColF_Onwards.length, dataForColF_Onwards[0].length).setValues(dataForColF_Onwards);

  console.log(`「${GET_MAIL_CONFIG.SHEET_NAME}」の ${startRow} 行目に ${data.length} 件追加しました。（A-C, E列は保持）`);
}

/**
 * F列以降のデータ配列を作成する関数
 * @param {Array} row - 元データ
 * @param {Number} rowIndex - 書き込まれる行番号
 */
function createRowFromF(row, rowIndex) {
  const [
    date, time, name, furi, phone, email, age, 
    _empty, location, remarks, subject, sender, msgId
  ] = row;

  // 備考欄の調整
  let finalRemarks = remarks;
  if (location && location !== "") {
    finalRemarks = (finalRemarks ? finalRemarks + "\n" : "") + `都道府県：${location}`;
  }

  // F列用数式
  const formulaF = `=IFERROR(VLOOKUP(AV${rowIndex},'パートナーID・メディアIDリスト'!$D$1:$G$99,4,0))`;

  // P列(16)〜AV列(48)の間を埋める空白 (Q〜AU = 31列分)
  const gapColumns = Array(31).fill(""); 

  return [
    formulaF,            // F: 集客経路 (数式)
    "",                  // G: (空白)
    "",                  // H: (空白)
    name,                // I: 求職者名
    furi,                // J: フリガナ
    phone,               // K: 電話番号
    email,               // L: メールアドレス
    age,                 // M: 年齢
    "",                  // N: 性別
    "",                  // O: (空白)
    finalRemarks,        // P: remarks 備考
    ...gapColumns,       // Q 〜 AU: 空白埋め
    subject,             // AV: LP/流入元
    sender,              // AW: 送信元
    msgId                // AX: MessageID
  ];
}

/**
 * 指定列の最終データ行を取得
 */
function getLastRowInColumn(sheet, columnNumber) {
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) return 0;
  const values = sheet.getRange(1, columnNumber, lastRow, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "" && values[i][0] != null) {
      return i + 1;
    }
  }
  return 0;
}

/* ==========================================
 * データ抽出ロジック
 * ========================================== */

function extractAgentNaviData(message) {
  return parseEmailBody(message, {
    name: /【 お名前 】\s*([\s\S]+?)\r?\n/,
    furi: /【 お名前（フリガナ） 】\s*([\s\S]+?)\r?\n/
  });
}

function extractFbLeadData(message) {
  return parseEmailBody(message, {
    name: /【氏名】\s*([\s\S]+?)\r?\n/,
    furi: /【 お名前（フリガナ） 】\s*([\s\S]+?)\r?\n/
  });
}

function parseEmailBody(message, patterns) {
  const body = message.getBody();
  const subject = message.getSubject();
  const dateObj = message.getDate();
  
  const getVal = (regex) => {
    const match = body.match(regex);
    return match ? match[1].trim() : "";
  };

  const name = getVal(patterns.name);
  const furi = getVal(patterns.furi);
  
  let age = getVal(/【 年齢 】\s*(\S+)/);
  if (age.includes("歳")) age = age.replace("歳", "");

  const email = getVal(/【\s*メールアドレス\s*】\s*(\S+)/);
  
  const rawPhone = getVal(/【\s*電話番号\s*】\s*(\S+)/) || getVal(/【 お電話番号 】\s*(\S+)/);
  const phone = rawPhone ? `=TEXT("${rawPhone}","0##########")` : "";

  const location = getVal(/【\s*希望勤務地\s*】\s*(\S+)/);
  const education = getVal(/【\s*最終学歴\s*】\s*(\S+)/);
  const hasChangedJob = getVal(/【 就職したことはありますか？ 】\s*(\S+)/);
  const seoId = getVal(/お問い合わせID(?:（識別番号）)?\s*[:：]\s*(\S+)/);
  const sender = getVal(/送信元：\s*(\S+)/);

  const remarksList = [];
  if (education) remarksList.push(education);
  if (hasChangedJob) remarksList.push("【 就職したことはありますか？ 】" + hasChangedJob);
  if (seoId) remarksList.push(seoId);

  return [
    Utilities.formatDate(dateObj, "JST", "yyyy-MM-dd"),
    Utilities.formatDate(dateObj, "JST", "HH:mm"),
    name,
    furi,
    phone,
    email,
    age,
    "",
    location,
    remarksList.join("\n"),
    subject,
    sender,
    message.getId()
  ];
}