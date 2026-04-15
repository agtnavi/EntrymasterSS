//20260415その回の重複なのか、以前の重複なのかで扱いを変える
//20260415重複検索範囲を全ての行に拡大

/**
 * ==========================================
 * 設定エリア
 * ==========================================
 */
const GET_MAIL_CONFIG = {
  // ★ true: テストモード（既読にしない・アーカイブしない）
  // ★ false: 本番モード（処理後に既読化・アーカイブする）
  IS_TEST_MODE: false, 

  SHEET_NAME: "求職者管理", // 書き込み先のシート名
  
  // 検索クエリの設定
  SEARCH: {
    AGENT_NAVI: 'is:unread from:agentnavi@circus-group.jp subject:"[自動送信 | 転職エージェントナビ ]" AND "自動送信" newer_than:2d',
    FB_LEAD: 'is:unread subject:FBリードからお問い合わせがありました newer_than:1d'
  },

  // 列の定義（0から数えたインデックス）
  COL: {
    NEXT_ACTION: 7, // H列: ネクストアクション（重複フラグを立てる列）
    PHONE: 10,      // K列: 既存データの電話番号
    EMAIL: 11,      // L列: 既存データのメールアドレス
    ID_BASE: 4      // D列: 書き込み開始行を判定する基準列
  }
};

/**
 * メイン実行関数
 * この関数を実行すると、メールを取得してスプレッドシートへ転記します
 */
function main_getEntryMail() {
  const modeName = GET_MAIL_CONFIG.IS_TEST_MODE ? "【テストモード】" : "【本番モード】";
  console.log(`=== 処理開始 ${modeName} ===`);

  console.log("--- AgentNavi 取得中 ---");
  processEmailGroup(GET_MAIL_CONFIG.SEARCH.AGENT_NAVI, extractAgentNaviData);

  console.log("--- FB Lead 取得中 ---");
  processEmailGroup(GET_MAIL_CONFIG.SEARCH.FB_LEAD, extractFbLeadData);
  
  console.log("=== 全工程 終了 ===");
}

/**
 * メールの検索・解析・保存を統括する処理
 */
function processEmailGroup(query, extractionCallback) {
  const threads = GmailApp.search(query, 0, 500);
  if (threads.length === 0) {
    console.log("対象のメールは見つかりませんでした。");
    return;
  }

  const newData = [];

  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      // 未読メッセージのみ処理
      if (message.isUnread()) {
        const extractedData = extractionCallback(message);
        if (extractedData) {
          newData.push(extractedData);
        }
        // テストモードでなければ既読にする
        if (!GET_MAIL_CONFIG.IS_TEST_MODE) message.markRead(); 
      }
    });
    // テストモードでなければアーカイブする
    if (!GET_MAIL_CONFIG.IS_TEST_MODE) thread.moveToArchive();
  });

  if (newData.length > 0) {
    saveEntryDataToSheet(newData);
  }
}

/**
 * データをシートに書き込む（重複チェック機能付き）
 */
/**
 * データをシートに書き込む
 * ロジック:
 * 1. 同一実行内の重複（バグ） -> シートに書き込まずスキップ
 * 2. 過去データとの重複（再入会） -> H列に「重複」と記載して書き込む
 */
function saveEntryDataToSheet(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GET_MAIL_CONFIG.SHEET_NAME);
  if (!sheet) {
    console.error(`エラー: シート「${GET_MAIL_CONFIG.SHEET_NAME}」が見つかりません。`);
    return;
  }

  // --- 1. 【過去データ】の読込（再入会判定用） ---
  const allValues = sheet.getDataRange().getValues();
  const historicalPhones = new Set();
  const historicalEmails = new Set();

  allValues.forEach((row, i) => {
    if (i === 0) return; // ヘッダー
    const p = String(row[GET_MAIL_CONFIG.COL.PHONE] || "").replace(/[^0-9]/g, "");
    const e = String(row[GET_MAIL_CONFIG.COL.EMAIL] || "").trim().toLowerCase();
    if (p) historicalPhones.add(p);
    if (e) historicalEmails.add(e);
  });

  // --- 2. 【今回の実行データ】の選別 ---
  const batchPhones = new Set();
  const batchEmails = new Set();
  const finalDataToWrite = [];

  data.forEach(row => {
    const newPhoneClean = String(row[4] || "").replace(/[^0-9]/g, "");
    const newEmail = String(row[5] || "").trim().toLowerCase();

    // A. 同一バッチ内の重複チェック（バグ対策：書き込まない）
    if (batchPhones.has(newPhoneClean) || batchEmails.has(newEmail)) {
      console.log(`[同一バッチ重複スキップ] ${row[2]} 様（システム重複のため除外）`);
      return; // この行の処理を飛ばす
    }

    // B. 過去データとの重複チェック（再入会対策：書き込むが「重複」ラベル付与）
    if (historicalPhones.has(newPhoneClean) || historicalEmails.has(newEmail)) {
      row[GET_MAIL_CONFIG.COL.NEXT_ACTION] = "重複";
      console.log(`[過去データ重複検知] ${row[2]} 様（再入会として処理）`);
    }

    // 次のループ判定のために、今回の処理済みリストに追加
    if (newPhoneClean) batchPhones.add(newPhoneClean);
    if (newEmail) batchEmails.add(newEmail);
    
    finalDataToWrite.push(row);
  });

  // 書き込むべきデータがゼロなら終了
  if (finalDataToWrite.length === 0) {
    console.log("書き込み対象の新規データはありませんでした。");
    return;
  }

  // --- 3. 書き込み処理 ---
  const lastRow = getLastRowInColumn(sheet, GET_MAIL_CONFIG.COL.ID_BASE);
  const startRow = lastRow === 0 ? 1 : lastRow + 1;

  // D列用
  const dataForColD = finalDataToWrite.map(row => [`${row[0]} ${row[1]}`]);
  // F列以降用
  const dataForColF = finalDataToWrite.map((row, index) => createRowFromF(row, startRow + index));

  sheet.getRange(startRow, 4, dataForColD.length, 1).setValues(dataForColD);
  sheet.getRange(startRow, 6, dataForColF.length, dataForColF[0].length).setValues(dataForColF);

  console.log(`${finalDataToWrite.length} 件のデータを保存しました。`);
}

/**
 * F列以降の書き込み用配列を作成
 */
function createRowFromF(row, rowIndex) {
  const [
    date, time, name, furi, phone, email, age, 
    nextAction, // 重複フラグ
    location, remarks, subject, sender, msgId
  ] = row;

  let finalRemarks = remarks;
  if (location) finalRemarks = (finalRemarks ? finalRemarks + "\n" : "") + `都道府県：${location}`;

  // 集客経路を判別するVLOOKUP数式（F列）
  const formulaF = `=IFERROR(VLOOKUP(AV${rowIndex},'パートナーID・メディアIDリスト'!$D$1:$G$99,4,0))`;

  return [
    formulaF,       // F: 集客経路
    "",             // G: (空白)
    nextAction,     // H: ネクストアクション（重複フラグ）
    name,           // I: 求職者名
    furi,           // J: フリガナ
    phone,          // K: 電話番号
    email,          // L: メールアドレス
    age,            // M: 年齢
    "",             // N: 性別
    "",             // O: (空白)
    finalRemarks,   // P: 備考
    "未",           // Q
    "未",           // R
    ...Array(12).fill(""), // S 〜 AD (空白埋め)
    "自動・初回メール",      // AE
    ...Array(16).fill(""), // AF 〜 AU (空白埋め)
    subject,        // AV: LP/流入元
    sender,         // AW: 送信元
    msgId           // AX: MessageID
  ];
}

/**
 * 特定の列の最終行を取得
 */
function getLastRowInColumn(sheet, columnNumber) {
  const maxRows = sheet.getMaxRows();
  if (sheet.getRange(maxRows, columnNumber).getValue() !== "") return maxRows;
  const lastRow = sheet.getRange(maxRows, columnNumber).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  if (lastRow === 1 && sheet.getRange(1, columnNumber).isBlank()) return 0;
  return lastRow;
}

/* --- データ抽出用ロジック --- */

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
  const getVal = (regex) => {
    const match = body.match(regex);
    return match ? match[1].trim() : "";
  };
  
  const rawPhone = getVal(/【\s*電話番号\s*】\s*(\S+)/) || getVal(/【 お電話番号 】\s*(\S+)/);
  
  return [
    Utilities.formatDate(message.getDate(), "JST", "yyyy-MM-dd"),
    Utilities.formatDate(message.getDate(), "JST", "HH:mm"),
    getVal(patterns.name),
    getVal(patterns.furi),
    rawPhone ? `=TEXT("${rawPhone}","0##########")` : "",
    getVal(/【\s*メールアドレス\s*】\s*(\S+)/),
    getVal(/【 年齢 】\s*(\S+)/).replace("歳", ""),
    "", // 重複フラグ用プレースホルダー
    getVal(/【\s*希望勤務地\s*】\s*(\S+)/),
    [getVal(/【\s*最終学歴\s*】\s*(\S+)/), getVal(/お問い合わせID(?:（識別番号）)?\s*[:：]\s*(\S+)/)].filter(Boolean).join("\n"),
    message.getSubject(),
    getVal(/送信元：\s*(\S+)/),
    message.getId()
  ];
}