/**
 * ==========================================
 * 設定エリア
 * ==========================================
 */
const GET_MAIL_CONFIG = {
  // ★ true: テストモード（既読にしない・アーカイブしない）
  // ★ false: 本番モード（処理後に既読化・アーカイブする）
  IS_TEST_MODE: false, 

  SHEET_NAME: "求職者管理", 
  
  SEARCH: {
    AGENT_NAVI: 'is:unread from:agentnavi@circus-group.jp subject:"[自動送信 | 転職エージェントナビ ]" AND "自動送信" newer_than:2d',
    FB_LEAD: 'is:unread subject:FBリードからお問い合わせがありました newer_than:1d'
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
    // 関数名を変更して呼び出し
    saveEntryDataToSheet(newData);
  } else {
    console.log("抽出対象のデータはありませんでした。");
  }
}

/**
 * 「シート7」に書き込む関数
 * 機能: 重複チェック機能付き（直前のデータと同じ電話番号ならスキップ）
 */
function saveEntryDataToSheet(data) {
  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GET_MAIL_CONFIG.SHEET_NAME);
  
  if (!targetSheet) {
    console.error(`エラー: シート「${GET_MAIL_CONFIG.SHEET_NAME}」が見つかりません。`);
    return;
  }

  // --- ▼▼▼ 重複チェックロジック開始 ▼▼▼ ---
  
  // 1. シート上の最新の電話番号を取得（K列=11列目と仮定）
  const lastRowK = getLastRowInColumn(targetSheet, 11); 
  let lastPhoneOnSheet = "";
  if (lastRowK > 0) {
    // 数式(=TEXT...)が入っている可能性があるため getFormula と getValue 両方を考慮
    const cell = targetSheet.getRange(lastRowK, 11);
    lastPhoneOnSheet = cell.getFormula() || cell.getValue();
    console.log("最終行の電話番号は"+lastPhoneOnSheet)
  }

  const filteredData = [];
  // 比較用に「直前の有効な電話番号」を保持する変数（初期値はシートの最終行）
  let previousPhone = lastPhoneOnSheet;

  data.forEach(row => {
    // parseEmailBodyの返り値配列で、電話番号は index 4
    const currentPhone = row[4]; 

    // 数字だけで比較して、同じならスキップ
    if (isSamePhone(previousPhone, currentPhone)) {
      console.log(`[重複スキップ] 電話番号が直前のデータと一致したため除外しました: ${row[2]} (${currentPhone})`);
    } else {
      // 重複していない場合のみリストに追加し、比較対象を更新
      filteredData.push(row);
      previousPhone = currentPhone;
    }
  });

  // 重複を除いた結果、書き込むデータがなくなったら終了
  if (filteredData.length === 0) {
    console.log("重複を除去した結果、追加すべきデータはありませんでした。");
    return;
  }
  
  // 書き込み対象をフィルタ後のデータに差し替え
  const dataToWrite = filteredData;
  // --- ▲▲▲ 重複チェックロジック終了 ▲▲▲ ---


  // D列(4列目)を見て書き込み開始行を判定
  const lastRowWithData = getLastRowInColumn(targetSheet, 4); 
  const startRow = lastRowWithData === 0 ? 1 : lastRowWithData + 1;

  // 1. 【D列】のデータを作成（日付＆時間）
  const dataForColD = dataToWrite.map(row => [`${row[0]} ${row[1]}`]);

  // 2. 【F列以降】のデータを作成
  //    行番号が必要なため、map内で計算
  const dataForColF_Onwards = dataToWrite.map((row, index) => {
    return createRowFromF(row, startRow + index);
  });

  // 書き込み実行
  // (1) D列 (4列目) に書き込み
  targetSheet.getRange(startRow, 4, dataForColD.length, 1).setValues(dataForColD);
  
  // (2) F列 (6列目) 以降に書き込み
  targetSheet.getRange(startRow, 6, dataForColF_Onwards.length, dataForColF_Onwards[0].length).setValues(dataForColF_Onwards);

  console.log(`「${GET_MAIL_CONFIG.SHEET_NAME}」の ${startRow} 行目に ${dataToWrite.length} 件追加しました。（A-C, E列は保持）`);
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

  // S列(19)〜AD列(30)の間を埋める空白 (S〜AD = 12列分)
  const gapColumns = Array(12).fill(""); 
  // AF列(32)〜AU列(47)の間を埋める空白 (AF〜AU = 16列分)
  const gapColumns2 = Array(16).fill(""); 

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
    "未",                //Q:未
    "未",                //R:未
    ...gapColumns,       // S 〜 AD: 空白埋め
    "自動・初回メール",    //AE:自動・初回メール
    ...gapColumns2,       // AF 〜 AU: 空白埋め
    subject,             // AV: LP/流入元
    sender,              // AW: 送信元
    msgId                // AX: MessageID
  ];
}

/**
 * 指定列の最終データ行を取得（Direction版）
 */
function getLastRowInColumn(sheet, columnNumber) {
  // シートの最大行数を取得
  const maxRows = sheet.getMaxRows();
  
  // その列の最終行が既に埋まっている場合のガード（稀なケースですが念のため）
  if (sheet.getRange(maxRows, columnNumber).getValue() !== "") {
    return maxRows;
  }

  // 一番下の行から上に向かって検索 (Ctrl + ↑)
  const lastRow = sheet.getRange(maxRows, columnNumber)
                       .getNextDataCell(SpreadsheetApp.Direction.UP)
                       .getRow();

  // データが1つもない場合、Direction.UPは「1」を返す仕様があるため、
  // 1行目が本当にデータ入りかチェックして、空白なら0を返す
  if (lastRow === 1 && sheet.getRange(1, columnNumber).isBlank()) {
    return 0;
  }

  return lastRow;
}

/**
 * 電話番号比較用ヘルパー関数
 * 数式(=TEXT...)やハイフン等の違いを無視して、数字のみで一致判定を行う
 */
function isSamePhone(val1, val2) {
  console.log(`【${val1}】と【${val2}】を比較`)
  if (!val1 || !val2) return false;
  // 文字列化して、数字以外を全て削除
  const num1 = String(val1).replace(/[^0-9]/g, "");
  const num2 = String(val2).replace(/[^0-9]/g, "");
  
  return num1 !== "" && num1 === num2;
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