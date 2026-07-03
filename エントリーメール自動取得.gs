//20260625 O列カラム用途変更の旨を記載
//20260520 同じ処理回に同じメールが複数件あった時→1メールずつ処理する際、重複メールを処理しないようにそのタイミングで既読化
//20260518 既読化先行によるデータロス対策: スプシ書込後に既読化・アーカイブする順序に変更
//20260424AKAHONの検索クエリを追加
//20260415その回の重複なのか、以前の重複なのかで扱いを変える
//20260415重複検索範囲を全ての行に拡大


/**
 * ==========================================
 * 設定エリア
 * ==========================================
 */
const GET_MAIL_CONFIG = {
  IS_TEST_MODE: false,
  SHEET_NAME: "求職者管理",

  SEARCH: {
    AGENT_NAVI: 'is:unread from:agentnavi@circus-group.jp subject:"[自動送信 | 転職エージェントナビ ]" AND "自動送信" newer_than:2d',
    FB_LEAD: 'is:unread subject:FBリードからお問い合わせがありました newer_than:1d',
    AKAHON: 'is:unread from:akahonmaster@jobakahon.com subject:"【転職アカホン】転職者情報をお送りします" newer_than:7d'
  },

  COL: {
    NEXT_ACTION: 7,
    PHONE: 10,
    EMAIL: 11,
    ID_BASE: 4
  }
};

/**
 * メイン実行関数
 */
function main_getEntryMail() {
  const modeName = GET_MAIL_CONFIG.IS_TEST_MODE ? "【テストモード】" : "【本番モード】";
  console.log(`=== 処理開始 ${modeName} ===`);

  console.log("--- AgentNavi 取得中 ---");
  processEmailGroup(GET_MAIL_CONFIG.SEARCH.AGENT_NAVI, extractAgentNaviData);

  console.log("--- FB Lead 取得中 ---");
  processEmailGroup(GET_MAIL_CONFIG.SEARCH.FB_LEAD, extractFbLeadData);

  console.log("--- Akahon 取得中 ---");
  processEmailGroup(GET_MAIL_CONFIG.SEARCH.AKAHON, extractAkahonData);

  console.log("=== 全工程 終了 ===");
}

/**
 * メールの検索・解析・保存を統括する処理
 * 順序: 抽出 → スプシ書込 → 既読化/アーカイブ
 * 理由: 既読化先行だと スプシ書込前のエラーでメールが消失する
 */
function processEmailGroup(query, extractionCallback) {
  const threads = GmailApp.search(query, 0, 500);
  if (threads.length === 0) {
    console.log("対象のメールは見つかりませんでした。");
    return;
  }

  // --- ① 抽出のみ。Gmail 副作用は実行しない ---
  const messageRefs = [];

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      if (message.isUnread()) {
        const extractedData = extractionCallback(message);
        if (extractedData) {
          messageRefs.push({ row: extractedData, message: message, thread: thread });
        }
      }
    });
  });

  if (messageRefs.length === 0) {
    console.log("対象の未読メッセージはありませんでした。");
    return;
  }

  // --- ② スプシ書込（戻り値: 実際に書き込めた row 配列） ---
  // --- ② スプシ書込（戻り値: 書込 row + バッチ重複スキップ row） ---
  const newData = messageRefs.map(ref => ref.row);
  const result = saveEntryDataToSheet(newData);
  const writtenRows = result.written || [];
  const skippedRows = result.skippedDuplicates || [];

  if (writtenRows.length === 0 && skippedRows.length === 0) {
    console.log("処理対象ゼロ。Gmail 副作用もスキップ。");
    return;
  }

  // --- ③ 書込分 + バッチ重複スキップ分 を 既読化・アーカイブ ---
  if (GET_MAIL_CONFIG.IS_TEST_MODE) return;

  const processedRows = new Set([...writtenRows, ...skippedRows]);  // ★両方含める
  const archivedThreadIds = new Set();

  messageRefs.forEach(ref => {
    if (!processedRows.has(ref.row)) return; // 処理されなかった row だけ除外
    try {
      ref.message.markRead();
    } catch (e) {
      console.warn(`markRead 失敗: ${e}`);
    }

    const tid = ref.thread.getId();
    if (!archivedThreadIds.has(tid)) {
      try {
        ref.thread.moveToArchive();
        archivedThreadIds.add(tid);
      } catch (e) {
        console.warn(`moveToArchive 失敗: ${e}`);
      }
    }
  });
}

/**
 * データをシートに書き込む
 * ロジック:
 * 1. 同一実行内の重複（バグ） -> シートに書き込まずスキップ
 * 2. 過去データとの重複（再入会） -> H列に「重複」と記載して書き込む
 * 戻り値: 実際に書き込んだ row 配列（呼出元が markRead 対象判定に使う）
 */
function saveEntryDataToSheet(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GET_MAIL_CONFIG.SHEET_NAME);
  if (!sheet) {
    console.error(`エラー: シート「${GET_MAIL_CONFIG.SHEET_NAME}」が見つかりません。`);
    return [];
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
  const skippedDuplicates = [];  // ★追加: バッチ重複でスキップした row（既読化対象）

  data.forEach(row => {
    const newPhoneClean = String(row[4] || "").replace(/[^0-9]/g, "");
    const newEmail = String(row[5] || "").trim().toLowerCase();

    // A. 同一バッチ内の重複チェック（バグ対策：書き込まない）
    if (batchPhones.has(newPhoneClean) || batchEmails.has(newEmail)) {
      console.log(`[同一バッチ重複スキップ] ${row[2]} 様（システム重複のため除外）`);
      skippedDuplicates.push(row);  // ★追加: 書込しないがGmail既読化はする
      return;
    }

    // B. 過去データとの重複チェック（再入会対策：書き込むが「重複」ラベル付与）
    if (historicalPhones.has(newPhoneClean) || historicalEmails.has(newEmail)) {
      row[GET_MAIL_CONFIG.COL.NEXT_ACTION] = "重複";
      console.log(`[過去データ重複検知] ${row[2]} 様（再入会として処理）`);
    }

    if (newPhoneClean) batchPhones.add(newPhoneClean);
    if (newEmail) batchEmails.add(newEmail);

    finalDataToWrite.push(row);
  });

  if (finalDataToWrite.length === 0) {
    console.log("書き込み対象の新規データはありませんでした。");
    return { written: [], skippedDuplicates: skippedDuplicates };
  }

  // --- 3. 書き込み処理 ---
  const lastRow = getLastRowInColumn(sheet, GET_MAIL_CONFIG.COL.ID_BASE);
  const startRow = lastRow === 0 ? 1 : lastRow + 1;

  const dataForColD = finalDataToWrite.map(row => [`${row[0]} ${row[1]}`]);
  const dataForColF = finalDataToWrite.map((row, index) => createRowFromF(row, startRow + index));

  sheet.getRange(startRow, 4, dataForColD.length, 1).setValues(dataForColD);
  sheet.getRange(startRow, 6, dataForColF.length, dataForColF[0].length).setValues(dataForColF);

  console.log(`${finalDataToWrite.length} 件のデータを保存しました。`);
  return { written: finalDataToWrite, skippedDuplicates: skippedDuplicates };
}

/**
 * F列以降の書き込み用配列を作成
 * ※AX列(msgId)書込みは廃止（既存別データとの衝突回避）
 */
function createRowFromF(row, rowIndex) {
  const [
    date, time, name, furi, phone, email, age,
    nextAction,
    location, remarks, subject, sender, msgId
  ] = row;

  let finalRemarks = remarks;
  if (location) finalRemarks = (finalRemarks ? finalRemarks + "\n" : "") + `都道府県：${location}`;

  const formulaF = `=IFERROR(VLOOKUP(AV${rowIndex},'パートナーID・メディアIDリスト'!$D$1:$G$99,4,0))`;

  return [
    formulaF,       // F: 集客経路
    "",             // G
    nextAction,     // H: 重複フラグ
    name,           // I
    furi,           // J
    phone,          // K
    email,          // L
    age,            // M
    "",             // N: 性別
    "未",            // O用途変更に伴い
    finalRemarks,   // P
    "未",           // Q
    "未",           // R
    ...Array(12).fill(""), // S 〜 AD
    "自動・初回メール",      // AE
    ...Array(16).fill(""), // AF 〜 AU
    subject,        // AV
    sender          // AW
    // AX: MessageID は廃止（既存別データと衝突）
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

function extractAkahonData(message) {
  const body = toPlainText_(message);
  const getVal = (regex) => {
    const match = body.match(regex);
    return match ? match[1].trim() : "";
  };

  const rawPhone = getVal(/■電話番号：\s*([\d\-()]+)/);
  const birthday = getVal(/■生年月日：\s*([0-9\/]+)/);

  return [
    Utilities.formatDate(message.getDate(), "JST", "yyyy-MM-dd"),
    Utilities.formatDate(message.getDate(), "JST", "HH:mm"),
    getVal(/■お名前：\s*([\s\S]+?)\r?\n/),
    getVal(/■ふりがな：\s*([\s\S]+?)\r?\n/),
    rawPhone ? `=TEXT("${rawPhone}","0##########")` : "",
    getVal(/■メールアドレス：\s*(\S+)/),
    calculateAgeFromBirthday_(birthday),
    "",
    getVal(/■転職希望勤務地：\s*([\s\S]+?)\r?\n/),
    [
      formatAkahonRemarks_(body),
      birthday ? `生年月日：${birthday}` : "",
      getVal(/■性別：\s*([\s\S]+?)\r?\n/) ? `性別：${getVal(/■性別：\s*([\s\S]+?)\r?\n/)}` : "",
      getVal(/最終学歴：\s*([\s\S]+?)\r?\n/) ? `最終学歴：${getVal(/最終学歴：\s*([\s\S]+?)\r?\n/)}` : "",
      getVal(/■転職希望年収：\s*([\s\S]+?)\r?\n/) ? `転職希望年収：${getVal(/■転職希望年収：\s*([\s\S]+?)\r?\n/)}` : "",
      getVal(/■転職経験回数：\s*([\s\S]+?)\r?\n/) ? `転職経験回数：${getVal(/■転職経験回数：\s*([\s\S]+?)\r?\n/)}` : ""
    ].filter(Boolean).join("\n"),
    message.getSubject(),
    "akahonmaster@jobakahon.com",
    message.getId()
  ];
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
    "",
    getVal(/【\s*希望勤務地\s*】\s*(\S+)/),
    [getVal(/【\s*最終学歴\s*】\s*(\S+)/), getVal(/お問い合わせID(?:（識別番号）)?\s*[:：]\s*(\S+)/)].filter(Boolean).join("\n"),
    message.getSubject(),
    getVal(/送信元：\s*(\S+)/),
    message.getId()
  ];
}

function formatAkahonRemarks_(body) {
  const address = (body.match(/■住所：\s*([\s\S]+?)\r?\n■電話番号：/) || [])[1] || "";
  return address.trim() ? `住所：${address.trim()}` : "";
}

function calculateAgeFromBirthday_(birthday) {
  if (!birthday) return "";

  const normalized = birthday.replace(/\./g, "/").trim();
  const parts = normalized.split("/");
  if (parts.length !== 3) return "";

  const year = Number(parts[0]);
  const month = Number(parts[1]);
  const day = Number(parts[2]);
  if (!year || !month || !day) return "";

  const today = new Date();
  let age = today.getFullYear() - year;
  const thisYearBirthday = new Date(today.getFullYear(), month - 1, day);
  if (today < thisYearBirthday) age -= 1;
  return age > 0 ? String(age) : "";
}

function toPlainText_(message) {
  const plain = message.getPlainBody();
  if (plain) return plain;
  return message.getBody()
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ");
}