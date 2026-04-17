//20260417　onEditTriggerを廃止しexecuteFromAppSheetを新設
//→https://script.google.com/home/projects/1p2nfN7OJ9eRt4HJAduiScAXYy_B8_kEaBQE9oIrKlwDOIx8xyTyDo-dV/edit こちらに行こう
/**
 * リリース時は★部分を修正
 */
// ================================================================
// 求職者PDF自動生成 GAS
// トリガー：B列「送客ステータス」が「面談予約中」になったとき
// ================================================================
const CFG = {
  SPREADSHEET_ID: "1yS4ua89_AxOgLh0Sf6rf8r1KxHPLXaty599ypzn3eBE",
  MENDAN_SHEET:    "送客管理",//"面談管理",
  CANDIDATE_SHEET: "求職者管理",//使ってなくね？
  TEMPLATE_SHEET:  "AGT送信用シート",
  AGT_SHEET:       "AGT法人管理", // ★追加：AGT管理用シート名
  NO_INPUT_CELL:   "C3",
  MENDAN_COL: {
    STATUS:         2,   // B: 送客ステータス ← トリガー
    CANDIDATE_NAME: 4,   // D: [自動反映]求職者氏名
    CANDIDATE_ID:   5,   // E: [自動反映]求職者ID
    AGT_COMPANY:    7,   // G: AGT会社
    INTERVIEW_DATE: 11,  // K: 面談予定日
    PDF_URL:        64,  // BL: 候補者情報PDF
  },
  // トリガーとなるステータス値（シートの表記と完全一致）
  TRIGGER_STATUS: "面談予約中"
  // AGT_FOLDERS は削除し、シートから動的取得に変更
};

// ▼ 新しくこれを追加します ▼
// ================================================================
// AppSheetから呼び出されるメイン処理
// ================================================================
function executeFromAppSheet(rowNumber) {
  console.log("AppSheetから起動しました。対象行: " + rowNumber);

  // AppSheetから渡された行番号を使って、そのままPDF作成処理へ投げる
  const result = generatePdf(rowNumber);

  if (!result.success) {
    // エラー時は今まで通りB列のセルにメモを追記する
    const ss = SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
    ss.getSheetByName(CFG.MENDAN_SHEET)
      .getRange(rowNumber, CFG.MENDAN_COL.STATUS)
      .setNote(`❌ PDF生成エラー: ${result.error}`);
    
    console.error("エラーが発生しました: " + result.error);
    return;
  }

  console.log("PDF生成成功: " + result.fileName);
}

// ================================================================
// 【廃止】トリガー：B列ステータスが「面談予約中」になったとき
// ================================================================
function onEditTrigger(e) {
  var source = e.source
  var range = e.range
  console.log("起動しました")
  console.log(source.getActiveSheet().getName(),range.getRow(),range.getColumn())

  try {
    const sheet = e.source.getActiveSheet();
    if (sheet.getName() !== CFG.MENDAN_SHEET) return;
    const col = e.range.getColumn();
    const row = e.range.getRow();
    if (col !== CFG.MENDAN_COL.STATUS) return;
    if (row <= 1) return;
    
    // e.valueに頼らず、直接セルの値を取得する（最強の解決策）
    const cellValue = e.range.getValue();
    console.log("入力された値: " + cellValue); // 確認用
    
    if (String(cellValue) !== CFG.TRIGGER_STATUS) return;
    // 二重実行防止：BL列にすでにURLが入っていればスキップ
    const pdfCell = sheet.getRange(row, CFG.MENDAN_COL.PDF_URL);
    if (pdfCell.getValue() !== "") {
      console.log(`行${row}はすでにPDF生成済みのためスキップ`);
      return;
    }

    const result = generatePdf(row);
    if (!result.success) {
      // エラー時はB列のセルにメモを追記
      sheet.getRange(row, CFG.MENDAN_COL.STATUS)
        .setNote(`❌ PDF生成エラー: ${result.error}`);
    }
  } catch(err) {
    console.error("onEditTrigger error:", err.message);
  }
}

// ================================================================
// AGT法人管理シートからフォルダIDを取得する関数
// ================================================================
function getAgtFolderId(agtCompany) {
  const ss = SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
  const agtSheet = ss.getSheetByName(CFG.AGT_SHEET);
  if (!agtSheet) return null;

  // データ範囲全体を取得 (ヘッダー含む)
  const data = agtSheet.getDataRange().getValues();

  // 2行目からループ (1行目はヘッダー行)
  for (let i = 1; i < data.length; i++) {
    const serviceName = String(data[i][1]).trim();  // B列: サービス名/呼称
    const officialName = String(data[i][2]).trim(); // C列: 正式名称
    const aliases = String(data[i][5]).trim();      // F列: エイリアス
    const targetFolderId = String(data[i][6]).trim(); // G列: フォルダID

    // B列 または C列 と完全一致するか確認
    if (agtCompany === serviceName || agtCompany === officialName) {
      return targetFolderId;
    }

    // F列（エイリアス）が含まれる場合、カンマ・読点・スペースなどで分割して判定
    if (aliases) {
      // 全角カンマ、読点、スペースなどを半角カンマに統一して分割
      const aliasArray = aliases.replace(/[、，\s　]+/g, ",").split(",");
      if (aliasArray.includes(agtCompany)) {
        return targetFolderId;
      }
    }
  }
  
  return null; // 見つからなかった場合
}

// ================================================================
// PDF生成メイン処理
// ================================================================
function generatePdf(mendanRow) {
    console.log("PDF作成し始めます")

  const ss          = SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
  const mendanSheet = ss.getSheetByName(CFG.MENDAN_SHEET);
  const rowValues   = mendanSheet.getRange(mendanRow, 1, 1, CFG.MENDAN_COL.PDF_URL).getValues()[0];
  const c           = CFG.MENDAN_COL;
  
  const candidateId   = String(rowValues[c.CANDIDATE_ID - 1]).trim();
  const candidateName = String(rowValues[c.CANDIDATE_NAME - 1]).trim();
  const agtCompany    = String(rowValues[c.AGT_COMPANY - 1]).trim();
  const interviewRaw  = rowValues[c.INTERVIEW_DATE - 1];

  if (!candidateId) return { error: "求職者IDが空です" };
  if (!agtCompany)  return { error: "AGT会社名が空です" };

  // ★ シートからフォルダIDを取得
  let folderId = getAgtFolderId(agtCompany);
  
  // ★ 本番リリースまでは以下の1行のコメントアウトを外して、テストフォルダに固定出力させる
  //folderId = "16sHxGs0tl3_fLqyu8lBZ00B9_2GEcpST"; 

  if (!folderId) return { error: `AGT「${agtCompany}」のフォルダIDがマスタ(AGT法人管理)に未設定、または見つかりません` };

  const interviewDate = interviewRaw ? new Date(interviewRaw) : null;
  const dateStr       = interviewDate ? Utilities.formatDate(interviewDate, "Asia/Tokyo", "yyMMdd") : "000000";
  //const fileName      = `${candidateId}_${candidateName.replace(/\s+/g, "")}_${dateStr}`;
  const fileName      = `${candidateId}_${dateStr}`;

  try {
      console.log("PDFそろそろ完成します")

    const templateSheet = ss.getSheetByName(CFG.TEMPLATE_SHEET);
    templateSheet.getRange(CFG.NO_INPUT_CELL).setValue(candidateId);
    SpreadsheetApp.flush();
    Utilities.sleep(1500);

    const pdfUrl = `https://docs.google.com/spreadsheets/d/${CFG.SPREADSHEET_ID}/export`
      + `?format=pdf&size=A4&portrait=true&fitw=true`
      + `&sheetnames=false&printtitle=false&pagenumbers=false`
      + `&gridlines=false&fzr=false&gid=${templateSheet.getSheetId()}`;

    const resp = UrlFetchApp.fetch(pdfUrl, {
      headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
      muteHttpExceptions: true,
    });

    if (resp.getResponseCode() !== 200) return { error: `PDF変換エラー（HTTP ${resp.getResponseCode()}）` };

    const savedFile = DriveApp.getFolderById(folderId).createFile(resp.getBlob().setName(`${fileName}.pdf`));
    
    //★リリースまで列へのリンク記載は行わない
     mendanSheet.getRange(mendanRow, CFG.MENDAN_COL.PDF_URL).setValue(savedFile.getUrl());
    
    return { success: true, fileName: `${fileName}.pdf`, agtCompany, candidateId, candidateName };
  } catch(err) {
    return { error: "処理中にエラーが発生しました: " + err.message };
  }
}

// ================================================================
// 手動実行（テスト・再実行用）
// ================================================================
function generatePdfManual() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== CFG.MENDAN_SHEET) {
    SpreadsheetApp.getUi().alert(`⚠️ ${CFG.MENDAN_SHEET}で実行してください。`); return;
  }
  const row = sheet.getActiveCell().getRow();
  if (row <= 1) {
    SpreadsheetApp.getUi().alert("⚠️ 2行目以降を選択してください。"); return;
  }
  const result = generatePdf(row);
  SpreadsheetApp.getUi().alert(result.success
    ? `✅ 完了！\nファイル名：${result.fileName}\n格納先：${result.agtCompany}`
    : `❌ エラー\n${result.error}`
  );
}

function showHelp() {
  SpreadsheetApp.getUi().alert(
    "【トリガー】B列「送客ステータス」を「面談予約中」にすると自動実行\n\n" +
    "【ファイル名】求職者ID_氏名_面談日(YYMMDD).pdf\n\n" +
    "【完了後】BL列にURLが自動入力されます\n\n" +
    "【二重実行防止】BL列にすでにURLがある場合はスキップされます\n\n" +
    "【再実行】BL列のURLを削除してからステータスを変更してください"
  );
}