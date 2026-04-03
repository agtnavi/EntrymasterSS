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
  TRIGGER_STATUS: "面談予約中",
  AGT_FOLDERS: {
    "株式会社キャリアデザインセンター": "16sHxGs0tl3_fLqyu8lBZ00B9_2GEcpST",
    "△△キャリア":      "ここにDriveフォルダIDを入力",
  },
};
// ================================================================
// トリガー：B列ステータスが「面談予約中」になったとき
// ================================================================
function onEditTrigger(e) {
  try {
    const sheet = e.source.getActiveSheet();
    if (sheet.getName() !== CFG.MENDAN_SHEET) return;//編集されたシートが送客管理かを判定
    const col = e.range.getColumn();
    const row = e.range.getRow();
    if (col !== CFG.MENDAN_COL.STATUS) return;//編集された列がBかを判定
    if (row <= 1) return;//編集された行が2行目以降かを判定
    if (String(e.value) !== CFG.TRIGGER_STATUS) return;//編集後の値が「面談予約中」かを判定
    // 二重実行防止：BL列にすでにURLが入っていればスキップ
    const pdfCell = sheet.getRange(row, CFG.MENDAN_COL.PDF_URL);
    if (pdfCell.getValue() !== "") {
      console.log(`行${row}はすでにPDF生成済みのためスキップ`);
      return;
    }
    const result = generatePdf(row);//メイン処理実行
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
// PDF生成メイン処理
// ================================================================
function generatePdf(mendanRow) {
  const ss          = SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
  const mendanSheet = ss.getSheetByName(CFG.MENDAN_SHEET);
  const rowValues   = mendanSheet.getRange(mendanRow, 1, 1, CFG.MENDAN_COL.PDF_URL).getValues()[0];
  const c             = CFG.MENDAN_COL;
  const candidateId   = String(rowValues[c.CANDIDATE_ID - 1]).trim();//ファイル名とテンプレへの書き込み値として利用
  const candidateName = String(rowValues[c.CANDIDATE_NAME - 1]).trim();//ファイル名として利用
  const agtCompany    = String(rowValues[c.AGT_COMPANY - 1]).trim();//格納先フォルダ判定に利用
  const interviewRaw  = rowValues[c.INTERVIEW_DATE - 1];//ファイル名として利用
  if (!candidateId) return { error: "求職者IDが空です" };
  if (!agtCompany)  return { error: "AGT会社名が空です" };
  const folderId = "16sHxGs0tl3_fLqyu8lBZ00B9_2GEcpST"//CFG.AGT_FOLDERS[agtCompany];一旦、全てのAGTがテストフォルダに入るように★
  if (!folderId) return { error: `AGT「${agtCompany}」のフォルダIDが未設定です` };
  const interviewDate = interviewRaw ? new Date(interviewRaw) : null;
  const dateStr       = interviewDate ? Utilities.formatDate(interviewDate, "Asia/Tokyo", "yyMMdd") : "000000";
  const fileName      = `${candidateId}_${candidateName.replace(/\s+/g, "")}_${dateStr}`;
  try {
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
    //★リリースまで列へのリンク記載は行わないmendanSheet.getRange(mendanRow, CFG.MENDAN_COL.PDF_URL).setValue(savedFile.getUrl());
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