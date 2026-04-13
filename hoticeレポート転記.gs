/**
 * 定期実行用のメイン関数
 */
function main_aggregateJobSeekerData() {
  // --- 設定項目 ---
  const SOURCE_SHEET_NAME = "求職者管理";
  const DEST_SS_ID = "1rEMhUh1qfu2JJVhvo_F5vAlgzmLK9hS1YmujY04KNCw"; // 転記先スプレッドシートID
  const DEST_SHEET_NAME_CV = "実CV";
  const DEST_SHEET_NAME_INTERVIEW = "実CV（面談予約）";

  // 日付の設定（実行日の前日を取得）
  const now = new Date();
  const yesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
  const yesterdayStr = Utilities.formatDate(yesterday, "JST", "yyyy-MM-dd");

  // --- 1. 元データ（求職者管理）の取得とフィルタリング ---
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  const data = sourceSheet.getDataRange().getValues();
  
  // 見出しが3行目（インデックス2）にあるとのことなので、データは4行目（インデックス3）から
  const rows = data.slice(3);

  let countMale = 0;
  let countFemale = 0;
  let countOtherGender = 0;
  let countInterviewReserved = 0;

  rows.forEach(row => {
    const entryDateValue = row[3]; // D列: エントリー日
    const channel = row[5];        // F列: 集客経路
    const nextAction = row[7];     // H列: ネクストアクション
    const gender = row[13];        // N列: 性別
    const interviewDate1 = row[18]; // S列: 個別面談日時
    const interviewDate2 = row[24]; // Y列: ナビ面談日時

    // 日付形式のチェックと変換
    if (!entryDateValue) return;
    const entryDate = new Date(entryDateValue);
    const entryDateStr = Utilities.formatDate(entryDate, "JST", "yyyy-MM-dd");

    // 条件判定: 
    // ①エントリー日が昨日 
    // ②集客経路が "hotice" 
    // ③ネクストアクションが "重複" 以外
    if (entryDateStr === yesterdayStr && channel === "hotice" && nextAction !== "重複") {
      
      // 性別の集計
      if (gender === "男性") {
        countMale++;
      } else if (gender === "女性") {
        countFemale++;
      } else {
        countOtherGender++;
      }

      // 面談予約の集計（S列またはY列に値がある場合）
      if (interviewDate1 || interviewDate2) {
        countInterviewReserved++;
      }
    }
  });
  console.log(`集計が完了しました\n男${countMale} 女${countFemale} 不明${countOtherGender} 面談予約${countInterviewReserved}`)
  // --- 2. 転記先スプレッドシートへの書き込み ---
  const destSs = SpreadsheetApp.openById(DEST_SS_ID);
  
  // 「実CV」シートへの書き込み
  updateDestSheet(destSs.getSheetByName(DEST_SHEET_NAME_CV), yesterday, [countMale, countFemale, countOtherGender], [2, 3, 4]);

  // 「実CV（面談予約）」シートへの書き込み
  updateDestSheet(destSs.getSheetByName(DEST_SHEET_NAME_INTERVIEW), yesterday, [countInterviewReserved], [3]);

  console.log(`処理完了: ${yesterdayStr} 分のデータを転記しました。`);
}

/**
 * 指定したシートの該当日付の行にデータを書き込む補助関数
 * @param {Sheet} sheet 対象シート
 * @param {Date} targetDate 昨日の日付
 * @param {Array} values 書き込む値の配列
 * @param {Array} columns 書き込む列番号の配列
 */
function updateDestSheet(sheet, targetDate, values, columns) {
  const data = sheet.getDataRange().getValues();
  const targetDateStr = Utilities.formatDate(targetDate, "JST", "yyyy/MM/dd"); // 転記先の形式に合わせて調整

  for (let i = 0; i < data.length; i++) {
    let rowDate = data[i][0]; // A列
    if (rowDate instanceof Date) {
      const rowDateStr = Utilities.formatDate(rowDate, "JST", "yyyy/MM/dd");
      if (rowDateStr === targetDateStr) {
        // 見つかった行の指定列に値を書き込む
        values.forEach((val, index) => {
          sheet.getRange(i + 1, columns[index]).setValue(val);
          console.log(`シート名${sheet.getName()}${i+1}行目の${columns[index]}列に${val}を追加しました`)
        });
        return;
      }
    }
  }
  console.warn(`${sheet.getName()} シートに ${targetDateStr} の行が見つかりませんでした。`);
}