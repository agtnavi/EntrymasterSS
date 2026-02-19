/**
 * ナビ担当者とメールアドレスの紐付け
 */
const NAVI_EMAIL_MAP = {
  '菅谷': 'naomisugaya721@gmail.com',
  '小木曽': 'yuki9517@gmail.com',
  '佐藤': 'nishi.shima217@gmail.com', // 「佐藤し」も「佐藤」で判定
  '浅富': 'naomi600143@gmail.com',
  '中村': 'hitomi10321107@gmail.com',
  '千田': 'shuntaro8710@gmail.com',
  '瀬戸': 'k-seto@circus-group.jp'
};

const CONFIG_NAVI = {
  SHEET_NAME: '📅面談予定',
  START_ROW: 76, // データ開始行
  COLUMN_INDICES: {
    NAVI_NAME: 12,      // M列: 明日のナビ面談 (【瀬戸】など)
    INTERVIEW_TIME: 13, // N列: 時間 (日付+時刻)
    CANDIDATE_ID: 14,   // O列: 求職者No
    CANDIDATE_NAME: 15, // P列: 氏名
    PHONE_NUMBER: 17,   // R列: 電話番号
    EMAIL: 18,          // S列: メールアドレス
    INTERVIEW_TYPE: 19  // T列: 面談種類
  }
};

const CC_ADDRESS = "agentnavi@circus-group.jp";

/**
 * メイン処理
 */
function main_sendNaviRemindEmails() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_NAVI.SHEET_NAME);
  
  if (!sheet) {
    ui.alert(`シート「${CONFIG_NAVI.SHEET_NAME}」が見つかりません。`);
    return;
  }

  // 明日の日付を取得
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowStr = Utilities.formatDate(tomorrow, 'JST', 'M/d');
  const tomorrowCompare = Utilities.formatDate(tomorrow, 'JST', 'yyyy/MM/dd');

  // データ取得
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG_NAVI.START_ROW) {
    ui.alert('データが存在しません。');
    return;
  }
  
  const data = sheet.getRange(CONFIG_NAVI.START_ROW, 1, lastRow - CONFIG_NAVI.START_ROW + 1, 21).getValues();
  const naviDataMap = {};

  // データ抽出
  data.forEach((row) => {
    const timeValue = row[CONFIG_NAVI.COLUMN_INDICES.INTERVIEW_TIME];
    if (!(timeValue instanceof Date)) return;

    // 日付が明日か判定
    const rowDate = Utilities.formatDate(timeValue, 'JST', 'yyyy/MM/dd');
    if (rowDate !== tomorrowCompare) return;

    // 担当者名を抽出 (例: 【瀬戸】 -> 瀬戸)
    let rawNaviName = row[CONFIG_NAVI.COLUMN_INDICES.NAVI_NAME].toString();
    let naviKey = rawNaviName.replace(/[【】]/g, '').replace(/し$/, ''); // 「佐藤し」対策

    if (!naviKey) return;

    if (!naviDataMap[naviKey]) {
      naviDataMap[naviKey] = [];
    }

    const timeStr = Utilities.formatDate(timeValue, 'JST', 'HH:mm');
    const candidateName = row[CONFIG_NAVI.COLUMN_INDICES.CANDIDATE_NAME];
    const candidateId = row[CONFIG_NAVI.COLUMN_INDICES.CANDIDATE_ID];
    const phone = row[CONFIG_NAVI.COLUMN_INDICES.PHONE_NUMBER];
    const email = row[CONFIG_NAVI.COLUMN_INDICES.EMAIL];
    const type = row[CONFIG_NAVI.COLUMN_INDICES.INTERVIEW_TYPE];

    naviDataMap[naviKey].push({
      time: timeStr,
      candidateName: candidateName,
      candidateId: candidateId,
      phone: phone,
      email: email,
      type: type
    });
  });

  // メール生成と送信
  const results = [];
  for (const naviName in naviDataMap) {
    const recipient = NAVI_EMAIL_MAP[naviName];
    if (!recipient) {
      results.push(`【送信不可】${naviName} さんのアドレスが未登録です。`);
      continue;
    }

    // 面談リストの作成
    let interviewListText = "";
    naviDataMap[naviName].forEach(item => {
      interviewListText += `${item.time}〜\n`;
      interviewListText += `${item.candidateId} ${item.candidateName} 様\n`;
      interviewListText += `${item.phone}  ${item.email}  ${item.type}\n\n`;
    });

    const subject = `明日の面談リマインドのご連絡`;
    const body = `お疲れ様です。
現時点での明日(${tomorrowStr})のナビ・個別面談のリマインドを送信いたします。

${interviewListText.trim()}

変更・追加があればLINEにてご連絡いたします。
宜しくお願いします！`;

    // 送信確認
    const confirm = ui.alert(
      `${naviName}様へ送信確認`,
      `宛先: ${recipient}\n\n${body}`,
      ui.ButtonSet.YES_NO
    );

    if (confirm === ui.Button.YES) {
      try {
        // GmailApp.sendEmail(recipient, subject, body, {
        //   name: "転職エージェントナビ事務局",
        //   cc: CC_ADDRESS
        // });
        results.push(`【送信完了】${naviName}様 (${recipient})`);
      } catch (e) {
        results.push(`【エラー】${naviName}様: ${e.message}`);
      }
    } else {
      results.push(`【スキップ】${naviName}様`);
    }
  }

  if (results.length > 0) {
    ui.alert("処理結果:\n\n" + results.join("\n"));
  } else {
    ui.alert("明日の面談予定が見つかりませんでした。");
  }
}