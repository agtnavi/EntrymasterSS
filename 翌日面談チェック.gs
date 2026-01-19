
// ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
// ★ 宛先リストが記載されているスプレッドシートの情報を設定します
// ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
const MAPPING_SHEET_CONFIG = {
  ID: '1aoSL_C-nmYdFJN7tysdREva81ByuXN3WX1rGdXfN_kM', // 宛先リストのスプレッドシートID
  SHEET_NAME: '新規既存結合',                        // シート名
  COMPANY_COL_INDEX: 3,                           // 会社名が入力されている列 (D列)
  EMAIL_COL_INDEX: 2                              // メールアドレスが入力されている列 (C列)
};

const CONFIG = {
  SHEET_NAME: '翌日AGT面談リマインド', // 面談データが入力されているシート名
  COLUMN_INDICES: {
    COMPANY_NAME: 9,      // J列: AGT名(株抜き)
    CONTACT_PERSON: 8,    // B列: ご担当者名 -> 260119　B列からI列に変更
    INTERVIEW_DATE: 2,    // C列: 面談予定日
    CANDIDATE_ID: 3,      // D列: 求職者No.
    CANDIDATE_NAME: 4,    // E列: 求職者名
    INTERVIEW_TIME: 5,    // F列: 時間
    NO_RESPONSE_FLAG: 6,  // G列: 「済」なら反応あり、それ以外なら「※」
  }
};

var cc = "agentnavi@circus-group.jp"

/**
 * 設定用の別シートから、会社名とメールアドレスの対応リストを取得する
 */
function getEmailMapFromSheet() {
  console.log('【宛先取得】宛先リストの読み込みを開始します...');
  try {
    const spreadsheet = SpreadsheetApp.openById(MAPPING_SHEET_CONFIG.ID);
    const sheet = spreadsheet.getSheetByName(MAPPING_SHEET_CONFIG.SHEET_NAME);
    if (!sheet) {
      console.error(`【エラー】シート「${MAPPING_SHEET_CONFIG.SHEET_NAME}」が見つかりません。`);
      SpreadsheetApp.getUi().alert(`宛先リストのシート「${MAPPING_SHEET_CONFIG.SHEET_NAME}」が見つかりませんでした。`);
      return null;
    }
    const data = sheet.getDataRange().getValues().slice(1);
    
    const emailMap = {};
    let count = 0;
    data.forEach(row => {
      const companyName = row[MAPPING_SHEET_CONFIG.COMPANY_COL_INDEX];
      const email = row[MAPPING_SHEET_CONFIG.EMAIL_COL_INDEX];

      if (companyName && email && typeof email === 'string' && email.includes('@')) {
        if (!emailMap[companyName]) {
          emailMap[companyName] = [];
        }
        emailMap[companyName].push(email);
        count++;
      }
    });
    console.log(`【宛先取得完了】有効な宛先データを ${count} 件（会社数: ${Object.keys(emailMap).length}社）読み込みました。`);
    return emailMap;
  } catch (e) {
    console.error(`【エラー】宛先取得中に例外が発生: ${e.message}`);
    SpreadsheetApp.getUi().alert(`宛先リストのスプレッドシートを開けませんでした。\nエラー: ${e.message}`);
    return null;
  }
}

function addSama(name) {
  if (!name || typeof name !== 'string') return '';
  const trimmedName = name.trim();
  if (trimmedName.endsWith('様') || trimmedName.endsWith('各位')) {
    return trimmedName;
  }
  return trimmedName + '様';
}

/**
 * メイン処理
 */
function generateCompanySpecificEmailsWithConfirmation() {
  console.log('▼▼▼ 処理開始 ▼▼▼');
  const ui = SpreadsheetApp.getUi();

  // 1. 宛先マップ取得
  const emailMap = getEmailMapFromSheet();
  if (!emailMap) return;

  // 2. データシート取得
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    console.error(`【エラー】シート「${CONFIG.SHEET_NAME}」が見つかりません。`);
    ui.alert(`シート「${CONFIG.SHEET_NAME}」が見つかりません。`);
    return;
  }

  // 3. データ読み込み
  const data = sheet.getDataRange().getValues().slice(1);
  console.log(`【データ読込】対象シートから ${data.length} 行のデータを取得しました。`);

  // 4. 日付設定
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  tomorrow.setHours(0, 0, 0, 0);
  console.log(`【日付設定】抽出対象日(明日): ${Utilities.formatDate(tomorrow, 'JST', 'yyyy/MM/dd')}`);

  const companiesData = {};
  let targetCount = 0;

  // 5. データ抽出ループ
  data.forEach((row, index) => {
    const rowNum = index + 2; // スプレッドシート上の行番号（ヘッダー分+1、0始まり補正+1）
    const interviewDateValue = row[CONFIG.COLUMN_INDICES.INTERVIEW_DATE];
    
    // 日付が入っていない行はスキップ（ログには出さない）
    if (!interviewDateValue || !(interviewDateValue instanceof Date)) return;

    const interviewDate = new Date(interviewDateValue);
    interviewDate.setHours(0, 0, 0, 0);

    // 日付判定
    if (interviewDate.getTime() === tomorrow.getTime()) {
      const companyName = row[CONFIG.COLUMN_INDICES.COMPANY_NAME];
      let contactPerson = row[CONFIG.COLUMN_INDICES.CONTACT_PERSON];
      
      // 必須データ欠けチェック
      if (!companyName || !contactPerson) {
        console.warn(`【スキップ】行${rowNum}: 日付は対象ですが、会社名または担当者名が不足しています。`);
        return;
      }

      // データの格納準備
      if (!companiesData[companyName]) {
        companiesData[companyName] = { interviewsByContact: {} };
      }
      
      contactPerson = contactPerson.toString().trim().replace(/様$/, '');

      if (!companiesData[companyName].interviewsByContact[contactPerson]) {
        companiesData[companyName].interviewsByContact[contactPerson] = [];
      }

      const timeValue = row[CONFIG.COLUMN_INDICES.INTERVIEW_TIME];
      const time = (timeValue instanceof Date) ? Utilities.formatDate(timeValue, 'JST', 'HH:mm') : timeValue;
      const candidateName = row[CONFIG.COLUMN_INDICES.CANDIDATE_NAME];
      const candidateId = row[CONFIG.COLUMN_INDICES.CANDIDATE_ID];
      
      // レスポンス状況確認
      const responseStatus = row[CONFIG.COLUMN_INDICES.NO_RESPONSE_FLAG];
      const note = (responseStatus === '済') ? '' : '※';
      
      // 詳細ログ出力
      console.log(`【抽出】行${rowNum}: [${companyName}] ${contactPerson}様 / 候補者:${candidateName} (${responseStatus || '未'})`);

      const interviewDetail = `${time} ${addSama(candidateName)} ${candidateId} ${note}`.trim();
      companiesData[companyName].interviewsByContact[contactPerson].push(interviewDetail);
      targetCount++;
    }
  });

  console.log(`【抽出完了】明日の面談は合計 ${targetCount} 件、${Object.keys(companiesData).length} 社分見つかりました。`);

  // 6. メール生成と送信確認ループ
  const sentSummaries = [];
  
  for (const companyName in companiesData) {
    console.log(`--- [${companyName}] メール生成プロセス開始 ---`);
    
    let interviewList = '';
    let containsNote = false;
    const interviewsByContact = companiesData[companyName].interviewsByContact;
    
    // 本文生成ロジック
    if (companyName === 'CDC') {
      const allInterviews = Object.values(interviewsByContact).flat();
      interviewList = allInterviews.join('\n');
      if (interviewList.includes('※')) containsNote = true;
    } else {
      for (const contactName in interviewsByContact) {
        interviewList += `\n■${addSama(contactName)}\n`;
        const details = interviewsByContact[contactName].join('\n');
        interviewList += details + '\n';
        if (details.includes('※')) containsNote = true;
      }
      interviewList = interviewList.trim();
    }

    const footerNote = containsNote ? 
      `\nなお「※」の表記が付いている候補者様につきましては、
本日お送りしたリマインドメッセージへの反応が無く、面談へご参加されない可能性がございます。
ただ候補者様へ面談のご案内自体は行っているため、恐れ入りますがご入室を頂きます様、` : '';

    const message = `${addSama(companyName)}
ご担当者 各位

いつもお世話になっております。
転職エージェントナビでございます。

以下、現時点で明日お願いしている面談案件一覧をお送りいたします。

${interviewList}
${footerNote}
引き続き何卒よろしくお願いいたします。

転職エージェントナビ`;
    
    const recipientList = emailMap[companyName];
    const recipients = recipientList ? recipientList.join(',') : '';

    if (recipients) {
      console.log(`【宛先確認】To: ${recipients}`);
      const subject = `【転職エージェントナビ】明日の面談のご連絡`;
      
      const response = ui.alert('メール送信の最終確認', `以下のメールを送信しますか？\n\n------------------------------------------------------\n■宛先 (To):\n${recipients}\n■宛先 (CC):\n${cc}}\n\n■件名 (Subject):\n${subject}\n\n■本文プレビュー:\n${message.substring(0, 400)}...\n------------------------------------------------------`, ui.ButtonSet.YES_NO);

      if (response == ui.Button.YES) {
        try {
          var options = {
            name: "転職エージェントナビ事務局",
            cc: cc,
          }
          // ★実際に送信する場合はコメントアウトを外してください
          GmailApp.sendEmail(recipients, subject, message, options);
          
          console.log(`【送信成功】${companyName} 宛に送信しました。`);
          sentSummaries.push(`送信成功: ${companyName} -> ${recipients}`);
        } catch (e) {
          console.error(`【送信失敗】${companyName}: ${e.message}`);
          sentSummaries.push(`送信失敗: ${companyName} -> ${recipients} (エラー: ${e.message})`);
        }
      } else {
        console.log(`【送信キャンセル】ユーザー操作により ${companyName} への送信をキャンセルしました。`);
        sentSummaries.push(`送信キャンセル: ${companyName}`);
      }
    } else {
      console.warn(`【宛先不明】${companyName} は設定シートにメールアドレスが見つかりません。`);
      sentSummaries.push(`送信スキップ: ${companyName} (設定シートに宛先が見つかりません)`);
    }
  }

  const summaryText = sentSummaries.join('\n');
  if (summaryText) {
    console.log('▼▼▼ 処理完了 ▼▼▼\n' + summaryText);
    ui.alert('処理が完了しました。\n\n【送信結果】\n' + summaryText);
  } else {
    console.log('明日の面談データがなかったため、処理を終了します。');
    ui.alert('明日の面談はありませんでした。');
  }
}