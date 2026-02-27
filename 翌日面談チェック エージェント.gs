// ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
// ★ 設定エリア
// ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
const MAPPING_SHEET_CONFIG = {
  ID: '1aoSL_C-nmYdFJN7tysdREva81ByuXN3WX1rGdXfN_kM',
  SHEET_NAME: '新規既存結合',
  COMPANY_COL_INDEX: 3, // D列
  EMAIL_COL_INDEX: 2    // C列
};

const CONFIG = {
  SHEET_NAME: '翌日AGT面談リマインド',
  COLUMN_INDICES: {
    COMPANY_NAME: 9,      // J列
    CONTACT_PERSON: 8,    // I列
    INTERVIEW_DATE: 2,    // C列
    CANDIDATE_ID: 3,      // D列
    CANDIDATE_NAME: 4,    // E列
    INTERVIEW_TIME: 5,    // F列
    NO_RESPONSE_FLAG: 6,  // G列
  }
};

const CC_ADDRESS = "agentnavi@circus-group.jp";

// ==========================================================
// 1. エントリーポイント
// ==========================================================

/**
 * 【手動用】UIでの確認あり
 */
function manual_generateAgtRemindEmails() {
  const ui = SpreadsheetApp.getUi();
  
  processAgtRemindLogic({
    isAuto: false,
    sendLogic: (companyName, recipients, subject, message) => {
      const response = ui.alert(
        'メール送信の最終確認',
        `以下のメールを送信しますか？\n\nTo: ${recipients}\nSubject: ${subject}\n\n${message.substring(0, 300)}...`,
        ui.ButtonSet.YES_NO
      );
      if (response == ui.Button.YES) {
        sendEmail(recipients, subject, message);
        return `【送信完了】${companyName}`;
      }
      return `【スキップ】${companyName}`;
    }
  });
}

/**
 * 【自動用】トリガー実行（確認なし・ログ強化）
 */
function auto_generateAgtRemindEmails() {
  console.info("--- [AGTリマインド自動実行開始] ---");
  let successCount = 0;
  let skipCount = 0;

  processAgtRemindLogic({
    isAuto: true,
    sendLogic: (companyName, recipients, subject, message) => {
      try {
        sendEmail(recipients, subject, message);
        console.info(`✅ 送信成功: ${companyName} (${recipients})`);
        successCount++;
      } catch (e) {
        console.error(`❌ 送信失敗: ${companyName} / エラー: ${e.message}`);
      }
      return null;
    },
    onSkip: (msg) => {
      console.warn(`⚠️ ${msg}`);
      skipCount++;
    },
    onNoData: () => console.info("明日の面談予定はありませんでした。")
  });

  console.info(`--- [自動実行終了] 成功: ${successCount}件 / スキップ: ${skipCount}件 ---`);
}

// ==========================================================
// 2. 共通ロジック（エンジン部分）
// ==========================================================

function processAgtRemindLogic({ isAuto, sendLogic, onSkip, onNoData }) {
  const emailMap = getEmailMapFromSheet();
  if (!emailMap) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues().slice(1);
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  tomorrow.setHours(0, 0, 0, 0);

  const companiesData = {};
  let foundTarget = false;

  // データ抽出
  data.forEach((row) => {
    const interviewDateValue = row[CONFIG.COLUMN_INDICES.INTERVIEW_DATE];
    if (!interviewDateValue || !(interviewDateValue instanceof Date)) return;

    const interviewDate = new Date(interviewDateValue);
    interviewDate.setHours(0, 0, 0, 0);

    if (interviewDate.getTime() === tomorrow.getTime()) {
      const companyName = row[CONFIG.COLUMN_INDICES.COMPANY_NAME];
      let contactPerson = row[CONFIG.COLUMN_INDICES.CONTACT_PERSON];
      
      if (!companyName || !contactPerson) return;

      if (!companiesData[companyName]) {
        companiesData[companyName] = { interviewsByContact: {} };
      }
      
      contactPerson = contactPerson.toString().trim().replace(/様$/, '');
      if (!companiesData[companyName].interviewsByContact[contactPerson]) {
        companiesData[companyName].interviewsByContact[contactPerson] = [];
      }

      const timeValue = row[CONFIG.COLUMN_INDICES.INTERVIEW_TIME];
      const time = (timeValue instanceof Date) ? Utilities.formatDate(timeValue, 'JST', 'HH:mm') : timeValue;
      const note = (row[CONFIG.COLUMN_INDICES.NO_RESPONSE_FLAG] === '済') ? '' : '※';
      const detail = `${time} ${addSama(row[CONFIG.COLUMN_INDICES.CANDIDATE_NAME])} ${row[CONFIG.COLUMN_INDICES.CANDIDATE_ID]} ${note}`.trim();
      
      companiesData[companyName].interviewsByContact[contactPerson].push(detail);
      foundTarget = true;
    }
  });

  if (!foundTarget) {
    if (onNoData) onNoData();
    return;
  }

  const results = [];
  for (const companyName in companiesData) {
    const recipientList = emailMap[companyName];
    if (!recipientList) {
      const msg = `宛先不明: ${companyName} (設定シートに未登録)`;
      if (onSkip) onSkip(msg);
      results.push(msg);
      continue;
    }

    const recipients = recipientList.join(',');
    const { subject, message } = generateMailContent(companyName, companiesData[companyName], isAuto);
    
    const res = sendLogic(companyName, recipients, subject, message);
    if (res) results.push(res);
  }

  return results.join('\n');
}

// ==========================================================
// 3. ヘルパー関数
// ==========================================================

function generateMailContent(companyName, companyObj, isAuto) {
  let interviewList = '';
  let containsNote = false;
  const interviewsByContact = companyObj.interviewsByContact;

  if (companyName === 'CDC') {
    const allInterviews = Object.values(interviewsByContact).flat();
    interviewList = allInterviews.join('\n');
    containsNote = interviewList.includes('※');
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
    `\nなお「※」の表記が付いている候補者様につきましては、\n本日お送りしたリマインドメッセージへの反応が無く、面談へご参加されない可能性がございます。\nただ候補者様へ面談のご案内自体は行っているため、恐れ入りますがご入室を頂きます様、` : '';

  const subject = isAuto ? 
    `【転職エージェントナビ】明日の面談のご連絡（自動配信）` : 
    `【転職エージェントナビ】明日の面談のご連絡`;

  const autoDisclaimer = isAuto ? `※本メールはシステムより自動配信しております。\n` : '';

  const message = `${addSama(companyName)}
ご担当者 各位

いつもお世話になっております。
転職エージェントナビでございます。

${isAuto ? '明日実施予定の面談案件一覧を、現時点で確定している内容にてお送りいたします。' : '以下、現時点で明日お願いしている面談案件一覧をお送りいたします。'}
${autoDisclaimer}
${interviewList}
${footerNote}
引き続き何卒よろしくお願いいたします。

転職エージェントナビ`;

  return { subject, message };
}

function sendEmail(to, subject, body) {
  GmailApp.sendEmail(to, subject, body, {
    name: "転職エージェントナビ事務局",
    cc: CC_ADDRESS
  });
}

function getEmailMapFromSheet() {
  try {
    const spreadsheet = SpreadsheetApp.openById(MAPPING_SHEET_CONFIG.ID);
    const sheet = spreadsheet.getSheetByName(MAPPING_SHEET_CONFIG.SHEET_NAME);
    if (!sheet) return null;
    const data = sheet.getDataRange().getValues().slice(1);
    const emailMap = {};
    data.forEach(row => {
      const companyName = row[MAPPING_SHEET_CONFIG.COMPANY_COL_INDEX];
      const email = row[MAPPING_SHEET_CONFIG.EMAIL_COL_INDEX];
      if (companyName && email && typeof email === 'string' && email.includes('@')) {
        if (!emailMap[companyName]) emailMap[companyName] = [];
        emailMap[companyName].push(email);
      }
    });
    return emailMap;
  } catch (e) {
    console.error(`宛先取得エラー: ${e.message}`);
    return null;
  }
}

function addSama(name) {
  if (!name || typeof name !== 'string') return '';
  const trimmed = name.trim();
  return (trimmed.endsWith('様') || trimmed.endsWith('各位')) ? trimmed : trimmed + '様';
}