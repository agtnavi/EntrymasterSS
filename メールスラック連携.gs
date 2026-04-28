//20260428このスクリプトを追加
//20260428ログを追加
const MAIL_SLACK_CONFIG = {
  webhookUrl:'',
  processedLabel: 'mail-router/processed',
  errorLabel: 'mail-router/error',
  slackNotify: [
    {
      key: 'gakujo',
      fromDomains: ['gakujo.ne.jp'],
    },
    {
      key: 'cdc',
      fromDomains: ['type.jp'],
    },
    {
      key: 'ardent_staff',
      fromDomains: ['ardent-staff.com'],
    },
    {
      key: 'access',
      fromEmails: ['ito.yuukaaccess@gmail.com', 'b080101148@gmail.com'],
    },
  ],
};

function main_notifySlackTargetMails() {
  const processedLabel = getOrCreateLabel_(MAIL_SLACK_CONFIG.processedLabel);
  const errorLabel = getOrCreateLabel_(MAIL_SLACK_CONFIG.errorLabel);
  const threads = GmailApp.search(`is:unread in:inbox -label:"${MAIL_SLACK_CONFIG.processedLabel}"`, 0, 500);
  console.log("結果"+threads.length+"件")

  threads.forEach(thread => {
    try {
      const handled = processSlackThread_(thread);
      if (!handled) return;

      thread.addLabel(processedLabel);
      thread.moveToArchive();
    } catch (error) {
      thread.addLabel(errorLabel);
      console.error(`[mail-router] thread failed: ${thread.getId()} ${error.message}`);
    }
  });
}

function processSlackThread_(thread) {
  const messages = thread.getMessages();
  let handled = false;

  messages.forEach(message => {
    if (!message.isUnread()) return;

    const route = matchSlackRoute_(message);
    if (!route) return;

    sendSlackNotification_(route, thread, message);
    message.markRead();
    handled = true;
  });

  return handled;
}


function matchSlackRoute_(message) {
  const normalized = normalizeFrom_(message.getFrom());

  const route = MAIL_SLACK_CONFIG.slackNotify.find(route => {
    const emailMatched = (route.fromEmails || []).some(email => normalized.email === email.toLowerCase());
    const domainMatched = (route.fromDomains || []).some(domain => normalized.domain === domain.toLowerCase());
    return emailMatched || domainMatched;
  });

  return route || null;
}

function sendSlackNotification_(route, thread, message) {
  const webhookUrl = MAIL_SLACK_CONFIG.webhookUrl;
  const plainBody = message.getPlainBody().trim();
  const bodyPreview = plainBody.length > 300 ? plainBody.slice(0, 300) + '\n...(truncated)' : plainBody;

  const payload = {
    text: [
      '新規メールを受信しました',
      'route: ' + route.key,
      'from: ' + normalizeFrom_(message.getFrom()).email,
      'subject: ' + message.getSubject(),
      'date: ' + Utilities.formatDate(message.getDate(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      '',
      '本文冒頭:',
      bodyPreview
    ].join('\n')
  };

  const response = UrlFetchApp.fetch(webhookUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  if (response.getResponseCode() >= 300) {
    throw new Error('Slack webhook failed: HTTP ' + response.getResponseCode());
  }

  Utilities.sleep(1000);
}



function normalizeFrom_(rawFrom) {
  const match = String(rawFrom).match(/<([^>]+)>/);
  const email = (match ? match[1] : rawFrom).trim().toLowerCase();
  const domain = email.includes('@') ? email.split('@')[1] : '';
  return { email, domain };
}

function getOrCreateLabel_(labelName) {
  return GmailApp.getUserLabelByName(labelName) || GmailApp.createLabel(labelName);
}

