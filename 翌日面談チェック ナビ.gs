//20260912 アンケ回答を本文に含められるように変更
//20260925 事前アンケスプシから 都道府県/生年月日/最終学歴/卒業年/経験社数 を引いて本文に追加

/**
 * 設定情報
 */
const NAVI_EMAIL_MAP = {
  '菅谷さん': 'naomisugaya721@gmail.com',
  '小木曽さん': 'yuki9517@gmail.com',
  '佐藤さん': 'nishi.shima217@gmail.com',
  '浅富さん': 'n-asatomi@circus-group.jp',
  '中村さん': 'hitomi10321107@gmail.com',
  '千田さん': 'shuntaro8710@gmail.com',
  '瀬戸さん': 'k-seto@circus-group.jp',
  '新国さん':'a-nikkuni@circus-group.jp',
  '國本さん':'y-kunimoto@circus-group.jp',
  '菊池さん':'k-kikuchi@circus-group.jp',
  '中原さん':'y-nakahara@circus-group.jp',
  '松本さん':'y-matsumoto@circus-group.jp',
};

const CONFIG_NAVI = {
  SHEET_NAME: '📅面談予定',
  START_ROW: 76,
  COLUMN_INDICES: {
    NAVI_NAME: 12,      // M列
    INTERVIEW_TIME: 13, // N列
    CANDIDATE_ID: 14,   // O列
    CANDIDATE_NAME: 15, // P列
    PHONE_NUMBER: 17,   // R列
    EMAIL: 18,          // S列
    INTERVIEW_TYPE: 19,  // T列
    JIZEN_SURVEY: 20          // U列（アンケ回答）★追加
  }
};

const CC_ADDRESS = "agentnavi@circus-group.jp";

/**
 * 事前アンケスプシ（別ファイル）の設定
 * ※ このスプシには別プロジェクトの sync_anke_status.js がバインドされているが、
 *   プロジェクトが分かれているため定数・関数は共有されない。ここで自己完結させる。
 */
const CONFIG_ANKE = {
  SS_ID: '1eKCrxnxVv_P7eDBkAJPKdo_naB7S1_L6LB1hIEuOWBo',
  SHEET_NAME: 'シート1',
  HEADER_ROW: 1,
  // ヘッダは「卒業年【4/6】」のように採番が付くため部分一致で解決する。
  // 複数列がヒットした場合は黙って別列を読まないようエラーで止める。
  HEADERS: {
    TIMESTAMP:    '回答日時',
    DISPLAY_NAME: 'システム表示名',
    BIRTHDAY:     '生年月日',
    AREA:         'お住まいの地域',
    SCHOOL:       '最終学歴の学校名',
    GRAD_YEAR:    '卒業年',
    COMPANIES:    '経験社数',
    EMAIL:        'メールアドレス'
  }
};

/** 引けなかった項目の表示 */
const ANKE_BLANK = '－';

// ==========================================================
// 1. エントリーポイント（ここを実行・トリガー設定する）
// ==========================================================

/**
 * 【手動用】ボタンやメニューから実行
 * 送信前に確認ダイアログを出します
 */
function manual_sendNaviRemindEmails() {
  const ui = SpreadsheetApp.getUi();

  const report = processNaviRemind((naviName, recipient, body, itemCount) => {
    // 送信確認ダイアログ
    const confirm = ui.alert(
      `${naviName}へ送信確認 (${itemCount}件)`,
      `宛先: ${recipient}\n\n${body}`,
      ui.ButtonSet.YES_NO
    );

    if (confirm === ui.Button.YES) {
      sendEmail(recipient, body);
      return `【完了】${naviName}に送信しました。`;
    } else {
      return `【スキップ】${naviName}の送信をキャンセルしました。`;
    }
  });

  if (report) ui.alert("実行結果:\n\n" + report);
}

/**
 * 【自動用】時間トリガーで実行
 * 確認なしで即送信し、詳細なログを残します
 */
function auto_sendNaviRemindEmails() {
  console.info("--- [自動リマインド開始] ---");

  let successCount = 0;
  let skipCount = 0;
  let errorCount = 0;

  const summary = processNaviRemind((naviName, recipient, body, itemCount) => {
    try {
      sendEmail(recipient, body);
      console.info(`✅ 送信成功: ${naviName} (${recipient}) / 面談数: ${itemCount}件`);
      successCount++;
      return null; // summaryには含めない
    } catch (e) {
      console.error(`❌ 送信失敗: ${naviName} / エラー: ${e.message}`);
      errorCount++;
      return `【エラー】${naviName}: ${e.message}`;
    }
  });

  // 異常終了（シート未検出など）を成功扱いで握り潰さない
  if (summary === "明日の予定はありません。") {
    console.info("明日の面談予定は0件でした。処理を終了します。");
    return;
  }
  if (summary === "データがありません。" || (summary && summary.indexOf("エラー:") === 0)) {
    console.error(`❌ 処理を中断しました: ${summary}`);
    return;
  }

  // スキップ（アドレス未登録など）の情報をログに吐き出す
  if (summary) {
    summary.split("\n").forEach(line => {
      if (line.includes("【不可】")) {
        console.warn(`⚠️ ${line}`);
        skipCount++;
      }
    });
  }

  console.info(`--- [自動リマインド終了] 成功: ${successCount}件 / スキップ: ${skipCount}件 / エラー: ${errorCount}件 ---`);
}


// ==========================================================
// 2. 共通ロジック
// ==========================================================

/**
 * データの抽出からループ処理までのメインロジック
 */
function processNaviRemind(sendLogic) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_NAVI.SHEET_NAME);
  if (!sheet) return "エラー: シートが見つかりません。";

  // 日付準備
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowStr = Utilities.formatDate(tomorrow, 'JST', 'M/d');
  const tomorrowCompare = Utilities.formatDate(tomorrow, 'JST', 'yyyy/MM/dd');

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG_NAVI.START_ROW) return "データがありません。";

  const data = sheet.getRange(CONFIG_NAVI.START_ROW, 1, lastRow - CONFIG_NAVI.START_ROW + 1, 21).getValues();

  // 担当者ごとにデータをまとめる
  const naviDataMap = groupDataByNavi(data, tomorrowCompare);
  const naviKeys = Object.keys(naviDataMap);

  if (naviKeys.length === 0) return "明日の予定はありません。";

  // 事前アンケの突合（失敗してもリマインド送信自体は止めない）
  attachAnkeInfo_(naviDataMap);

  const results = [];
  for (const naviName of naviKeys) {
    // 名寄せ（「佐藤し」→「佐藤」、「佐藤」→「佐藤さん」など）
    const recipient = NAVI_EMAIL_MAP[naviName] || NAVI_EMAIL_MAP[naviName + "さん"];
    const items = naviDataMap[naviName];

    if (!recipient) {
      results.push(`【不可】${naviName} のアドレスが未登録です。`);
      continue;
    }

    const body = generateMailBody(tomorrowStr, items);

    // 外部から渡された送信処理を実行
    const res = sendLogic(naviName, recipient, body, items.length);
    if (res) results.push(res);
  }

  return results.length > 0 ? results.join("\n") : "";
}

/**
 * データを担当者ごとにグルーピング
 */
function groupDataByNavi(data, dateStr) {
  const map = {};
  data.forEach(row => {
    const timeValue = row[CONFIG_NAVI.COLUMN_INDICES.INTERVIEW_TIME];
    if (!(timeValue instanceof Date)) return;

    // 日付チェック
    if (Utilities.formatDate(timeValue, 'JST', 'yyyy/MM/dd') !== dateStr) return;

    // --- 名前整形ロジック ---
        let name = row[CONFIG_NAVI.COLUMN_INDICES.NAVI_NAME].toString()
                    .replace(/[【】]/g, '')   // 1. カッコを消す → "佐藤しさん"
                    .replace(/さん$/, '')     // 2. 先に「さん」を消す → "佐藤し"
                    .replace(/し$/, '')       // 3. その後で末尾の「し」を消す → "佐藤"
                    .trim();                  // 4. 余計な空白を掃除
    // ----------------------------

    if (!name) return;
    if (!map[name]) map[name] = [];

    map[name].push({
      time: Utilities.formatDate(timeValue, 'JST', 'HH:mm'),
      candidateName: row[CONFIG_NAVI.COLUMN_INDICES.CANDIDATE_NAME],
      candidateId: row[CONFIG_NAVI.COLUMN_INDICES.CANDIDATE_ID],
      phone: row[CONFIG_NAVI.COLUMN_INDICES.PHONE_NUMBER],
      email: row[CONFIG_NAVI.COLUMN_INDICES.EMAIL],
      type: row[CONFIG_NAVI.COLUMN_INDICES.INTERVIEW_TYPE],
      survey: row[CONFIG_NAVI.COLUMN_INDICES.JIZEN_SURVEY] || ANKE_BLANK,
      // 突合用（CL-IDを正規化したもの）
      clKey: normalizeClId_(row[CONFIG_NAVI.COLUMN_INDICES.CANDIDATE_ID]),
      emailKey: normalizeEmailForMatch_(row[CONFIG_NAVI.COLUMN_INDICES.EMAIL]),
      anke: null,
      ankeMatchedBy: 'なし'
    });
  });
  return map;
}

/**
 * メール本文の生成
 */
function generateMailBody(dateStr, items) {

  const listText = items.map(item =>
    `${item.time}〜\n` +
    `${item.candidateId} ${item.candidateName} 様\n` +
    `${item.phone}  ${item.email}  ${item.type}\n` +
    `${formatAnkeLine_(item.anke)}`
  ).join("\n\n");

  return `お疲れ様です。
現時点での明日(${dateStr})のナビ・個別面談のリマインドを送信いたします。

${listText}

変更・追加があればLINEにてご連絡いたします。
宜しくお願いします！`;
}

/**
 * 事前アンケ情報の1行表示を組み立てる
 * 例: 埼玉県｜2002/10/22｜ミスパリビューティ専門学校（2023年卒）｜経験社数 1社
 */
function formatAnkeLine_(anke) {
  if (!anke) return `事前アンケ：${ANKE_BLANK}（未回答または未突合）`;

  const school = anke.school || ANKE_BLANK;
  const gradYear = anke.gradYear ? `（${anke.gradYear}年卒）` : '';

  return [
    anke.area || ANKE_BLANK,
    anke.birthday || ANKE_BLANK,
    school + gradYear,
    `経験社数 ${anke.companies || ANKE_BLANK}`
  ].join('｜');
}

/**
 * 実際のメール送信
 */
function sendEmail(to, body) {
  GmailApp.sendEmail(to, "明日の面談リマインドのご連絡", body, {
    name: "転職エージェントナビ事務局",
    cc: CC_ADDRESS
  });
}


// ==========================================================
// 3. 事前アンケ突合
// ==========================================================

/**
 * 担当者ごとのデータに事前アンケ情報を紐づける。
 * アンケスプシが読めなくてもリマインド送信自体は継続する（アドオン扱い）。
 */
function attachAnkeInfo_(naviDataMap) {
  let index;
  try {
    index = buildAnkeIndex_();
  } catch (e) {
    console.error(`⚠️ 事前アンケの読み込みに失敗したため、アンケ情報なしで送信します: ${e.message}`);
    return;
  }

  Object.keys(naviDataMap).forEach(naviName => {
    naviDataMap[naviName].forEach(item => {
      // CL-ID優先 → 空振りならメールアドレスで補完
      if (item.clKey && index.byCl[item.clKey]) {
        item.anke = index.byCl[item.clKey];
        item.ankeMatchedBy = 'CL-ID';
      } else if (item.emailKey && index.byEmail[item.emailKey]) {
        item.anke = index.byEmail[item.emailKey];
        item.ankeMatchedBy = 'メール';
      }
    });
  });
}

/**
 * アンケスプシを読み、CL-ID / メールアドレスの2本立てで引ける索引を作る。
 * 同一キーに複数回答がある場合は回答日時が最も新しいものを採用する。
 */
function buildAnkeIndex_() {
  const sheet = SpreadsheetApp.openById(CONFIG_ANKE.SS_ID).getSheetByName(CONFIG_ANKE.SHEET_NAME);
  if (!sheet) {
    throw new Error(`アンケスプシにシートが見つかりません: ${CONFIG_ANKE.SHEET_NAME}`);
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const firstDataRow = CONFIG_ANKE.HEADER_ROW + 1;
  if (lastRow < firstDataRow) {
    throw new Error('アンケスプシにデータ行がありません。');
  }

  const header = sheet.getRange(CONFIG_ANKE.HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const col = {};
  Object.keys(CONFIG_ANKE.HEADERS).forEach(key => {
    col[key] = findAnkeColumn_(header, CONFIG_ANKE.HEADERS[key]);
  });

  const values = sheet.getRange(firstDataRow, 1, lastRow - CONFIG_ANKE.HEADER_ROW, lastCol).getValues();

  const byCl = {};
  const byEmail = {};
  let multiClRows = 0;

  values.forEach((row, i) => {
    const answeredAt = toDateOrNull_(row[col.TIMESTAMP]);
    const record = {
      area:      toText_(row[col.AREA]),
      birthday:  formatBirthday_(row[col.BIRTHDAY]),
      school:    toText_(row[col.SCHOOL]),
      gradYear:  toText_(row[col.GRAD_YEAR]),
      companies: toText_(row[col.COMPANIES]),
      answeredAt: answeredAt,
      rowNumber: firstDataRow + i
    };

    const clIds = extractClIds_(row[col.DISPLAY_NAME]);
    if (clIds.length > 1) {
      multiClRows++;
      console.warn(`⚠️ CL-IDが複数入っている行があります（${clIds.join(' / ')}）: ${CONFIG_ANKE.SHEET_NAME} ${record.rowNumber}行目。先頭を採用します。`);
    }
    if (clIds.length > 0) {
      upsertNewer_(byCl, clIds[0], record);
    }

    const email = normalizeEmailForMatch_(row[col.EMAIL]);
    if (email) {
      upsertNewer_(byEmail, email, record);
    }
  });

  return { byCl: byCl, byEmail: byEmail, rowCount: values.length, multiClRows: multiClRows };
}

/**
 * 同一キーは回答日時が新しい方を残す。
 * （同一人物が連続で回答し、古い方に生年月日の誤入力が入っている実例があるため）
 */
function upsertNewer_(store, key, record) {
  const current = store[key];
  if (!current) {
    store[key] = record;
    return;
  }
  const a = current.answeredAt ? current.answeredAt.getTime() : -1;
  const b = record.answeredAt ? record.answeredAt.getTime() : -1;
  if (b >= a) store[key] = record;
}

/**
 * ヘッダを部分一致で解決する。
 * ヒットが0件、または2件以上のときは黙って別列を読まずエラーで止める。
 */
function findAnkeColumn_(header, needle) {
  const target = normalizeText_(needle);
  const hits = [];
  header.forEach((h, i) => {
    if (normalizeText_(h).indexOf(target) !== -1) hits.push(i);
  });

  if (hits.length === 0) {
    throw new Error(`アンケスプシに列が見つかりません: 「${needle}」\n実際のヘッダ: ${header.join(' | ')}`);
  }
  if (hits.length > 1) {
    const names = hits.map(i => header[i]).join(' / ');
    throw new Error(`アンケスプシで列が一意に決まりません: 「${needle}」が ${hits.length}列に一致しました（${names}）`);
  }
  return hits[0];
}

/**
 * 「岩崎 莉央 CL40090」のような表示名からCL-IDを全部抜く。
 * 全角（ＣＬ４００９０）や小文字（cl40090）も拾えるようにNFKC正規化してから照合する。
 */
function extractClIds_(value) {
  const text = normalizeText_(value).toUpperCase();
  const matches = text.match(/CL\s*\d+/g);
  if (!matches) return [];
  return matches.map(m => canonicalClId_(m));
}

/**
 * 面談予定シートO列（基本「CL40090」形式）を突合キーに正規化する。
 * CLが付いていない数値だけの値も同じキーに寄せる。
 */
function normalizeClId_(value) {
  const text = normalizeText_(value).toUpperCase();
  if (!text) return '';

  const matched = text.match(/CL\s*\d+/);
  if (matched) return canonicalClId_(matched[0]);

  const digitsOnly = text.replace(/[^\d]/g, '');
  if (digitsOnly && /^\d+$/.test(text.replace(/\s/g, ''))) {
    return canonicalClId_('CL' + digitsOnly);
  }
  return '';
}

/** CL-IDの表記ゆれ（空白・ゼロ埋め）を吸収した比較用の形にする */
function canonicalClId_(raw) {
  const digits = raw.replace(/[^\d]/g, '').replace(/^0+/, '');
  return digits ? 'CL' + digits : '';
}

/** メールアドレスの正規化（空白除去 + 小文字化） */
function normalizeEmailForMatch_(value) {
  return normalizeText_(value).replace(/\s+/g, '').toLowerCase();
}

/** 全角/半角・連続空白の正規化 */
function normalizeText_(value) {
  if (value === null || value === undefined) return '';
  const text = (value instanceof Date)
    ? Utilities.formatDate(value, 'JST', 'yyyy/MM/dd')
    : String(value);
  return text.normalize('NFKC').replace(/\s+/g, ' ').trim();
}

/** セル値を表示用テキストに（空なら空文字） */
function toText_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, 'JST', 'yyyy/MM/dd');
  return (value === null || value === undefined) ? '' : String(value).trim();
}

/** 生年月日の表示整形（Date型・文字列入力の両方に対応） */
function formatBirthday_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, 'JST', 'yyyy/MM/dd');
  return toText_(value);
}

/** 回答日時をDateとして取り出す（文字列で入っている場合も拾う） */
function toDateOrNull_(value) {
  if (value instanceof Date) return value;
  const text = toText_(value);
  if (!text) return null;
  const parsed = new Date(text.replace(/-/g, '/'));
  return isNaN(parsed.getTime()) ? null : parsed;
}


// ==========================================================
// 4. 動作確認用（送信しない）
// ==========================================================

/**
 * 明日の面談に対して、事前アンケが何件引けたかだけをログに出す。
 * 本番の文面を変える前に必ずこれを実行してヒット率を確認する。
 */
function dryRun_checkAnkeMatch() {
  const result = collectTomorrowItems_();
  if (typeof result === 'string') {
    console.info(result);
    return;
  }

  const items = result.items;
  attachAnkeInfo_(result.map);

  let byCl = 0, byEmail = 0, unmatched = 0;
  const unmatchedList = [];

  items.forEach(item => {
    if (item.ankeMatchedBy === 'CL-ID') byCl++;
    else if (item.ankeMatchedBy === 'メール') byEmail++;
    else {
      unmatched++;
      unmatchedList.push(`  ${item.candidateId} ${item.candidateName} / CLキー:${item.clKey || '取得不可'} / メール:${item.emailKey || '空'}`);
    }
  });

  console.info(`明日(${result.dateStr})の面談 ${items.length}件`);
  console.info(`→ CL-IDで突合 ${byCl}件 / メール補完 ${byEmail}件 / 未突合 ${unmatched}件`);
  if (unmatchedList.length > 0) {
    console.warn('未突合の一覧:\n' + unmatchedList.join('\n'));
  }
}

/**
 * 実際に送られる本文を送信せずにログへ出す。
 * dryRun_checkAnkeMatch でヒット率を確認したあと、文面の見た目を確認する用。
 */
function dryRun_previewMailBody() {
  const result = collectTomorrowItems_();
  if (typeof result === 'string') {
    console.info(result);
    return;
  }

  attachAnkeInfo_(result.map);

  Object.keys(result.map).forEach(naviName => {
    const recipient = NAVI_EMAIL_MAP[naviName] || NAVI_EMAIL_MAP[naviName + 'さん'] || '(アドレス未登録)';
    const body = generateMailBody(result.dateStr, result.map[naviName]);
    console.info(`===== ${naviName} → ${recipient} =====\n${body}\n`);
  });
}

/**
 * アンケスプシ側だけの状態を確認する。列解決とキーの作成数を見る。
 */
function dryRun_checkAnkeSheet() {
  const index = buildAnkeIndex_();
  console.info(
    `アンケスプシ ${index.rowCount}行 → CL-IDキー ${Object.keys(index.byCl).length}件 / メールキー ${Object.keys(index.byEmail).length}件`
  );
  console.info(`CL-IDが複数入っていた行: ${index.multiClRows}件`);

  const sampleKeys = Object.keys(index.byCl).slice(0, 5);
  sampleKeys.forEach(key => {
    const r = index.byCl[key];
    console.info(`  ${key} (${r.rowNumber}行目): ${formatAnkeLine_(r)}`);
  });
}

/** 明日の面談データを集める（ドライラン用の共通処理） */
function collectTomorrowItems_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_NAVI.SHEET_NAME);
  if (!sheet) return 'エラー: シートが見つかりません。';

  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const dateStr = Utilities.formatDate(tomorrow, 'JST', 'M/d');
  const compareStr = Utilities.formatDate(tomorrow, 'JST', 'yyyy/MM/dd');

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG_NAVI.START_ROW) return 'データがありません。';

  const data = sheet.getRange(CONFIG_NAVI.START_ROW, 1, lastRow - CONFIG_NAVI.START_ROW + 1, 21).getValues();
  const map = groupDataByNavi(data, compareStr);

  const items = [];
  Object.keys(map).forEach(naviName => {
    map[naviName].forEach(item => items.push(item));
  });

  if (items.length === 0) return '明日の予定はありません。';
  return { map: map, items: items, dateStr: dateStr };
}
