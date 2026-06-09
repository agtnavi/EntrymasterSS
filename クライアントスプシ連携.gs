//20260609 クライアントスプシ連携を新設
/**
 * ===== AGT送客配信ツール（全コード1本まとめ・手貼り用） =====
 * 使い方:
 *   1) 基幹スプシ → 拡張機能 → Apps Script を wagatsuma@lmi で開く
 *   2) デフォルトの「コード.gs」の中身を全消し → このファイルを丸ごと貼る
 *   3) プロジェクト設定で appsscript.json を表示ONにし、別途マニフェストを貼る
 *   4) 保存 → スプシ再読込 → メニュー「AGT配信」
 *
 * 編集の正本は projects/active/agt-distribution/src/ 配下。
 * このファイルは src/ を連結しただけのコピー（編集はsrc側で行い再連結する）。
 */

/* ============================================================
 * Config
 * ============================================================ */
var CONFIG_CLIENTSYNC = {
  // ── 基幹スプシ内のシート名（実物に合わせて確認） ──
  SOUKYAKU_SHEET: '送客管理', // 送客データ本体
  AGT_SHEET: 'AGT法人管理',               // AGT会社台帳
  LOG_SHEET: '配信ログ',                  // 同期結果ログ（GASが自動生成）

  // ── 各AGTファイル内に作る「出力タブ」名 ──
  OUTPUT_TAB: '送客リスト（自動）',

  // ── 同期間隔（時間トリガー） ──
  SYNC_INTERVAL_MIN: 5,

  // ── 送客マスタのヘッダー名（列順が変わっても名前で引くので壊れにくい） ──
  COLS: {
    soukyakuId:   '送客ID',
    status:       '送客ステータス',
    jobseekerName:'[自動反映]求職者氏名',
    jobseekerId:  '[自動反映]求職者ID',
    agentStaff:   'AGT　担当者',          // 注: 全角スペース入り
    agentCompany: 'AGT会社',
    amount:       '送客金額',
    meetingDate:  '面談予定日',
    jisshibi:     '面談実施日',
    pdf:          '候補者情報PDF',
    yoteiCleaned: '面談予定日(文字を除外)'  // 注: 半角括弧
  },

  // ── フィルタ条件 ──
  STATUS_ALLOW: ['面談実施済み', '面談実施済み（クローズ）', '面談予約中'], // 括弧は全角
  JISSHIBI_EXCLUDE: '面談未実施',

  // ── AGT法人管理シートのヘッダー名 ──
  AGT_COLS: {
    seishiki: '正式名称',
    call:     'サービス名/呼称',
    alias:    'エイリアス',
    fileId:   'ファイルID'  // 新設列。なければ ensureFileIdColumn_() が作る
  },

  // ── 出力先ファイルID収集（Gmail走査）の設定 ──
  SHARE_FROM: '',  // 共有元アドレス（要確認: 例 y-terakami@circus-group.jp）。空でもキーワード検索は動くがノイズ増
  SHARE_SUBJECT_KEYWORD: '求職者管理シート',
  GMAIL_SEARCH_MAX: 500,

  // ── 出力タブのレイアウト（左→右の表示列） ──
  OUTPUT_HEADERS: ['頁番', '求職者No.', '求職者名', 'ご担当者', '金額', '面談日時', '候補者情報PDF', '状態', '送客ID'],
  STATE_ACTIVE:  '有効',
  STATE_INACTIVE:'対象外',
  COLOR_INACTIVE:'#d9d9d9',
  COLOR_ACTIVE:  '#ffffff'
};


/* ============================================================
 * Resolver — 名寄せ + 出力先ファイル解決
 * ============================================================ */

/** 名前の正規化: トリム + 全角/半角統一 + 小文字化 + 空白除去 */
function normalizeName_(s) {
  if (s === null || s === undefined) return '';
  var t = String(s);
  t = t.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(c){ return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });
  t = t.replace(/　/g, ' ').replace(/\s+/g, '');
  return t.trim().toLowerCase();
}

/** AGT法人管理シートを読み、正規化キー → エントリ のマップを作る */
function buildAgentDirectory_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(CONFIG_CLIENTSYNC.AGT_SHEET);
  if (!sh) throw new Error('AGT法人管理シートが見つからない: ' + CONFIG_CLIENTSYNC.AGT_SHEET);

  ensureFileIdColumn_(sh);

  var values = sh.getDataRange().getValues();
  var header = values[0];
  var idx = headerIndexMap_(header, CONFIG_CLIENTSYNC.AGT_COLS);

  var map = {};
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var seishiki = String(row[idx.seishiki] || '').trim();
    if (!seishiki) continue;
    var fileId = idx.fileId >= 0 ? String(row[idx.fileId] || '').trim() : '';
    var entry = { rowIndex: r + 1, seishiki: seishiki, fileId: fileId };

    var keys = [seishiki];
    if (idx.call >= 0 && row[idx.call]) keys.push(String(row[idx.call]).trim());
    if (idx.alias >= 0 && row[idx.alias]) {
      String(row[idx.alias]).split(',').forEach(function(a){
        a = a.trim();
        if (a && a !== '-') keys.push(a);
      });
    }
    keys.forEach(function(k){
      var nk = normalizeName_(k);
      if (nk && !map[nk]) map[nk] = entry;
    });
  }
  return map;
}

/** AGT会社名 → エントリ（見つからなければ null） */
function resolveAgent_(directory, agentCompanyName) {
  var nk = normalizeName_(agentCompanyName);
  return directory[nk] || null;
}

/** 出力先ファイルID（キャッシュ済みなら返す。無ければ null） */
function resolveOutputFileId_(entry) {
  return entry.fileId || null;
}

/** ファイル名からAGT名部分を抽出: 「JOB PALETTE様【求職者管理シート】...」→「JOB PALETTE」 */
function extractAgentFromFilename_(fname) {
  if (!fname) return '';
  var head = String(fname).split('【')[0];
  head = head.replace(/様\s*$/, '');
  return head.trim();
}

/** 共有通知メールを走査し、出力先ファイルIDを法人管理「ファイルID」列にキャッシュ */
function collectOutputFileIds_(dryRun) {
  var directory = buildAgentDirectory_();

  var q = '"' + CONFIG_CLIENTSYNC.SHARE_SUBJECT_KEYWORD + '"';
  if (CONFIG_CLIENTSYNC.SHARE_FROM) q = 'from:(' + CONFIG_CLIENTSYNC.SHARE_FROM + ') ' + q;

  var threads = GmailApp.search(q, 0, CONFIG_CLIENTSYNC.GMAIL_SEARCH_MAX);
  var logs = [];
  var pickedByEntry = {};

  threads.forEach(function(t){
    t.getMessages().forEach(function(m){
      var subject = m.getSubject() || '';
      var fname = extractQuotedName_(subject, CONFIG_CLIENTSYNC.SHARE_SUBJECT_KEYWORD);
      var body = m.getBody() || '';
      var idMatch = body.match(/spreadsheets\/d\/([a-zA-Z0-9_\-]+)/);
      var fileId = idMatch ? idMatch[1] : '';
      if (!fname || !fileId) return;

      var agentHead = extractAgentFromFilename_(fname);
      var entry = resolveAgent_(directory, agentHead);
      var isCopy = /コピー/.test(fname);

      if (!entry) {
        logs.push([fname, fileId, '名寄せ失敗（AGT名=' + agentHead + '）', '']);
        return;
      }
      var cur = pickedByEntry[entry.seishiki];
      if (!cur || (cur.isCopy && !isCopy)) {
        pickedByEntry[entry.seishiki] = { fileId: fileId, fname: fname, isCopy: isCopy, entry: entry };
      }
    });
  });

  Object.keys(pickedByEntry).forEach(function(seishiki){
    var p = pickedByEntry[seishiki];
    if (!dryRun) writeFileIdCache_(p.entry.rowIndex, p.fileId);
    logs.push([p.fname, p.fileId, dryRun ? '収集OK（書き込みなし）' : 'ファイルID記録' + (p.isCopy ? '（※コピー名）' : ''), seishiki]);
  });

  return logs;
}

/** 文字列中の「…」のうち keyword を含むものを返す（無ければ最初の「」） */
function extractQuotedName_(s, keyword) {
  var re = /「([^」]*)」/g, m, first = null;
  while ((m = re.exec(s)) !== null) {
    if (first === null) first = m[1];
    if (m[1].indexOf(keyword) >= 0) return m[1];
  }
  return first;
}

/** 法人管理シートにファイルID列が無ければ末尾に追加 */
function ensureFileIdColumn_(sh) {
  var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (header.indexOf(CONFIG_CLIENTSYNC.AGT_COLS.fileId) >= 0) return;
  var col = sh.getLastColumn() + 1;
  sh.getRange(1, col).setValue(CONFIG_CLIENTSYNC.AGT_COLS.fileId);
}

/** 法人管理シートの指定行のファイルID列に値を書く */
function writeFileIdCache_(rowIndex, fileId) {
  var sh = SpreadsheetApp.getActive().getSheetByName(CONFIG_CLIENTSYNC.AGT_SHEET);
  var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var col = header.indexOf(CONFIG_CLIENTSYNC.AGT_COLS.fileId);
  if (col < 0) { ensureFileIdColumn_(sh); col = sh.getLastColumn() - 1; }
  sh.getRange(rowIndex, col + 1).setValue(fileId);
}

/** ヘッダー配列と {key:ヘッダー名} から {key:colIndex(0始まり,無ければ-1)} */
function headerIndexMap_(header, colNameMap) {
  var out = {};
  Object.keys(colNameMap).forEach(function(k){
    out[k] = header.indexOf(colNameMap[k]);
  });
  return out;
}


/* ============================================================
 * Distributor — 差分配信（1 AGTファイルへの書き込み）
 * ============================================================ */

function distributeToAgent_(fileId, matchedRows) {
  var ss = SpreadsheetApp.openById(fileId);
  var sh = ss.getSheetByName(CONFIG_CLIENTSYNC.OUTPUT_TAB) || createOutputTab_(ss);

  var numCols = CONFIG_CLIENTSYNC.OUTPUT_HEADERS.length;
  var idColIndex = numCols - 1;
  var stateColIndex = numCols - 2;

  var lastRow = sh.getLastRow();
  var arr = (lastRow >= 2) ? sh.getRange(2, 1, lastRow - 1, numCols).getValues() : [];

  var idxBySid = {};
  var maxPage = 0;
  for (var i = 0; i < arr.length; i++) {
    var sid = String(arr[i][idColIndex] || '').trim();
    if (sid) idxBySid[sid] = i;
    var pg = Number(arr[i][0]) || 0;
    if (pg > maxPage) maxPage = pg;
  }

  var matchedSet = {};
  matchedRows.forEach(function(m){ matchedSet[m.soukyakuId] = m; });

  var stats = { added: 0, updated: 0, deactivated: 0 };

  // 既存行: 差分のみ書き込み
  for (var j = 0; j < arr.length; j++) {
    var sheetRow = j + 2;
    var sid2 = String(arr[j][idColIndex] || '').trim();
    if (!sid2) continue;

    if (matchedSet[sid2]) {
      var newRow = buildRow_(arr[j][0], matchedSet[sid2].cells, CONFIG_CLIENTSYNC.STATE_ACTIVE, sid2);
      var wasInactive = (String(arr[j][stateColIndex]) === CONFIG_CLIENTSYNC.STATE_INACTIVE);
      if (!rowsEqual_(arr[j], newRow)) {
        sh.getRange(sheetRow, 1, 1, numCols).setValues([newRow]);
        if (wasInactive) setRowBg_(sh, sheetRow, numCols, CONFIG_CLIENTSYNC.COLOR_ACTIVE);
        stats.updated++;
      }
    } else {
      if (String(arr[j][stateColIndex]) !== CONFIG_CLIENTSYNC.STATE_INACTIVE) {
        sh.getRange(sheetRow, stateColIndex + 1).setValue(CONFIG_CLIENTSYNC.STATE_INACTIVE);
        setRowBg_(sh, sheetRow, numCols, CONFIG_CLIENTSYNC.COLOR_INACTIVE);
        stats.deactivated++;
      }
    }
  }

  // 新規合致 → 末尾にまとめて1回追記
  var appends = [];
  matchedRows.forEach(function(m){
    if (idxBySid.hasOwnProperty(m.soukyakuId)) return;
    maxPage++;
    appends.push(buildRow_(maxPage, m.cells, CONFIG_CLIENTSYNC.STATE_ACTIVE, m.soukyakuId));
  });
  if (appends.length) {
    var startRow = arr.length + 2;
    sh.getRange(startRow, 1, appends.length, numCols).setValues(appends);
    stats.added = appends.length;
  }

  return stats;
}

/** 出力タブを新規作成しヘッダー設定 + 送客ID列を非表示 */
function createOutputTab_(ss) {
  var sh = ss.insertSheet(CONFIG_CLIENTSYNC.OUTPUT_TAB);
  var numCols = CONFIG_CLIENTSYNC.OUTPUT_HEADERS.length;
  sh.getRange(1, 1, 1, numCols).setValues([CONFIG_CLIENTSYNC.OUTPUT_HEADERS])
    .setFontWeight('bold').setBackground('#9bbb59').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.hideColumns(numCols);
  return sh;
}

/** 出力1行を組む */
function buildRow_(page, cells, state, soukyakuId) {
  return [page, cells[0], cells[1], cells[2], cells[3], cells[4], cells[5], state, soukyakuId];
}

/** 2行の全セルを文字列比較 */
function rowsEqual_(a, b) {
  if (a.length !== b.length) return false;
  for (var i = 0; i < a.length; i++) {
    if (String(a[i]) !== String(b[i])) return false;
  }
  return true;
}

/** 1行ぶんの背景色をまとめて設定 */
function setRowBg_(sh, row, numCols, color) {
  var colors = [[]];
  for (var k = 0; k < numCols; k++) colors[0].push(color);
  sh.getRange(row, 1, 1, numCols).setBackgrounds(colors);
}


/* ============================================================
 * Main — 全AGT同期 / onEdit / フィルタ判定
 * ============================================================ */

/** 送客マスタを読み、AGT会社名 → 合致行配列 のマップを作る */
function buildMatchedByAgent_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(CONFIG_CLIENTSYNC.SOUKYAKU_SHEET);
  if (!sh) throw new Error('送客シートが見つからない: ' + CONFIG_CLIENTSYNC.SOUKYAKU_SHEET);

  var values = sh.getDataRange().getValues();
  var headerRow = findHeaderRow_(values, CONFIG_CLIENTSYNC.COLS.soukyakuId);
  if (headerRow < 0) throw new Error('ヘッダー行が見つからない（送客ID列なし）');
  var header = values[headerRow];
  var c = headerIndexMap_(header, CONFIG_CLIENTSYNC.COLS);

  var byAgent = {};
  for (var r = headerRow + 1; r < values.length; r++) {
    var row = values[r];
    if (!isMatch_(row, c)) continue;
    var agent = String(row[c.agentCompany] || '').trim();
    if (!agent) continue;
    if (!byAgent[agent]) byAgent[agent] = [];
    byAgent[agent].push({
      soukyakuId: String(row[c.soukyakuId]).trim(),
      cells: [
        row[c.jobseekerId],
        row[c.jobseekerName],
        row[c.agentStaff],
        row[c.amount],
        row[c.meetingDate],
        row[c.pdf]
      ]
    });
  }
  return byAgent;
}

/** フィルタ判定 */
function isMatch_(row, c) {
  var status = String(row[c.status] || '').trim();
  if (CONFIG_CLIENTSYNC.STATUS_ALLOW.indexOf(status) < 0) return false;
  if (String(row[c.yoteiCleaned] || '').trim() === '') return false;
  if (String(row[c.jisshibi] || '').trim() === CONFIG_CLIENTSYNC.JISSHIBI_EXCLUDE) return false;
  return true;
}

/** 全AGT同期（時間トリガー / 手動） */
function syncAllAgents() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) { return; }
  try {
    var directory = buildAgentDirectory_();
    var byAgent = buildMatchedByAgent_();
    var props = PropertiesService.getScriptProperties();
    var logs = [];

    var agentSet = {};
    Object.keys(byAgent).forEach(function(a){ agentSet[a] = true; });
    props.getKeys().forEach(function(k){ if (k.indexOf('h_') === 0) agentSet[k.substring(2)] = true; });

    Object.keys(agentSet).forEach(function(agent){
      var matched = byAgent[agent] || [];
      var hash = hashMatched_(matched);
      var prev = props.getProperty('h_' + agent);
      if (hash === prev) return; // 変化なし → ファイルを開かない

      var entry = resolveAgent_(directory, agent);
      if (!entry) { logs.push([agent, matched.length, '名寄せ失敗', '']); return; }
      var fileId = resolveOutputFileId_(entry);
      if (!fileId) { logs.push([agent, matched.length, '出力先未収集（共有メール収集が必要）', entry.seishiki]); return; }
      try {
        var stats = distributeToAgent_(fileId, matched);
        props.setProperty('h_' + agent, hash);
        logs.push([agent, matched.length, '成功 追加' + stats.added + '/更新' + stats.updated + '/対象外化' + stats.deactivated, entry.seishiki]);
      } catch (e) {
        logs.push([agent, matched.length, 'エラー: ' + e.message, entry.seishiki]);
      }
    });

    writeLog_(logs);
  } finally {
    lock.releaseLock();
  }
}

/** 合致行集合のハッシュ（変化検知用） */
function hashMatched_(matched) {
  if (!matched.length) return 'EMPTY';
  var s = JSON.stringify(matched);
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s);
  return Utilities.base64Encode(d);
}

/** installable onEdit: 編集行のAGTだけ即時同期 */
function onMasterEdit(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (sh.getName() !== CONFIG_CLIENTSYNC.SOUKYAKU_SHEET) return;

    var values = sh.getDataRange().getValues();
    var headerRow = findHeaderRow_(values, CONFIG_CLIENTSYNC.COLS.soukyakuId);
    if (headerRow < 0) return;
    var c = headerIndexMap_(values[headerRow], CONFIG_CLIENTSYNC.COLS);

    var editedRow = e.range.getRow();
    if (editedRow <= headerRow + 1) return;
    var agent = String(values[editedRow - 1][c.agentCompany] || '').trim();
    if (!agent) return;

    var matched = [];
    for (var r = headerRow + 1; r < values.length; r++) {
      var row = values[r];
      if (String(row[c.agentCompany] || '').trim() !== agent) continue;
      if (!isMatch_(row, c)) continue;
      matched.push({
        soukyakuId: String(row[c.soukyakuId]).trim(),
        cells: [row[c.jobseekerId], row[c.jobseekerName], row[c.agentStaff], row[c.amount], row[c.meetingDate], row[c.pdf]]
      });
    }

    var directory = buildAgentDirectory_();
    var entry = resolveAgent_(directory, agent);
    if (!entry) return;
    var fileId = resolveOutputFileId_(entry);
    if (!fileId) return;
    distributeToAgent_(fileId, matched);
  } catch (err) {
    Logger.log('onMasterEdit error: ' + err.message);
  }
}

/** ヘッダー行のindexを返す（指定キー列名を含む最初の行）。無ければ-1 */
function findHeaderRow_(values, keyName) {
  for (var i = 0; i < Math.min(values.length, 10); i++) {
    if (values[i].indexOf(keyName) >= 0) return i;
  }
  return -1;
}

/** 配信ログシートに結果を書く（毎回上書き） */
function writeLog_(rows) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(CONFIG_CLIENTSYNC.LOG_SHEET) || ss.insertSheet(CONFIG_CLIENTSYNC.LOG_SHEET);
  sh.clearContents();
  var header = ['AGT会社（送客マスタ表記）', '合致件数', '結果', '正式名称'];
  var out = [header].concat(rows);
  out.unshift(['最終実行', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'), '', '']);
  sh.getRange(1, 1, out.length, 4).setValues(out);
}


/* ============================================================
 * Setup — メニュー / トリガー / ドライラン
 * ============================================================ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AGT配信')
    .addItem('① 名寄せ確認（書き込みなし）', 'dryRun')
    .addItem('② 出力先ID収集プレビュー（書き込みなし）', 'collectFileIdsDryRun')
    .addItem('③ 出力先ファイルID収集（共有メール走査・記録）', 'collectFileIds')
    .addItem('④ 今すぐ全AGT同期', 'syncAllAgents')
    .addSeparator()
    .addItem('⑤ 自動同期をON（トリガー登録）', 'setupTriggers')
    .addItem('   自動同期をOFF（トリガー削除）', 'removeTriggers')
    .addToUi();
}

function setupTriggers() {
  removeTriggers();
  ScriptApp.newTrigger('syncAllAgents')
    .timeBased()
    .everyMinutes(roundToAllowed_(CONFIG_CLIENTSYNC.SYNC_INTERVAL_MIN))
    .create();
  ScriptApp.newTrigger('onMasterEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
  SpreadsheetApp.getActive().toast('自動同期をONにした（' + CONFIG_CLIENTSYNC.SYNC_INTERVAL_MIN + '分間隔 + 編集時）');
}

function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    var fn = t.getHandlerFunction();
    if (fn === 'syncAllAgents' || fn === 'onMasterEdit') ScriptApp.deleteTrigger(t);
  });
}

function roundToAllowed_(min) {
  var allowed = [1, 5, 10, 15, 30];
  for (var i = 0; i < allowed.length; i++) if (min <= allowed[i]) return allowed[i];
  return 30;
}

function collectFileIds() {
  var logs = collectOutputFileIds_(false);
  writeLog_(logs);
  SpreadsheetApp.getActive().toast('出力先ID収集 完了。「' + CONFIG_CLIENTSYNC.LOG_SHEET + '」シートを確認');
}

function collectFileIdsDryRun() {
  var logs = collectOutputFileIds_(true);
  writeLog_(logs);
  SpreadsheetApp.getActive().toast('収集プレビュー 完了（書き込みなし）。「' + CONFIG_CLIENTSYNC.LOG_SHEET + '」シート確認');
}

function dryRun() {
  var directory = buildAgentDirectory_();
  var byAgent = buildMatchedByAgent_();
  var logs = [];
  Object.keys(byAgent).forEach(function(agent){
    var cnt = byAgent[agent].length;
    var entry = resolveAgent_(directory, agent);
    if (!entry) { logs.push([agent, cnt, '名寄せ失敗', '']); return; }
    if (entry.fileId) { logs.push([agent, cnt, '配信対象OK（出力先ID有）', entry.seishiki]); return; }
    logs.push([agent, cnt, '出力先ID未収集（③で共有メール収集が必要）', entry.seishiki]);
  });
  writeLog_(logs);
  SpreadsheetApp.getActive().toast('名寄せ確認 完了。「' + CONFIG_CLIENTSYNC.LOG_SHEET + '」シート確認');
}