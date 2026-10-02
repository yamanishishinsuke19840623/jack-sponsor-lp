// =============================================
//  着物マジシャン Jack LP — スポンサー管理 GAS
//  フォーム自動返信 / Stripe通知 / スプレッドシート記録
//  タスクチェックリスト / LP支援者一覧
// =============================================

var JACK_EMAIL = 'kimonomagician@gmail.com';
var CC_EMAIL   = 'yamanishishinsuke19840623@gmail.com';
var SHEET_ID   = '1BUdt7GVFGvPhzFMOM6trGvqfeHLwWjAJ3o5Cen4uCb4';

// =============================================
//  エントリーポイント
// =============================================

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.type && data.data && data.data.object) {
      return handleStripeWebhook(data);
    }
    var ledgerRow = logToSheet(data);
    try { createBenefitTasks_(data, false, ledgerRow); } catch (err) {}
    sendAutoReplyToApplicant(data);
    sendTaskChecklist(data);
    return res({ok: true});
  } catch (err) {
    return res({ok: false, error: err.toString()});
  }
}

// LP からスポンサー一覧 / 訪問カウントを取得（JSONP）
function doGet(e) {
  var params   = e.parameter || {};
  var callback = params.callback || 'cb';
  var action   = params.action  || 'sponsors';

  // 訪問カウンター
  if (action === 'visit') {
    var count = incrementVisitorCount();
    return ContentService.createTextOutput(callback + '(' + JSON.stringify({count: count}) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  // スポンサー一覧
  var sponsors = [];
  if (SHEET_ID) {
    try {
      var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('台帳');
      var rows  = sheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][13] === 'はい') { // N列: LP掲載
          sponsors.push({
            name:    rows[i][4]  || rows[i][1], // 掲載希望名 or お名前
            plan:    rows[i][3],
            message: rows[i][11] || ''
          });
        }
      }
    } catch(err) {}
  }
  var json = JSON.stringify({sponsors: sponsors});
  return ContentService.createTextOutput(callback + '(' + json + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// スプレッドシートの「訪問数」シートに累計を記録して返す
function incrementVisitorCount() {
  try {
    var ss    = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('訪問数') || ss.insertSheet('訪問数');
    var cell  = sheet.getRange('A1');
    var count = (parseInt(cell.getValue()) || 0) + 1;
    cell.setValue(count);
    return count;
  } catch(err) {
    return 0;
  }
}

function res(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheetUrl() {
  return SHEET_ID ? 'https://docs.google.com/spreadsheets/d/' + SHEET_ID : '（SHEET_ID未設定）';
}

// =============================================
//  2026-09-14 Stripe再送重複の一括整理（1回限り実行）
//  ・「お名前+メール+プラン」が同一で振込確認=未確認のグループは最古の1行だけ残す
//  ・お名前が「【テスト】」で始まる行（接続テスト用ダミー）は無条件で削除
//  Apps Scriptエディタで本関数を選択して実行する。実行後は削除してよい。
// =============================================

function dedupeStripeDuplicates_20260914() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('台帳');
  var rows  = sheet.getDataRange().getValues();
  var seen  = {};
  var rowsToDelete = [];

  for (var i = 1; i < rows.length; i++) {
    var name    = rows[i][1];
    var email   = rows[i][2];
    var plan    = rows[i][3];
    var confirm = rows[i][12];

    if (String(name).indexOf('【テスト】') === 0) {
      rowsToDelete.push(i + 1); // シート上の行番号（1始まり＋ヘッダー分）
      continue;
    }

    if (confirm === '未確認') {
      var key = name + '|' + email + '|' + plan;
      if (seen[key]) {
        rowsToDelete.push(i + 1);
      } else {
        seen[key] = true;
      }
    }
  }

  // 行番号が大きい方から削除（小さい方から消すと後続の行番号がずれるため）
  rowsToDelete.sort(function(a, b){ return b - a; });
  rowsToDelete.forEach(function(r){ sheet.deleteRow(r); });

  return {deletedRows: rowsToDelete.length, deletedRowNumbers: rowsToDelete};
}

// =============================================
//  初期セットアップ（1回だけ実行）
// =============================================

function setupSheet() {
  var ss      = SpreadsheetApp.create('Jack LP スポンサー台帳');
  var sheet   = ss.getActiveSheet();
  sheet.setName('台帳');

  var headers = [
    '申込日時','お名前','メール','プラン','掲載希望名',
    'Instagram','X(Twitter)','URL','企業紹介文','ブランドストーリー',
    'Powered_by','応援メッセージ','振込確認','LP掲載','確認メール送信日'
  ];
  var hRange = sheet.getRange(1, 1, 1, headers.length);
  hRange.setValues([headers]);
  hRange.setFontWeight('bold');
  hRange.setBackground('#1a3a1a');
  hRange.setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  // 振込確認・LP掲載 のドロップダウン
  var confirmRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未確認', '確認済'], true).build();
  var publishRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['いいえ', 'はい'], true).build();
  sheet.getRange('M2:M1000').setDataValidation(confirmRule);
  sheet.getRange('N2:N1000').setDataValidation(publishRule);

  sheet.autoResizeColumns(1, headers.length);

  Logger.log('✅ セットアップ完了');
  Logger.log('Sheet ID: ' + ss.getId());
  Logger.log('Sheet URL: ' + ss.getUrl());
  Logger.log('');
  Logger.log('↑ このSheet IDをGASコード冒頭の SHEET_ID に貼り付けてください');
}

// =============================================
//  スプレッドシート記録
// =============================================

function logToSheet(d, autoConfirmed) {
  if (!SHEET_ID) return;
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('台帳');
  sheet.appendRow([
    d.date            || new Date(),
    d.name            || '',
    d.email           || '',
    d.plan            || '',
    d['掲載希望名']   || '',
    d['Instagram']    || '',
    d['X(Twitter)']   || '',
    d['ウェブサイトURL']    || '',
    d['企業・活動紹介文']   || '',
    d['ブランドストーリー'] || '',
    d['Powered_by表記']    || '',
    d['応援メッセージ']    || '',
    autoConfirmed ? '確認済' : '未確認',   // 振込確認：クレカ決済は即時確定するため自動で確認済に
    autoConfirmed ? 'はい'   : 'いいえ',   // LP掲載：クレカ決済は自動でLPの支援総額・一覧に反映
    autoConfirmed ? new Date() : ''
  ]);
  return sheet.getLastRow();
}

// =============================================
//  プラン別 特典・タスク定義
// =============================================

var PLAN_NAMES = {
  a: 'コーヒー1杯のエール', b: '旅の相棒（日本国旗へお名前記入）', c: '現地直通ラジオ生電話権',
  d: '旅の拠点に泊まる（ブリッジ宿泊）', e: 'ネームロケーション写真',
  f: '名前を刻む（YouTube概要欄）', g: '荒野の生還パーツ', h: 'アメリカからの生還（ルート66）',
  i: 'レジェンド集結（オンライン飲み会）', j: 'あなたの街に直撃！', k: '出張講演会プラン',
  l: '伝説の相棒譲渡（リアル・リヤカー永久所有権）', m: 'ブリッジ懇親会（食べ飲み放題）', n: 'プライベート・キャンプ会',
  o: '秘伝マンツーマンレッスン（60分）', p: '北米大陸ロゴ掲載（小）＋概要欄クレジット',
  q: '北米横断パートナーコース（ロゴ中）', r: '北米横断メインパートナーコース（ロゴ大）'
};

var PLAN_BENEFITS = {
  a: ['御礼メッセージの送付','支援者ページへのお名前掲載'],
  b: ['日本国旗へお名前記入（寄せ書き）','御礼メッセージの送付'],
  c: ['現地からの直通ラジオ生電話','御礼メッセージの送付'],
  d: ['ブリッジ（下関）宿泊1泊','御礼メッセージの送付'],
  e: ['砂漠・荒野のネームロケーション写真','御礼メッセージの送付'],
  f: ['YouTube概要欄へのお名前掲載','支援者ページへのお名前掲載','御礼メッセージの送付'],
  g: ['旅で使用した私物の欠片の送付','直筆のお手紙','御礼メッセージの送付'],
  h: ['ルート66からの直筆エアメール','限定ステッカーの送付','御礼メッセージの送付'],
  i: ['ゴッチさん・うすくくん・こたろうさんも参加のオンライン飲み会（2時間）へのご招待','御礼メッセージの送付'],
  j: ['ご自宅・お店への訪問','一緒に飲みに行く権利','御礼メッセージの送付'],
  k: ['出張講演（交通費込）','支援者ページへのお名前掲載','御礼メッセージの送付'],
  l: ['リヤカー本体の永久譲渡','支援者ページへのお名前掲載','御礼メッセージの送付'],
  m: ['食べ飲み放題懇親会へのご招待','御礼メッセージの送付'],
  n: ['プライベートキャンプ会へのご招待','御礼メッセージの送付'],
  o: ['マンツーマンレッスン（60分）','御礼メッセージの送付'],
  p: ['ウェア＆リヤカーへロゴ掲載（小）','YouTube概要欄クレジット','御礼メッセージの送付'],
  q: ['ウェア＆リヤカーへロゴ掲載（中）','YouTube概要欄掲載','SNSでのご紹介','御礼メッセージの送付'],
  r: ['ウェア＆リヤカー特等席にロゴ掲載（大）','YouTube概要欄トップに継続掲載','御礼メッセージの送付']
};

function getPlanKey(planStr) {
  var m = (planStr || '').match(/^([a-r])[：:]/i);
  return m ? m[1].toLowerCase() : null;
}

// プランごとのタスクリスト（コピペテンプレート付き）
// [{title, template}] の配列で返す。メール本文と「特典タスク」シートの両方で使う。
// isCard: クレカ決済（振込確認・LP掲載は自動で済んでいるため、その手順は出さない）
var NAME_DISPLAY_PLANS = {a:1, b:1, e:1, f:1, k:1, l:1, p:1, q:1, r:1};

function getPlanTasks(d, isCard) {
  var key      = getPlanKey(d.plan);
  var dispName = d['掲載希望名'] || d.name;
  var tasks    = [];
  var add = function(title, template) { tasks.push({title: title, template: template || ''}); };

  // プランが金額からの推定のとき（同額プランが複数ある）
  if (String(d.plan).indexOf('【要確認') >= 0) {
    add('Stripeダッシュボードで実際のプランを確認し、台帳のプラン欄を直す\n   → https://dashboard.stripe.com/payments');
  }
  // クレカ決済は掲載希望名を受け取っていない（Stripeのカード名義のまま）
  if (isCard && NAME_DISPLAY_PLANS[key]) {
    add('掲載・記入するお名前を本人に確認する（今はカード名義「' + d.name + '」のまま）\n   → 連絡先: ' + d.email);
  }

  if (key === 'a') {
    add('支援者ページにお名前が出ているか確認する');
  }
  if (key === 'b') {
    add('日本国旗の寄せ書きにお名前を記入する\n   → 記入名: ' + dispName);
  }
  if (key === 'c') {
    add('生電話の日程調整メールを送る\n   → 連絡先: ' + d.email);
  }
  if (key === 'd') {
    add('宿泊日程の調整メールを送る\n   → 連絡先: ' + d.email);
    add('ブリッジ（下関）の予約枠を確保する');
  }
  if (key === 'e') {
    add('現地で名前を書いた撮影を行う\n   → 記入名: ' + dispName);
    add('撮影した写真を送付する\n   → 送信先: ' + d.email);
  }
  if (key === 'f') {
    add('YouTube概要欄に追加する',
      '── サポーター ──\n' +
      dispName + '\n' +
      '──────────\n' +
      '↑ YouTubeの各動画の概要欄に追記してください'
    );
  }
  if (key === 'g') {
    add('旅で使用した私物の欠片を用意する（20個限定・在庫管理）');
    add('直筆のお手紙を書く\n   → 宛名: ' + dispName);
    add('発送する\n   → 送付先住所を確認: ' + d.email);
  }
  if (key === 'h') {
    add('現地からエアメールを投函する\n   → 宛名: ' + dispName);
    add('限定ステッカーを同封して発送する\n   → 送付先住所を確認: ' + d.email);
  }
  if (key === 'i') {
    add('オンライン飲み会の日程調整メールを送る（先着3名・ゴッチさん/うすくくん/こたろうさんも参加）\n   → 連絡先: ' + d.email);
  }
  if (key === 'j') {
    add('帰国後の日本縦断ルートと訪問希望地の照合・日程調整メールを送る\n   → 連絡先: ' + d.email);
  }
  if (key === 'k') {
    add('講演日程・会場・交通費の調整メールを送る\n   → 連絡先: ' + d.email);
  }
  if (key === 'l') {
    add('帰国・譲渡時期と受け渡し方法の調整メールを送る\n   → 連絡先: ' + d.email);
    add('リヤカーの譲渡・名義変更手続きを行う\n   → 宛名: ' + dispName);
  }
  if (key === 'm') {
    add('開催日程の調整メールを送る\n   → 連絡先: ' + d.email);
    add('ブリッジ（下関）の懇親会枠を確保する');
  }
  if (key === 'n') {
    add('キャンプ会の日程・場所の調整メールを送る\n   → 連絡先: ' + d.email);
  }
  if (key === 'o') {
    add('帰国後のレッスン日程調整メールを送る\n   → 連絡先: ' + d.email);
  }
  if (key === 'p') {
    add('ロゴデータをリクエストする\n   → 送付先: kimonomagician@gmail.com（申込者へ案内）');
    add('ウェア＆リヤカーにロゴ（小）を掲載する');
    add('YouTube概要欄にクレジットを追加する\n   → 掲載名: ' + dispName);
  }
  if (key === 'q') {
    add('ロゴデータをリクエストする\n   → 送付先: kimonomagician@gmail.com（申込者へ案内）');
    add('ウェア＆リヤカーにロゴ（中）を掲載する');
    add('YouTube概要欄に掲載する\n   → 掲載名: ' + dispName);
    add('SNSで紹介投稿する');
  }
  if (key === 'r') {
    add('ロゴデータをリクエストする\n   → 送付先: kimonomagician@gmail.com（申込者へ案内）');
    add('ウェア＆リヤカーの特等席にロゴ（大）を掲載する');
    add('YouTube概要欄トップに継続掲載する\n   → 掲載名: ' + dispName);
  }

  // 旧プラン（7月時点のLP）PLAN B — ジャーニースポンサー（¥30,000）
  if (/^PLAN B/.test(String(d.plan))) {
    add('Instagram・Xで支援者として紹介投稿する\n   → 紹介名: ' + dispName);
    add('各到達地点での活動報告投稿（旅の間ずっと続ける。仕組みを決めたら✓）');
    add('徒歩地球一周の公式サポーターとして支援者ページに名前が出ているか確認する');
    add('限定オープンチャットに招待する\n   → 連絡先: ' + d.email);
  }

  // 全プラン共通の特典「御礼メッセージの送付」
  add('御礼メッセージを送る\n   → 連絡先: ' + d.email);

  if (!isCard) {
    add('入金を確認したら、台帳の行を選んでメニュー「🌟 振込確認＋LP掲載」を実行');
  }
  return tasks;
}

function buildTaskBody(d, isCard) {
  var lines = [];
  getPlanTasks(d, isCard).forEach(function(t, i) {
    lines.push('□ ' + (i + 1) + '. ' + t.title);
    if (t.template) {
      lines.push('   ┌── コピペ用 ──────────────────');
      t.template.split('\n').forEach(function(l){ lines.push('   │ ' + l); });
      lines.push('   └──────────────────────────────');
    }
    lines.push('');
  });
  lines.push('▶ 終わったら「特典タスク」シートのチェックボックスに ✓ を入れてください');
  lines.push('   ' + taskSheetUrl_());
  return lines.join('\n');
}

// =============================================
//  特典タスク シート（チェックボックスで完了管理）
//  台帳 P列「特典対応」に進捗（未完了 1/3 / ✅ 完了）、Q列「管理ID」に紐付けIDを書く
// =============================================

var TASK_SHEET_NAME = '特典タスク';
var TASK_HEADERS    = ['完了','申込日時','お名前','メール','プラン','やること','コピペ用','完了日','管理ID'];
var LEDGER_STATUS_COL = 16; // P列：特典対応
var LEDGER_ID_COL     = 17; // Q列：管理ID

function taskSheet_(ss) {
  ss = ss || SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TASK_SHEET_NAME);
  if (sheet) return sheet;

  sheet = ss.insertSheet(TASK_SHEET_NAME);
  var h = sheet.getRange(1, 1, 1, TASK_HEADERS.length);
  h.setValues([TASK_HEADERS]).setFontWeight('bold').setBackground('#1a3a1a').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(6, 420);
  sheet.setColumnWidth(7, 260);
  sheet.getRange('F:G').setWrap(true);
  sheet.hideColumns(9);

  // 完了した行はグレー＋取り消し線
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$A2=TRUE')
    .setFontColor('#999999').setStrikethrough(true).setBackground('#eeeeee')
    .setRanges([sheet.getRange('A2:H')]).build();
  sheet.setConditionalFormatRules([rule]);

  // 台帳側の見出し
  var ledger = ss.getSheetByName('台帳');
  ledger.getRange(1, LEDGER_STATUS_COL).setValue('特典対応');
  ledger.getRange(1, LEDGER_ID_COL).setValue('管理ID');
  ledger.getRange(1, LEDGER_STATUS_COL, 1, 2)
    .setFontWeight('bold').setBackground('#1a3a1a').setFontColor('#ffffff');
  return sheet;
}

function taskSheetUrl_() {
  try {
    return sheetUrl() + '/edit#gid=' + taskSheet_().getSheetId();
  } catch (err) {
    return sheetUrl();
  }
}

// 台帳の1行分の特典タスクを「特典タスク」シートに追加する
function createBenefitTasks_(d, isCard, ledgerRow) {
  if (!SHEET_ID) return;
  var ss     = SpreadsheetApp.openById(SHEET_ID);
  var sheet  = taskSheet_(ss);
  var ledger = ss.getSheetByName('台帳');
  var id     = 'S' + Utilities.getUuid().replace(/-/g, '').slice(0, 10);
  var date   = d.date || new Date();

  var tasks = getPlanTasks(d, isCard);
  var rows  = tasks.map(function(t) {
    return [false, date, d.name || '', d.email || '', d.plan || '', t.title, t.template, '', id];
  });
  var start = sheet.getLastRow() + 1;
  sheet.getRange(start, 1, rows.length, TASK_HEADERS.length).setValues(rows);
  sheet.getRange(start, 1, rows.length, 1).insertCheckboxes();

  if (ledgerRow) {
    ledger.getRange(ledgerRow, LEDGER_ID_COL).setValue(id);
    ledger.getRange(ledgerRow, LEDGER_STATUS_COL).setValue('未完了 0/' + tasks.length);
  }
  return id;
}

// チェックボックスが押されたら完了日を記録し、台帳の「特典対応」を更新する
// （インストール型トリガー：setupBenefitTasks() で登録）
function onTaskEdit(e) {
  var sheet = e.range.getSheet();
  if (sheet.getName() !== TASK_SHEET_NAME) return;
  if (e.range.getColumn() !== 1) return;

  var ids = {};
  for (var r = e.range.getRow(); r <= e.range.getLastRow(); r++) {
    if (r <= 1) continue;
    var done = sheet.getRange(r, 1).getValue() === true;
    sheet.getRange(r, 8).setValue(done ? new Date() : '');
    var id = sheet.getRange(r, 9).getValue();
    if (id) ids[id] = true;
  }
  Object.keys(ids).forEach(function(id){ refreshLedgerStatus_(sheet.getParent(), id); });
}

function refreshLedgerStatus_(ss, id) {
  var tasks = ss.getSheetByName(TASK_SHEET_NAME).getDataRange().getValues();
  var total = 0, done = 0;
  for (var i = 1; i < tasks.length; i++) {
    if (tasks[i][8] !== id) continue;
    total++;
    if (tasks[i][0] === true) done++;
  }
  var status = (total && done === total) ? '✅ 完了' : '未完了 ' + done + '/' + total;

  var ledger = ss.getSheetByName('台帳');
  var ids    = ledger.getRange(1, LEDGER_ID_COL, ledger.getLastRow(), 1).getValues();
  for (var j = 1; j < ids.length; j++) {
    if (ids[j][0] === id) { ledger.getRange(j + 1, LEDGER_STATUS_COL).setValue(status); return; }
  }
}

// 台帳の行（1始まり）から特典タスクを作る。既に管理IDがあれば何もしない。
function createTasksForLedgerRow_(ledger, row) {
  var v = ledger.getRange(row, 1, 1, LEDGER_ID_COL).getValues()[0];
  if (!v[1] || v[LEDGER_ID_COL - 1]) return null;
  if (String(v[1]).indexOf('【テスト】') === 0) return null;
  var plan = String(v[3]);
  createBenefitTasks_({
    date: v[0], name: v[1], email: v[2], plan: plan, '掲載希望名': v[4]
  }, plan.indexOf('【クレカ決済】') >= 0, row);
  return v[1] + '（' + plan + '）';
}

// 初期セットアップ（1回だけ、Apps Scriptエディタで実行）
//  ・「特典タスク」シートを作る
//  ・チェックで完了を記録するトリガーを登録
//  ・BENEFIT_TASK_BACKFILL_FROM 以降の申込で、まだタスクが無いものを一括作成
var BENEFIT_TASK_BACKFILL_FROM = new Date('2026-09-21T00:00:00+09:00');

function setupBenefitTasks() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  taskSheet_(ss);

  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'onTaskEdit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onTaskEdit').forSpreadsheet(ss).onEdit().create();

  var ledger = ss.getSheetByName('台帳');
  var dates  = ledger.getRange(1, 1, ledger.getLastRow(), 1).getValues();
  var made   = [];
  for (var r = 2; r <= dates.length; r++) {
    var t = new Date(dates[r - 1][0]);
    if (isNaN(t) || t < BENEFIT_TASK_BACKFILL_FROM) continue;
    var m = createTasksForLedgerRow_(ledger, r);
    if (m) made.push(m);
  }
  Logger.log('特典タスク作成: ' + made.length + '件\n' + made.join('\n'));
  return made;
}

// 2026-10-03 山西さんの初期支援（旧PLAN B ジャーニースポンサー ¥30,000）を台帳に追加し、特典タスクを作る
// 支援日は台帳に記録が無いため、実行日時で記帳する。Apps Scriptエディタで1回だけ実行。
function addYamanishiSponsor_20261003() {
  var ss     = SpreadsheetApp.openById(SHEET_ID);
  var ledger = ss.getSheetByName('台帳');
  var rows   = ledger.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][2] === CC_EMAIL && String(rows[i][3]).indexOf('PLAN B') === 0) {
      Logger.log('既に台帳にあります（' + (i + 1) + '行目）');
      return;
    }
  }
  var d = {
    name: '山西 伸典', email: CC_EMAIL,
    plan: 'PLAN B（¥30,000）ジャーニースポンサー【初期支援・後日記帳】',
    '掲載希望名':'', 'Instagram':'', 'X(Twitter)':'',
    'ウェブサイトURL':'', '企業・活動紹介文':'', 'ブランドストーリー':'',
    'Powered_by表記':'', '応援メッセージ':''
  };
  var row = logToSheet(d, true); // 入金済み：振込確認＝確認済、LP掲載＝はい
  createBenefitTasks_(d, true, row);
  Logger.log('✅ 台帳 ' + row + '行目に追加し、特典タスクを作りました');
}

// メニュー：台帳で選んだ行（複数可）の特典タスクを作る
function menuCreateTasksForSelection() {
  var ui    = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== '台帳') { ui.alert('「台帳」シートで、タスクを作りたい行を選んでから実行してください'); return; }
  var range = sheet.getActiveRange();
  var made  = [];
  for (var r = range.getRow(); r <= range.getLastRow(); r++) {
    if (r <= 1) continue;
    var m = createTasksForLedgerRow_(sheet, r);
    if (m) made.push(m);
  }
  ui.alert(made.length
    ? '✅ ' + made.length + '件の特典タスクを作りました\n\n' + made.join('\n')
    : '新しく作るタスクはありませんでした（既に作成済みです）');
}

// =============================================
//  Jackへのタスクチェックリスト送信
// =============================================

function sendTaskChecklist(d) {
  var taskBody = buildTaskBody(d, false);

  var body = '━━━━━━━━━━━━━━━━━━━━\n';
  body += '【タスクリスト】' + d.plan + '\n';
  body += '申込者: ' + d.name + ' ／ ' + d.email + '\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n\n';
  body += taskBody + '\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n';
  body += 'スプレッドシート: ' + sheetUrl() + '\n';

  MailApp.sendEmail({
    to:      JACK_EMAIL,
    cc:      CC_EMAIL,
    subject: '【タスク】' + d.plan + ' — ' + d.name,
    body:    body
  });
}

// =============================================
//  スプレッドシート カスタムメニュー
// =============================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎩 Jack LP')
    .addItem('✅ 振込確認メールを送信（振込確認のみ）', 'menuSendConfirmation')
    .addItem('🌟 振込確認＋LP掲載（メーターに反映）', 'menuSendConfirmationAndPublish')
    .addSeparator()
    .addItem('📋 選んだ行の特典タスクを作る', 'menuCreateTasksForSelection')
    .addSeparator()
    .addItem('🔄 Stripe決済を今すぐ同期（漏れ回収）', 'menuSyncStripe')
    .addItem('🔑 Stripe読み取りキーを登録＋自動同期ON', 'menuSetupStripeSync')
    .addToUi();
}

// 振込確認のみ：M列（振込確認）を更新。N列（LP掲載）には触れない。
// クレカ決済の旧データなど「既にメーターの基準値に含まれている支援」を
// 二重カウントせずに確認だけ済ませたい場合に使う。
function menuSendConfirmation() {
  menuSendConfirmation_(false);
}

// 振込確認＋LP掲載：M列とN列（LP掲載＝はい）を同時に更新し、支援者一覧・
// メーターに即反映させる。通常の銀行振込の新規申し込みはこちらを使う。
function menuSendConfirmationAndPublish() {
  menuSendConfirmation_(true);
}

function menuSendConfirmation_(alsoPublish) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var row   = sheet.getActiveCell().getRow();
  if (row <= 1) { SpreadsheetApp.getUi().alert('データ行を選択してください'); return; }

  var data  = sheet.getRange(row, 1, 1, 15).getValues()[0];
  var name  = data[1], email = data[2], plan = data[3];

  if (!email) { SpreadsheetApp.getUi().alert('メールアドレスが空です'); return; }

  var msg = name + ' 様（' + plan + '）に振込確認メールを送信しますか？';
  if (alsoPublish) msg += '\n（LP掲載も「はい」にして、支援者一覧・メーターに反映します）';

  var confirm = SpreadsheetApp.getUi().alert(msg, SpreadsheetApp.getUi().ButtonSet.YES_NO);
  if (confirm !== SpreadsheetApp.getUi().Button.YES) return;

  sendPaymentConfirmation({name: name, email: email, plan: plan});

  sheet.getRange(row, 13).setValue('確認済');
  if (alsoPublish) sheet.getRange(row, 14).setValue('はい');
  sheet.getRange(row, 15).setValue(new Date());

  SpreadsheetApp.getUi().alert('✅ 送信完了しました！');
}

function sendPaymentConfirmation(d) {
  var key      = getPlanKey(d.plan);
  var benefits = key ? PLAN_BENEFITS[key] : [];

  var body = d.name + ' 様\n\n';
  body += 'ご入金を確認いたしました。\n';
  body += '改めて、応援ありがとうございます！\n\n';
  if (benefits && benefits.length) {
    body += '━━━━━━━━━━━━━━━━━━━━\n';
    body += '■ ご支援いただいた特典の実施について\n';
    body += '━━━━━━━━━━━━━━━━━━━━\n';
    benefits.forEach(function(b){ body += '  ✓ ' + b + '\n'; });
    body += '\n順次対応いたします。\n';
    body += 'ロゴデータ等が必要な特典は別途ご連絡します。\n\n';
  }
  body += 'ご不明な点は kimonomagician@gmail.com までご連絡ください。\n\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n';
  body += '一歩一歩、共に世界へ。\n\n';
  body += '着物マジシャン Jack（佐藤 基）\n';
  body += 'https://kimonomagician.jp\n';

  MailApp.sendEmail({
    to:      d.email,
    replyTo: JACK_EMAIL,
    subject: '【着物マジシャン Jack】ご入金確認のご連絡',
    body:    body
  });
}

// =============================================
//  自動返信メール（申し込み者向け）
// =============================================

function sendAutoReplyToApplicant(d) {
  var name     = d.name || 'スポンサー様';
  var plan     = d.plan || '';
  var key      = getPlanKey(plan);
  var benefits = key ? PLAN_BENEFITS[key] : null;

  var body = name + ' 様\n\n';
  body += 'この度は、着物マジシャン Jackの挑戦を応援してくださり、\n';
  body += '本当にありがとうございます！\n\n';
  body += '着物を身にまとい、世界を歩き続けるこの旅に、\n';
  body += 'あなたが加わってくださったこと、心から嬉しく思っています。\n\n';

  if (benefits) {
    body += '━━━━━━━━━━━━━━━━━━━━\n';
    body += '■ ' + plan + ' の特典はこちらです！\n';
    body += '━━━━━━━━━━━━━━━━━━━━\n';
    benefits.forEach(function(b){ body += '  ✓ ' + b + '\n'; });
    body += '\nご入金確認後、順次ご対応いたします。\n';
    body += '（ロゴ・紹介文等が必要な特典は別途ご連絡します）\n\n';
  }

  body += '━━━━━━━━━━━━━━━━━━━━\n';
  body += '■ 振込先情報\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n';
  body += '銀行名　：楽天銀行\n';
  body += '支店名　：コード支店\n';
  body += '口座種別：普通\n';
  body += '口座番号：2397746\n';
  body += '口座名義：サトウ　モトキ\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n\n';
  body += '【振込時のご注意】\n';
  body += '・振込名義は、お申し込みのお名前でお願いします。\n\n';
  body += 'ご不明な点がございましたら、このメールへ返信いただくか、\n';
  body += 'kimonomagician@gmail.com までご連絡ください。\n\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n';
  body += '一歩一歩、共に世界へ。\n\n';
  body += '着物マジシャン Jack（佐藤 基）\n';
  body += 'kimonomagician@gmail.com\n';
  body += 'https://kimonomagician.jp\n';

  MailApp.sendEmail({
    to: d.email, replyTo: JACK_EMAIL,
    subject: '【着物マジシャン Jack】応援ありがとうございます！特典・振込先のご案内',
    body: body
  });
}

// =============================================
//  Jack + やまちゃんへの通知メール
// =============================================

function sendNotificationToJack(d) {
  var fields = [
    ['申込者', d.name], ['メール', d.email], ['プラン', d.plan],
    ['掲載希望名', d['掲載希望名']], ['Instagram', d['Instagram']],
    ['X(Twitter)', d['X(Twitter)']], ['URL', d['ウェブサイトURL']],
    ['企業紹介文', d['企業・活動紹介文']], ['ブランドストーリー', d['ブランドストーリー']],
    ['Powered_by', d['Powered_by表記']], ['メッセージ', d['応援メッセージ']]
  ];
  var body = '新しいスポンサー申し込みがありました！\n\n━━━━━━━━━━━━━━━━━━━━\n';
  fields.forEach(function(f){ if (f[1]) body += f[0] + '：' + f[1] + '\n'; });
  body += '━━━━━━━━━━━━━━━━━━━━\n\nスプレッドシート: ' + sheetUrl();

  MailApp.sendEmail({
    to: JACK_EMAIL, cc: CC_EMAIL,
    subject: '【スポンサー申し込み】' + d.plan + ' — ' + d.name,
    body: body
  });
}

// =============================================
//  Stripe 決済完了通知
// =============================================

// GASのWebアプリは応答に必ず302リダイレクトを挟む仕様のため、Stripeからは
// 「配信失敗」と誤認識され続け、最大3日間・間隔を空けて同じイベントを再送してくる。
// event.id / Checkout Session ID を処理済みシートに記録し、既出なら何もせず終了する（多重記帳・多重メール防止）。
// さらに失敗扱いが続くとStripeはエンドポイント自体を自動で無効化する（2026-09-21以降の取りこぼしの原因）。
// そのため webhook だけに頼らず、syncStripePayments() が10分ごとにStripe APIから決済を取りに行く。
function processedSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName('Stripeイベント処理済み') || ss.insertSheet('Stripeイベント処理済み');
}

function isProcessedId_(id) {
  if (!id) return false;
  var sheet = processedSheet_();
  var ids   = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) return true;
  }
  return false;
}

function markProcessedId_(id) {
  if (id) processedSheet_().appendRow([id, new Date()]);
}

function isDuplicateStripeEvent(eventId) {
  if (!eventId) return false;
  if (isProcessedId_(eventId)) return true;
  markProcessedId_(eventId);
  return false;
}

function handleStripeWebhook(event) {
  if (event.type !== 'checkout.session.completed' &&
      event.type !== 'payment_intent.succeeded') {
    return res({ok: true, skipped: true});
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (isDuplicateStripeEvent(event.id)) {
      return res({ok: true, duplicate: true});
    }
    var obj = event.data.object;
    // 自動同期で既に記帳済みのセッションなら記帳しない
    if (obj.id && isProcessedId_(obj.id)) {
      return res({ok: true, duplicate: true});
    }
    markProcessedId_(obj.id);
    recordCardPayment_(obj, false);
  } finally {
    lock.releaseLock();
  }
  return res({ok: true});
}

// Checkout Session（または PaymentIntent）1件を台帳に記帳し、Jackへ通知する
function recordCardPayment_(obj, fromSync) {
  var detail = obj.customer_details || {};
  var cName  = detail.name  || '不明';
  var cEmail = detail.email || '不明';
  var amount = obj.amount_total || obj.amount_received || 0;

  // LPの決済ボタンが付与する client_reference_id（プラン記号 a-r）を優先。
  // 無い場合は金額から推定（¥10,000・¥30,000は同額プランが複数あるため要確認扱い）。
  var refKey = (obj.client_reference_id || '').toLowerCase();
  var amountFallback = {1000:'a', 3000:'b', 5000:'d', 10000:'f', 15000:'g', 20000:'i', 30000:'h', 50000:'j', 100000:'k', 300000:'q', 500000:'l'};
  var ambiguousAmounts = {
    3000:'b 国旗 or c ラジオ生電話 or e ネームロケーション写真',
    10000:'f 名前を刻む or m ブリッジ懇親会',
    30000:'h ルート66 or n キャンプ会 or o マンツーマンレッスン',
    100000:'k 出張講演会 or p ロゴ小+概要欄クレジット',
    500000:'l 相棒譲渡 or r メインパートナーコース'
  };
  var key  = PLAN_NAMES[refKey] ? refKey : (amountFallback[amount] || '');
  var plan = key ? (key + '：' + PLAN_NAMES[key] + '（¥' + amount.toLocaleString() + '）') : ('¥' + amount.toLocaleString());
  if (!PLAN_NAMES[refKey] && ambiguousAmounts[amount]) plan += '【要確認：' + ambiguousAmounts[amount] + '】';
  if (!key) plan += '【要確認：プラン不明（複数口の可能性）】';

  // スプレッドシートに記録（クレカ決済は即時確定するため自動で振込確認済・LP掲載＝はい）
  var d = {
    date: obj.created ? new Date(obj.created * 1000) : new Date(),
    name: cName, email: cEmail, plan: plan + '【クレカ決済】',
    '掲載希望名':'', 'Instagram':'', 'X(Twitter)':'',
    'ウェブサイトURL':'', '企業・活動紹介文':'', 'ブランドストーリー':'',
    'Powered_by表記':'', '応援メッセージ':''
  };
  var ledgerRow = logToSheet(d, true);
  try { createBenefitTasks_(d, true, ledgerRow); } catch (err) {} // 失敗しても通知メールは必ず送る

  var body = '【クレカ決済完了】スポンサー申し込みがありました！\n\n';
  if (fromSync) body += '※Stripeとの自動同期で記帳しました（決済日時は台帳の申込日時をご確認ください）\n\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n';
  body += '顧客名：' + cName + '\n';
  body += 'メール：' + cEmail + '\n';
  body += '金　額：¥' + amount.toLocaleString() + '\n';
  body += 'プラン：' + plan + '\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n\n';
  body += '■ やること（特典タスク）\n\n';
  body += buildTaskBody(d, true) + '\n\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n';
  body += '▶ Stripeダッシュボード: https://dashboard.stripe.com/payments\n';
  body += '▶ スプレッドシート: ' + sheetUrl();

  MailApp.sendEmail({
    to: JACK_EMAIL, cc: CC_EMAIL,
    subject: '【クレカ決済完了】' + cName + ' 様（¥' + amount.toLocaleString() + '）',
    body: body
  });
}

// =============================================
//  Stripe 自動同期（webhook取りこぼし対策）
//  スクリプトプロパティ STRIPE_READ_KEY に「制限付きキー（Checkout Sessions：読み取り）」を登録して使う
// =============================================

// 初回はこの時刻以降に作成されたセッションを取り込む（webhook最終受信 2026-09-21 22:45 の当日から）
var STRIPE_SYNC_BACKFILL_FROM = new Date('2026-09-21T00:00:00+09:00');

function syncStripePayments() {
  var props  = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('STRIPE_READ_KEY');
  if (!apiKey) throw new Error('STRIPE_READ_KEY が未登録です（メニュー「Stripe読み取りキーを登録」から登録）');

  var since = Number(props.getProperty('STRIPE_SYNC_SINCE')) ||
              Math.floor(STRIPE_SYNC_BACKFILL_FROM.getTime() / 1000);
  var startedAt = Math.floor(Date.now() / 1000);

  var sessions = [];
  var startingAfter = '';
  do {
    var url = 'https://api.stripe.com/v1/checkout/sessions?limit=100&status=complete&created[gte]=' + since +
              (startingAfter ? '&starting_after=' + startingAfter : '');
    var resp = UrlFetchApp.fetch(url, {
      headers: {Authorization: 'Bearer ' + apiKey},
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) {
      throw new Error('Stripe API エラー ' + resp.getResponseCode() + ': ' + resp.getContentText());
    }
    var page = JSON.parse(resp.getContentText());
    sessions = sessions.concat(page.data);
    startingAfter = page.has_more && page.data.length ? page.data[page.data.length - 1].id : '';
  } while (startingAfter);

  // 古い順に記帳する
  sessions.sort(function(a, b){ return a.created - b.created; });

  var added = [];
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ledger = SpreadsheetApp.openById(SHEET_ID).getSheetByName('台帳').getDataRange().getValues();
    sessions.forEach(function(s) {
      if (s.payment_status !== 'paid') return;
      if (isProcessedId_(s.id)) return;
      // 処理済みシートにセッションIDが無い過去分（event.idでのみ記録）は、台帳のメール＋金額＋日時で照合する
      if (existsInLedger_(ledger, s)) { markProcessedId_(s.id); return; }
      markProcessedId_(s.id);
      recordCardPayment_(s, true);
      added.push(((s.customer_details && s.customer_details.name) || '不明') + ' ¥' + (s.amount_total || 0).toLocaleString());
    });
  } finally {
    lock.releaseLock();
  }

  // 次回は直近48時間分だけ見る（Checkout Sessionの有効期限は24時間なので取りこぼさない）
  props.setProperty('STRIPE_SYNC_SINCE', String(startedAt - 2 * 86400));
  Logger.log('Stripe同期: 取得 ' + sessions.length + '件 / 新規記帳 ' + added.length + '件 ' + added.join(', '));
  return added;
}

// 台帳に同じメール・同じ金額の【クレカ決済】行が、セッション作成〜24時間以内の日時で既にあるか
function existsInLedger_(ledger, s) {
  var email   = ((s.customer_details && s.customer_details.email) || '').toLowerCase();
  var amount  = s.amount_total || 0;
  var created = s.created * 1000;
  for (var i = 1; i < ledger.length; i++) {
    var row = ledger[i];
    if (String(row[2]).toLowerCase() !== email) continue;
    var plan = String(row[3]);
    if (plan.indexOf('【クレカ決済】') < 0) continue;
    var m = plan.match(/¥([\d,]+)/);
    if (!m || Number(m[1].replace(/,/g, '')) !== amount) continue;
    var t = new Date(row[0]).getTime();
    if (t >= created - 10 * 60000 && t <= created + 24 * 3600000) return true;
  }
  return false;
}

function installStripeSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'syncStripePayments') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncStripePayments').timeBased().everyMinutes(10).create();
}

function menuSetupStripeSync() {
  var ui = SpreadsheetApp.getUi();
  var r  = ui.prompt('Stripe読み取りキーの登録',
    'Stripeの制限付きキー（rk_live_ で始まる）を貼り付けてください。\n権限は「Checkout Sessions：読み取り」のみでOKです。',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var key = r.getResponseText().trim();
  if (key.indexOf('rk_') !== 0) { ui.alert('rk_ で始まる制限付きキーを貼り付けてください（sk_ の秘密キーは使いません）'); return; }
  PropertiesService.getScriptProperties().setProperty('STRIPE_READ_KEY', key);
  installStripeSyncTrigger();
  ui.alert('✅ 登録しました。10分ごとの自動同期をONにしました。\n続けて「Stripe決済を今すぐ同期」で漏れ分を回収してください。');
}

function menuSyncStripe() {
  var added = syncStripePayments();
  SpreadsheetApp.getUi().alert(added.length
    ? '✅ ' + added.length + '件を台帳に追加しました\n\n' + added.join('\n')
    : '漏れはありませんでした（台帳とStripeは一致しています）');
}
