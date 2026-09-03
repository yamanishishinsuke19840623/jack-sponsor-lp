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
    logToSheet(data);
    sendAutoReplyToApplicant(data);
    sendTaskChecklist(data);
    return res({ok: true});
  } catch (err) {
    return res({ok: false, error: err.toString()});
  }
}

// LP からスポンサー一覧を取得（JSONP）
function doGet(e) {
  var callback = (e.parameter || {}).callback || 'cb';
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

function res(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheetUrl() {
  return SHEET_ID ? 'https://docs.google.com/spreadsheets/d/' + SHEET_ID : '（SHEET_ID未設定）';
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

function logToSheet(d) {
  if (!SHEET_ID) return;
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('台帳');
  sheet.appendRow([
    new Date(),
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
    '未確認',
    'いいえ',
    ''
  ]);
}

// =============================================
//  プラン別 特典・タスク定義
// =============================================

var PLAN_NAMES = {
  a: 'コーヒー1杯のエール', b: '旅の相棒（日本国旗へお名前記入）', d: '旅の拠点に泊まる（ブリッジ宿泊）',
  f: '名前を刻む（YouTube概要欄）', g: '荒野の生還パーツ', h: 'アメリカからの生還（ルート66）',
  i: 'レジェンド集結（オンライン飲み会）', j: 'あなたの街に直撃！', k: '出張講演会プラン'
};

var PLAN_BENEFITS = {
  a: ['御礼メッセージの送付','支援者ページへのお名前掲載'],
  b: ['日本国旗へお名前記入（寄せ書き）','御礼メッセージの送付'],
  d: ['ブリッジ（下関）宿泊1泊','御礼メッセージの送付'],
  f: ['YouTube概要欄へのお名前掲載','支援者ページへのお名前掲載','御礼メッセージの送付'],
  g: ['旅で使用した私物の欠片の送付','直筆のお手紙','御礼メッセージの送付'],
  h: ['ルート66からの直筆エアメール','限定ステッカーの送付','御礼メッセージの送付'],
  i: ['ゴッチさん・うすくくん・こたろうさんも参加のオンライン飲み会（2時間）へのご招待','御礼メッセージの送付'],
  j: ['ご自宅・お店への訪問','一緒に飲みに行く権利','御礼メッセージの送付'],
  k: ['出張講演（交通費込）','支援者ページへのお名前掲載','御礼メッセージの送付']
};

function getPlanKey(planStr) {
  var m = (planStr || '').match(/^([a-k])[：:]/i);
  return m ? m[1].toLowerCase() : null;
}

// プランごとのタスクリスト（コピペテンプレート付き）
function buildTaskBody(d) {
  var key      = getPlanKey(d.plan);
  var dispName = d['掲載希望名'] || d.name;

  var lines = [];

  var add = function(num, title, template) {
    lines.push('□ ' + num + '. ' + title);
    if (template) {
      lines.push('   ┌── コピペ用 ──────────────────');
      template.split('\n').forEach(function(l){ lines.push('   │ ' + l); });
      lines.push('   └──────────────────────────────');
    }
    lines.push('');
  };

  // === a: コーヒー1杯のエール ===
  if (key === 'a') {
    add(1, 'スプレッドシートで「振込確認」→「確認済」に変更', null);
    add(2, 'LP掲載欄を「いいえ」→「はい」に変更（支援者ページに名前が自動表示）', null);
  }

  // === b: 旅の相棒（日本国旗へお名前記入） ===
  if (key === 'b') {
    add(1, '日本国旗の寄せ書きにお名前を記入する\n   → 記入名: ' + dispName, null);
    add(2, 'スプレッドシートで「振込確認」→「確認済」・「LP掲載」→「はい」に変更', null);
  }

  // === d: 旅の拠点に泊まる（ブリッジ宿泊） ===
  if (key === 'd') {
    add(1, '宿泊日程の調整メールを送る\n   → 連絡先: ' + d.email, null);
    add(2, 'ブリッジ（下関）の予約枠を確保する', null);
    add(3, 'スプレッドシートで「振込確認」→「確認済」・「LP掲載」→「はい」に変更', null);
  }

  // === f: 名前を刻む（YouTube概要欄） ===
  if (key === 'f') {
    add(1, 'YouTube概要欄に追加する',
      '── サポーター ──\n' +
      dispName + '\n' +
      '──────────\n' +
      '↑ YouTubeの各動画の概要欄に追記してください'
    );
    add(2, 'スプレッドシートで「振込確認」→「確認済」・「LP掲載」→「はい」に変更', null);
  }

  // === g: 荒野の生還パーツ ===
  if (key === 'g') {
    add(1, '旅で使用した私物の欠片を用意する（20個限定・在庫管理）', null);
    add(2, '直筆のお手紙を書く\n   → 宛名: ' + dispName, null);
    add(3, '発送する\n   → 送付先住所を確認: ' + d.email, null);
    add(4, 'スプレッドシートで「振込確認」→「確認済」・「LP掲載」→「はい」に変更', null);
  }

  // === h: アメリカからの生還（ルート66） ===
  if (key === 'h') {
    add(1, '現地からエアメールを投函する\n   → 宛名: ' + dispName, null);
    add(2, '限定ステッカーを同封して発送する\n   → 送付先住所を確認: ' + d.email, null);
    add(3, 'スプレッドシートで「振込確認」→「確認済」・「LP掲載」→「はい」に変更', null);
  }

  // === i: レジェンド集結（オンライン飲み会） ===
  if (key === 'i') {
    add(1, 'オンライン飲み会の日程調整メールを送る（先着3名・ゴッチさん/うすくくん/こたろうさんも参加）\n   → 連絡先: ' + d.email, null);
    add(2, 'スプレッドシートで「振込確認」→「確認済」・「LP掲載」→「はい」に変更', null);
  }

  // === j: あなたの街に直撃！ ===
  if (key === 'j') {
    add(1, '帰国後の日本縦断ルートと訪問希望地の照合・日程調整メールを送る\n   → 連絡先: ' + d.email, null);
    add(2, 'スプレッドシートで「振込確認」→「確認済」・「LP掲載」→「はい」に変更', null);
  }

  // === k: 出張講演会プラン ===
  if (key === 'k') {
    add(1, '講演日程・会場・交通費の調整メールを送る\n   → 連絡先: ' + d.email, null);
    add(2, 'スプレッドシートで「振込確認」→「確認済」・「LP掲載」→「はい」に変更', null);
  }

  return lines.join('\n');
}

// =============================================
//  Jackへのタスクチェックリスト送信
// =============================================

function sendTaskChecklist(d) {
  var taskBody = buildTaskBody(d);
  if (!taskBody) return;

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
    .addItem('✅ 振込確認メールを送信（選択行）', 'menuSendConfirmation')
    .addToUi();
}

function menuSendConfirmation() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var row   = sheet.getActiveCell().getRow();
  if (row <= 1) { SpreadsheetApp.getUi().alert('データ行を選択してください'); return; }

  var data  = sheet.getRange(row, 1, 1, 15).getValues()[0];
  var name  = data[1], email = data[2], plan = data[3];

  if (!email) { SpreadsheetApp.getUi().alert('メールアドレスが空です'); return; }

  var confirm = SpreadsheetApp.getUi().alert(
    name + ' 様（' + plan + '）に振込確認メールを送信しますか？',
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  if (confirm !== SpreadsheetApp.getUi().Button.YES) return;

  sendPaymentConfirmation({name: name, email: email, plan: plan});

  sheet.getRange(row, 13).setValue('確認済');
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

function handleStripeWebhook(event) {
  if (event.type !== 'checkout.session.completed' &&
      event.type !== 'payment_intent.succeeded') {
    return res({ok: true, skipped: true});
  }
  var obj    = event.data.object;
  var detail = obj.customer_details || {};
  var cName  = detail.name  || '不明';
  var cEmail = detail.email || '不明';
  var amount = obj.amount_total || obj.amount_received || 0;

  // LPの決済ボタンが付与する client_reference_id（プラン記号 a-k）を優先。
  // 無い場合は金額から推定（¥30,000は h/i の2プランがあるため要確認扱い）。
  var refKey = (obj.client_reference_id || '').toLowerCase();
  var amountFallback = {1000:'a', 3000:'b', 5000:'d', 10000:'f', 15000:'g', 30000:'h', 50000:'j', 100000:'k'};
  var key  = PLAN_NAMES[refKey] ? refKey : (amountFallback[amount] || '');
  var plan = key ? (key + '：' + PLAN_NAMES[key] + '（¥' + amount.toLocaleString() + '）') : ('¥' + amount.toLocaleString());
  if (amount === 30000 && !PLAN_NAMES[refKey]) plan += '【要確認：h ルート66 or i オンライン飲み会】';

  // スプレッドシートに記録
  logToSheet({
    name: cName, email: cEmail, plan: plan + '【クレカ決済】',
    '掲載希望名':'', 'Instagram':'', 'X(Twitter)':'',
    'ウェブサイトURL':'', '企業・活動紹介文':'', 'ブランドストーリー':'',
    'Powered_by表記':'', '応援メッセージ':''
  });

  var body = '【クレカ決済完了】スポンサー申し込みがありました！\n\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n';
  body += '顧客名：' + cName + '\n';
  body += 'メール：' + cEmail + '\n';
  body += '金　額：¥' + amount.toLocaleString() + '\n';
  body += 'プラン：' + plan + '\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n\n';
  body += '▶ Stripeダッシュボード: https://dashboard.stripe.com/payments\n';
  body += '▶ スプレッドシート: ' + sheetUrl();

  MailApp.sendEmail({
    to: JACK_EMAIL, cc: CC_EMAIL,
    subject: '【クレカ決済完了】' + cName + ' 様（¥' + amount.toLocaleString() + '）',
    body: body
  });
  return res({ok: true});
}
