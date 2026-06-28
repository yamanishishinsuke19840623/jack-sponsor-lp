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

var PLAN_BENEFITS = {
  A: ['御礼メッセージの送付','支援者ページへのお名前掲載','徒歩地球一周挑戦メンバーとして永久記録'],
  B: ['Instagram・Xでのスポンサー紹介投稿','各到達地点での活動報告投稿','徒歩地球一周公式サポーターとして掲載','限定オープンチャットへのご招待'],
  C: ['YouTube概要欄への社名掲載（半年間）','旅の限定写真・動画の共有','活動レポートの定期共有'],
  D: ['日本国旗へのロゴ・お名前掲載','YouTube動画内での企業・活動紹介','スタート・各地到達・ゴール時の継続露出','アメリカ各地の絶景地からの限定動画送付','SNSでの特別スポンサー紹介'],
  E: ['リヤカー側面へのロゴ掲載','YouTube動画内での企業・活動紹介','ドキュメンタリー作品への企業名掲載','ブランド紹介を含む映像演出','YouTubeエンディングでの社名・ロゴ掲載','特別スポンサーインタビューへの対応'],
  F: ['ウェア最上部へのロゴ掲載','リヤカー大型ロゴ掲載','動画内での継続的な特別紹介','ゴール達成時のメインクレジット掲載','4K映像素材の商用二次利用（無制限・永続）','「Powered by ○○」表記対応','徒歩地球一周プロジェクト公式パートナー認定','特別協賛企業として各媒体で優先掲載']
};

function getPlanKey(planStr) {
  var m = (planStr || '').match(/PLAN\s*([A-F])/i);
  return m ? m[1].toUpperCase() : null;
}

// プランごとのタスクリスト（コピペテンプレート付き）
function buildTaskBody(d) {
  var key     = getPlanKey(d.plan);
  var dispName = d['掲載希望名'] || d.name;
  var ig      = d['Instagram']  || '';
  var tw      = d['X(Twitter)'] || '';
  var url     = d['ウェブサイトURL'] || '';
  var intro   = d['企業・活動紹介文'] || '';

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

  // === PLAN A ===
  if (key === 'A') {
    add(1, 'スプレッドシートで「振込確認」→「確認済」に変更', null);
    add(2, 'LP掲載欄を「いいえ」→「はい」に変更（支援者ページに名前が自動表示）', null);
  }

  // === PLAN B ===
  if (key === 'B') {
    add(1, 'Instagram で紹介投稿',
      '🎉 新しいジャーニースポンサーが加わりました！\n' +
      dispName + ' さん' + (ig ? ' ' + ig : '') + '、ありがとうございます！🙏\n' +
      'あなたのサポートが、Jackをアメリカ大陸へ運びます✨\n' +
      '#着物マジシャンジャック #徒歩地球一周 #スポンサー'
    );
    add(2, 'X（Twitter）で紹介投稿',
      '🎉 ' + dispName + ' さん' + (tw ? ' ' + tw : '') + ' が\n' +
      'ジャーニースポンサーとして応援してくださいました！\n' +
      'ありがとうございます🙏 一緒に世界を歩きましょう！\n' +
      '#着物マジシャンジャック #徒歩地球一周'
    );
    add(3, 'LINEオープンチャットに招待する\n   → 招待先: ' + d.email, null);
    add(4, 'スプレッドシートで「振込確認」→「確認済」・「LP掲載」→「はい」に変更', null);
  }

  // === PLAN C ===
  if (key === 'C') {
    add(1, 'YouTube概要欄に追加する（6ヶ月掲載）',
      '── スポンサー ──\n' +
      dispName + (url ? ' ' + url : '') + '\n' +
      '──────────\n' +
      '↑ YouTubeの各動画の概要欄に追記してください'
    );
    add(2, '限定写真・動画を共有する\n   → 送信先: ' + d.email, null);
    add(3, '活動レポートを月1回送付する（送信先: ' + d.email + '）', null);
    add(4, 'スプレッドシートで「振込確認」→「確認済」・「LP掲載」→「はい」に変更', null);
  }

  // === PLAN D ===
  if (key === 'D') {
    add(1, 'ロゴデータをリクエストする',
      d.name + ' 様\n\n' +
      'この度はフラッグスポンサーとしてご支援いただき、\n' +
      '誠にありがとうございます！\n\n' +
      '日本国旗・SNSへのロゴ掲載のため、\n' +
      'ロゴ画像（PNG・AI・SVG推奨）を\n' +
      'kimonomagician@gmail.com までお送りください。\n\n' +
      '着物マジシャン Jack'
    );
    add(2, '日本国旗にロゴ・名前を手書き or 印刷して貼る\n   → 掲載名: ' + dispName, null);
    add(3, 'YouTube動画内で紹介する',
      '【企業紹介テキスト】\n' + (intro || dispName + ' 様にご支援いただいています。ありがとうございます！') + '\n' +
      (url ? 'URL: ' + url : '')
    );
    add(4, 'SNSで特別スポンサー紹介投稿',
      '🏳️ フラッグスポンサー紹介 🏳️\n' +
      dispName + ' さん' + (ig ? ' ' + ig : '') + ' に\n' +
      '特別スポンサーとしてご支援いただいています！\n' +
      '日本国旗を背負ってアメリカを歩きます🇯🇵\n' +
      '#着物マジシャンジャック #フラッグスポンサー'
    );
    add(5, 'スプレッドシートで「振込確認」→「確認済」・「LP掲載」→「はい」に変更', null);
  }

  // === PLAN E ===
  if (key === 'E') {
    add(1, 'ロゴデータをリクエストする',
      d.name + ' 様\n\n' +
      'この度はドキュメンタリースポンサーとしてご支援いただき、\n' +
      '誠にありがとうございます！\n\n' +
      'リヤカー・YouTube エンディングへのロゴ掲載のため、\n' +
      'ロゴ画像（PNG・AI・SVG推奨）を\n' +
      'kimonomagician@gmail.com までお送りください。\n' +
      '映像演出の詳細は別途ご相談させてください。\n\n' +
      '着物マジシャン Jack'
    );
    add(2, 'リヤカー側面にロゴを貼る\n   → ロゴデータ届いたら実施', null);
    add(3, 'YouTubeエンディングに社名・ロゴを追加する\n   → 掲載名: ' + dispName, null);
    add(4, 'スポンサーインタビューの日程を調整する\n   → 連絡先: ' + d.email, null);
    add(5, 'スプレッドシートで「振込確認」→「確認済」・「LP掲載」→「はい」に変更', null);
  }

  // === PLAN F ===
  if (key === 'F') {
    var powered = d['Powered_by表記'] || dispName;
    add(1, 'ロゴデータ・「Powered by」表記をリクエストする',
      d.name + ' 様\n\n' +
      'この度はレジェンドスポンサーとしてご支援いただき、\n' +
      '誠にありがとうございます！\n\n' +
      'ウェア・リヤカー・動画への掲載のため、\n' +
      'ロゴ画像（PNG・AI・SVG推奨）を\n' +
      'kimonomagician@gmail.com までお送りください。\n' +
      '「Powered by ' + powered + '」の表記確認もお願いします。\n\n' +
      '着物マジシャン Jack'
    );
    add(2, 'ウェア最上部にロゴを掲載する', null);
    add(3, 'リヤカーに大型ロゴを掲載する', null);
    add(4, '4K映像素材の共有方法を案内する\n   → 送信先: ' + d.email, null);
    add(5, '公式パートナー認定書を作成して送付する\n   → 宛名: ' + dispName, null);
    add(6, '動画・各媒体に「Powered by ' + powered + '」表記を追加する', null);
    add(7, 'スプレッドシートで「振込確認」→「確認済」・「LP掲載」→「はい」に変更', null);
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

  // 金額からプランを推定
  var planMap = {10000:'PLAN A（¥10,000）エールスポンサー', 30000:'PLAN B（¥30,000）ジャーニースポンサー',
    50000:'PLAN C（¥50,000）メディアスポンサー', 100000:'PLAN D（¥100,000）フラッグスポンサー',
    300000:'PLAN E（¥300,000）ドキュメンタリースポンサー', 1000000:'PLAN F（¥1,000,000）レジェンドスポンサー'};
  var plan = planMap[amount] || '¥' + amount.toLocaleString();

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
