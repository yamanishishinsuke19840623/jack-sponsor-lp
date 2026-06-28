// =============================================
//  着物マジシャン Jack LP スポンサーフォーム
//  自動返信 + 通知メール GAS
// =============================================

var JACK_EMAIL = 'kimonomagician@gmail.com';
var CC_EMAIL   = 'yamanishishinsuke19840623@gmail.com';

// ---- エントリーポイント ----
function doPost(e) {
  var data = {};
  try {
    data = JSON.parse(e.postData.contents);
    sendNotificationToJack(data);
    sendAutoReplyToApplicant(data);
    return res({ok: true});
  } catch (err) {
    return res({ok: false, error: err.toString()});
  }
}

function res(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- Jack + やまちゃんへの通知メール ----
function sendNotificationToJack(d) {
  var fields = [
    ['申込者',             d.name],
    ['メール',             d.email],
    ['プラン',             d.plan],
    ['掲載希望名',         d['掲載希望名']],
    ['Instagram',          d['Instagram']],
    ['X(Twitter)',         d['X(Twitter)']],
    ['ウェブサイトURL',    d['ウェブサイトURL']],
    ['企業・活動紹介文',   d['企業・活動紹介文']],
    ['ブランドストーリー', d['ブランドストーリー']],
    ['Powered_by表記',     d['Powered_by表記']],
    ['応援メッセージ',     d['応援メッセージ']]
  ];

  var body = '新しいスポンサー申し込みが届きました！\n\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n';
  fields.forEach(function(f) {
    if (f[1]) body += f[0] + '：' + f[1] + '\n';
  });
  body += '━━━━━━━━━━━━━━━━━━━━\n\n';
  body += '振込確認後、掲載・特典対応をお願いします。\n';
  body += 'https://kimonomagician.jp';

  MailApp.sendEmail({
    to:      JACK_EMAIL,
    cc:      CC_EMAIL,
    subject: '【スポンサー申し込み】' + d.plan + ' — ' + d.name,
    body:    body
  });
}

// ---- 申し込み者への自動返信メール ----
function sendAutoReplyToApplicant(d) {
  var name = d.name || 'スポンサー様';
  var plan = d.plan || 'ご選択のプラン';

  var body = name + ' 様\n\n';
  body += 'この度は、着物マジシャン Jackの挑戦を応援してくださり、\n';
  body += '本当にありがとうございます！\n\n';
  body += '着物を身にまとい、世界を歩き続けるこの旅に、\n';
  body += 'あなたが加わってくださったこと、心から嬉しく思っています。\n\n';
  body += 'ご支援いただいたサポートは、旅の装備・映像制作・\n';
  body += '活動発信に大切に使わせていただきます。\n\n';
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
  body += '・振込名義は、お申し込みのお名前でお願いします。\n';
  body += '・ご入金確認後、' + plan + ' の特典について\n';
  body += '  別途ご案内メールをお送りします。\n\n';
  body += 'ご不明な点がございましたら、このメールへ返信いただくか、\n';
  body += 'kimonomagician@gmail.com までご連絡ください。\n\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n';
  body += '一歩一歩、共に世界へ。\n\n';
  body += '着物マジシャン Jack（佐藤 基）\n';
  body += 'kimonomagician@gmail.com\n';
  body += 'https://kimonomagician.jp\n';
  body += '━━━━━━━━━━━━━━━━━━━━\n';

  MailApp.sendEmail({
    to:      d.email,
    replyTo: JACK_EMAIL,
    subject: '【着物マジシャン Jack】応援ありがとうございます！振込先のご案内',
    body:    body
  });
}
