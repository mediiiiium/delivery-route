/**
 * GAS Webアプリ：メッセージから店名抽出 → ルート計算
 */

// ─── Webアプリのエントリーポイント ────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('配達ルート作成')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── Claude APIで店名を抽出（セクション分割対応） ──────────────
function extractShopsFromMessage(message) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) throw new Error('ANTHROPIC_API_KEY が設定されていません');

  // りくり・はいたつ でセクションを分割
  var rikuriIdx = message.indexOf('りくり');
  var haitasuIdx = message.indexOf('はいたつ');

  var rikuriText = '';
  var haitasuText = '';

  if (rikuriIdx !== -1 && haitasuIdx !== -1) {
    if (rikuriIdx < haitasuIdx) {
      rikuriText  = message.substring(rikuriIdx + 3, haitasuIdx).trim();
      haitasuText = message.substring(haitasuIdx + 4).trim();
    } else {
      haitasuText = message.substring(haitasuIdx + 4, rikuriIdx).trim();
      rikuriText  = message.substring(rikuriIdx + 3).trim();
    }
  } else if (rikuriIdx !== -1) {
    rikuriText  = message.substring(rikuriIdx + 3).trim();
  } else if (haitasuIdx !== -1) {
    haitasuText = message.substring(haitasuIdx + 4).trim();
  }

  var results = [];

  // りくりセクション → @RIKRI BREWING + *店名
  if (rikuriText) {
    var rikuriShops = callClaude_(rikuriText, apiKey);
    var withStar = rikuriShops.split('\n')
      .map(function(l) { return l.trim(); })
      .filter(function(l) { return l && l !== 'なし'; })
      .map(function(l) { return (l.startsWith('@') || l.startsWith('*')) ? l : '*' + l; })
      .join('\n');
    results.push('@RIKRI BREWING');
    if (withStar) results.push(withStar);
  }

  // はいたつセクション → 箱舟 + 店名
  if (haitasuText) {
    var haitasuShops = callClaude_(haitasuText, apiKey);
    var shops = haitasuShops.split('\n')
      .map(function(l) { return l.trim(); })
      .filter(function(l) { return l && l !== 'なし'; })
      .join('\n');
    results.push('箱舟（はこぶね）');
    if (shops) results.push(shops);
  }

  return results.join('\n');
}

// ─── Claude API呼び出し（内部共通処理） ────────────────────────
function callClaude_(text, apiKey) {

  var prompt = [
    'あなたはクラフトビール配達業務のアシスタントです。',
    '以下のメッセージから配達先の店舗名だけを抽出してください。',
    '',
    'メッセージには以下のフォーマットがあります：',
    '',
    '【@ の使われ方は2種類ある。必ず区別すること】',
    '  ① 行の先頭に @ → ピックアップ優先マーク。@ごと抽出する',
    '     例：「@RIKRI BREWING」→「@RIKRI BREWING」',
    '  ② 店舗名の後ろに @場所 → 場所を示すだけ。@以降を削除して店舗名だけ抽出',
    '     例：「アイリッシュパブタラモア@代々木八幡」→「アイリッシュパブタラモア」',
    '     例：「天沼酒場@荻窪」→「天沼酒場」',
    '',
    '【フォーマットA：箇条書き形式】',
    '  ・店舗名@場所　※メモ',
    '  　ビール名×本数　← 除外',
    '  例：「・ビアボム@西新宿」→「ビアボム」',
    '',
    '【フォーマットB：シンプル形式】',
    '  店舗名@場所',
    '  　ビール名×本数　← 除外',
    '  例：「天沼酒場@荻窪」→「天沼酒場」',
    '',
    '【フォーマットC：住所ブロック形式】',
    '  〒郵便番号／都道府県／番地／ビル名',
    '  店舗名　← これだけ抽出',
    '  （空行）',
    '  ビール名・keg情報　← 除外',
    '  例：「セントラル経堂B1Fマジックアワー」→「マジックアワー」',
    '',
    '【フォーマットD：@*記号形式（行頭に@または*）】',
    '  @店舗名　← @をつけたまま抽出（ピックアップ優先）',
    '  *店舗名　← *をつけたまま抽出（後回し）',
    '',
    '除外するもの：',
    '- 店舗名の後の@場所（例：@荻窪、@渋谷）',
    '- 住所情報（〒、日本、都道府県、市区町村、番地、ビル名）',
    '- ビール名・keg情報（ONEWAY、#番号など）',
    '- 挨拶・数量・日時・メモ（※〜）・記号（・）',
    '',
    '出力ルール：',
    '- 1行1店舗',
    '- 該当なければ「なし」',
    '',
    'メッセージ：',
    text
  ].join('\n');

  var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'content-type': 'application/json'
    },
    payload: JSON.stringify({
      model: 'claude-haiku-4-5',
      max_tokens: 512,
      messages: [{ role: 'user', content: prompt }]
    }),
    muteHttpExceptions: true
  });

  var result = JSON.parse(response.getContentText());
  if (result.error) throw new Error(result.error.message);

  return result.content[0].text.trim();
}

// ─── ルート計算を実行してplanシートに書き込む ─────────────────
function runRouteFromWebApp(shopText, startTimeStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var planSheet = ss.getSheetByName('plan');

  planSheet.getRange('B2').setValue(shopText);

  if (startTimeStr) {
    var startTime = new Date(startTimeStr);
    planSheet.getRange('C2').setValue(startTime);
    planSheet.getRange('C2').setNumberFormat('yyyy/MM/dd HH:mm');
  }

  try {
    processAdvancedDeliveryRoute();
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}
