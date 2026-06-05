/**
 * GAS Webアプリ：メッセージから店名抽出 → ルート計算 → 配達ログ書き込み
 *
 * キーワード → シート・Brewery対応
 * りくり → RIKRI sheet / Brewery: RIKRI
 * レッツ → let's sheet / Brewery: Let's Beer Works
 * カンパイ → kanpai sheet / Brewery: カンパイ！ブルーイング
 */

// ─── キーワード設定 ───────────────────────────────────────────
var SECTION_CONFIG = {
  'りくり': { sheet: 'RIKRI',   brewery: 'RIKRI',             extractKeg: true  },
  'レッツ':  { sheet: "let's",  brewery: "Let's Beer Works",  extractKeg: false },
  'カンパイ': { sheet: 'kanpai', brewery: 'カンパイ！ブルーイング', extractKeg: false }
};

// ─── セクション分割ヘルパー ────────────────────────────────────
function parseSections_(message) {
  var keywords = Object.keys(SECTION_CONFIG);
  var positions = [];

  keywords.forEach(function(kw) {
    var idx = message.indexOf(kw);
    if (idx !== -1) positions.push({ kw: kw, idx: idx });
  });

  positions.sort(function(a, b) { return a.idx - b.idx; });

  var sections = {};
  positions.forEach(function(pos, i) {
    var start = pos.idx + pos.kw.length;
    var end = (i + 1 < positions.length) ? positions[i + 1].idx : message.length;
    sections[pos.kw] = message.substring(start, end).trim();
  });

  return sections;
}

// ─── Webアプリのエントリーポイント ────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('配達ルート作成')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── GitHub PagesからのAPIリクエストを処理 ─────────────────────
// LINEのwebhook(doPost)と共存するため、bodyにpinがあればAPIとして処理
function doPost(e) {
  // APIリクエスト判定（pinフィールドがあればGitHub Pagesからの呼び出し）
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.hasOwnProperty('pin')) {
      return handleApiRequest_(body);
    }
  } catch(err) {}

  // それ以外はLINE webhook
  return handleLineWebhook_(e);
}

function handleApiRequest_(body) {
  var correctPin = PropertiesService.getScriptProperties().getProperty('WEBAPP_PIN');
  if (String(body.pin) !== String(correctPin)) {
    return jsonResponse_({ error: '認証エラー' });
  }
  try {
    switch(body.action) {
      case 'checkPin':
        return jsonResponse_({ success: true });  // PINが正しければここに到達
      case 'extractShops':
        return jsonResponse_({ success: true, result: extractShopsFromMessage(body.message) });
      case 'writeLog':
        return jsonResponse_(writeToDeliveryLog(body.message, body.startTime));
      case 'calcRoute':
        return jsonResponse_(runRouteFromWebApp(body.shops, body.startTime));
      default:
        return jsonResponse_({ error: '不明なアクション' });
    }
  } catch(err) {
    return jsonResponse_({ error: err.message });
  }
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleLineWebhook_(e) {
  var events = JSON.parse(e.postData.contents).events;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName('LINE_IDs') || ss.insertSheet('LINE_IDs');
  events.forEach(function(event) {
    var source = event.source;
    var type = source.type;
    var id = type === 'user'  ? source.userId
           : type === 'group' ? source.groupId
           : type === 'room'  ? source.roomId
           : null;
    if (!id) return;
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
    logSheet.appendRow([timestamp, type, id]);
  });
}

// ─── Claude APIで店名を抽出（ルート計算用） ───────────────────
function extractShopsFromMessage(message) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) throw new Error('ANTHROPIC_API_KEY が設定されていません');

  var sections = parseSections_(message);
  var results = [];

  // りくりセクション → @RIKRI BREWING + *店名
  if (sections['りくり']) {
    var shops = callClaude_(sections['りくり'], apiKey);
    var withStar = shops.split('\n')
      .map(function(l) { return l.trim(); })
      .filter(function(l) { return l && l !== 'なし'; })
      .map(function(l) { return (l.startsWith('@') || l.startsWith('*')) ? l : '*' + l; })
      .join('\n');
    results.push('@RIKRI BREWING');
    if (withStar) results.push(withStar);
  }

  // レッツ・カンパイセクション → 箱舟 + 店名
  ['レッツ', 'カンパイ'].forEach(function(kw) {
    if (sections[kw]) {
      var shops = callClaude_(sections[kw], apiKey);
      var lines = shops.split('\n')
        .map(function(l) { return l.trim(); })
        .filter(function(l) { return l && l !== 'なし'; })
        .join('\n');
      if (results.indexOf('箱舟（はこぶね）') === -1) results.push('箱舟（はこぶね）');
      if (lines) results.push(lines);
    }
  });

  return results.join('\n');
}

// ─── 配達ログに書き込む ────────────────────────────────────────
function writeToDeliveryLog(message, startTimeStr) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  var logId  = PropertiesService.getScriptProperties().getProperty('KEG_DELIVERY_LOG_ID');
  if (!apiKey) throw new Error('ANTHROPIC_API_KEY が設定されていません');
  if (!logId)  throw new Error('KEG_DELIVERY_LOG_ID が設定されていません');

  var ss = SpreadsheetApp.openById(logId);
  var sections = parseSections_(message);
  var written = 0;

  Object.keys(sections).forEach(function(kw) {
    var config = SECTION_CONFIG[kw];
    if (!config || !sections[kw]) return;

    var sheet = ss.getSheetByName(config.sheet);
    if (!sheet) return;

    // Claude で配達データをJSON抽出
    var rows = extractDeliveryRows_(sections[kw], apiKey, config.extractKeg);
    if (rows.length === 0) return;

    // 出発日時の日付を使用（未入力の場合は今日）
    var baseDate = startTimeStr ? new Date(startTimeStr) : new Date();
    var today = Utilities.formatDate(baseDate, Session.getScriptTimeZone(), 'yyyyMMdd');
    var colC = sheet.getRange('C:C').getValues();
    var maxSeq = 0;
    colC.forEach(function(cell) {
      var match = String(cell[0]).match(/^(\d{8})-(\d+)$/);
      if (match && match[1] === today) maxSeq = Math.max(maxSeq, parseInt(match[2]));
    });

    // Destination → 🚚# のマッピング（同じDestinationは同じ番号）
    var destToBatch = {};
    rows.forEach(function(row) {
      if (!destToBatch[row.destination]) {
        maxSeq++;
        destToBatch[row.destination] = today + '-' + maxSeq;
      }
      row.batchId = destToBatch[row.destination];
    });

    // E列で実際のデータがある最終行を探す
    var colE = sheet.getRange('E:E').getValues();
    var lastDataRow = 1;
    for (var r = colE.length - 1; r >= 0; r--) {
      if (colE[r][0] !== '') { lastDataRow = r + 1; break; }
    }

    rows.forEach(function(row) {
      var writeRow = lastDataRow + 1;
      lastDataRow++;
      sheet.getRange(writeRow, 2).setValue(config.brewery);  // B: Brewery
      sheet.getRange(writeRow, 3).setValue(row.batchId);     // C: 🚚#
      sheet.getRange(writeRow, 5).setValue(row.destination); // E: Destination
      sheet.getRange(writeRow, 6).setValue(row.beer);        // F: Beer name
      sheet.getRange(writeRow, 7).setValue(row.amount || '1'); // G: amount
      sheet.getRange(writeRow, 8).setValue(row.keg);         // H: keg#
      sheet.getRange(writeRow, 4).insertCheckboxes();        // D: チェックボックス
      sheet.getRange(writeRow, 1, 1, 9).setBackground(null); // 白背景にリセット
      written++;
    });
  });

  return { success: true, written: written };
}

// ─── デバッグ用（スクリプトエディタから直接実行） ──────────────
function debugWriteToDeliveryLog() {
  var logId = PropertiesService.getScriptProperties().getProperty('KEG_DELIVERY_LOG_ID');
  console.log('KEG_DELIVERY_LOG_ID:', logId);

  if (!logId) { console.log('IDが設定されていません'); return; }

  try {
    var ss = SpreadsheetApp.openById(logId);
    console.log('スプレッドシート名:', ss.getName());
    console.log('シート一覧:', ss.getSheets().map(function(s) { return s.getName(); }).join(', '));
  } catch(e) {
    console.log('スプレッドシートアクセスエラー:', e.message);
  }

  // テストメッセージで実行
  var testMessage = 'レッツ\n天沼酒場@荻窪\n　HEFE WEISSE 15L × 1本';
  console.log('テストメッセージ:', testMessage);
  var sections = parseSections_(testMessage);
  console.log('セクション:', JSON.stringify(sections));

  var result = writeToDeliveryLog(testMessage);
  console.log('結果:', JSON.stringify(result));
}

// ─── Claude で配達行データをJSON抽出 ─────────────────────────
function extractDeliveryRows_(text, apiKey, extractKeg) {
  var kegRule = extractKeg
    ? 'keg欄：keg番号（例：#12→"12"）またはONEWAY/one way→"one way"、なければ空文字'
    : 'keg欄：ONEWAY/one wayの記載があれば"one way"、なければ空文字';

  var prompt = [
    'あなたはクラフトビール配達業務のアシスタントです。',
    '以下のメッセージから配達データを抽出してJSON配列で返してください。',
    '',
    '抽出するフィールド：',
    '- destination: 店舗名（@場所は除く、・記号も除く）',
    '- beer: ビール名（サイズ表記15Lなどは除く）',
    '- amount: 本数（数字のみ。記載がなければ"1"）',
    '- keg: ' + kegRule,
    '',
    '出力形式（JSON配列のみ、他のテキスト不要）：',
    '[{"destination":"店名","beer":"ビール名","amount":"1","keg":""}]',
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
      max_tokens: 1024,
      messages: [{ role: 'user', content: prompt }]
    }),
    muteHttpExceptions: true
  });

  var result = JSON.parse(response.getContentText());
  if (result.error) throw new Error(result.error.message);

  var responseText = result.content[0].text.trim();
  // JSON部分だけ抽出
  var match = responseText.match(/\[[\s\S]*\]/);
  if (!match) return [];
  try { return JSON.parse(match[0]); } catch(e) { return []; }
}

// ─── Claude API呼び出し（店名抽出用） ────────────────────────
function callClaude_(text, apiKey) {
  var prompt = [
    'あなたはクラフトビール配達業務のアシスタントです。',
    '以下のメッセージから配達先の店舗名だけを抽出してください。',
    '',
    '【@ の使われ方は2種類ある。必ず区別すること】',
    '  ① 行の先頭に @ → ピックアップ優先マーク。@ごと抽出する',
    '  ② 店舗名の後ろに @場所 → 場所を示すだけ。@以降を削除して店舗名だけ抽出',
    '     例：「天沼酒場@荻窪」→「天沼酒場」',
    '',
    '【フォーマットA：箇条書き形式】',
    '  ・店舗名@場所　→ 店舗名のみ抽出',
    '',
    '【フォーマットB：シンプル形式】',
    '  店舗名@場所　→ 店舗名のみ抽出',
    '',
    '【フォーマットC：住所ブロック形式】',
    '  〒住所ブロックの直後の行が店舗名',
    '  例：「セントラル経堂B1Fマジックアワー」→「マジックアワー」',
    '',
    '除外するもの：',
    '- 店舗名の後の@場所、住所情報、ビール名、keg情報（ONEWAY・#番号）',
    '- 挨拶・数量・日時・メモ・記号（・）',
    '',
    '出力ルール：1行1店舗。該当なければ「なし」',
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
