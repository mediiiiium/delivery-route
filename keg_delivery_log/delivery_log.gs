// ==========================================
// 配達ログ管理スクリプト（複数シート対応）
// 列構成: A=配達日時, B=Brewery, C=🚚#, D=チェックボックス,
//         E=Destination, F=Beer name, G=amount, H=keg#, I=memo
// ==========================================

// ▼ 管理対象のシート一覧（LINE通知なしでもチェックボックス・色変えは動く）
var SHEET_NAMES = ["let's", "kanpai", "Drinkuppers", "RIKRI"];

// ▼ LINE Messaging API 設定（トークンはScriptPropertiesに保存）
// LINE_CHANNEL_TOKEN はsendLine_()内で都度取得

// ▼ シートごとの送信先ID（設定がないシートはLINE通知なし）
var NOTIFICATION_CONFIG = {
  "let's":       "C563f4f2c938e01d7ca384102cd410187",
  "kanpai":      "C563f4f2c938e01d7ca384102cd410187",
  "RIKRI":       "Cc4d3c0f2a63082be175defe02d18b60d"
  // 新しいシートを追加するときはここに足す
};

var HEADER_ROW  = 1;
var COL_DATE    = 1;  // A: 配達日時
var COL_BATCH   = 3;  // C: 🚚#
var CHECKBOX_COL = 4; // D: チェックボックス
var COL_DEST    = 5;  // E: Destination
var COL_BEER    = 6;  // F: Beer name
var COL_AMOUNT  = 7;  // G: amount
var COL_KEG     = 8;  // H: keg#
var TOTAL_COLS  = 9;  // 総列数（I列まで）
var DELIVERED_COLOR = "#D4CFC8";

// ─── 初期セットアップ（全シート一括） ────────────────────────────
function setup() {
  var logId = PropertiesService.getScriptProperties().getProperty('KEG_DELIVERY_LOG_ID');
  var ss = logId ? SpreadsheetApp.openById(logId) : SpreadsheetApp.getActiveSpreadsheet();
  var missing = [];

  SHEET_NAMES.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) { missing.push(name); return; }
    setupSheet_(sheet);
  });

  deleteTrigger_();
  ScriptApp.newTrigger("onEditTrigger").forSpreadsheet(ss).onEdit().create();

  var msg = "セットアップ完了！";
  if (missing.length > 0) msg += "\n⚠️ 見つからなかったシート: " + missing.join(", ");
  ss.toast(msg, "✅ 完了", 7);
}

function setupSheet_(sheet) {
  sheet.getRange(HEADER_ROW, COL_DATE).setValue("配達日時");

  // H列：🚚#ヘッダー
  var batchHeader = sheet.getRange(HEADER_ROW, COL_BATCH);
  batchHeader.setValue("🚚#");
  batchHeader.setFontWeight("bold");
  batchHeader.setBackground("#D9D9D9");

  // I列：チェックボックスヘッダー
  var checkHeader = sheet.getRange(HEADER_ROW, CHECKBOX_COL);
  checkHeader.setValue("配達済み");
  checkHeader.setFontWeight("bold");
  checkHeader.setBackground("#D9D9D9");

  var lastRow = sheet.getLastRow();
  if (lastRow > HEADER_ROW) {
    sheet.getRange(HEADER_ROW + 1, CHECKBOX_COL, lastRow - HEADER_ROW, 1).insertCheckboxes();
  }
}

// ─── 新しい行にチェックボックスを追加（全シート一括） ──────────────
function addCheckboxToNewRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  SHEET_NAMES.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    var lastRow = sheet.getLastRow();
    for (var i = HEADER_ROW + 1; i <= lastRow; i++) {
      var cell = sheet.getRange(i, CHECKBOX_COL);
      if (cell.getValue() === "") cell.insertCheckboxes();
    }
  });
  ss.toast("全シートのチェックボックスを更新しました。", "✅ 完了", 3);
}

// ─── onEdit トリガー ─────────────────────────────────────────────
function onEditTrigger(e) {
  if (!e || !e.source || !e.range) return;
  var sheet = e.source.getActiveSheet();
  var sheetName = sheet.getName();

  if (SHEET_NAMES.indexOf(sheetName) === -1) return;

  var range    = e.range;
  var col      = range.getColumn();
  var startRow = range.getRow();
  var numRows  = range.getNumRows();

  if (col !== CHECKBOX_COL) return;

  var now       = new Date();
  var timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");

  // 範囲内の全行に色・日時を適用
  for (var i = 0; i < numRows; i++) {
    var r = startRow + i;
    if (r <= HEADER_ROW) continue;
    if (sheet.getRange(r, CHECKBOX_COL).getValue() === true) {
      sheet.getRange(r, 1, 1, TOTAL_COLS).setBackground(DELIVERED_COLOR);
      sheet.getRange(r, COL_DATE).setValue(timestamp);
    } else {
      sheet.getRange(r, COL_DATE).clearContent();
      sheet.getRange(r, 1, 1, TOTAL_COLS).setBackground(null);
    }
  }

  // チェックONの行があれば通知判定
  var firstCheckedRow = startRow;
  if (firstCheckedRow <= HEADER_ROW) return;
  if (sheet.getRange(firstCheckedRow, CHECKBOX_COL).getValue() !== true) return;

  var batchId = sheet.getRange(firstCheckedRow, COL_BATCH).getDisplayValue();
  if (!batchId) {
    e.source.toast("🚚# を入力してからチェックしてください。", "⚠️ 配送番号なし", 5);
    return;
  }

  // 同じ🚚#の行を全て取得
  var lastRow = sheet.getLastRow();
  var allRows = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, TOTAL_COLS).getValues();

  var batchRows  = [];
  var allChecked = true;

  allRows.forEach(function(r) {
    if (String(r[COL_BATCH - 1]) !== String(batchId)) return;
    batchRows.push(r);
    if (r[CHECKBOX_COL - 1] !== true) allChecked = false;
  });

  // 全行チェック済みになったらLINE送信（30秒以内の重複送信を防止）
  if (allChecked && batchRows.length > 0) {
    var props    = PropertiesService.getScriptProperties();
    var sentKey  = sheetName + "_" + batchId;
    var lastSent = props.getProperty(sentKey);

    if (!lastSent || (Date.now() - parseInt(lastSent)) > 30000) {
      props.setProperty(sentKey, String(Date.now()));

      var destination = batchRows[0][COL_DEST - 1];
      var brewery = (sheetName === "let's" || sheetName === "kanpai")
                    ? batchRows[0][1]
                    : null;
      var beers = batchRows.map(function(r) {
        return { beer: r[COL_BEER - 1], amount: r[COL_AMOUNT - 1], keg: r[COL_KEG - 1] };
      });
      var message = buildDeliveryMessage_(timestamp, destination, beers, brewery);

      var lineId = NOTIFICATION_CONFIG[sheetName];
      if (lineId) sendLine_(lineId, message);

      e.source.toast(message, "🍺 配達完了通知", 10);
    }
  }
}

// ─── LINE Messaging API 送信 ──────────────────────────────────
function sendLine_(to, message) {
  var token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_TOKEN');
  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
    method: "post",
    headers: {
      "Authorization": "Bearer " + token,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify({
      to: to,
      messages: [{ type: "text", text: message }]
    })
  });
}

// ─── 送信先IDを確認するためのヘルパー関数 ────────────────────────
function doPost(e) {
  var events = JSON.parse(e.postData.contents).events;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("LINE_IDs") || ss.insertSheet("LINE_IDs");

  events.forEach(function(event) {
    var source = event.source;
    var type = source.type;
    var id = type === "user"  ? source.userId
           : type === "group" ? source.groupId
           : type === "room"  ? source.roomId
           : null;
    if (!id) return;
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
    logSheet.appendRow([timestamp, type, id]);
  });
}

// ─── LINEテスト送信 ───────────────────────────────────────────
function testSendLine() {
  var sheetName = "RIKRI";
  var to = NOTIFICATION_CONFIG[sheetName];
  var message = "テスト送信です";

  try {
    var response = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
      method: "post",
      headers: {
        "Authorization": "Bearer " + LINE_CHANNEL_TOKEN,
        "Content-Type": "application/json"
      },
      payload: JSON.stringify({
        to: to,
        messages: [{ type: "text", text: message }]
      }),
      muteHttpExceptions: true
    });
    Logger.log("ステータス: " + response.getResponseCode());
    Logger.log("レスポンス: " + response.getContentText());
  } catch(err) {
    Logger.log("エラー: " + err);
  }
}

// ─── メッセージ文面（複数ビール対応） ─────────────────────────────
function buildDeliveryMessage_(date, destination, beers, brewery) {
  var lines = [];
  lines.push("配達が完了しました。");
  lines.push("");
  if (brewery)     lines.push("醸造所：" + brewery);
  if (destination) lines.push("配達先：" + destination);
  lines.push("");
  beers.forEach(function(b) {
    var line = "・" + b.beer;
    if (b.amount) line += "　" + b.amount + "本";
    if (b.keg)    line += "　#" + b.keg;
    lines.push(line);
  });
  lines.push("");
  if (date) lines.push("配達日時：" + date);
  return lines.join("\n");
}

// ─── 🚚#を選択行に割り当て ───────────────────────────────
function assignBatchNumber() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  if (SHEET_NAMES.indexOf(sheet.getName()) === -1) {
    ss.toast("このシートは対象外です。", "⚠️ エラー", 4);
    return;
  }

  var selection = sheet.getActiveRange();
  var startRow  = selection.getRow();
  var numRows   = selection.getNumRows();

  if (startRow <= HEADER_ROW) {
    ss.toast("ヘッダー行は除いて選択してください。", "⚠️ エラー", 4);
    return;
  }

  // 今日の日付部分（MMdd）
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");

  // シート内で今日使われている🚚#の最大連番を取得
  var lastRow  = sheet.getLastRow();
  var maxSeq   = 0;
  if (lastRow > HEADER_ROW) {
    var batchVals = sheet.getRange(HEADER_ROW + 1, COL_BATCH, lastRow - HEADER_ROW, 1).getValues();
    batchVals.forEach(function(r) {
      var val = String(r[0]);
      var match = val.match(/^(\d{8})-(\d+)$/);
      if (match && match[1] === today) {
        maxSeq = Math.max(maxSeq, parseInt(match[2]));
      }
    });
  }

  var newBatch = today + "-" + (maxSeq + 1);

  // 選択行に🚚#を入力
  for (var i = 0; i < numRows; i++) {
    sheet.getRange(startRow + i, COL_BATCH).setValue(newBatch);
  }

  ss.toast("🚚# " + newBatch + " を " + numRows + " 行に割り当てました。", "✅ 完了", 4);
}

// ─── カスタムメニュー ─────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📦 配達ログ")
    .addItem("初期セットアップ（初回のみ・全シート）", "setup")
    .addItem("新しい行にチェックボックスを追加（全シート）", "addCheckboxToNewRows")
    .addSeparator()
    .addItem("選択行に🚚#を割り当て", "assignBatchNumber")
    .addToUi();
}

// ─── 内部：既存のonEditトリガーを削除 ─────────────────────────
function deleteTrigger_() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === "onEditTrigger") ScriptApp.deleteTrigger(t);
  });
}
