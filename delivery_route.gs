/**
 * 設定：Google Cloud Consoleで取得したAPIキーと、拠点の住所を入力
 */
const GOOGLE_MAPS_API_KEY = 'AIzaSyArZEK07d7b6tzcmvsIpMYmpaTN06vYsHQ';
const START_POINT = '東京都文京区水道２丁目１２−１０ 北島ビル'; // 拠点の住所を入力

const VALID_AREAS = ["東京都", "神奈川県", "埼玉県", "千葉県", "Tokyo", "Kanagawa", "Saitama", "Chiba"];
const SPLIT_WAIT_MIN = 30; // この分数以上の待機でルート②を生成
const DEFAULT_SHOPS = '@RIKRI BREWING\n箱舟（はこぶね）'; // フォームのデフォルト入力値

/**
 * 0-import. import_listシートのA列から店舗名を読み込んでtaproom_masterに一括登録
 */
function importFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName('import_list');
  if (!importSheet) { Browser.msgBox('import_list シートが見つかりません。'); return; }

  const masterSheet = ss.getSheetByName('taproom_master');
  deduplicateMaster();
  const masterCache = loadMasterCache(masterSheet);

  const values = importSheet.getRange('A1:A' + importSheet.getLastRow()).getValues();
  const names = values.map(r => String(r[0]).trim()).filter(n => n);
  if (names.length === 0) { Browser.msgBox('import_list のA列にデータがありません。'); return; }

  let added = 0, skipped = 0, failed = 0;
  const logs = [];

  names.forEach(name => {
    // 既にキャッシュにあればスキップ
    const cleanName = name.replace(/\s+/g, '').toLowerCase();
    const alreadyExists = masterCache.rows.some(r => {
      const mn = String(r.name).replace(/\s+/g, '').toLowerCase();
      return mn.includes(cleanName) || cleanName.includes(mn);
    });
    if (alreadyExists) { skipped++; return; }

    const shopData = getOrUpdateStoreMaster(name, masterSheet, masterCache, logs);
    if (shopData) { added++; } else { failed++; }
    Utilities.sleep(250);
  });

  const detail = logs.length > 0 ? '\n\n詳細:\n' + logs.slice(0, 20).join('\n') : '';
  Browser.msgBox(`インポート完了\n✅ 追加: ${added}件\nスキップ（既存）: ${skipped}件\n❌ 取得失敗: ${failed}件${detail}`);
}

/**
 * 0-debug. Saved Places.json の構造確認用（実行ログで確認）
 */
function debugTakeoutJson() {
  const files = DriveApp.getFilesByName('Delivary_Taproom.json');
  if (!files.hasNext()) { Logger.log('ファイルが見つかりません'); return; }
  const json = JSON.parse(files.next().getBlob().getDataAsString());
  const features = json.features || [];
  Logger.log('総件数: ' + features.length);
  // 最初の3件の構造を出力
  features.slice(0, 3).forEach((f, i) => {
    Logger.log('--- ' + i + ' ---');
    Logger.log(JSON.stringify(f.properties, null, 2));
  });
}

/**
 * 0. Takeoutインポート：Google Driveに "Saved Places.json" をアップロードしてから実行
 *    takeout.google.com → マップ（マイプレイス）→ エクスポート → 解凍 → Driveにアップ
 */
function importFromTakeout() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('taproom_master');

  // Drive から Saved Places.json を検索
  const files = DriveApp.getFilesByName('Delivary_Taproom.json');
  if (!files.hasNext()) {
    Browser.msgBox('Delivary_Taproom.json がGoogle Driveに見つかりません。\nGoogle TakeoutでエクスポートしたファイルをDriveにアップロードしてください。');
    return;
  }

  const json = JSON.parse(files.next().getBlob().getDataAsString());
  const features = json.features || [];
  if (features.length === 0) {
    Browser.msgBox('JSONにデータが見つかりません。');
    return;
  }

  // 重複チェック用にキャッシュを構築
  deduplicateMaster();
  const masterCache = loadMasterCache(masterSheet);

  let added = 0, skipped = 0, failed = 0;
  const logs = [];

  features.forEach(feature => {
    const props    = feature.properties || {};
    const location = props.location || props.Location || {};
    const title    = props.Title || location.name || location['Business Name'] || '';
    const address  = location.address || location.Address || '';
    if (!title) { skipped++; return; }

    // 対象エリア外はスキップ
    if (address && !VALID_AREAS.some(a => address.includes(a))) {
      logs.push(`「${title}」→ 対象エリア外 (${address})`);
      skipped++;
      return;
    }

    // 既にキャッシュにあればスキップ
    const cleanTitle = title.replace(/\s+/g, '').toLowerCase();
    const alreadyExists = masterCache.rows.some(r => {
      const mn = String(r.name).replace(/\s+/g, '').toLowerCase();
      return mn.includes(cleanTitle) || cleanTitle.includes(mn);
    });
    if (alreadyExists) { skipped++; return; }

    // Places API で詳細を取得してマスターに追加
    const shopData = getOrUpdateStoreMaster(title, masterSheet, masterCache, logs);
    if (shopData) { added++; } else { failed++; }

    Utilities.sleep(250); // レート制限対策
  });

  const detail = logs.length > 0 ? '\n\n詳細:\n' + logs.slice(0, 20).join('\n') : '';
  Browser.msgBox(`インポート完了\n✅ 追加: ${added}件\n⏭ スキップ（既存・エリア外）: ${skipped}件\n❌ 取得失敗: ${failed}件${detail}`);
}

/**
 * 0b. フォームセットアップ：スクリプトエディタから一度だけ手動実行する
 */
function setupDeliveryForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // フォーム作成
  const form = FormApp.create('配送ルート入力');
  form.setDescription('店舗名を入力して送信すると、自動でルートが計算されます');
  form.setCollectEmail(false);
  form.setLimitOneResponsePerUser(false);

  // 店舗名（段落テキスト：1行1店舗、件数制限なし）
  const shopItem = form.addParagraphTextItem();
  shopItem.setTitle('店舗名（1行に1店舗）');
  shopItem.setHelpText('例:\n@RIKRI BREWING\n箱舟\n*カンパイスタンド SHIBUYA\n\n@ = ピックアップ優先、* = @より後に訪問');
  shopItem.setRequired(true);

  // 出発日時（カレンダー＋時間ピッカー）
  const timeItem = form.addDateTimeItem();
  timeItem.setTitle('出発日時');
  timeItem.setRequired(false);


  // フォームをスプレッドシートにリンク（回答シートは使わない）
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  // onFormSubmit トリガーを設定
  ScriptApp.newTrigger('onFormSubmit')
    .forForm(form)
    .onFormSubmit()
    .create();

  // フォームIDをScriptPropertiesに保存（updatePrefilledUrl で使用）
  PropertiesService.getScriptProperties().setProperty('FORM_ID', form.getId());

  const prefilledUrl = form.getPublishedUrl();

  // フォームURLをplanシートに記録
  const planSheet = ss.getSheetByName('plan');
  planSheet.getRange('E2').setValue('フォームURL');
  planSheet.getRange('F2').setFormula(`=HYPERLINK("${prefilledUrl}", "📝 入力フォームを開く")`);

  Browser.msgBox('フォームを作成しました。\nF2セルのリンクからアクセスできます。');
}

/**
 * 0b. 事前入力URL更新：setupDeliveryForm の後に一度だけ実行する
 */
function updatePrefilledUrl() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planSheet = ss.getSheetByName('plan');

  const formId = PropertiesService.getScriptProperties().getProperty('FORM_ID');
  if (!formId) { Browser.msgBox('フォームIDが見つかりません。先にsetupDeliveryFormを実行してください。'); return; }

  const form = FormApp.openById(formId);
  const items = form.getItems();
  if (items.length === 0) { Browser.msgBox('フォームにアイテムが見つかりません。'); return; }

  // FormResponse.toPrefilledUrl() で事前入力URLを生成
  const response = form.createResponse();
  response.withItemResponse(items[0].asParagraphTextItem().createResponse(DEFAULT_SHOPS));
  const prefilledUrl = response.toPrefilledUrl();

  planSheet.getRange('F2').setFormula(`=HYPERLINK("${prefilledUrl}", "📝 入力フォームを開く")`);
  Browser.msgBox('事前入力URLを設定しました。F2のリンクを確認してください。');
}

/**
 * 0c. フォーム送信トリガー
 */
function onFormSubmit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planSheet = ss.getSheetByName('plan');
  const responses = e.response.getItemResponses();

  const shopNames = [];
  let startTime = null;

  responses.forEach(r => {
    const title = r.getItem().getTitle();
    const val = String(r.getResponse() || '').trim();
    if (!val) return;

    if (title.startsWith('店舗名')) {
      // 段落テキスト：改行・読点・カンマで分割して複数店舗に対応
      val.split(/[\n、,]+/).map(s => s.trim()).filter(s => s).forEach(s => shopNames.push(s));
    } else if (title.startsWith('出発日時')) {
      startTime = r.getResponse(); // Date オブジェクトで返る
    }
  });

  if (shopNames.length === 0) return;

  // plan シートに書き込み
  planSheet.getRange('B2').setValue(shopNames.join('\n'));

  if (startTime) {
    // DateTimeItem のレスポンスは Date オブジェクト
    planSheet.getRange('C2').setValue(startTime);
    planSheet.getRange('C2').setNumberFormat('yyyy/MM/dd HH:mm');
  }

  // チェックボックスをONにしてルート計算を発火
  planSheet.getRange('A2').setValue(true);

  try {
    planSheet.getRange('A2').setValue(false);
    processAdvancedDeliveryRoute();
  } catch (err) {
    console.error(err.stack);
  }
}

/**
 * 1. 編集トリガー
 */
function installedOnEdit(e) {
  if (!e || !e.source) return;
  const sheet = e.source.getActiveSheet();
  const range = e.range;

  if (sheet.getName() === "plan" && range.getA1Notation() === "A2" && range.getValue() === true) {
    try {
      range.setValue(false);
      processAdvancedDeliveryRoute();
    } catch (err) {
      console.error(err.stack);
      Browser.msgBox("⚠️ ルート計算を中断しました:\n" + err.message);
    }
  }
}

/**
 * 2. メインロジック
 */
function processAdvancedDeliveryRoute() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("plan");
  const masterSheet = ss.getSheetByName("taproom_master");

  const inputText = inputSheet.getRange("B2").getValue();
  if (!inputText) return;

  let startTimeValue = inputSheet.getRange("C2").getValue();
  let startTime = (startTimeValue instanceof Date) ? startTimeValue : new Date();
  if (startTime.getFullYear() < 2000) {
    let now = new Date();
    startTime.setFullYear(now.getFullYear(), now.getMonth(), now.getDate());
  }

  let waitThresholdMin = inputSheet.getRange("D2").getValue() || 30;
  const rawNames = inputText.split(/[,\n、]+/).map(n => n.trim()).filter(n => n.length > 0);
  const dayIdx = startTime.getDay();

  // マスター重複を自動削除してからキャッシュ構築
  deduplicateMaster();
  const masterCache = loadMasterCache(masterSheet);

  let targetShops = [];
  let seenPlaceIds = new Set();
  let searchLogs = [];

  rawNames.forEach(name => {
    let isPickup = name.includes("@");
    let isMemo = name.trimStart().startsWith("*");
    let shopData = getOrUpdateStoreMaster(name, masterSheet, masterCache, searchLogs);

    if (shopData && shopData.placeId && !seenPlaceIds.has(shopData.placeId)) {
      if (shopData.address && !["日本", "Japan"].includes(shopData.address)) {
        shopData.isPickup = isPickup;
        shopData.isMemo = isMemo;
        shopData.slots = generateDeliverySlots(shopData.openingHours[dayIdx], startTime);
        targetShops.push(shopData);
        seenPlaceIds.add(shopData.placeId);
      } else {
        searchLogs.push(`「${name}」→ 住所が無効 (${shopData.address})`);
      }
    }
  });

  if (targetShops.length === 0) {
    const detail = searchLogs.length > 0 ? "\n\n詳細:\n" + searchLogs.join("\n") : "";
    Browser.msgBox("有効な店舗が見つかりませんでした。" + detail);
    return;
  }

  const dataMatrix = fetchDistanceMatrix(START_POINT, targetShops);

  // ルート① 計算（マルチスタート: 5通りの出発順序で最良を選択）
  const { route: route1, failedShops } = multiStartOptimize(targetShops, START_POINT, startTime, dataMatrix, waitThresholdMin);

  // シートをクリアして描画
  inputSheet.getRange("A5:H200").clearContent();

  const route1EndRow = renderRoute(inputSheet, route1, failedShops, startTime, dataMatrix, 5, "【ルート①】");

  // ルート②：SPLIT_WAIT_MIN 以上の待機が発生するか確認して生成
  const splitResult = findSplitPoint(route1, startTime, dataMatrix);
  if (splitResult) {
    const { splitIndex, returnTime } = splitResult;
    const firstLeg = route1.slice(0, splitIndex);
    const secondLegShops = route1.slice(splitIndex);

    const { route: secondBase } = buildRoute(secondLegShops, START_POINT, returnTime, dataMatrix, waitThresholdMin);
    const secondLeg = optimizeOrOpt(optimize2Opt(secondBase, returnTime, dataMatrix), returnTime, dataMatrix);

    const route2StartRow = route1EndRow + 2;
    renderRoute2(inputSheet, firstLeg, secondLeg, failedShops, startTime, returnTime, dataMatrix, route2StartRow);

    // G2:H5 のサマリー行とデータ行の高さを統一
    const lastDataRow = route2StartRow + firstLeg.length + secondLeg.length + 2;
    inputSheet.setRowHeights(2, 4, 25);                          // G2:H5（サマリー）
    inputSheet.setRowHeights(5, lastDataRow - 5, 25);            // A5以降（データ）
  } else {
    // ルート①のみの場合
    const lastDataRow = route1EndRow;
    inputSheet.setRowHeights(2, 4, 25);
    inputSheet.setRowHeights(5, lastDataRow - 5, 25);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('ルート計算が完了しました ✅', '完了', 5);
}

/**
 * 3. グリーディ ルート構築（共通処理）
 */
function buildRoute(shops, startPos, startTime, dataMatrix, waitThresholdMin) {
  let currentPos = startPos;
  let currentTime = new Date(startTime.getTime());
  let unvisited = [...shops];
  let deferred = [];
  let route = [];
  let failedShops = [];

  while (unvisited.length > 0 || deferred.length > 0) {
    let bestCandidate = null;
    let bestScore = Infinity;
    let bestWait = 0;
    let bestTravel = 0;
    let hasRemainingPickup = unvisited.some(s => s.isPickup) || deferred.some(s => s.isPickup);

    for (let i = 0; i < unvisited.length; i++) {
      let shop = unvisited[i];
      if (shop.isMemo && hasRemainingPickup) continue;

      let travelSec = dataMatrix[currentPos]?.[shop.address]?.duration || 0;
      let travelMin = Math.ceil(travelSec / 60);
      let arrivalTime = new Date(currentTime.getTime() + travelMin * 60 * 1000);
      let arrivalMin = arrivalTime.getHours() * 60 + arrivalTime.getMinutes();
      let status = checkSlots(arrivalMin, shop.slots);

      if (status.type !== "FAILED") {
        let waitMin = status.type === "WAIT" ? status.waitTime : 0;
        let score = travelMin + waitMin;
        if (status.type === "AVAILABLE" || waitMin <= waitThresholdMin) {
          if (score < bestScore) {
            bestScore = score;
            bestCandidate = { shop };
            bestWait = waitMin;
            bestTravel = travelMin;
          }
        } else { shop._tempDeferred = true; }
      } else { shop._failed = true; }
    }

    for (let i = unvisited.length - 1; i >= 0; i--) {
      if (unvisited[i]._tempDeferred) {
        delete unvisited[i]._tempDeferred;
        deferred.push(unvisited.splice(i, 1)[0]);
      } else if (unvisited[i]._failed) {
        delete unvisited[i]._failed;
        failedShops.push(unvisited.splice(i, 1)[0]);
      }
    }

    if (bestCandidate) {
      let chosen = bestCandidate.shop;
      route.push(chosen);
      currentTime = new Date(currentTime.getTime() + (bestTravel + bestWait + 5) * 60 * 1000);
      currentPos = chosen.address;
      unvisited.splice(unvisited.indexOf(chosen), 1);
    } else if (deferred.length > 0) {
      let nowMin = currentTime.getHours() * 60 + currentTime.getMinutes();
      // 閉店前（end > nowMin）の店はまだ訪問可能
      let stillReachable = deferred.filter(s => s.slots.some(sl => sl.end > nowMin));
      let unreachable    = deferred.filter(s => !s.slots.some(sl => sl.end > nowMin));
      failedShops = failedShops.concat(unreachable);
      if (stillReachable.length > 0) {
        // 現在時刻より未来に開く店があれば時刻を進める（既に営業中なら進めない）
        let futureOpens = stillReachable.flatMap(s => s.slots.map(sl => sl.start).filter(t => t > nowMin));
        if (futureOpens.length > 0) {
          let earliestOpen = Math.min(...futureOpens);
          currentTime.setHours(Math.floor(earliestOpen / 60), earliestOpen % 60, 0, 0);
        }
        unvisited = unvisited.concat(stillReachable);
        deferred = [];
      } else {
        deferred = [];
      }
    } else { break; }
  }

  return { route, failedShops };
}

/**
 * 4. 分割点の検出：SPLIT_WAIT_MIN 以上の待機が最初に発生するインデックスを返す
 */
function findSplitPoint(route, startTime, dataMatrix) {
  let currentPos = START_POINT;
  let currentTime = new Date(startTime.getTime());

  for (let i = 0; i < route.length; i++) {
    let shop = route[i];
    let travelMin = Math.ceil((dataMatrix[currentPos]?.[shop.address]?.duration || 0) / 60);
    let arrivalTime = new Date(currentTime.getTime() + travelMin * 60 * 1000);
    let status = checkSlots(arrivalTime.getHours() * 60 + arrivalTime.getMinutes(), shop.slots);
    let wait = status.type === "WAIT" ? status.waitTime : 0;

    if (wait >= SPLIT_WAIT_MIN) {
      // 帰着時刻 = 前の停留点からの帰着移動時間
      let returnMin = Math.ceil((dataMatrix[currentPos]?.[START_POINT]?.duration || 0) / 60);
      let returnTime = new Date(currentTime.getTime() + returnMin * 60 * 1000);
      return { splitIndex: i, returnTime };
    }

    currentTime = new Date(arrivalTime.getTime() + (wait + 5) * 60 * 1000);
    currentPos = shop.address;
  }
  return null;
}

/**
 * 5a. マスター重複削除（手動実行）：同じplaceIdの行を古い方から削除する
 */
function deduplicateMaster() {
  const masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("taproom_master");
  const data = masterSheet.getDataRange().getValues();
  const seenPlaceIds = new Set();
  const rowsToDelete = [];

  for (let i = 1; i < data.length; i++) {
    const placeId = String(data[i][2]).trim();
    if (!placeId) continue;
    if (seenPlaceIds.has(placeId)) {
      rowsToDelete.push(i + 1); // シート行番号（1始まり）
    } else {
      seenPlaceIds.add(placeId);
    }
  }

  // 下から消さないとインデックスがずれる
  rowsToDelete.reverse().forEach(row => masterSheet.deleteRow(row));
  Browser.msgBox(`重複削除完了：${rowsToDelete.length} 行削除しました。`);
}

/**
 * 5b. マスターキャッシュ構築（実行開始時に1回だけ呼ぶ）
 */
function loadMasterCache(masterSheet) {
  const data = masterSheet.getDataRange().getValues();
  const getMap = (row) => ({ 0: row[9], 1: row[3], 2: row[4], 3: row[5], 4: row[6], 5: row[7], 6: row[8] });
  const rows = data.slice(1)
    .filter(row => row[0])
    .map(row => ({ name: row[0], address: row[1], placeId: row[2], openingHours: getMap(row) }));
  return { rows, placeIds: new Set(rows.map(r => r.placeId).filter(Boolean)) };
}

/**
 * 5b. マスター管理（キャッシュ経由で重複を防ぐ）
 */
function getOrUpdateStoreMaster(shopName, masterSheet, masterCache, logs) {
  const baseName = shopName.replace(/^[@*]+/, "").replace(/\*.*$/, "").trim();
  const cleanSearchName = baseName.replace(/\s+/g, "").toLowerCase();
  if (!cleanSearchName) return null;

  // ① キャッシュから名前で双方向部分一致（表記ゆれ対応）
  for (const entry of masterCache.rows) {
    const masterName = String(entry.name).replace(/\s+/g, "").toLowerCase();
    if (masterName.length >= 2 && (masterName.includes(cleanSearchName) || cleanSearchName.includes(masterName))) {
      return entry;
    }
  }

  // ② Places API で検索（候補を順に試す）
  const query = `${baseName} クラフトビール 東京`;
  try {
    const findUrl = `https://maps.googleapis.com/maps/api/place/findplacefromtext/json?input=${encodeURIComponent(query)}&inputtype=textquery&fields=place_id,formatted_address&key=${GOOGLE_MAPS_API_KEY}`;
    const res = JSON.parse(UrlFetchApp.fetch(findUrl).getContentText());

    if (res.status && res.status !== "OK" && res.status !== "ZERO_RESULTS") {
      logs && logs.push(`「${shopName}」→ Places API エラー: ${res.status} (${res.error_message || "詳細なし"})`);
      return null;
    }
    if (!res.candidates || res.candidates.length === 0) {
      logs && logs.push(`「${shopName}」→ 候補が見つかりません (クエリ: ${query})`);
      return null;
    }

    for (const cand of res.candidates) {
      if (!VALID_AREAS.some(area => cand.formatted_address.includes(area))) continue;

      // ③ placeId がキャッシュに既存なら重複追加せず返す（同一実行内も含む）
      if (masterCache.placeIds.has(cand.place_id)) {
        return masterCache.rows.find(r => r.placeId === cand.place_id);
      }

      const det = JSON.parse(UrlFetchApp.fetch(`https://maps.googleapis.com/maps/api/place/details/json?place_id=${cand.place_id}&fields=name,opening_hours,geometry&key=${GOOGLE_MAPS_API_KEY}&language=ja`).getContentText());
      const result = det.result;

      // ④ 名称チェック：不一致でも次の候補へ（完全に別の場所だけ除外）
      const resultNameClean = result.name.replace(/\s+/g, "").toLowerCase();
      const isValidName = resultNameClean.includes(cleanSearchName) || cleanSearchName.includes(resultNameClean);
      if (!isValidName && res.candidates.length > 1) {
        logs && logs.push(`「${shopName}」→ 名称不一致: API結果="${result.name}" → 次の候補を確認`);
        continue;
      }

      let h = {0:"不明", 1:"不明", 2:"不明", 3:"不明", 4:"不明", 5:"不明", 6:"不明"};
      if (result.opening_hours && result.opening_hours.weekday_text) {
        result.opening_hours.weekday_text.forEach((t, idx) => {
          let timePart = t.split(/:\s+/)[1] || "不明";
          h[(idx + 1) % 7] = normalizeHours(timePart);
        });
      }
      const newEntry = { name: result.name, address: cand.formatted_address, placeId: cand.place_id, openingHours: h };
      masterSheet.appendRow([result.name, cand.formatted_address, cand.place_id, h[1], h[2], h[3], h[4], h[5], h[6], h[0], result.geometry.location.lat, result.geometry.location.lng]);
      // キャッシュに即時追加（同一実行内の後続呼び出しで重複しない）
      masterCache.rows.push(newEntry);
      masterCache.placeIds.add(cand.place_id);
      return newEntry;
    }

    logs && logs.push(`「${shopName}」→ 有効な候補が見つかりませんでした (クエリ: ${query})`);

  } catch (e) {
    console.error(e);
    logs && logs.push(`「${shopName}」→ 例外: ${e.message}`);
  }
  return null;
}

/**
 * 6. 距離マトリクス取得
 */
function fetchDistanceMatrix(origin, shops) {
  let addresses = [origin, ...shops.map(s => s.address)];
  let matrix = {};
  addresses.forEach(a => matrix[a] = {});
  const payload = {
    "origins": addresses.map(a => ({ "waypoint": { "address": a } })),
    "destinations": addresses.map(a => ({ "waypoint": { "address": a } })),
    "travelMode": "DRIVE", "routingPreference": "TRAFFIC_AWARE"
  };
  const response = UrlFetchApp.fetch('https://routes.googleapis.com/distanceMatrix/v2:computeRouteMatrix', {
    "method": "post", "contentType": "application/json",
    "headers": { "X-Goog-Api-Key": GOOGLE_MAPS_API_KEY, "X-Goog-FieldMask": "originIndex,destinationIndex,duration,distanceMeters,status,travelAdvisory" },
    "payload": JSON.stringify(payload), "muteHttpExceptions": true
  });
  if (response.getResponseCode() !== 200) throw new Error(`APIエラー: ${response.getResponseCode()} - ${response.getContentText().slice(0, 300)}`);
  const data = JSON.parse(response.getContentText());
  data.forEach(e => {
    if (e.duration) {
      let dur;
      if (typeof e.duration === "string") {
        dur = parseInt(e.duration.replace("s", ""), 10);
      } else if (typeof e.duration === "object" && e.duration !== null && e.duration.seconds != null) {
        dur = e.duration.seconds;
      } else {
        dur = Number(e.duration);
      }
      const hasToll = !!(e.travelAdvisory?.tollInfo?.estimatedPrice?.length);
      matrix[addresses[e.originIndex]][addresses[e.destinationIndex]] = { duration: dur, distance: e.distanceMeters || 0, hasToll };
    }
  });
  return matrix;
}

/**
 * 7. ルート①描画（ヘッダー行 + 停留点リスト）
 *    返値: 次に書き込めるシート行番号
 */
function renderRoute(sheet, route, failedShops, startTime, dataMatrix, startRow, label) {
  // ヘッダー行（青背景）
  const hdr1 = sheet.getRange(startRow, 1, 1, 6);
  hdr1.setValues([["#", "店舗名", "到着", "移動", "待機", ""]]);
  hdr1.setFontWeight("bold");
  hdr1.setBackground("#CFE2F3");

  let currentTime = new Date(startTime.getTime());
  let currentPos = START_POINT;
  let formulaOutput = [];
  let valueOutput = [];
  let waypoints = [];
  let dataRow = startRow + 1;

  route.forEach((shop, index) => {
    let travelSec = dataMatrix[currentPos]?.[shop.address]?.duration || 0;
    let travelMin = Math.ceil(travelSec / 60);
    let arrivalTime = new Date(currentTime.getTime() + travelMin * 60 * 1000);
    let status = checkSlots(arrivalTime.getHours() * 60 + arrivalTime.getMinutes(), shop.slots);
    let wait = status.type === "WAIT" ? status.waitTime : 0;
    let displayArrival = new Date(arrivalTime.getTime() + wait * 60 * 1000);
    let roundedMin = Math.ceil(displayArrival.getMinutes() / 10) * 10;
    if (roundedMin === 60) {
      displayArrival.setHours(displayArrival.getHours() + 1, 0, 0, 0);
    } else if (displayArrival.getMinutes() % 10 !== 0) {
      displayArrival.setMinutes(roundedMin, 0, 0);
    } else {
      displayArrival.setSeconds(0, 0);
    }

    let hasToll = !!(dataMatrix[currentPos]?.[shop.address]?.hasToll);
    let individualUrl = `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(shop.name + " " + shop.address)}&query_place_id=${shop.placeId}`;
    formulaOutput.push([`=HYPERLINK("${individualUrl}", "${index + 1}")`]);
    valueOutput.push([
      (shop.isPickup ? "🛒 " : "") + shop.name,
      Utilities.formatDate(displayArrival, "JST", "HH:mm"),
      travelMin + " 分",
      wait + " 分",
      (wait >= 15 ? "⚠️" : "") + (hasToll ? "🛣️" : "")
    ]);

    waypoints.push(shop.address);
    currentTime = new Date(arrivalTime.getTime() + (wait + 5) * 60 * 1000);
    currentPos = shop.address;
  });

  if (formulaOutput.length > 0) {
    sheet.getRange(dataRow, 1, formulaOutput.length, 1).setFormulas(formulaOutput);
    sheet.getRange(dataRow, 2, valueOutput.length, 5).setValues(valueOutput);
  }

  // 拠点への帰着時間を加算
  let returnMin = Math.ceil((dataMatrix[currentPos]?.[START_POINT]?.duration || 0) / 60);
  currentTime = new Date(currentTime.getTime() + returnMin * 60 * 1000);

  // 総所要時間 + Mapsリンク：ヘッダー行（startRow）のG・H列に出力
  let totalMin = Math.ceil((currentTime - startTime) / 60000);
  if (waypoints.length > 0) {
    let totalUrl = `https://maps.google.com/maps/dir/?api=1&origin=${encodeURIComponent(START_POINT)}&destination=${encodeURIComponent(START_POINT)}&waypoints=${waypoints.map(w => encodeURIComponent(w)).join('|')}&travelmode=driving`;
    sheet.getRange(startRow, 7).setFormula(`=HYPERLINK("${totalUrl}", "ROUTE 1  🗺️")`);
  } else {
    sheet.getRange(startRow, 7).setValue("ROUTE 1");
  }
  sheet.getRange(startRow, 8).setValue(`${Math.floor(totalMin/60)}h ${totalMin%60}m`);
  if (failedShops.length > 0) {
    sheet.getRange(startRow + 1, 7).setValue("⚠️ " + failedShops.map(s => s.name).join(", "));
  }

  return dataRow + route.length; // 次の空き行
}

/**
 * 8. ルート②描画（前半 → 帰着 → 後半）
 */
function renderRoute2(sheet, firstLeg, secondLeg, failedShops, startTime, returnTime, dataMatrix, startRow) {
  // ヘッダー行（緑背景）
  const hdr2 = sheet.getRange(startRow, 1, 1, 6);
  hdr2.setValues([["#", "店舗名", "到着", "移動", "待機", ""]]);
  hdr2.setFontWeight("bold");
  hdr2.setBackground("#D9EAD3");

  let currentTime = new Date(startTime.getTime());
  let currentPos = START_POINT;
  let dataRow = startRow + 1;
  let allWaypoints = [];

  // 前半を描画
  firstLeg.forEach((shop, index) => {
    let travelSec = dataMatrix[currentPos]?.[shop.address]?.duration || 0;
    let travelMin = Math.ceil(travelSec / 60);
    let arrivalTime = new Date(currentTime.getTime() + travelMin * 60 * 1000);
    let status = checkSlots(arrivalTime.getHours() * 60 + arrivalTime.getMinutes(), shop.slots);
    let wait = status.type === "WAIT" ? status.waitTime : 0;
    let displayArrival = new Date(arrivalTime.getTime() + wait * 60 * 1000);
    let roundedMin = Math.ceil(displayArrival.getMinutes() / 10) * 10;
    if (roundedMin === 60) {
      displayArrival.setHours(displayArrival.getHours() + 1, 0, 0, 0);
    } else if (displayArrival.getMinutes() % 10 !== 0) {
      displayArrival.setMinutes(roundedMin, 0, 0);
    } else {
      displayArrival.setSeconds(0, 0);
    }

    let hasToll1 = !!(dataMatrix[currentPos]?.[shop.address]?.hasToll);
    let individualUrl = `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(shop.name + " " + shop.address)}&query_place_id=${shop.placeId}`;
    sheet.getRange(dataRow, 1).setFormula(`=HYPERLINK("${individualUrl}", "${index + 1}")`);
    sheet.getRange(dataRow, 2, 1, 5).setValues([[
      (shop.isPickup ? "🛒 " : "") + shop.name,
      Utilities.formatDate(displayArrival, "JST", "HH:mm"),
      travelMin + " 分",
      wait + " 分",
      (wait >= 15 ? "⚠️" : "") + (hasToll1 ? "🛣️" : "")
    ]]);

    allWaypoints.push(shop.address);
    currentTime = new Date(arrivalTime.getTime() + (wait + 5) * 60 * 1000);
    currentPos = shop.address;
    dataRow++;
  });

  // 帰着行
  let returnTravelMin = Math.ceil((dataMatrix[currentPos]?.[START_POINT]?.duration || 0) / 60);
  let returnArrival = Utilities.formatDate(returnTime, "JST", "HH:mm");
  sheet.getRange(dataRow, 1, 1, 5).setValues([["", "↩ 北島ビルへ戻る", returnArrival, returnTravelMin + " 分", ""]]);
  sheet.getRange(dataRow, 1, 1, 5).setFontStyle("italic");
  dataRow++;

  // 後半を描画
  let leg2CurrentTime = new Date(returnTime.getTime());
  let leg2CurrentPos = START_POINT;
  secondLeg.forEach((shop, index) => {
    let travelSec = dataMatrix[leg2CurrentPos]?.[shop.address]?.duration || 0;
    let travelMin = Math.ceil(travelSec / 60);
    let arrivalTime = new Date(leg2CurrentTime.getTime() + travelMin * 60 * 1000);
    let status = checkSlots(arrivalTime.getHours() * 60 + arrivalTime.getMinutes(), shop.slots);
    let wait = status.type === "WAIT" ? status.waitTime : 0;
    let displayArrival = new Date(arrivalTime.getTime() + wait * 60 * 1000);
    let roundedMin = Math.ceil(displayArrival.getMinutes() / 10) * 10;
    if (roundedMin === 60) {
      displayArrival.setHours(displayArrival.getHours() + 1, 0, 0, 0);
    } else if (displayArrival.getMinutes() % 10 !== 0) {
      displayArrival.setMinutes(roundedMin, 0, 0);
    } else {
      displayArrival.setSeconds(0, 0);
    }

    let hasToll2 = !!(dataMatrix[leg2CurrentPos]?.[shop.address]?.hasToll);
    let num = firstLeg.length + 1 + index;
    let individualUrl = `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(shop.name + " " + shop.address)}&query_place_id=${shop.placeId}`;
    sheet.getRange(dataRow, 1).setFormula(`=HYPERLINK("${individualUrl}", "${num}")`);
    sheet.getRange(dataRow, 2, 1, 5).setValues([[
      (shop.isPickup ? "🛒 " : "") + shop.name,
      Utilities.formatDate(displayArrival, "JST", "HH:mm"),
      travelMin + " 分",
      wait + " 分",
      (wait >= 15 ? "⚠️" : "") + (hasToll2 ? "🛣️" : "")
    ]]);

    allWaypoints.push(shop.address);
    leg2CurrentTime = new Date(arrivalTime.getTime() + (wait + 5) * 60 * 1000);
    leg2CurrentPos = shop.address;
    dataRow++;
  });

  // ルート②の総所要時間 + Mapsリンク（拠点帰着を含む）
  let returnMin2 = Math.ceil((dataMatrix[leg2CurrentPos]?.[START_POINT]?.duration || 0) / 60);
  leg2CurrentTime = new Date(leg2CurrentTime.getTime() + returnMin2 * 60 * 1000);
  let totalMin2 = Math.ceil((leg2CurrentTime - startTime) / 60000);
  if (allWaypoints.length > 0) {
    let totalUrl2 = `https://maps.google.com/maps/dir/?api=1&origin=${encodeURIComponent(START_POINT)}&destination=${encodeURIComponent(START_POINT)}&waypoints=${allWaypoints.join('|')}&travelmode=driving`;
    sheet.getRange(startRow, 7).setFormula(`=HYPERLINK("${totalUrl2}", "ROUTE 2  🗺️")`);
  } else {
    sheet.getRange(startRow, 7).setValue("ROUTE 2");
  }
  sheet.getRange(startRow, 8).setValue(`${Math.floor(totalMin2/60)}h ${totalMin2%60}m`);
}

/**
 * 9. 営業スロット生成
 */
function generateDeliverySlots(hourStr, startTime) {
  if (!hourStr || hourStr === "定休日") return [];
  if (hourStr === "不明") return [{ start: 0, end: 1440 }]; // 営業時間不明は終日対応とみなす
  if (hourStr === "24 時間営業") return [{ start: 0, end: 1440 }];

  let cleanStr = String(hourStr)
    .replace(/[０-９：]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
    .replace(/時/g, ":").replace(/分/g, "").replace(/翌/g, "")
    .replace(/～|ー|–|—|〜/g, "-").replace(/\s/g, "");

  let slots = [];
  let startDayMin = startTime.getHours() * 60 + startTime.getMinutes();

  cleanStr.split(/[、, ]+/).forEach(slotStr => {
    let match = slotStr.match(/(\d{1,2}):(\d{2})-(\d{1,2}):(\d{2})/);
    if (match) {
      let openH = parseInt(match[1], 10), openM = parseInt(match[2], 10);
      let closeH = parseInt(match[3], 10), closeM = parseInt(match[4], 10);
      let open = openH * 60 + openM, close = closeH * 60 + closeM;

      if (close <= open && close < 720) close += 1440;

      let s = open - 30, e = close + 30;

      if (e > startDayMin && Math.max(s, startDayMin) <= e) {
        slots.push({ start: Math.max(s, startDayMin), end: e });
      }
    }
  });
  return slots;
}

/**
 * 10. 補助関数
 */
function normalizeHours(rawStr) {
  if (!rawStr) return "不明";
  const s = String(rawStr);
  if (s === "不明" || s === "定休日") return s;
  if (/24\s*時間営業/.test(s)) return "00:00-24:00";

  let work = s
    .replace(/[０-９：]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .replace(/時/g, ":").replace(/分/g, "")
    .replace(/翌/g, "")
    .replace(/～|ー|–|—|〜/g, "-")
    .replace(/\s/g, "");

  work = work.replace(/(\d{1,2}):(\d{2})([Aa][Mm])/g, (_, h, m) =>
    `${String(parseInt(h) === 12 ? 0 : parseInt(h)).padStart(2,"0")}:${m}`
  );
  work = work.replace(/(\d{1,2}):(\d{2})([Pp][Mm])/g, (_, h, m) =>
    `${String(parseInt(h) === 12 ? 12 : parseInt(h) + 12).padStart(2,"0")}:${m}`
  );

  const slots = [...work.matchAll(/(\d{1,2}):(\d{2})-(\d{1,2}):(\d{2})/g)].map(m =>
    `${String(m[1]).padStart(2,"0")}:${m[2]}-${String(m[3]).padStart(2,"0")}:${m[4]}`
  );
  return slots.length > 0 ? slots.join(",") : s;
}

function checkSlots(arrivalMin, slots) {
  if (slots.length === 0) return { type: "FAILED" };
  let minWait = Infinity;
  for (let slot of slots) {
    if (arrivalMin >= slot.start && arrivalMin <= slot.end) return { type: "AVAILABLE" };
    if (arrivalMin < slot.start) minWait = Math.min(minWait, slot.start - arrivalMin);
  }
  return (minWait !== Infinity) ? { type: "WAIT", waitTime: minWait } : { type: "FAILED" };
}

/**
 * マルチスタート + ILS（反復局所探索）：摂動→再最適化を繰り返して局所最適を脱出
 */
function multiStartOptimize(shops, startPos, startTime, dataMatrix, waitThresholdMin, numRestarts = 5) {
  let bestRoute = null;
  let bestFailedShops = [];
  let bestCost = Infinity;

  for (let attempt = 0; attempt < numRestarts; attempt++) {
    const shuffled = attempt === 0 ? [...shops] : shuffleArray([...shops]);
    const { route: base, failedShops } = buildRoute(shuffled, startPos, startTime, dataMatrix, waitThresholdMin);
    let current = optimizeOrOpt(optimize2Opt(base, startTime, dataMatrix), startTime, dataMatrix);
    let currentCost = evaluateRouteCost(current, startTime, dataMatrix);

    // ILS: 摂動（ルートを大きく崩す）→ 再最適化 を繰り返す
    for (let iter = 0; iter < 5; iter++) {
      const perturbed = perturbRoute(current);
      const reopt = optimizeOrOpt(optimize2Opt(perturbed, startTime, dataMatrix), startTime, dataMatrix);
      const reoptCost = evaluateRouteCost(reopt, startTime, dataMatrix);
      if (reoptCost.isValid && reoptCost.totalTime < currentCost.totalTime) {
        current = reopt;
        currentCost = reoptCost;
      }
    }

    if (currentCost.isValid && currentCost.totalTime < bestCost) {
      bestRoute = current;
      bestFailedShops = failedShops;
      bestCost = currentCost.totalTime;
    }
  }
  return { route: bestRoute || [], failedShops: bestFailedShops };
}

// ルートを大きく崩す摂動（セグメント入れ替え）：2-opt/Or-optでは到達できない解空間へ
function perturbRoute(route) {
  const n = route.length;
  if (n < 4) return [...route];

  for (let tries = 0; tries < 20; tries++) {
    const cuts = new Set();
    while (cuts.size < 3) cuts.add(1 + Math.floor(Math.random() * (n - 1)));
    const [c1, c2, c3] = [...cuts].sort((a, b) => a - b);
    // A + C + B + D（セグメントBとCを入れ替え）
    const newRoute = [
      ...route.slice(0, c1),
      ...route.slice(c2, c3),
      ...route.slice(c1, c2),
      ...route.slice(c3)
    ];
    if (!violatesPickupMemoOrder(newRoute)) return newRoute;
  }
  return [...route];
}

function shuffleArray(arr) {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}

/**
 * Or-opt：1〜3店舗のブロックを別の位置に移動して改善を探す（2-optより広い探索）
 */
function optimizeOrOpt(route, startTime, dataMatrix) {
  if (route.length < 3) return route;
  let bestRoute = [...route];
  let bestEval = evaluateRouteCost(bestRoute, startTime, dataMatrix);
  let improved = true;

  while (improved) {
    improved = false;
    for (let segLen = 1; segLen <= Math.min(3, bestRoute.length - 1); segLen++) {
      for (let i = 0; i <= bestRoute.length - segLen; i++) {
        const segment = bestRoute.slice(i, i + segLen);
        const remaining = bestRoute.slice(0, i).concat(bestRoute.slice(i + segLen));

        for (let j = 0; j <= remaining.length; j++) {
          if (j === i) continue; // 元の位置と同じなのでスキップ
          const newRoute = remaining.slice(0, j).concat(segment, remaining.slice(j));
          if (violatesPickupMemoOrder(newRoute)) continue;
          const newEval = evaluateRouteCost(newRoute, startTime, dataMatrix);
          if (newEval.isValid && newEval.totalTime < bestEval.totalTime) {
            bestRoute = newRoute;
            bestEval = newEval;
            improved = true;
          }
        }
      }
    }
  }
  return bestRoute;
}

function optimize2Opt(route, startTime, dataMatrix) {
  let bestRoute = [...route], bestEval = evaluateRouteCost(bestRoute, startTime, dataMatrix);
  if (bestRoute.length < 3) return bestRoute;
  let improved = true;
  while (improved) {
    improved = false;
    for (let i = 0; i < bestRoute.length - 1; i++) {
      for (let j = i + 1; j < bestRoute.length; j++) {
        let newRoute = bestRoute.slice(0, i).concat(bestRoute.slice(i, j + 1).reverse()).concat(bestRoute.slice(j + 1));
        if (violatesPickupMemoOrder(newRoute)) continue; // @/* 順序を守る
        let newEval = evaluateRouteCost(newRoute, startTime, dataMatrix);
        if (newEval.isValid && newEval.totalTime < bestEval.totalTime) {
          bestRoute = newRoute; bestEval = newEval; improved = true;
        }
      }
    }
  }
  return bestRoute;
}

// @ 店舗が * 店舗より後に来るルートはNG
function violatesPickupMemoOrder(route) {
  const lastPickupIdx = route.reduce((acc, s, i) => s.isPickup ? i : acc, -1);
  const firstMemoIdx = route.findIndex(s => s.isMemo);
  if (lastPickupIdx === -1 || firstMemoIdx === -1) return false;
  return lastPickupIdx > firstMemoIdx;
}

function evaluateRouteCost(route, startTime, dataMatrix) {
  let currentPos = START_POINT, currentTime = new Date(startTime.getTime()), isValid = true;
  for (let shop of route) {
    let travelMin = Math.ceil((dataMatrix[currentPos]?.[shop.address]?.duration || 0) / 60);
    let arrivalTime = new Date(currentTime.getTime() + travelMin * 60 * 1000);
    let status = checkSlots(arrivalTime.getHours()*60 + arrivalTime.getMinutes(), shop.slots);
    if (status.type === "FAILED") { isValid = false; break; }
    currentTime = new Date(arrivalTime.getTime() + ((status.type === "WAIT" ? status.waitTime : 0) + 5) * 60 * 1000);
    currentPos = shop.address;
  }
  // 拠点への帰着時間を含めて評価（始点→全店舗→終点の最短）
  let returnMin = Math.ceil((dataMatrix[currentPos]?.[START_POINT]?.duration || 0) / 60);
  currentTime = new Date(currentTime.getTime() + returnMin * 60 * 1000);
  return { isValid: isValid, totalTime: (currentTime - startTime) };
}
