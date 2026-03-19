/**
 * sedori_appsheet_tool (プロ仕様リサーチツール v2.2)
 * 
 * AppSheet × GAS × Keepa API 連携
 * JANコードからKeepa APIを呼び出し、正規JANを除いた全32項目を取得・出力
 */

// ==========================================
// 設定エリア
// ==========================================

const KEEPA_API_KEY = 'psr7fkj9soadqmmqptf70e34bs6t317ujjptvul3vfi5i5hvcmrst4p8hf22aqmb';
const TARGET_SHEET_NAME = "リサーチリスト";

// APIキー群 (スクリプトプロパティから取得)
function getApiTokens() {
  const props = PropertiesService.getScriptProperties();
  return {
    LINE_ACCESS_TOKEN: props.getProperty('LINE_ACCESS_TOKEN') || '',
    LINE_USER_ID: props.getProperty('LINE_USER_ID') || '',
    KV_REST_API_URL: props.getProperty('KV_REST_API_URL') || '',
    KV_REST_API_TOKEN: props.getProperty('KV_REST_API_TOKEN') || ''
  };
}

const WANTED_LIST_TARGETS = [
  { sheetName: 'Keepa指名手配リスト', flag: '🚓 逮捕(keepa)' },
  { sheetName: 'ニッチ指名手配リスト', flag: '🕸️ 捕獲(ニッチ)' },
  { sheetName: 'ASIN指名手配リスト', flag: '🎯 発見(ASIN)' },
  { sheetName: 'cosme_価格チェック_0224', flag: '💄 発見(cosme)' }
];

// 列番号の設定
const COL_JAN_INPUT    = 2;   // B: JAN
const COL_TITLE        = 3;   // C: 商品名
const COL_ASIN         = 4;   // D: ASIN
const COL_BRAND        = 5;   // E: ブランド
const COL_RANK         = 6;   // F: 売れ筋ランキング
const COL_MONTHLY_SOLD = 7;   // G: 先月購入数
const COL_CATEGORY     = 8;   // H: カテゴリー
const COL_SELLER_COUNT = 9;   // I: セラー数
const COL_VARIATION    = 10;  // J: バリエーション
const COL_BUYBOX       = 11;  // K: カート価格
const COL_NEW_PRICE    = 12;  // L: 新品現在価格
const COL_FBA_LOWEST   = 13;  // M: FBA最安値
const COL_PURCHASE     = 14;  // N: 仕入価格
const COL_PROFIT       = 15;  // O: 粗利益
const COL_BREAK_EVEN   = 16;  // P: 損益分岐(仕入上限)
const COL_PROFIT_RATE  = 17;  // Q: 利益率
const COL_ROI          = 18;  // R: ROI
const COL_JUDGMENT     = 19;  // S: 仕入判定
const COL_FBA_FEE      = 20;  // T: FBA手数料
const COL_REF_RATE     = 21;  // U: 紹介料率
const COL_SIZE_WEIGHT  = 22;  // V: 重量・サイズ
const COL_HAZMAT       = 23;  // W: 危険物
const COL_LINK_AMAZON  = 24;  // X: Amazonリンク
const COL_LINK_KEEPA   = 25;  // Y: Keepaリンク
const COL_LINK_POI     = 26;  // Z: sedori_poipoi
const COL_IMAGE        = 27;  // AA: 画像URL
const COL_AMAZON_SELL  = 28;  // AB: Amazon本体有無
const COL_RESTRICTION  = 29;  // AC: 出品制限フラグ
const COL_DROPS_30     = 30;  // AD: 下落回数(30日)
const COL_SHIPPING     = 31;  // AE: 納品送料概算
const COL_RESEARCH_DT  = 32;  // AF: リサーチ日時
const COL_SKU          = 36;  // AJ: SKU
const COL_WANTED_FLAG  = 37;  // AK: 手配書フラグ

const DATA_COL_COUNT = COL_RESEARCH_DT - COL_TITLE + 1;

// ==========================================
// 入力判定ヘルパー
// ==========================================
function detectInputType(input) {
  if (!input) return null;
  const s = String(input).trim();
  if (/^[A-Z0-9]{10}$/.test(s) && /^B[0-9]/.test(s)) return 'asin';
  if (/^\d{8}$/.test(s) || /^\d{13}$/.test(s)) return 'jan';
  return 'jan';
}

// ==========================================
// 手配書マッチング・逮捕処理
// ==========================================
function getWantedJanMap() {
  const map = new Map();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  WANTED_LIST_TARGETS.forEach(t => {
    const sheet = ss.getSheetByName(t.sheetName);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colJanIdx = headers.findIndex(h => String(h).toUpperCase() === 'JAN');
    if (colJanIdx === -1) return;
    const colJan = colJanIdx + 1;
    
    const values = sheet.getRange(2, colJan, lastRow - 1, 1).getValues();
    values.forEach((row, r) => {
      const jan = String(row[0]).trim();
      if (jan && jan !== "") {
        if (!map.has(jan)) {
          map.set(jan, [{ sheetName: t.sheetName, rowIdx: r + 2, flag: t.flag }]);
        } else {
          const arr = map.get(jan);
          const exists = arr.find(item => item.sheetName === t.sheetName && item.flag === t.flag);
          if (!exists) arr.push({ sheetName: t.sheetName, rowIdx: r + 2, flag: t.flag });
        }
      }
    });
  });
  return map;
}

function getJudgmentFlagString(matches) {
  if (!matches || matches.length === 0) return "";
  const flags = matches.map(m => m.flag);
  return [...new Set(flags)].join(' | ');
}

function markWantedListAsArrested(ss, matches) {
  if (!matches || matches.length === 0) return;
  matches.forEach(match => {
    const sheet = ss.getSheetByName(match.sheetName);
    if (!sheet) return;
    sheet.getRange(match.rowIdx, 1).setValue("✅逮捕済");
    const lastCol = sheet.getLastColumn();
    if (lastCol > 0) sheet.getRange(match.rowIdx, 1, 1, lastCol).setBackground("#e0e0e0");
  });
}

// ==========================================
// トリガー関数群
// ==========================================

function onChange(e) {
  console.log("🌸 onChange トリガー発火 (AppSheet連携)");
  // AppSheet側の同期書き込み処理が完全に終わるのを待つ（競合での上書き防止）
  Utilities.sleep(3000);
  autoResearch(null);
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== TARGET_SHEET_NAME) return;

  const row = e.range.getRow();
  if (row < 2) return; // ヘッダーは除外

  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const colPurchased = headers.indexOf('仕入済') + 1;
  const colTransferred = headers.indexOf('転送済フラグ') + 1;

  if (colPurchased === 0 || colTransferred === 0) {
    autoResearch(e);
    return;
  }

  const editedCol = e.range.getColumn();

  if (editedCol === colPurchased) {
    const isPurchased = sheet.getRange(row, colPurchased).getValue() === true;
    const isTransferred = sheet.getRange(row, colTransferred).getValue() === true;

    if (isPurchased && !isTransferred) {
      transferToPurchaseData(sheet, row);
      sheet.getRange(row, colTransferred).setValue(true);
    }
  } else {
    autoResearch(e);
  }
}

function testRun() {
  autoResearch(null);
}

// ==========================================
// メイン処理 (AppSheet対応版)
// ==========================================

function autoResearch(e) {
  console.log("🌸 autoResearch 開始...");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const lastCol = sheet.getLastColumn();
  if (lastCol < COL_JAN_INPUT) return;

  const fullRange = sheet.getRange(1, COL_JAN_INPUT, lastRow, DATA_COL_COUNT + 1);
  const allValues = fullRange.getValues();

  const cacheMap = new Map();
  // 1. すでにタイトルが入っている行のデータをローカルキャッシュに蓄積
  for (let r = 1; r < allValues.length; r++) {
    const input = String(allValues[r][0]).trim();
    const title = allValues[r][1];
    if (input && input !== "" && title && title !== "" && !title.startsWith("[Error") && title !== "見つかりませんでした") {
      const rowData = allValues[r].slice(1, DATA_COL_COUNT + 1);
      cacheMap.set(input, rowData);
    }
  }

  // グローバル変数に依存せず、1回だけ手配書リストを取得する
  const wantedMap = getWantedJanMap();

  // 2. タイトルが空でJANのみ入力されている行に対して処理を行う
  for (let i = 1; i < allValues.length; i++) {
    const currentRow = i + 1;
    const input = String(allValues[i][0]).trim();
    const title = allValues[i][1];

    if (input && input !== "" && (!title || title === "")) {
      if (cacheMap.has(input)) {
        console.log(`🌸 キャッシュヒット: ${input} -> 行: ${currentRow} に書き込みます`);
        const cachedData = cacheMap.get(input);
        sheet.getRange(currentRow, COL_TITLE, 1, DATA_COL_COUNT).setValues([cachedData]);
        SpreadsheetApp.flush(); // 即時反映して競合を防ぐ
        
        const matches = wantedMap.get(input) || [];
        if (matches.length > 0) markWantedListAsArrested(ss, matches);
        continue;
      }
      
      const canContinue = fetchProductData(input, currentRow, sheet, ss, wantedMap);
      if (canContinue === false) break;
      Utilities.sleep(1000);
    }
  }
}

// ==========================================
// データ取得・書き込みロジック
// ==========================================

function fetchProductData(barcode, row, sheet, ss, wantedMap) {
  try {
    let product = getFromKvCache(barcode);
    let usedCache = !!product;

    if (!usedCache) {
      const inputType = detectInputType(barcode);
      const paramKey = (inputType === 'asin') ? 'asin' : 'code';
      const url = `https://api.keepa.com/product?key=${KEEPA_API_KEY}&domain=5&type=product&${paramKey}=${barcode}&stats=1`;
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const json = JSON.parse(response.getContentText());

      if (json.error || !json.products || json.products.length === 0) {
        if (json.error) setError(sheet, row, `API Error: ${json.error.message}`);
        else setNotFound(sheet, row);
        return true;
      }
      product = json.products[0];
      setToKvCache(barcode, product);
    }

    const asin = product.asin;
    const title = product.title || "Unknown";
    const brand = product.brand || "";
    let categoryName = "";
    if (product.categoryTree && product.categoryTree.length > 0) {
      const rootCat = product.categoryTree.find(c => c.catId === product.rootCategory);
      categoryName = rootCat ? rootCat.name : product.categoryTree[0].name;
    }
    const imageCsv = product.imagesCSV || "";
    const imageUrl = imageCsv ? `https://m.media-amazon.com/images/I/${imageCsv.split(",")[0]}` : "";
    const stats = product.stats || {};
    const current = stats.current || [];
    const rank = current[3] || "";
    let newPrice = current[1] || current[0] || "";
    let buyBox = current[18] || stats.buyBoxPrice || newPrice;
    let fbaLowest = current[7] || "";
    let sellerCount = current[11] || "";
    const isVariation = (product.variationCSV && product.variationCSV.length > 0) ? "有" : "無";
    const isAmazonSelling = (current[0] && current[0] > 0) ? "有" : "無";
    let monthlySold = product.monthlySold !== undefined ? product.monthlySold : stats.salesRankDrops30 || "";
    let drops30 = stats.salesRankDrops30 || "";
    let fbaFee = (product.fbaFees && product.fbaFees.pickAndPackFee) ? product.fbaFees.pickAndPackFee : "";
    let refRate = product.referralFeePercent || "";

    let breakEven = "";
    if (buyBox && buyBox > 0 && fbaFee) {
      const rate = refRate ? (parseFloat(refRate) / 100) : 0.15;
      breakEven = Math.floor(buyBox - fbaFee - (buyBox * rate));
    }

    let sizeWeight = "";
    if (product.packageLength > 0) sizeWeight = `${product.packageLength}x${product.packageWidth}x${product.packageHeight}mm ${product.packageWeight}g`;

    let isHazmat = product.hazardousMaterialType ? "Yes" : "No";
    if (title.toLowerCase().includes("battery") || title.includes("電池")) isHazmat = "Yes";

    const amazonLink = `https://www.amazon.co.jp/dp/${asin}`;
    const keepaLink = `https://keepa.com/#!product/5-${asin}`;
    const poiLink = `http://localhost:3000/?q=${asin}`;
    const now = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");

    const matches = wantedMap.get(String(barcode).trim()) || [];
    const judgmentFlag = getJudgmentFlagString(matches);
    const existingPurchasePrice = sheet.getRange(row, COL_PURCHASE).getValue() || "";
    const sku = `${existingPurchasePrice || 0}_${Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd")}_${breakEven || 0}_${asin}`;

    const values = [[
      title, asin, brand, rank, monthlySold, categoryName, sellerCount, isVariation,
      buyBox, newPrice, fbaLowest, existingPurchasePrice, "", breakEven, "", "", "",
      fbaFee, refRate, sizeWeight, isHazmat, amazonLink, keepaLink, poiLink, imageUrl,
      isAmazonSelling, "", drops30, "", now
    ]];

    sheet.getRange(row, COL_TITLE, 1, values[0].length).setValues(values);
    sheet.getRange(row, COL_SKU).setValue(sku);
    sheet.getRange(row, COL_WANTED_FLAG).setValue(judgmentFlag);
    SpreadsheetApp.flush(); // 即時反映して競合を防ぐ

    if (matches.length > 0 && ss) markWantedListAsArrested(ss, matches);
    console.log(`🌸 API リサーチ完了: ${title} -> 行: ${row} に書き込みました`);
    return true;
  } catch (e) { setError(sheet, row, e.toString()); return true; }
}

// ==========================================
// キャッシュ / LINE / ユーティリティ
// ==========================================

function getFromKvCache(barcode) {
  const tokens = getApiTokens();
  if (!tokens.KV_REST_API_URL || !tokens.KV_REST_API_TOKEN) return null;
  try {
    const response = UrlFetchApp.fetch(tokens.KV_REST_API_URL, {
      method: "post",
      headers: { "Authorization": `Bearer ${tokens.KV_REST_API_TOKEN}` },
      payload: JSON.stringify(["GET", `product:keepa:${barcode}`]),
      contentType: "application/json", muteHttpExceptions: true
    });
    if (response.getResponseCode() === 200) {
      const json = JSON.parse(response.getContentText());
      if (json && json.result) return JSON.parse(json.result);
    }
  } catch(e) { console.warn("KV GET Error"); }
  return null; 
}

function setToKvCache(barcode, productData) {
  const tokens = getApiTokens();
  if (!tokens.KV_REST_API_URL || !tokens.KV_REST_API_TOKEN) return;
  try {
    UrlFetchApp.fetch(tokens.KV_REST_API_URL, {
      method: "post",
      headers: { "Authorization": `Bearer ${tokens.KV_REST_API_TOKEN}` },
      payload: JSON.stringify(["SET", `product:keepa:${barcode}`, JSON.stringify(productData), "EX", 259200]), 
      contentType: "application/json", muteHttpExceptions: true
    });
  } catch(e) { console.warn("KV SET Error"); }
}

function sendLineNotification(message) {
  const tokens = getApiTokens();
  if (!tokens.LINE_ACCESS_TOKEN || !tokens.LINE_USER_ID) return;
  const options = {
    "method": "post",
    "headers": { "Content-Type": "application/json", "Authorization": "Bearer " + tokens.LINE_ACCESS_TOKEN },
    "payload": JSON.stringify({ "to": tokens.LINE_USER_ID, "messages": [{ "type": "text", "text": message.toString() }] }),
    "muteHttpExceptions": true
  };
  try { UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', options); } catch(e) {}
}

function setNotFound(sheet, row) { sheet.getRange(row, COL_TITLE).setValue("見つかりませんでした"); }
function setError(sheet, row, message) { sheet.getRange(row, COL_TITLE).setValue(`[Error] ${message}`); }

// ==========================================
// 外部連携 (zaiko_tool / doPost)
// ==========================================
const ZAIKO_TOOL_SPREADSHEET_ID = '1EIYt3IP7FidK-RbmNj2MYNamIbVvbt_VMVs89FMYkKY';
const ZAIKO_SHEET_NAME = '仕入れデータ';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📱 AppSheet連携')
    .addItem('📥 AppSheetの「仕入済」をzaiko_toolへ送信', 'syncAppSheetPurchases')
    .addToUi();
}

function syncAppSheetPurchases() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colPurchased = headers.indexOf('仕入済') + 1;
  const colTransferred = headers.indexOf('転送済フラグ') + 1;
  if (colPurchased === 0 || colTransferred === 0) return;
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  let count = 0;
  for (let i = 0; i < data.length; i++) {
    if (data[i][colPurchased - 1] === true && data[i][colTransferred - 1] !== true) {
      transferToPurchaseData(sheet, i + 2);
      sheet.getRange(i + 2, colTransferred).setValue(true);
      count++;
    }
  }
  if (count > 0) SpreadsheetApp.getUi().alert('転送完了', `${count}件送信しました。`, SpreadsheetApp.getUi().ButtonSet.OK);
}

function transferToPurchaseData(sourceSheet, row) {
  try {
    const dataRow = sourceSheet.getRange(row, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
    const appId = dataRow[0] || `ID-${new Date().getTime()}`;
    const researchDate = dataRow[COL_RESEARCH_DT - 1] || Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");
    const janCode = dataRow[COL_JAN_INPUT - 1] || "";
    const itemName = dataRow[COL_TITLE - 1] || "";
    const unitPrice = dataRow[COL_PURCHASE - 1] || 0;
    const defaultQty = 1;
    const totalPrice = Number(unitPrice) * defaultQty;
    const shopName = "";
    const condition = "";
    const imageUrl = dataRow[COL_IMAGE - 1] || "";
    const receiptImage = "";

    const zaikoSs = SpreadsheetApp.openById(ZAIKO_TOOL_SPREADSHEET_ID);
    const targetSheet = zaikoSs.getSheetByName(ZAIKO_SHEET_NAME);
    
    if (targetSheet) {
      targetSheet.appendRow([
        appId, researchDate, janCode, itemName, unitPrice, defaultQty, totalPrice, shopName, condition, imageUrl, receiptImage
      ]);
    }

    const dateObj = new Date(researchDate);
    const isValidDate = !isNaN(dateObj.getTime());
    const dateStr = isValidDate ? Utilities.formatDate(dateObj, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') : Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
    const dateIdPart = isValidDate ? Utilities.formatDate(dateObj, 'Asia/Tokyo', 'yyyyMMdd') : Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    const totalAmountStr = totalPrice + '円';
    const strAppId = String(appId);
    const shortId = strAppId.length > 4 ? strAppId.slice(-4) : strAppId.padStart(4, '0');

    const ecHistorySheet = zaikoSs.getSheetByName('EC注文履歴');
    if (ecHistorySheet) {
      const lastRow = ecHistorySheet.getLastRow();
      let isDuplicateEC = false;
      if (lastRow > 1) {
        const ids = ecHistorySheet.getRange(2, 9, lastRow - 1, 1).getValues().flat();
        if (ids.includes(appId)) isDuplicateEC = true;
      }
      if (!isDuplicateEC) {
        ecHistorySheet.appendRow([dateStr, shopName, itemName, unitPrice, defaultQty, totalAmountStr, appId, now, appId]);
      }
    }

    const invMasterSheet = zaikoSs.getSheetByName('在庫管理マスタ');
    if (invMasterSheet) {
      const lastRow = invMasterSheet.getLastRow();
      let isDuplicateInv = false;
      if (lastRow > 1) {
        const ids = invMasterSheet.getRange(2, 12, lastRow - 1, 1).getValues().flat();
        if (ids.includes(appId)) isDuplicateInv = true;
      }
      if (!isDuplicateInv) {
        const newRows = [];
        const cleanDateStr = isValidDate ? Utilities.formatDate(dateObj, 'Asia/Tokyo', 'yyyy-MM-dd') : Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
        for (let i = 1; i <= defaultQty; i++) {
          const invId = `INV-${dateIdPart}-${shortId}-${i}`;
          newRows.push([invId, itemName, cleanDateStr, shopName, unitPrice, "出品待ち", "", "", "", "", "", appId]);
        }
        if (newRows.length > 0) invMasterSheet.getRange(invMasterSheet.getLastRow() + 1, 1, newRows.length, 12).setValues(newRows);
      }
    }
  } catch(e) { console.error("転送エラー: " + e.toString()); }
}

const POST_API_KEY = "muqzyz-tonhyz-Disry7";

function doPost(e) {
  try {
    const json = JSON.parse(e.postData.contents);
    if (json.api_key !== POST_API_KEY) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Invalid API Key" })).setMimeType(ContentService.MimeType.JSON);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = json.sheet_name || "results";
    let sheet = ss.getSheetByName(sheetName);

    const startCol = json.start_col || 1;

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    if (json.clear === true) {
      sheet.clear();
    }

    if (sheet.getLastRow() === 0 && json.headers && Array.isArray(json.headers)) {
      sheet.getRange(1, startCol, 1, json.headers.length).setValues([json.headers]);
    }

    const rows = json.data;

    if (Array.isArray(rows) && rows.length > 0) {
      const lastRow = sheet.getLastRow();
      const numRows = rows.length;
      const numCols = rows[0].length;
      if (numCols > 0) {
        sheet.getRange(lastRow + 1, startCol, numRows, numCols).setValues(rows);
      }
    }

    if (sheetName === TARGET_SHEET_NAME) {
      try {
        autoResearch(null);
      } catch (researchErr) {
        console.warn("doPost後の自動リサーチでエラー: " + researchErr.toString());
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ status: "success", message: `${rows ? rows.length : 0} rows added` })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}
