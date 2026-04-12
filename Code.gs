/**
 * sedori_appsheet_tool (プロ仕様リサーチツール v3.0 - 動的カラムマッピング対応版)
 * 
 * AppSheet × GAS × Keepa API 連携
 * シートごとの異なるカラム配置をヘッダーから自動検知して処理します。
 */

// ==========================================
// 設定エリア
// ==========================================

// 処理対象となるシート名のリスト（これ以外のシートで編集されても発火しない）
const TARGET_SHEET_NAMES = [
  "リサーチリスト", 
  "プレ値リスト", 
  "Keepa指名手配リスト", 
  "ニッチ指名手配リスト", 
  "ASIN指名手配リスト",
  "ターゲットリスト",
  "cosme_価格チェック_0224"
];

// APIキー群 (スクリプトプロパティから取得)
function getApiTokens() {
  const props = PropertiesService.getScriptProperties();
  return {
    KEEPA_API_KEY: props.getProperty('KEEPA_API_KEY') || '',
    POST_API_KEY: props.getProperty('POST_API_KEY') || '',
    LINE_ACCESS_TOKEN: props.getProperty('LINE_ACCESS_TOKEN') || '',
    LINE_USER_ID: props.getProperty('LINE_USER_ID') || '',
    KV_REST_API_URL: props.getProperty('KV_REST_API_URL') || '',
    KV_REST_API_TOKEN: props.getProperty('KV_REST_API_TOKEN') || ''
  };
}

function getRequiredApiToken(keyName) {
  const value = getApiTokens()[keyName];
  if (!value) {
    throw new Error(`Script Properties に ${keyName} が未設定です`);
  }
  return value;
}

const WANTED_LIST_TARGETS = [
  { sheetName: 'Keepa指名手配リスト', flag: '🚓 逮捕(keepa)' },
  { sheetName: 'ニッチ指名手配リスト', flag: '🕸️ 捕獲(ニッチ)' },
  { sheetName: 'ASIN指名手配リスト', flag: '🎯 発見(ASIN)' },
  { sheetName: 'cosme_価格チェック_0224', flag: '💄 発見(cosme)' }
];

// ==========================================
// 動的カラムマッピング (全シート対応)
// ==========================================

const COL_ALIASES = {
  JAN: ['JAN', 'JANコード', 'jan'],
  TITLE: ['商品名', 'タイトル'],
  ASIN: ['ASIN', 'asin'],
  BRAND: ['ブランド'],
  RANK: ['売れ筋ランキング', 'ランキング'],
  MONTHLY_SOLD: ['先月購入数', '月間販売数'],
  CATEGORY: ['カテゴリー', 'ランキングカテゴリ'],
  SELLER_COUNT: ['セラー数'],
  VARIATION: ['バリエーション', 'バリエーション数'],
  BUYBOX: ['カート価格', 'カート価格(基準)'],
  NEW_PRICE: ['新品現在価格'],
  FBA_LOWEST: ['FBA最安値'],
  PURCHASE: ['仕入価格', '仕入価格(入力用)'],
  PROFIT: ['粗利益'],
  BREAK_EVEN: ['損益分岐(仕入上限)', '損益分岐点'],
  PROFIT_RATE: ['利益率'],
  ROI: ['ROI'],
  LIST_PRICE: ['定価/通常価格', '定価'],
  PREMIUM_JUDGE: ['プレ値判定', 'プレ値'],
  JUDGMENT: ['仕入判定', '判定(自動計算)', '判定'],
  FBA_FEE: ['FBA手数料'],
  REF_RATE: ['紹介料率'],
  SIZE_WEIGHT: ['重量・サイズ', 'サイズ・重量'],
  HAZMAT: ['危険物'],
  LINK_AMAZON: ['Amazonリンク'],
  LINK_KEEPA: ['Keepaリンク'],
  LINK_POI: ['sedori_poipoi', 'sedori_poipoiリンク'],
  IMAGE: ['画像URL'],
  AMAZON_SELL: ['Amazon本体有無', 'Amazon本体'],
  RESTRICTION: ['出品制限フラグ'],
  DROPS_30: ['下落回数(30日)', '下落回数'],
  SHIPPING: ['納品送料概算', '納品送料'],
  RESEARCH_DT: ['リサーチ日時', '取得日時'],
  PURCHASED: ['仕入済'],
  SHIP_METHOD: ['発送方法'],
  TRANSFERRED: ['転送済フラグ'],
  SKU: ['SKU'],
  WANTED_FLAG: ['手配書フラグ'],
  QTY: ['仕入個数'],
  SHOP: ['店舗', '仕入店舗', 'ショップ名', '店名']
};

/**
 * シートの1行目を読み取り、キー名と列番号(1-based)のペアを返す
 */
function getColMap(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return { colMap: {}, lastCol: 0, headers: [] };
  
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const colMap = {};
  
  headers.forEach((header, index) => {
    if (!header) return;
    const hStr = String(header).trim();
    for (const [key, aliases] of Object.entries(COL_ALIASES)) {
      // 既にマッチ済みならスキップしない（重複がある場合は最初に見つかったものを優先したいが、後の列で上書きされる。とりあえずそのまま）
      if (!colMap[key] && aliases.includes(hStr)) {
        colMap[key] = index + 1; // 1-based index
      }
    }
  });
  
  return { colMap, lastCol, headers };
}

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
    const { colMap } = getColMap(sheet);
    if (!colMap.JAN) return; // JAN列がないシートはスキップ
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    
    const values = sheet.getRange(2, colMap.JAN, lastRow - 1, 1).getValues();
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
  Utilities.sleep(3000);
  autoResearch(null); // 以降でアクティブシートを判別
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (!TARGET_SHEET_NAMES.includes(sheet.getName())) return;

  const row = e.range.getRow();
  if (row < 2) return; // ヘッダーは除外

  const { colMap } = getColMap(sheet);
  if (!colMap.JAN) return; // JAN列がないシートは無視

  const colPurchased = colMap.PURCHASED;
  const colTransferred = colMap.TRANSFERRED;
  const editedCol = e.range.getColumn();

  if (colPurchased && editedCol === colPurchased) {
    const isPurchased = sheet.getRange(row, colPurchased).getValue() === true;
    const isTransferred = colTransferred ? sheet.getRange(row, colTransferred).getValue() : false;

    if (isPurchased && !isTransferred) {
      transferToPurchaseData(sheet, row, colMap);
      if (colTransferred) sheet.getRange(row, colTransferred).setValue(true);
    }
  } else if (editedCol === colMap.PURCHASE || editedCol === colMap.QTY || editedCol === colMap.BUYBOX || editedCol === colMap.FBA_FEE) {
    // 💡 価格や個数が変更された場合の処理
    recalculateRow(sheet, row, colMap);
    
    // 🛡️ もし商品名が空（＝過去に取得失敗している）なら、ついでに再リサーチを試みる
    const title = colMap.TITLE ? sheet.getRange(row, colMap.TITLE).getValue() : "exists";
    if (!title || title === "" || String(title).startsWith("[Error]")) {
      console.log(`🔎 行:${row} の商品名が空のため、価格編集をトリガーに再取得を試みます...`);
      autoResearch(e);
    }
  } else {
    autoResearch(e);
  }
}

/**
 * 指定した行の利益・ROI・判定を動的カラムマップに基づいて再計算する（Keepa通信なし）
 */
function recalculateRow(sheet, row, colMap) {
  if (!colMap) colMap = getColMap(sheet).colMap;
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;

  const range = sheet.getRange(row, 1, 1, lastCol);
  const values = range.getValues()[0];

  // 🧮 数値のクリーンアップ（¥やカンマ、スペースを除去）
  const parseNum = (val) => {
    if (typeof val === 'number') return val;
    if (!val || val === "圏外" || val === "データなし") return 0;
    const clean = String(val).replace(/[¥, 位]/g, '').trim();
    if (clean === "" || isNaN(Number(clean))) return 0;
    return Number(clean);
  };

  const getNum = (colKey) => colMap[colKey] ? parseNum(values[colMap[colKey] - 1]) : 0;
  const getVal = (colKey) => colMap[colKey] ? values[colMap[colKey] - 1] : "";

  const buyBox = getNum('BUYBOX');
  const fbaFee = getNum('FBA_FEE');
  const refRate = getNum('REF_RATE');
  const purchasePrice = getNum('PURCHASE');
  
  let qty = getNum('QTY');
  if (qty === 0 && colMap.QTY) {
      qty = 1;
  } else if (!colMap.QTY) {
      qty = 1;
  }

  const fbaLowest = getNum('FBA_LOWEST');
  const monthlySold = getNum('MONTHLY_SOLD');
  const sellerCount = getNum('SELLER_COUNT');

  if (buyBox > 0) {
    const rate = refRate > 0 ? (refRate / 100) : 0.15; // 紹介料率のデフォルト15%
    const breakEven = Math.floor(buyBox - fbaFee - (buyBox * rate));
    const profitPerUnit = breakEven - purchasePrice;
    const totalProfit = profitPerUnit * qty;
    const profitRateValue = profitPerUnit / buyBox; // 利益率（少数点）
    const roi = purchasePrice > 0 ? Math.round((profitPerUnit / purchasePrice) * 100) : 0;

    // プレ値判定の更新
    let premiumLabel = "ー";
    const listPrice = getNum('LIST_PRICE');
    if (listPrice > 0 && buyBox > 0) {
      const ratio = ((buyBox - listPrice) / listPrice) * 100;
      if (ratio >= 30) premiumLabel = `🔥 +${Math.round(ratio)}%`;
      else if (ratio >= 20) premiumLabel = `🔸 +${Math.round(ratio)}%`;
      else if (ratio < -10) premiumLabel = `📉 ${Math.round(ratio)}%`;
    }

    // 判定ロジック（なぎさんのAppSheet数式と100%同期）
    let judgment = "✖️ 見送り";

    if (buyBox < 2000 && buyBox > 0 && profitPerUnit >= 100 && profitRateValue >= 0.10 && monthlySold >= 10) {
      judgment = "○ 仕入れ対象";
    } else if (profitPerUnit >= 500 && monthlySold >= 30 && sellerCount <= 5) {
      judgment = "◎ 超即買い";
    } else if (profitPerUnit >= 300 && monthlySold >= 10) {
      judgment = "○ 仕入れ対象";
    } else if (profitPerUnit >= 1000 && monthlySold <= 3) {
      judgment = "△ 要確認";
    }
    
    // 手配書マッチがある場合は優先的に表示
    const wantedFlag = getVal('WANTED_FLAG');
    if (wantedFlag && wantedFlag !== "") judgment = `🚓 ${judgment}`;

    // 反映
    if (colMap.BREAK_EVEN) sheet.getRange(row, colMap.BREAK_EVEN).setValue(breakEven);
    if (colMap.PROFIT) sheet.getRange(row, colMap.PROFIT).setValue(totalProfit);
    if (colMap.PROFIT_RATE) sheet.getRange(row, colMap.PROFIT_RATE).setValue(Math.round(profitRateValue * 100) + "%");
    if (colMap.ROI) sheet.getRange(row, colMap.ROI).setValue(roi + "%");
    if (colMap.PREMIUM_JUDGE) sheet.getRange(row, colMap.PREMIUM_JUDGE).setValue(premiumLabel);
    if (colMap.JUDGMENT) sheet.getRange(row, colMap.JUDGMENT).setValue(judgment);

    // SKUも更新（仕入価格や損益分岐点が変わるため）
    const asin = getVal('ASIN');
    if (asin && colMap.SKU) {
      const dateIdPart = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd");
      const sku = `${purchasePrice}_${dateIdPart}_${breakEven}_${asin}`;
      sheet.getRange(row, colMap.SKU).setValue(sku);
    }
    
    console.log(`🔄 再計算完了 (行:${row}): ${judgment} / 利益 ${totalProfit}円`);
  }
}

function testRun() {
  autoResearch(null);
}

// ==========================================
// メイン処理 (AppSheet対応版)
// ==========================================

function autoResearch(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet;
  let targetRow = null;

  if (e && e.range) {
    sheet = e.range.getSheet();
    targetRow = e.range.getRow();
  } else {
    sheet = ss.getActiveSheet();
  }

  if (!TARGET_SHEET_NAMES.includes(sheet.getName())) {
    sheet = ss.getSheetByName("リサーチリスト") || sheet;
    if (!TARGET_SHEET_NAMES.includes(sheet.getName())) return;
  }

  const { colMap, lastCol } = getColMap(sheet);
  if (!colMap.JAN) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // 🚀 スキャン範囲の決定 (全件スキャン or ターゲット行のみ)
  // eが存在する場合は基本ターゲット行のみ、eがない場合は未取得全件
  const startIdx = (targetRow && targetRow > 1) ? targetRow - 1 : 1;
  const endIdx = (targetRow && targetRow > 1) ? targetRow - 1 : lastRow - 1;

  const fullRange = sheet.getRange(1, 1, lastRow, lastCol);
  const allValues = fullRange.getValues();

  // 1回だけ手配書リストを取得
  const wantedMap = getWantedJanMap();

  // 2. データの取得と再計算
  for (let i = startIdx; i <= endIdx; i++) {
    const currentRow = i + 1;
    const input = String(allValues[i][colMap.JAN - 1]).trim();
    const title = colMap.TITLE ? String(allValues[i][colMap.TITLE - 1]) : "";
    const sizeWeight = colMap.SIZE_WEIGHT ? String(allValues[i][colMap.SIZE_WEIGHT - 1]) : "";

    if (!input || input === "") continue;

    const isBroken = title === "" || title.startsWith("[Error]") || title === "見つかりませんでした" || sizeWeight.includes("nan");

    if (isBroken) {
      // 🔎 新規または欠損データの取得
      const canContinue = fetchProductData(input, currentRow, sheet, ss, wantedMap, colMap);
      if (canContinue === false) break;
      Utilities.sleep(500); 
    } else {
      // 🔄 既存行の再計算
      recalculateRow(sheet, currentRow, colMap);
    }
  }
}

// ==========================================
// データ取得・書き込みロジック
// ==========================================

function fetchProductData(barcode, row, sheet, ss, wantedMap, colMap) {
  try {
    const keepaApiKey = getRequiredApiToken('KEEPA_API_KEY');
    let product = getFromKvCache(barcode);
    let usedCache = !!product;

    if (!usedCache) {
      const inputType = detectInputType(barcode);
      const paramKey = (inputType === 'asin') ? 'asin' : 'code';
      let url = `https://api.keepa.com/product?key=${keepaApiKey}&domain=5&type=product&${paramKey}=${barcode}&stats=1`;
      let response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      let json = JSON.parse(response.getContentText());

      // 🛒 JANで失敗した場合にASINでリトライするロジックを追加
      if (paramKey === 'code' && (json.error || !json.products || json.products.length === 0)) {
        const existingAsin = colMap.ASIN ? sheet.getRange(row, colMap.ASIN).getValue() : '';
        if (existingAsin && String(existingAsin).trim() !== '') {
          console.log(`🔎 JANでの取得に失敗したため、ASIN(${existingAsin})でリトライします...`);
          url = `https://api.keepa.com/product?key=${keepaApiKey}&domain=5&type=product&asin=${existingAsin}&stats=1`;
          response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
          json = JSON.parse(response.getContentText());
        }
      }

      // 🔋 トークン残量チェック（20以下で緊急通知、0で処理停止）
      const tokensLeft = json.tokensLeft;
      const refillIn = json.refillIn || 0;
      if (tokensLeft !== undefined && tokensLeft <= 20) {
        const refillMinutes = Math.ceil(refillIn / 60000);
        const warnMsg = `⚠️ Keepaトークン残量警告 ⚠️\n現在の残りトークンが「${tokensLeft}」になりました。\n約${refillMinutes}分後に回復予定です。`;
        sendLineNotification(warnMsg);
        console.warn(`🔋 トークン残量警告: 残り${tokensLeft}トークン / 約${refillMinutes}分後に回復`);
        if (tokensLeft <= 0) {
          console.error("🚫 KEEPトークンが枯渇したため、処理を中断します。");
          return false;
        }
      }

      // 🛡️ 既存データ保護: ASINが入っている行はAPIエラーでも絶対上書きしない
      const existingAsinFinal = colMap.ASIN ? sheet.getRange(row, colMap.ASIN).getValue() : '';
      const hasExistingData = existingAsinFinal && String(existingAsinFinal).trim() !== '';

      if (json.error || !json.products || json.products.length === 0) {
        if (hasExistingData) {
          console.warn(`🛡️ データ保護: 行${row}はASIN取得済みのため、エラーによる上書きをスキップしました。`);
          return true;
        }
        if (json.error) setError(sheet, row, colMap, `API Error: ${json.error.message}`);
        else setNotFound(sheet, row, colMap);
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
    const rawRank = current[3];
    let rank = (rawRank === -1 || rawRank === undefined || rawRank === null) ? "圏外" : rawRank;
    let newPrice = (current[1] > 0 ? current[1] : (current[0] > 0 ? current[0] : ""));
    
    // カート価格 (buyBoxPrice や current[18] が -1 や -2 の場合は除外する)
    let buyBoxCandidate = current[18] > 0 ? current[18] : (stats.buyBoxPrice > 0 ? stats.buyBoxPrice : "");
    let buyBox = buyBoxCandidate || newPrice; // カート価格がない場合は新品価格を採用

    let fbaLowest = current[7] > 0 ? current[7] : "";
    let sellerCount = current[11] >= 0 ? current[11] : "";
    const isVariation = (product.variationCSV && product.variationCSV.length > 0) ? "有" : "無";
    const isAmazonSelling = (current[0] && current[0] > 0) ? "有" : "無";
    let monthlySold = product.monthlySold !== undefined ? product.monthlySold : stats.salesRankDrops30 || "";
    let drops30 = stats.salesRankDrops30 >= 0 ? stats.salesRankDrops30 : "";
    let fbaFee = (product.fbaFees && product.fbaFees.pickAndPackFee) ? product.fbaFees.pickAndPackFee : "";
    
    // 定価の取得
    let listPrice = (stats.listPrice > 0) ? stats.listPrice : (stats.suggestedPrice > 0 ? stats.suggestedPrice : 0);
    
    // 紹介料率の取得（小数点以下も保持）
    let refRate = product.referralFeePercent || "";
    if (refRate === "" && product.fbaFees && product.fbaFees.referralFeePercent !== undefined) {
        refRate = product.fbaFees.referralFeePercent;
    }

    let breakEven = "";
    if (buyBox && buyBox > 0 && fbaFee) {
      const rate = refRate !== "" ? (parseFloat(refRate) / 100) : 0.15; // デフォルト15%
      breakEven = Math.floor(buyBox - fbaFee - (buyBox * rate));
    }

    // プレ値判定
    let premiumJudge = "ー";
    if (listPrice > 0 && buyBox > 0) {
      const ratio = ((buyBox - listPrice) / listPrice) * 100;
      if (ratio >= 30) premiumJudge = `🔥 +${Math.round(ratio)}%`;
      else if (ratio >= 20) premiumJudge = `🔸 +${Math.round(ratio)}%`;
      else if (ratio < -10) premiumJudge = `📉 ${Math.round(ratio)}%`;
    }

    // サイズ・重量のフォーマット（cm、g表記。Keepaエクスポートに合わせる）
    let sizeWeight = "";
    const pL = product.packageLength || product.itemLength || 0;
    const pW = product.packageWidth || product.itemWidth || 0;
    const pH = product.packageHeight || product.itemHeight || 0;
    const pWeight = product.packageWeight || product.itemWeight || 0;

    if (pL > 0) {
      // APIの mm を cm に変換
      const cmL = (pL / 10).toFixed(1);
      const cmW = (pW / 10).toFixed(1);
      const cmH = (pH / 10).toFixed(1);
      sizeWeight = `${cmL}x${cmW}x${cmH}cm ${pWeight}g`;
    } else if (pWeight > 0) {
      sizeWeight = `${pWeight}g`;
    }

    let isHazmat = product.hazardousMaterialType ? "Yes" : "No";
    if (title.toLowerCase().includes("battery") || title.includes("電池")) isHazmat = "Yes";

    const amazonLink = `https://www.amazon.co.jp/dp/${asin}`;
    const keepaLink = `https://keepa.com/#!product/5-${asin}`;
    const poiLink = `http://localhost:3000/?q=${asin}`;
    const now = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");

    // 現在のシートと同じシート由来のマッチは自己検知なので除外
    const matches = (wantedMap.get(String(barcode).trim()) || []).filter(m => m.sheetName !== sheet.getName());
    const judgmentFlag = getJudgmentFlagString(matches);
    
    // 現在の行データを取得し、必要な列だけを安全に上書き
    const currentValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    const outRow = [...currentValues]; 

    const setCol = (colKey, val) => {
        if (colMap[colKey]) outRow[colMap[colKey] - 1] = val;
    };

    setCol('TITLE', title);
    setCol('ASIN', asin);
    setCol('BRAND', brand);
    setCol('RANK', rank);
    setCol('MONTHLY_SOLD', monthlySold);
    setCol('CATEGORY', categoryName);
    setCol('SELLER_COUNT', sellerCount);
    setCol('VARIATION', isVariation);
    setCol('BUYBOX', buyBox);
    setCol('NEW_PRICE', newPrice);
    setCol('FBA_LOWEST', fbaLowest);
    setCol('LIST_PRICE', listPrice);
    setCol('PREMIUM_JUDGE', premiumJudge);
    setCol('BREAK_EVEN', breakEven);
    setCol('FBA_FEE', fbaFee);
    setCol('REF_RATE', refRate);
    setCol('SIZE_WEIGHT', sizeWeight);
    setCol('HAZMAT', isHazmat);
    setCol('LINK_AMAZON', amazonLink);
    setCol('LINK_KEEPA', keepaLink);
    setCol('LINK_POI', poiLink);
    setCol('IMAGE', imageUrl);
    setCol('AMAZON_SELL', isAmazonSelling);
    setCol('DROPS_30', drops30);
    setCol('RESEARCH_DT', now);
    setCol('WANTED_FLAG', judgmentFlag);
    
    // 店舗名（既存の値があれば保持、なければ空。ポイポイからの書き込み等で拡張可能）
    const existingShop = colMap.SHOP ? (currentValues[colMap.SHOP - 1] || "") : "";
    setCol('SHOP', existingShop);

    const existingPurchasePrice = colMap.PURCHASE ? (currentValues[colMap.PURCHASE - 1] || "") : "";

    sheet.getRange(row, 1, 1, outRow.length).setValues([outRow]);
    
    // 💡 最後に利益・ROI・判定・SKUを一括計算して上書き
    recalculateRow(sheet, row, colMap);
    SpreadsheetApp.flush(); // 即時反映して競合を防ぐ

    if (matches.length > 0 && ss) markWantedListAsArrested(ss, matches);

    // 🔔 LINE リッチ通知（利益100円以上 or 手配書マッチ時）
    const MIN_PROFIT_FOR_ALERT = 100;
    const numBreakEven = Number(breakEven) || 0;
    const numPurchase = Number(existingPurchasePrice) || 0;
    const profit = numPurchase > 0 ? (numBreakEven - numPurchase) : numBreakEven;
    if (profit >= MIN_PROFIT_FOR_ALERT || matches.length > 0) {
      try {
        sendLineFlexNotification({
          title, imageUrl, asin, buyBox, rank, breakEven,
          sellerCount, monthlySold, isVariation, isHazmat,
          isAmazonSelling, amazonLink, keepaLink, judgmentFlag, profit,
          purchasePrice: existingPurchasePrice
        });
        console.log(`📱 LINE通知送信: ${title} (利益: ¥${profit})`);
      } catch (lineErr) {
        console.warn(`LINE通知エラー（処理は継続）: ${lineErr}`);
      }
    }

    console.log(`🌸 API リサーチ完了: ${title} -> 行: ${row} に書き込みました`);
    return true;
  } catch (e) {
    // 🛡️ 既存データ保護: ASINが入っている行は例外エラーでも絶対上書きしない
    const existingAsinOnCatch = colMap.ASIN ? sheet.getRange(row, colMap.ASIN).getValue() : '';
    if (!existingAsinOnCatch || String(existingAsinOnCatch).trim() === '') {
      setError(sheet, row, colMap, e.toString());
    } else {
      console.warn(`🛡️ データ保護(catch): 行${row}はASIN取得済みのため、例外エラーによる上書きをスキップしました。`);
    }
    return true;
  }
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

// ==========================================
// LINE Flex Message リッチ通知 🔔
// ==========================================

function sendLineFlexNotification(data) {
  const props = PropertiesService.getScriptProperties();
  const accessToken = props.getProperty('MERUPOI_LINE_ACCESS_TOKEN') || '';
  const userId = props.getProperty('MERUPOI_LINE_USER_ID') || '';
  if (!accessToken || !userId) return;

  const {
    title, imageUrl, asin, buyBox, rank, breakEven,
    sellerCount, monthlySold, isVariation, isHazmat,
    isAmazonSelling, amazonLink, keepaLink, judgmentFlag, profit,
    purchasePrice
  } = data;

  // メルカリ検索用: 商品名の先頭40文字をURLエンコード
  const mercariQuery = encodeURIComponent((title || "").substring(0, 40));
  const mercariLink = `https://jp.mercari.com/search?keyword=${mercariQuery}`;

  // ヘッダーの色（手配書マッチ=赤、通常=緑）
  const headerColor = judgmentFlag ? "#DC3545" : "#28A745";
  const headerText = judgmentFlag ? `🚨 手配書マッチ！ ${judgmentFlag}` : "🔥 利益商品を発見！";

  // 危険物・Amazon本体の警告色
  const hazmatColor = isHazmat === "Yes" ? "#DC3545" : "#888888";
  const amazonSellColor = isAmazonSelling === "有" ? "#DC3545" : "#28A745";

  const flexMessage = {
    "type": "flex",
    "altText": `${headerText} ${(title || "").substring(0, 20)}... ¥${profit}`,
    "contents": {
      "type": "bubble",
      "size": "giga",
      "header": {
        "type": "box",
        "layout": "vertical",
        "contents": [{
          "type": "text",
          "text": headerText,
          "color": "#FFFFFF",
          "weight": "bold",
          "size": "md"
        }],
        "backgroundColor": headerColor,
        "paddingAll": "12px"
      },
      "hero": imageUrl ? {
        "type": "image",
        "url": imageUrl,
        "size": "full",
        "aspectRatio": "1:1",
        "aspectMode": "fit",
        "backgroundColor": "#FFFFFF"
      } : undefined,
      "body": {
        "type": "box",
        "layout": "vertical",
        "spacing": "md",
        "contents": [
          {
            "type": "text",
            "text": `📦 ${title || "Unknown"}`,
            "weight": "bold",
            "size": "sm",
            "wrap": true,
            "maxLines": 2
          },
          {
            "type": "separator"
          },
          {
            "type": "box",
            "layout": "vertical",
            "spacing": "sm",
            "contents": [
              makeInfoRow("💰 カート価格", `¥${Number(buyBox || 0).toLocaleString()}`),
              makeInfoRow("📊 ランキング", `${Number(rank || 0).toLocaleString()}位`),
              makeInfoRow("🏷️ 仕入上限", `¥${Number(breakEven || 0).toLocaleString()}`),
              makeInfoRow("🛒 仕入価格", purchasePrice ? `¥${Number(purchasePrice).toLocaleString()}` : "-"),
              makeInfoRow("💵 見込利益", `¥${Number(profit || 0).toLocaleString()}`),
              makeInfoRow("👥 セラー数", `${sellerCount || "-"}人`),
              makeInfoRow("📈 月間販売", `${monthlySold || "-"}個`),
              makeInfoRow("🔀 バリエ", isVariation || "-"),
              makeInfoRowColored("⚠️ 危険物", isHazmat || "No", hazmatColor),
              makeInfoRowColored("🏢 Amazon様", isAmazonSelling || "無", amazonSellColor)
            ]
          }
        ]
      },
      "footer": {
        "type": "box",
        "layout": "horizontal",
        "spacing": "sm",
        "contents": [
          {
            "type": "button",
            "action": { "type": "uri", "label": "📦Amazon", "uri": amazonLink },
            "style": "primary",
            "color": "#FF9900",
            "height": "sm"
          },
          {
            "type": "button",
            "action": { "type": "uri", "label": "📈Keepa", "uri": keepaLink },
            "style": "primary",
            "color": "#1E88E5",
            "height": "sm"
          },
          {
            "type": "button",
            "action": { "type": "uri", "label": "🔴メルカリ", "uri": mercariLink },
            "style": "primary",
            "color": "#E53935",
            "height": "sm"
          }
        ]
      }
    }
  };

  // hero が undefined の場合は削除（画像なし対応）
  if (!flexMessage.contents.hero) delete flexMessage.contents.hero;

  const payload = {
    "to": userId,
    "messages": [flexMessage]
  };

  const options = {
    "method": "post",
    "headers": {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + accessToken
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', options);
    if (res.getResponseCode() !== 200) {
      console.warn(`LINE Flex送信エラー: ${res.getContentText()}`);
    }
  } catch(e) {
    console.warn(`LINE Flex通信エラー: ${e}`);
  }
}

// Flex Message 用ヘルパー関数
function makeInfoRow(label, value) {
  return {
    "type": "box",
    "layout": "horizontal",
    "contents": [
      { "type": "text", "text": label, "size": "xs", "color": "#888888", "flex": 4 },
      { "type": "text", "text": String(value), "size": "xs", "weight": "bold", "align": "end", "flex": 3 }
    ]
  };
}

function makeInfoRowColored(label, value, color) {
  return {
    "type": "box",
    "layout": "horizontal",
    "contents": [
      { "type": "text", "text": label, "size": "xs", "color": "#888888", "flex": 4 },
      { "type": "text", "text": String(value), "size": "xs", "weight": "bold", "color": color, "align": "end", "flex": 3 }
    ]
  };
}

function setNotFound(sheet, row, colMap) { if (colMap.TITLE) sheet.getRange(row, colMap.TITLE).setValue("見つかりませんでした"); }
function setError(sheet, row, colMap, message) { if (colMap.TITLE) sheet.getRange(row, colMap.TITLE).setValue(`[Error] ${message}`); }

// ==========================================
// 外部連携 (zaiko_tool / doPost)
// ==========================================
const ZAIKO_TOOL_SPREADSHEET_ID = '1EIYt3IP7FidK-RbmNj2MYNamIbVvbt_VMVs89FMYkKY';
const ZAIKO_SHEET_NAME = '仕入れデータ';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📱 AppSheet連携')
    .addItem('📥 AppSheetの「仕入済」をzaiko_toolへ送信', 'syncAppSheetPurchases')
    .addSeparator()
    .addItem('🔄 未取得データを一括リサーチ (取得漏れの修復)', 'retryMissingResearch')
    .addToUi();
}

function syncAppSheetPurchases() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet(); // 汎用的にするためにアクティブシートを使う
  if (!TARGET_SHEET_NAMES.includes(sheet.getName())) {
    SpreadsheetApp.getUi().alert('エラー', '対象外のシートです。リサーチリストなどを開いた状態で実行してください。', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const { colMap, lastCol } = getColMap(sheet);
  const colPurchased = colMap.PURCHASED;
  const colTransferred = colMap.TRANSFERRED;
  
  if (!colPurchased || !colTransferred) {
    SpreadsheetApp.getUi().alert('エラー', '「仕入済」または「転送済フラグ」の列が見つかりません。', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  let count = 0;
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][colPurchased - 1] === true && data[i][colTransferred - 1] !== true) {
      transferToPurchaseData(sheet, i + 2, colMap);
      sheet.getRange(i + 2, colTransferred).setValue(true);
      count++;
    }
  }
  
  if (count > 0) SpreadsheetApp.getUi().alert('転送完了', `${count}件送信しました。`, SpreadsheetApp.getUi().ButtonSet.OK);
}

function transferToPurchaseData(sourceSheet, row, colMap) {
  try {
    const dataRow = sourceSheet.getRange(row, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
    const appId = dataRow[0] || `ID-${new Date().getTime()}`;
    
    // 動的マップから値を取り出すヘルパー
    const getVal = (key, defaultVal = "") => colMap[key] ? (dataRow[colMap[key] - 1] || defaultVal) : defaultVal;
    const getNum = (key, defaultVal = 0) => colMap[key] ? (Number(dataRow[colMap[key] - 1]) || defaultVal) : defaultVal;

    const researchDate = getVal('RESEARCH_DT', Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm"));
    const janCode = getVal('JAN');
    const itemName = getVal('TITLE');
    const unitPrice = getNum('PURCHASE', 0);
    
    let defaultQty = getNum('QTY', 1);
    if (defaultQty <= 0) defaultQty = 1;

    const totalPrice = unitPrice * defaultQty;
    const shopName = getVal('SHOP', '');
    const condition = "";
    const imageUrl = getVal('IMAGE');
    const receiptImage = "";

    // ★ 新規追加: ASIN, SKU, 損益分岐点をリサーチリストから取得
    const asin = getVal('ASIN');
    const sku = getVal('SKU');
    const breakEven = getNum('BREAK_EVEN', 0);

    const zaikoSs = SpreadsheetApp.openById(ZAIKO_TOOL_SPREADSHEET_ID);
    const targetSheet = zaikoSs.getSheetByName(ZAIKO_SHEET_NAME);
    
    // ★ 仕入れデータシートへの書き込み (14列: A-N)
    // 既存11列 + L:ASIN, M:SKU, N:損益分岐点
    if (targetSheet) {
      targetSheet.appendRow([
        appId, researchDate, janCode, itemName, unitPrice, defaultQty, totalPrice, shopName, condition, imageUrl, receiptImage,
        asin, sku, breakEven
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

    // ★ EC注文履歴への書き込み (12列: A-L)
    // 既存10列 + K:ASIN, L:SKU
    const ecHistorySheet = zaikoSs.getSheetByName('EC注文履歴');
    if (ecHistorySheet) {
      const lastRow = ecHistorySheet.getLastRow();
      let isDuplicateEC = false;
      if (lastRow > 1) {
        const ids = ecHistorySheet.getRange(2, 9, lastRow - 1, 1).getValues().flat();
        if (ids.includes(appId)) isDuplicateEC = true;
      }
      if (!isDuplicateEC) {
        ecHistorySheet.appendRow([
          dateStr, shopName, itemName, unitPrice, defaultQty, totalAmountStr, appId, now, appId, imageUrl,
          asin, sku
        ]);
      }
    }

    // ★ 在庫管理マスタへの書き込み (18列: A-R)
    // カラムレイアウト：
    // A:管理ID, B:商品名, C:仕入日, D:仕入先, E:仕入額(単価), F:ステータス,
    // G:出品日, H:回転日数, I:販売日, J:売上額, K:手数料, L:送料, M:純利益,
    // N:元注文番号, O:商品画像, P:SKU, Q:ASIN, R:損益分岐点
    const invMasterSheet = zaikoSs.getSheetByName('在庫管理マスタ');
    if (invMasterSheet) {
      const lastRow = invMasterSheet.getLastRow();
      let isDuplicateInv = false;
      if (lastRow > 1) {
        // ★ 修正: N列(14列目)で重複チェック（旧版は12列目を見ていたバグを修正）
        // 後方互換性のため、L列(12)とN列(14)の両方をチェック
        const idsN = invMasterSheet.getRange(2, 14, lastRow - 1, 1).getValues().flat();
        const idsL = invMasterSheet.getRange(2, 12, lastRow - 1, 1).getValues().flat();
        if (idsN.includes(appId) || idsL.includes(appId)) isDuplicateInv = true;
      }
      if (!isDuplicateInv) {
        const newRows = [];
        const cleanDateStr = isValidDate ? Utilities.formatDate(dateObj, 'Asia/Tokyo', 'yyyy-MM-dd') : Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
        for (let i = 1; i <= defaultQty; i++) {
          const invId = `INV-${dateIdPart}-${shortId}-${i}`;
          // ★ 18列フォーマット: SalesData.jsと統一
          newRows.push([
            invId, itemName, cleanDateStr, shopName, unitPrice,
            "出品待ち", "", "", "", "", "", "", "",
            appId, imageUrl, sku, asin, breakEven
          ]);
        }
        if (newRows.length > 0) {
          invMasterSheet.getRange(invMasterSheet.getLastRow() + 1, 1, newRows.length, 18).setValues(newRows);
        }
      }
    }
  } catch(e) { console.error("転送エラー: " + e.toString()); }
}

function doPost(e) {
  try {
    const json = JSON.parse(e.postData.contents);
    const postApiKey = getRequiredApiToken('POST_API_KEY');
    if (json.api_key !== postApiKey) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Invalid API Key" })).setMimeType(ContentService.MimeType.JSON);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = json.sheet_name || "リサーチリスト"; // デフォルトをリサーチリストに
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

    // 更新対象のシートであれば自動リサーチをトリガー
    if (TARGET_SHEET_NAMES.includes(sheetName)) {
      if (json.skip_research === true || json.skip_research === "true") {
        console.log("python側からのフラグにより自動リサーチをスキップしました");
      } else if (Array.isArray(rows) && rows.length > 5) {
        console.warn("大量データ追加のため、同期的な自動リサーチをスキップします。");
      } else {
        try {
          // eオブジェクトのモックを作成し、該当のシートでautoResearchが実行されるようにする
          autoResearch({ range: sheet.getRange(1, 1) });
        } catch (researchErr) {
          console.warn("doPost後の自動リサーチでエラー: " + researchErr.toString());
        }
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ status: "success", message: `${rows ? rows.length : 0} rows added` })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * JANはあるがタイトル等の情報が取得できていない行を一括リサーチする (メニュー用)
 */
function retryMissingResearch() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  if (!TARGET_SHEET_NAMES.includes(sheet.getName())) {
    SpreadsheetApp.getUi().alert('対象外のシートです。');
    return;
  }
  
  const { colMap } = getColMap(sheet);
  if (!colMap.JAN || !colMap.TITLE) return;
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  const allValues = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
  const wantedMap = getWantedJanMap(); // 1回だけ取得

  let count = 0;
  for (let i = 1; i < allValues.length; i++) {
    const jan = String(allValues[i][colMap.JAN - 1]).trim();
    const title = String(allValues[i][colMap.TITLE - 1]).trim();
    const isBroken = title === "" || title.startsWith("[Error") || title === "見つかりませんでした";

    if (jan !== "" && isBroken) {
      count++;
      const canContinue = fetchProductData(jan, i + 1, sheet, ss, wantedMap, colMap);
      if (canContinue === false) {
        SpreadsheetApp.getUi().alert('APIトークンが不足したため中断しました。');
        break;
      }
      Utilities.sleep(500); 
    }
  }
  
  if (count > 0) {
    SpreadsheetApp.getUi().alert(`✅ 完了: ${count}件のデータを取得・修復しました。`);
  } else {
    SpreadsheetApp.getUi().alert('💡 すべてのデータは取得済みです。');
  }
}
