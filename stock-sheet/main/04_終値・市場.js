const MARKET_VALUE_UNIT_DIVISOR = 1000; // PDF上は百万円単位。これで割って10億円単位(百万円の桁をさらに3桁省略)にする

/**
 * 商況プリント２(後場)PDFのテキストから、取引代金概算(立会内)のプライム/スタンダード/グロースを取り出す。
 * 単位は百万円のまま(呼び出し側のfetchMarketTradingValuesForDateで10億円単位に変換する)。
 */
function parseMarketTradingValuesFromPdfText(text) {
  const idx = text.indexOf('取引代金概算');
  if (idx === -1) return null;
  const area = text.substring(idx, idx + 1500);

  function extract(label) {
    const re = new RegExp(label + '\\s*([\\d,]+)\\s*百万円');
    const m = area.match(re);
    return m ? parseInt(m[1].replace(/,/g, ''), 10) : null;
  }

  const prime = extract('内国株式・プライム市場');
  const standard = extract('内国株式・スタンダード市場');
  const growth = extract('内国株式・グロース市場');
  if (prime === null || standard === null || growth === null) return null;
  return { '1001': prime, '1002': standard, '1003': growth };
}

/**
 * 商況プリント２(後場)PDFのテキストから、取引高概算(立会内)のプライム/スタンダード/グロースを百万株単位で取り出す。
 * PDF上の単位は市場ごとに「万株」「株」「百株」と異なるため、単位を見て百万株に換算する。
 */
function parseMarketVolumesFromPdfText(text) {
  const idx = text.search(/取引高概算|売買高概算/);
  if (idx === -1) return null;
  const area = text.substring(idx, idx + 1500);
  const unitToShares = { '万株': 10000, '百株': 100, '株': 1 };

  function extract(label) {
    const re = new RegExp(label + '\\s*([\\d,]+)\\s*(万株|百株|株)');
    const m = area.match(re);
    if (!m) return null;
    const shares = parseInt(m[1].replace(/,/g, ''), 10) * unitToShares[m[2]];
    return Math.round(shares / 1000000);
  }

  return {
    '1001': extract('内国株式・プライム市場'),
    '1002': extract('内国株式・スタンダード市場'),
    '1003': extract('内国株式・グロース市場')
  };
}

/**
 * 「商況プリント２(後場)」PDFのURLを組み立てる(日付部分のみ変化し、添付ID部分は固定という前提)。
 * 直接アクセスが失敗した場合は、一覧ページ(volume-and-value/index.html)から該当日のリンクを探すフォールバックを行う。
 */
function marketValuePdfUrl_(dateStr) {
  const ymd = dateStr.replace(/-/g, '');
  return 'https://www.jpx.co.jp/markets/equities/volume-and-value/tvdivq000000derc-att/2_' + ymd + '.pdf';
}

function findMarketValuePdfUrlFromIndex_(dateStr) {
  const ymd = dateStr.replace(/-/g, '');
  const response = UrlFetchApp.fetch('https://www.jpx.co.jp/markets/equities/volume-and-value/index.html', { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) return null;
  const html = response.getContentText();
  const re = new RegExp('href="((?:https://www\\.jpx\\.co\\.jp)?/markets/equities/volume-and-value/[^"]+2_' + ymd + '\\.pdf)"');
  const m = html.match(re);
  if (!m) return null;
  return m[1].startsWith('http') ? m[1] : 'https://www.jpx.co.jp' + m[1];
}

/**
 * 指定日のプライム/スタンダード/グロースの売買代金(10億円単位)と出来高(百万株)を
 * {コード: {value, volume}} で返す。
 * 休場日・未公開日はnullを返す(呼び出し側でスキップ・失敗扱いにする)。
 */
function fetchMarketTradingValuesForDate(dateStr) {
  try {
    let text = fetchPdfTextFromUrl_(marketValuePdfUrl_(dateStr));
    if (!text) {
      const fallbackUrl = findMarketValuePdfUrlFromIndex_(dateStr);
      if (fallbackUrl) text = fetchPdfTextFromUrl_(fallbackUrl);
    }
    if (!text) return null;
    const millionYenValues = parseMarketTradingValuesFromPdfText(text);
    if (!millionYenValues) return null;
    const volumes = parseMarketVolumesFromPdfText(text) || {};

    const result = {};
    Object.keys(millionYenValues).forEach(code => {
      const v = millionYenValues[code];
      result[code] = {
        value: (v === null || v === undefined) ? null : Math.round(v / MARKET_VALUE_UNIT_DIVISOR),
        volume: (volumes[code] === undefined) ? null : volumes[code]
      };
    });
    return result;
  } catch (e) {
    Logger.log('市場別売買代金PDFの取得に失敗(' + dateStr + '): ' + e);
    return null;
  }
}

// ==== 株価取得 ====
function parseChartClose(text) {
  try {
    const json = JSON.parse(text);
    const result = json.chart && json.chart.result && json.chart.result[0];
    if (!result) return null;

    const meta = result.meta;
    const price = (meta && meta.regularMarketPrice !== undefined) ? meta.regularMarketPrice : result.indicators.quote[0].close.slice(-1)[0];
    if (price == null) return null;

    const prevClose = (meta && meta.chartPreviousClose !== undefined) ? meta.chartPreviousClose : null;
    const volume = (meta && meta.regularMarketVolume !== undefined) ? meta.regularMarketVolume : (result.indicators.quote[0].volume ? result.indicators.quote[0].volume.slice(-1)[0] : null);

    // 株式分割イベント(events=splitパラメータを付けているため、追加の通信なしで同じレスポンスに含まれる)
    const splits = [];
    if (result.events && result.events.splits) {
      Object.keys(result.events.splits).forEach(key => {
        const sp = result.events.splits[key];
        if (sp && sp.numerator && sp.denominator) {
          splits.push({
            dateStr: Utilities.formatDate(new Date(sp.date * 1000), 'Asia/Tokyo', 'yyyy-MM-dd'),
            factor: sp.numerator / sp.denominator // 例: 1→2の分割なら2。分割日より前の価格をこの値で割ると連続する
          });
        }
      });
    }

    return { price: price, prevClose: prevClose, volume: volume, splits: splits };
  } catch (e) {
    return null;
  }
}

function fetchClose(code) {
  for (let attempt = 0; attempt < 2; attempt++) {
    try {
      const url = 'https://query1.finance.yahoo.com/v8/finance/chart/' + code + '.T?range=5d&events=split';
      const options = { method: 'get', muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } };
      const response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() === 200) {
        const parsed = parseChartClose(response.getContentText());
        if (parsed) return parsed;
      }
      Utilities.sleep(1000);
    } catch (e) {
      Utilities.sleep(1000);
    }
  }
  return null;
}

/**
 * 複数銘柄の終値をまとめて並列取得する({コード: 結果 or null} を返す)。
 * 並列で失敗した銘柄だけ、時間に余裕があれば従来の逐次取得(リトライ付き)でやり直す。
 */
function fetchClosesBulk(codes, startMs) {
  const results = {};
  const BATCH = 10;
  for (let i = 0; i < codes.length; i += BATCH) {
    const batch = codes.slice(i, i + BATCH);
    const requests = batch.map(code => ({
      url: 'https://query1.finance.yahoo.com/v8/finance/chart/' + code + '.T?range=5d&events=split',
      method: 'get', muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' }
    }));
    let responses = [];
    try { responses = UrlFetchApp.fetchAll(requests); } catch (e) { responses = []; }
    batch.forEach((code, k) => {
      const r = responses[k];
      results[code] = (r && r.getResponseCode() === 200) ? parseChartClose(r.getContentText()) : null;
    });
    if (i + BATCH < codes.length) Utilities.sleep(300);
  }
  codes.forEach(code => {
    if (!results[code] && (Date.now() - startMs) < RUN_TIME_LIMIT_MS) results[code] = fetchClose(code);
  });
  return results;
}

// ==== 株式分割の自動調整(DBのみ。当月シートの実際の取引価格は変更しない) ====
const SPLIT_PROCESSED_PROPERTY = 'PROCESSED_SPLITS'; // {コード: [処理済み分割日,...]} をJSONで保持
// 統合前(株式情報_DB側)で処理済みの分割の記録はこのプロジェクトに引き継がれないため、
// この日付より前の分割は「処理済み」とみなして、データベースを二重に調整しないようにする。
const SPLIT_ADJUST_FROM_DATE = '2026-10-11';

/**
 * 分割日より前のDB上の過去データを分割比率で割り、スパークラインが連続するように揃える。
 * 同じ分割は一度だけ処理される(処理済みはスクリプトプロパティに記録)。
 */
function applyStockSplitsIfNeeded(code, splits) {
  if (!splits || splits.length === 0) return;

  const props = PropertiesService.getScriptProperties();
  const store = JSON.parse(props.getProperty(SPLIT_PROCESSED_PROPERTY) || '{}');
  const done = new Set(store[code] || []);

  splits.forEach(sp => {
    if (done.has(sp.dateStr) || !sp.factor || sp.factor === 1) return;
    if (sp.dateStr < SPLIT_ADJUST_FROM_DATE) return; // 統合前に処理済みの分割(上の定数の説明を参照)

    const adjusted = adjustDatabaseForSplit(code, sp.dateStr, sp.factor);
    done.add(sp.dateStr);
    store[code] = Array.from(done);
    props.setProperty(SPLIT_PROCESSED_PROPERTY, JSON.stringify(store));

    Logger.log('株式分割を検知しました: コード' + code + ' / 分割日=' + sp.dateStr + ' / 比率=' + sp.factor
      + (adjusted ? ' → データベースの過去データを調整しました(当月シートの実際の取引価格は変更していません)' : ' → データベースに調整対象のデータがありませんでした'));
  });
}

/** 指定コードのDB列のうち、splitDateStrより前の日付の値を factor で割る */
function adjustDatabaseForSplit(code, splitDateStr, factor) {
  const dbSheet = getOrCreateDatabaseSheet();
  const lastCol = Math.max(dbSheet.getLastColumn(), 1);
  const headerRow = dbSheet.getRange(1, 1, 1, lastCol).getValues()[0];

  let col = -1;
  for (let c = 2; c <= lastCol; c++) {
    if (String(headerRow[c - 1]) === String(code)) { col = c; break; }
  }
  if (col === -1) return false;

  const lastRow = dbSheet.getLastRow();
  if (lastRow < 2) return false;

  const dates = dbSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const values = dbSheet.getRange(2, col, lastRow - 1, 1).getValues();

  let changed = false;
  for (let i = 0; i < dates.length; i++) {
    const d = dates[i][0];
    if (d instanceof Date && Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd') < splitDateStr && typeof values[i][0] === 'number') {
      values[i][0] = values[i][0] / factor;
      changed = true;
    }
  }
  if (changed) dbSheet.getRange(2, col, lastRow - 1, 1).setValues(values);
  return changed;
}

/**
 * 指定コードの指定期間の日次終値・出来高・株式分割情報を取得する。
 * 戻り値: { dateMap: {日付: {close, volume}}, splits: [{dateStr, factor}, ...] } または取得失敗時null。
 */
function fetchHistoricalCloses(code, startDate, endDate) {
  for (let attempt = 0; attempt < 2; attempt++) {
    if (attempt > 0) Utilities.sleep(1000);
    try {
      const period1 = Math.floor(startDate.getTime() / 1000);
      const period2 = Math.floor(endDate.getTime() / 1000) + 86400;
      const url = 'https://query1.finance.yahoo.com/v8/finance/chart/' + code + '.T?period1=' + period1 + '&period2=' + period2 + '&interval=1d&events=split';
      const response = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
      if (response.getResponseCode() !== 200) continue;
      const json = JSON.parse(response.getContentText());
      const result = json.chart && json.chart.result && json.chart.result[0];
      if (!result) continue;
      const timestamps = result.timestamp, closes = result.indicators.quote[0].close, volumes = result.indicators.quote[0].volume;
      if (!timestamps || !closes) continue;
      const dateDataMap = {};
      for (let i = 0; i < timestamps.length; i++) {
        if (closes[i] != null) dateDataMap[Utilities.formatDate(new Date(timestamps[i] * 1000), 'Asia/Tokyo', 'yyyy-MM-dd')] = { close: closes[i], volume: volumes && volumes[i] != null ? volumes[i] : null };
      }

      const splits = [];
      if (result.events && result.events.splits) {
        Object.keys(result.events.splits).forEach(key => {
          const sp = result.events.splits[key];
          if (sp && sp.numerator && sp.denominator) {
            splits.push({ dateStr: Utilities.formatDate(new Date(sp.date * 1000), 'Asia/Tokyo', 'yyyy-MM-dd'), factor: sp.numerator / sp.denominator });
          }
        });
      }

      return { dateMap: dateDataMap, splits: splits };
    } catch (e) {}
  }
  return null;
}

// ==== 日次更新(トリガー本体) ====
function updateDailyClose() {
  const startMs = Date.now();
  const today = new Date();
  const sheet = getTargetSheet(today);

  if (!sheet) {
    Logger.log('今月に対応するシートが見つかりません。');
    return '今月に対応するシートが見つかりません';
  }

  const year = today.getFullYear();
  const month = today.getMonth() + 1;
  const day = today.getDate();
  const targetRow = CLOSE_FIRST_ROW + (day - 1);

  const lastCol = sheet.getLastColumn();
  const codeRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];

  const dow = today.getDay();
  const holiday = (dow === 0 || dow === 6) || isJapaneseHoliday(today);
  const blockWidth = lastCol - 2;

  if (holiday) {
    if (blockWidth > 0) {
      const blockValues = new Array(blockWidth).fill('');
      const blockColors = new Array(blockWidth).fill(null);
      MARKET_VALUE_ITEMS.forEach(item => {
        for (let k = 0; k < MARKET_VALUE_COL_WIDTH; k++) {
          blockValues[item.col - 3 + k] = HOLIDAY_LABEL;
          blockColors[item.col - 3 + k] = HOLIDAY_COLOR;
        }
      });
      for (let col = STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) {
        if (!codeRow[col - 1]) continue;
        for (let k = 0; k < STOCK_COL_WIDTH; k++) {
          blockValues[col - 3 + k] = HOLIDAY_LABEL;
          blockColors[col - 3 + k] = HOLIDAY_COLOR;
        }
      }
      const blockRange = sheet.getRange(targetRow, 3, 1, blockWidth);
      blockRange.setValues([blockValues]);
      blockRange.setBackgrounds([blockColors]);
    }
    createNextMonthSheetIfNeeded(year, month, startMs); // 月末が近ければ来月シートを作成/再開(休場日でも実行)
    return '本日は休場日のため、休マークを入力しました';
  }

  const rowValues = sheet.getRange(targetRow, 1, 1, lastCol).getValues()[0];
  const priceMap = {};
  const todayStr = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy-MM-dd');

  // 通常銘柄の終値をまとめて並列取得
  const stockCols = [];
  for (let col = STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) {
    const code = codeRow[col - 1];
    if (code) stockCols.push({ col: col, code: String(code) });
  }
  const normalCodes = stockCols.map(s => s.code);
  const closes = fetchClosesBulk(normalCodes, startMs);
  // 前日比はYahoo APIのchartPreviousCloseではなく、DBに記録済みの前営業日の実終値から計算する
  // (chartPreviousCloseは銘柄によって実際の前営業日と一致しないことがあるため)
  const prevCloseMap = getPreviousCloseMap(getOrCreateDatabaseSheet(), todayStr);

  const failedCodes = [];

  // プライム/スタンダード/グロースの売買代金(商況プリント２(後場)PDFより)
  const marketValues = fetchMarketTradingValuesForDate(todayStr);
  writeMarketValuesToRow_(rowValues, marketValues, prevCloseMap, priceMap, failedCodes);

  stockCols.forEach(s => {
    const col = s.col;
    const result = closes[s.code];
    if (!result) { failedCodes.push(s.code); return; }
    rowValues[col - 1] = result.price;
    const prevClose = prevCloseMap[s.code];
    // prevCloseが取れない場合も前日比セルは明示的に空欄にする(古い値が残って誤表示される事故を防ぐ)
    rowValues[col] = (typeof prevClose === 'number' && prevClose !== 0) ? (result.price / prevClose - 1) : '';
    if (result.volume !== null) {
      rowValues[col + 1] = roundTradingValueToOkuUnit_(result.price, result.volume); // 売買代金(終値×出来高)を10億円単位で表示
    }
    priceMap[s.code] = result.price;
    if (result.splits && result.splits.length > 0) applyStockSplitsIfNeeded(s.code, result.splits);
  });

  sheet.getRange(targetRow, 1, 1, lastCol).setValues([rowValues]);

  try {
    writeClosesToDatabaseForDate(today, priceMap);
  } catch (e) {
    Logger.log('データベースへの一括書き込みに失敗しました: ' + e);
  }

  createNextMonthSheetIfNeeded(year, month, startMs); // 月末が近ければ来月シートを作成/再開

  const totalItems = stockCols.length + MARKET_VALUE_ITEMS.length;
  const message = '本日の終値取得処理が完了しました(' + Object.keys(priceMap).length + '/' + totalItems + '項目)'
    + (failedCodes.length > 0 ? '\n取得できなかった項目: ' + failedCodes.join(', ') : '');
  Logger.log(message);
  return message;
}

/**
 * 市場ブロック(プライム/スタンダード/グロース)の出来高・売買代金・前日比を1行分の配列に書き込む。
 * 前日比は売買代金の前営業日比(DBに記録済みの前営業日の売買代金と比べる)。
 */
function writeMarketValuesToRow_(rowValues, marketValues, prevCloseMap, priceMap, failedCodes) {
  MARKET_VALUE_ITEMS.forEach(item => {
    const mv = marketValues ? marketValues[item.code] : null;
    const value = mv ? mv.value : null;
    if (value === null || value === undefined) { failedCodes.push(item.code); return; }
    const base = item.col - 1;
    if (mv.volume !== null && mv.volume !== undefined) rowValues[base + MARKET_VOLUME_OFFSET] = mv.volume;
    rowValues[base + MARKET_VALUE_OFFSET] = value;
    const prevClose = prevCloseMap[item.code];
    // prevCloseが取れない場合も前日比セルは明示的に空欄にする(古い値が残って誤表示される事故を防ぐ)
    rowValues[base + MARKET_CHANGE_OFFSET] = (typeof prevClose === 'number' && prevClose !== 0) ? (value / prevClose - 1) : '';
    priceMap[item.code] = value;
  });
}

/** 「市場別売買代金を更新(日付入力)」メニューから本日分だけを単独実行したい場合用 */
function updateMarketTradingValues() {
  return updateCloseForSpecificDate(new Date());
}

/** 指定期間の日次チャート(終値・出来高・分割情報)を取得するYahoo APIのURLを組み立てる */
function fetchHistoricalChart(code, startDate, endDate) {
  const period1 = Math.floor(startDate.getTime() / 1000);
  const period2 = Math.floor(endDate.getTime() / 1000) + 86400;
  return 'https://query1.finance.yahoo.com/v8/finance/chart/' + code + '.T?period1=' + period1 + '&period2=' + period2 + '&interval=1d&events=split';
}

function parseHistoricalChart(text) {
  try {
    const json = JSON.parse(text);
    const result = json.chart && json.chart.result && json.chart.result[0];
    if (!result) return null;
    const timestamps = result.timestamp, closes = result.indicators.quote[0].close, volumes = result.indicators.quote[0].volume;
    if (!timestamps || !closes) return null;

    const dateMap = {};
    for (let i = 0; i < timestamps.length; i++) {
      if (closes[i] != null) dateMap[Utilities.formatDate(new Date(timestamps[i] * 1000), 'Asia/Tokyo', 'yyyy-MM-dd')] = { close: closes[i], volume: volumes && volumes[i] != null ? volumes[i] : null };
    }

    const splits = [];
    if (result.events && result.events.splits) {
      Object.keys(result.events.splits).forEach(key => {
        const sp = result.events.splits[key];
        if (sp && sp.numerator && sp.denominator) {
          splits.push({ dateStr: Utilities.formatDate(new Date(sp.date * 1000), 'Asia/Tokyo', 'yyyy-MM-dd'), factor: sp.numerator / sp.denominator });
        }
      });
    }
    return { dateMap: dateMap, splits: splits };
  } catch (e) {
    return null;
  }
}

/** 複数銘柄の日次チャートを並列取得する({コード: {dateMap, splits} or null}) */
function fetchHistoricalChartsBulk(codes, startDate, endDate, startMs) {
  const results = {};
  const BATCH = 10;
  for (let i = 0; i < codes.length; i += BATCH) {
    const batch = codes.slice(i, i + BATCH);
    const requests = batch.map(code => ({
      url: fetchHistoricalChart(code, startDate, endDate),
      method: 'get', muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' }
    }));
    let responses = [];
    try { responses = UrlFetchApp.fetchAll(requests); } catch (e) { responses = []; }
    batch.forEach((code, k) => {
      const r = responses[k];
      results[code] = (r && r.getResponseCode() === 200) ? parseHistoricalChart(r.getContentText()) : null;
    });
    if (i + BATCH < codes.length) Utilities.sleep(300);
  }
  codes.forEach(code => {
    if (!results[code] && startMs && (Date.now() - startMs) < RUN_TIME_LIMIT_MS) {
      try {
        const response = UrlFetchApp.fetch(fetchHistoricalChart(code, startDate, endDate), { method: 'get', muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
        if (response.getResponseCode() === 200) results[code] = parseHistoricalChart(response.getContentText());
      } catch (e) {}
    }
  });
  return results;
}

/**
 * 指定した1日分の終値・前日比・売買代金(銘柄別は終値×出来高の概算、プライム/スタンダード/グロースは実際の売買代金)を
 * 当月シートとDBに書き込む(過去日付にも対応。商況プリントPDFは日付指定で取得できるため)。
 */
function updateCloseForSpecificDate(targetDate) {
  const startMs = Date.now();
  const targetDateStr = Utilities.formatDate(targetDate, 'Asia/Tokyo', 'yyyy-MM-dd');
  const sheet = getTargetSheet(targetDate);

  if (!sheet) return targetDateStr + ': 対応するシートが見つかりません';

  const day = targetDate.getDate();
  const targetRow = CLOSE_FIRST_ROW + (day - 1);
  const lastCol = sheet.getLastColumn();
  const codeRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];

  const dow = targetDate.getDay();
  if ((dow === 0 || dow === 6) || isJapaneseHoliday(targetDate)) {
    return targetDateStr + ': 休場日のためスキップしました';
  }

  const rowValues = sheet.getRange(targetRow, 1, 1, lastCol).getValues()[0];
  const priceMap = {};

  const stockCols = [];
  for (let col = STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) {
    const code = codeRow[col - 1];
    if (code) stockCols.push({ col: col, code: String(code) });
  }
  const normalCodes = stockCols.map(s => s.code);

  // 前後2日の余裕を持たせて取得し、ちょうどtargetDateStrのデータを使う(休場日や祝日のズレ対策)
  const fetchStart = new Date(targetDate); fetchStart.setDate(fetchStart.getDate() - 3);
  const fetchEnd = new Date(targetDate); fetchEnd.setDate(fetchEnd.getDate() + 1);
  const charts = fetchHistoricalChartsBulk(normalCodes, fetchStart, fetchEnd, startMs);
  const prevCloseMap = getPreviousCloseMap(getOrCreateDatabaseSheet(), targetDateStr);

  const failedCodes = [];

  const marketValues = fetchMarketTradingValuesForDate(targetDateStr);
  writeMarketValuesToRow_(rowValues, marketValues, prevCloseMap, priceMap, failedCodes);

  stockCols.forEach(s => {
    const col = s.col;
    const chart = charts[s.code];
    const dayData = chart && chart.dateMap[targetDateStr];
    if (!dayData) { failedCodes.push(s.code); return; }

    rowValues[col - 1] = dayData.close;
    const prevClose = prevCloseMap[s.code];
    // prevCloseが取れない場合も前日比セルは明示的に空欄にする(古い値が残って誤表示される事故を防ぐ)
    rowValues[col] = (typeof prevClose === 'number' && prevClose !== 0) ? (dayData.close / prevClose - 1) : '';
    if (dayData.volume !== null) {
      const splitFactor = computeSplitCorrectionFactor_(targetDateStr, chart.splits);
      rowValues[col + 1] = roundTradingValueToOkuUnit_(dayData.close, dayData.volume, splitFactor);
    }
    priceMap[s.code] = dayData.close;
    if (chart.splits && chart.splits.length > 0) applyStockSplitsIfNeeded(s.code, chart.splits);
  });

  sheet.getRange(targetRow, 1, 1, lastCol).setValues([rowValues]);
  try { writeClosesToDatabaseForDate(targetDate, priceMap); } catch (e) { Logger.log(targetDateStr + ': データベース書き込みに失敗: ' + e); }

  const totalItems = stockCols.length + MARKET_VALUE_ITEMS.length;
  let message = targetDateStr + ': ' + Object.keys(priceMap).length + '/' + totalItems + '項目を更新しました';
  if (failedCodes.length > 0) message += ' / 取得できなかった項目: ' + failedCodes.join(', ');
  return message;
}

/** ダイアログからの単日更新 */
function updateCloseForDateInput(dateStr) {
  const message = updateCloseForSpecificDate(parseDateStr_(dateStr));
  Logger.log(message);
  return message;
}

/** ダイアログからの期間指定更新。時間切れが近づいたら安全に止まり、続きは再実行で進む */
function updateCloseForDateRangeInput(startDateStr, endDateStr) {
  const rangeStartMs = Date.now();
  const start = parseDateStr_(startDateStr);
  const end = parseDateStr_(endDateStr);
  if (start > end) return '開始日は終了日より前の日付にしてください';

  const messages = [];
  for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
    if (Date.now() - rangeStartMs > RUN_TIME_LIMIT_MS) {
      messages.push('(実行時間の上限のため、' + Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd') + '以降は未処理です。同じ期間でもう一度実行すると続きから処理されます)');
      break;
    }
    messages.push(updateCloseForSpecificDate(new Date(d)));
  }
  const message = messages.join('\n');
  Logger.log(message);
  return message;
}

/**
 * 来月シートの作成・休日マーク入力を行う。中断しても安全に再開できる作り:
 * ・シートが既にあれば作り直さず、未入力の日だけ続きから埋める
 * ・実行時間の上限(RUN_TIME_LIMIT_MS)が近づいたら、その日までの状態を残して安全に抜ける
 * 月末に限らず毎日の更新の最後に呼ばれる想定(完成していれば一瞬で終わるため無駄にならない)。
 */
function createNextMonthSheetIfNeeded(currentYear, currentMonth, startMs) {
  // 月末が近づく(残り6日以内)までは何もしない
  if (getDaysInMonth(currentYear, currentMonth) - new Date().getDate() > 5) return;

  let nextYear = currentYear;
  let nextMonth = currentMonth + 1;
  if (nextMonth > 12) { nextMonth = 1; nextYear += 1; }
  const newSheetName = nextMonth + '月' + String(nextYear).slice(-2) + '年';

  const ss = getMainSpreadsheet();
  let newSheet = ss.getSheetByName(newSheetName);
  const daysInMonth = getDaysInMonth(nextYear, nextMonth);

  if (!newSheet) {
    const template = ss.getSheetByName(TEMPLATE_SHEET_NAME);
    if (!template) return;
    newSheet = template.copyTo(ss);
    newSheet.setName(newSheetName);
    ss.setActiveSheet(newSheet);
    ss.moveActiveSheet(template.getIndex() + 1);
    newSheet.setFrozenRows(5);
  }

  const dayLabels = newSheet.getRange(CLOSE_FIRST_ROW, 1, daysInMonth, 1).getValues();
  if (dayLabels[daysInMonth - 1][0] !== '') return; // 最終日まで入力済み=作成完了済み

  const lastCol = newSheet.getLastColumn();
  const codeRow = newSheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const maxRows = newSheet.getMaxRows();
  // 休場マークを入れる列ブロック(市場ブロック+登録済み銘柄)
  const holidayBlocks = MARKET_VALUE_ITEMS.map(item => ({ col: item.col, width: MARKET_VALUE_COL_WIDTH }));
  for (let col = STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) {
    if (codeRow[col - 1]) holidayBlocks.push({ col: col, width: STOCK_COL_WIDTH });
  }

  for (let d = 1; d <= daysInMonth; d++) {
    if (dayLabels[d - 1][0] !== '') continue; // 前回までに処理済みの日はスキップ(再開時)
    if (startMs && (Date.now() - startMs) > RUN_TIME_LIMIT_MS) {
      Logger.log(newSheetName + ' シートの作成が途中(' + (d - 1) + '/' + daysInMonth + '日)で時間切れになりました。翌日以降の更新で続きが作成されます。');
      return;
    }

    const date = new Date(nextYear, nextMonth - 1, d);
    const dow = date.getDay();
    const isHoliday = (dow === 0 || dow === 6) || isJapaneseHoliday(date);

    // 終値の日次行(6行目〜)と信用残の日次行(39行目〜)の両方に日付・曜日・休場マークを入れる。
    // 終値側のA列(完成判定に使う)は最後に書くので、途中で止まっても次回この日からやり直せる。
    [MARGIN_FIRST_ROW + (d - 1), CLOSE_FIRST_ROW + (d - 1)].forEach(row => {
      if (row > maxRows) return;
      newSheet.getRange(row, 2).setValue(MARGIN_WEEKDAY_NAMES_[dow]);
      if (isHoliday) {
        // 休場日は全ブロックをまとめて1回で「休」+グレーにする
        const a1List = holidayBlocks.map(b => columnToLetter(b.col) + row + ':' + columnToLetter(b.col + b.width - 1) + row);
        newSheet.getRangeList(a1List).setValue(HOLIDAY_LABEL).setBackground(HOLIDAY_COLOR);
      }
      newSheet.getRange(row, 1).setValue(d + '日');
    });
  }
  Logger.log(newSheetName + ' シートの作成が完了しました');
}