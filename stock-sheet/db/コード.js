/**
 * 日本株 終値自動取得システム(サブファイル/ライブラリ、プロジェクト名: 株式情報_DB)
 * ------------------------------------------------------------
 * このプロジェクトの役割:
 *  1) 日次・週次の時間主導トリガーで、メインファイルのデータを自動更新する
 *     (終値・信用残・投資部門別売買状況、およびデータベースシートへの書き込み)
 *  2) メインファイルの「銘柄追加」機能から呼び出される取得関数(fetchClose等)を
 *     ライブラリとして提供する
 *
 * データ(当月シート・データベースシート・テンプレート・売買状況シート)自体は
 * すべてメインファイル側に置いたまま。このプロジェクトはSpreadsheetApp.openByIdで
 * メインファイルを直接開いて読み書きする。
 *
 * 【セットアップ手順】
 * 1. MAIN_SPREADSHEET_ID にメインファイルのスプレッドシートIDを設定する
 *    (メインファイルのURL: https://docs.google.com/spreadsheets/d/【ここがID】/edit)
 * 2. Apps Scriptエディタの左メニュー「サービス」から
 *    Google Sheets API(並び替え機能を使わないならこのプロジェクトでは不要)、
 *    Drive API(JPXのPDF→テキスト変換に必要)を有効化する
 * 3. 「デプロイ」→「新しいデプロイ」→種類「ライブラリ」を選んで公開し、
 *    表示されるスクリプトIDを控える
 * 4. メインファイル側のApps Scriptエディタで、左メニュー「ライブラリ」の+ボタンから
 *    3のスクリプトIDを追加し、識別子(ID)を "StockLib" にする
 *    (メイン側のコードは StockLib.fetchClose(...) のように呼び出す想定)
 * 5. このプロジェクトの「トリガー」ページで、以下の時間主導トリガーを設定する
 *      - updateDailyClose            : 毎日(平日大引け後、例 15:30以降)
 *      - updateMarginData            : 毎日(JPXの日次信用残。前日までの未入力日を埋める。例 18:00以降)
 *      - updateJpxStatsForLatestWeek : 週1回(金〜土あたり)
 *      - updateMarketTradingValues   : 毎日(商況プリント２(後場)発表後、例 16:15以降)
 *      - updateShortSellForToday     : 毎日(空売り集計PDF公開後、例 17:30以降)
 *
 * 注意: このプロジェクトはスプレッドシートに紐付いていないため SpreadsheetApp.getUi() は
 * 使用できない。実行結果・エラーはすべて実行ログ(Logger.log / 表示 > 実行数)で確認する。
 * 個別機能(initializeDatabaseForAllStocksなど)はメインファイル(株式情報)の
 * 「機能」→「データ更新(株式情報_DB)」メニューからシート画面上で実行できる。
 * このプロジェクトのエディタから直接関数を選んで実行することもできる。
 */

// ==== 設定項目(メインファイルのIDを必ず書き換えること) ====
const MAIN_SPREADSHEET_ID = '1WE6PgUNDmF-FkuidsgyRSaDKQXdkTngHZArOz1optZg';

const TEMPLATE_SHEET_NAME = 'テンプレート';
const STOCK_COL_WIDTH = 3;
const STOCK_START_COL = 12; // 通常銘柄はL列から(C〜K列はプライム/スタンダード/グロースの市場ブロックに固定で使う)
// 行構成: 5行目=見出し / 6〜36行目=日次の終値等(1日〜31日) / 38行目=信用残見出し / 39〜69行目=日次の信用残(1日〜31日)
const CLOSE_FIRST_ROW = 6;
const MARGIN_HEADER_ROW = 38;
const MARGIN_FIRST_ROW = 39;
const DB_SHEET_NAME = 'データベース';
const DB_BACKFILL_YEARS = 3; // データベースに保持する過去データの年数
const RUN_TIME_LIMIT_MS = 4.5 * 60 * 1000; // 6分制限の手前で安全に止めるための実行時間上限

// ==== 市場ブロック(プライム/スタンダード/グロース。C:E・F:H・I:Kの3列ブロック) ====
// テンプレート・当月シートのC〜K列に固定配置。通常銘柄(3列ブロック)とは別枠として扱う。
//   日次行(6〜36行目): [出来高(百万株), 売買代金(10億円), 前日比(売買代金の前日比)]
//   信用残行(39〜69行目): [売残(千株), 買残(千株), 信用倍率]
const MARKET_VALUE_COL_WIDTH = 3;
const MARKET_VOLUME_OFFSET = 0;
const MARKET_VALUE_OFFSET = 1;
const MARKET_CHANGE_OFFSET = 2;
const MARKET_VALUE_ITEMS = [
  { code: '1001', name: 'プライム', col: 3 },
  { code: '1002', name: 'スタンダード', col: 6 },
  { code: '1003', name: 'グロース', col: 9 }
];

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

/** PDFのBlobをGoogle ドキュメントに変換してテキストを取り出す(変換用ファイルは処理後に削除) */
function convertPdfBlobToText_(blob) {
  const resource = { name: 'temp_pdf_convert_' + new Date().getTime(), mimeType: MimeType.GOOGLE_DOCS };
  const file = Drive.Files.create(resource, blob, { convert: true });
  try {
    return DocumentApp.openById(file.id).getBody().getText();
  } finally {
    DriveApp.getFileById(file.id).setTrashed(true);
  }
}

function fetchPdfTextFromUrl_(url) {
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) return null;
  return convertPdfBlobToText_(response.getBlob());
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

const HOLIDAY_LABEL = '休';
const HOLIDAY_COLOR = '#999999';

function getMainSpreadsheet() {
  return SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
}

// ==== シート特定・日付ユーティリティ ====
function getTargetSheet(today) {
  const targetYearStr = String(today.getFullYear()).slice(-2);
  const targetMonth = today.getMonth() + 1;
  const targetName = targetMonth + '月' + targetYearStr + '年';
  return getMainSpreadsheet().getSheetByName(targetName);
}

function getAdjacentMonthSheet(year, month, offset) {
  let targetYear = year;
  let targetMonth = month + offset;

  while (targetMonth <= 0) { targetMonth += 12; targetYear -= 1; }
  while (targetMonth > 12) { targetMonth -= 12; targetYear += 1; }

  const targetYearStr = String(targetYear).slice(-2);
  const targetName = targetMonth + '月' + targetYearStr + '年';
  return getMainSpreadsheet().getSheetByName(targetName);
}

function isLastDayOfMonth(date) {
  const tomorrow = new Date(date);
  tomorrow.setDate(date.getDate() + 1);
  return tomorrow.getMonth() !== date.getMonth();
}

function getDaysInMonth(year, month) {
  return new Date(year, month, 0).getDate();
}

// Googleの「日本の祝日」カレンダーは、ある時期から七夕・節分・七五三・クリスマスのような
// 国民の祝日ではない行事もイベントとして登録するようになったため、イベントの有無だけで
// 判定すると7/7(七夕)のような平日を休場日と誤判定してしまう。
// そのため、実際の「国民の祝日に関する法律」に基づく祝日名(+振替休日)のみを休日として扱う。
const JP_PUBLIC_HOLIDAY_NAMES_ = [
  '元日', '成人の日', '建国記念の日', '天皇誕生日', '春分の日', '昭和の日',
  '憲法記念日', 'みどりの日', 'こどもの日', '海の日', '山の日', '敬老の日',
  '秋分の日', 'スポーツの日', '体育の日', '文化の日', '勤労感謝の日',
  '振替休日', '国民の休日'
];

function isJapaneseHoliday(date) {
  const month = date.getMonth() + 1;
  const day = date.getDate();
  if ((month === 1 && day <= 3) || (month === 12 && day === 31)) return true;

  try {
    const holidayCalendarId = 'ja.japanese#holiday@group.v.calendar.google.com';
    const events = CalendarApp.getCalendarById(holidayCalendarId).getEventsForDay(date);
    return events.some(function (e) {
      const title = e.getTitle();
      return JP_PUBLIC_HOLIDAY_NAMES_.indexOf(title) !== -1;
    });
  } catch (e) {
    return false;
  }
}

function roundVolumeToThousands(volume) {
  if (volume === null || volume === undefined) return null;
  return Math.round(volume / 1000);
}

/**
 * 株式分割の補正係数を計算する。
 * YahooのチャートAPIは、過去の「終値」を直近までに起きた株式分割の比率に遡って調整して返すが、
 * 「出来高」は調整しない(発生当時の実株数のまま)。そのため分割発生日より前の日付について
 * 終値×出来高で売買代金を概算すると、分割後に取得し直した場合に実際より小さい値になってしまう。
 * この関数は、対象日より後(=対象日に対してYahooが遡って調整をかけた分)に発生した分割の比率を
 * すべて掛け合わせて返す。これを終値×出来高の結果に掛けることで実際の売買代金に戻す。
 */
function computeSplitCorrectionFactor_(targetDateStr, splits) {
  if (!splits || splits.length === 0) return 1;
  let factor = 1;
  splits.forEach(function (sp) {
    if (sp && sp.dateStr && sp.factor && sp.dateStr > targetDateStr) factor *= sp.factor;
  });
  return factor;
}

/**
 * 終値×出来高から、その日の概算売買代金を求め、10億円単位(十億円単位)に丸めて返す。
 * (セルに入る数値は「10億円の個数」。例: 12 → 120億円)
 * splitCorrectionFactorは、computeSplitCorrectionFactor_で求めた株式分割補正係数(省略時は1)。
 */
function roundTradingValueToOkuUnit_(close, volume, splitCorrectionFactor) {
  if (close === null || close === undefined || volume === null || volume === undefined) return null;
  const factor = splitCorrectionFactor || 1;
  return Math.round((close * volume * factor) / 1000000000);
}

// ==== 株価取得(ライブラリとしてメインファイルからも呼び出される) ====
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

// ==== 日付指定更新(ダイアログから特定日・期間の終値・前日比を更新) ====

function parseDateStr_(dateStr) {
  const parts = String(dateStr).split('-').map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

/** 複数銘柄の、指定期間内の日次終値・出来高(売買代金算出用)・分割情報をまとめて取得する */
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
  const weekdayNames = ['日曜日', '月曜日', '火曜日', '水曜日', '木曜日', '金曜日', '土曜日'];

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
      if (row > newSheet.getMaxRows()) return;
      newSheet.getRange(row, 2).setValue(weekdayNames[dow]);
      if (isHoliday) {
        MARKET_VALUE_ITEMS.forEach(item => {
          newSheet.getRange(row, item.col, 1, MARKET_VALUE_COL_WIDTH).setValue(HOLIDAY_LABEL).setBackground(HOLIDAY_COLOR);
        });
        for (let col = STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) {
          if (codeRow[col - 1]) newSheet.getRange(row, col, 1, STOCK_COL_WIDTH).setValue(HOLIDAY_LABEL).setBackground(HOLIDAY_COLOR);
        }
      }
      newSheet.getRange(row, 1).setValue(d + '日');
    });
  }
  Logger.log(newSheetName + ' シートの作成が完了しました');
}

// ==== JPX 銘柄別信用取引残高(mtall.pdf)の解析 ====
// 109ページ程度の大きなPDFのため、Googleドキュメント変換(ページ数制限・時間)に頼らず、
// PDFの中身(圧縮されたテキスト描画命令)を直接展開して文字を取り出す。
// PDFは「描画順」に文字が並ぶため、1銘柄が
//   「コード ISIN 株数 Shs. 売残 前日比 上場比 買残 前日比 上場比 …」
// の順で連続する(実ファイルで確認済み)。この並びから売残・買残を読む。

// ---- Deflate(Flate)展開(zlib形式のストリームを展開する最小実装) ----
function mtInflate_(src) {
  // zlibヘッダ(2バイト)を飛ばす。ヘッダがない(生deflate)場合はそのまま
  let pos = 0;
  if (src.length > 2 && (src[0] & 0x0f) === 8 && ((src[0] << 8) | src[1]) % 31 === 0) pos = 2;

  let out = new Uint8Array(Math.max(src.length * 4, 1024));
  let outLen = 0;
  function ensure(n) {
    if (outLen + n <= out.length) return;
    let size = out.length * 2;
    while (size < outLen + n) size *= 2;
    const b = new Uint8Array(size);
    b.set(out.subarray(0, outLen));
    out = b;
  }

  let bitBuf = 0, bitCnt = 0;
  function bits(n) {
    while (bitCnt < n) {
      if (pos >= src.length) throw new Error('inflate: unexpected end');
      bitBuf |= src[pos++] << bitCnt;
      bitCnt += 8;
    }
    const v = bitBuf & ((1 << n) - 1);
    bitBuf >>>= n;
    bitCnt -= n;
    return v;
  }

  // Huffman表: counts[len], symbols[]
  function buildTree(lengths, n) {
    const counts = new Uint16Array(16), offs = new Uint16Array(16), symbols = new Uint16Array(n);
    for (let i = 0; i < n; i++) counts[lengths[i]]++;
    counts[0] = 0;
    for (let i = 1; i < 16; i++) offs[i] = offs[i - 1] + counts[i - 1];
    for (let i = 0; i < n; i++) if (lengths[i]) symbols[offs[lengths[i]]++] = i;
    return { counts: counts, symbols: symbols };
  }
  function decodeSym(t) {
    let code = 0, first = 0, index = 0;
    for (let len = 1; len < 16; len++) {
      code |= bits(1);
      const count = t.counts[len];
      if (code - count < first) return t.symbols[index + (code - first)];
      index += count;
      first += count;
      first <<= 1;
      code <<= 1;
    }
    throw new Error('inflate: bad code');
  }

  const LBASE = [3,4,5,6,7,8,9,10,11,13,15,17,19,23,27,31,35,43,51,59,67,83,99,115,131,163,195,227,258];
  const LEXT = [0,0,0,0,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,5,5,5,5,0];
  const DBASE = [1,2,3,4,5,7,9,13,17,25,33,49,65,97,129,193,257,385,513,769,1025,1537,2049,3073,4097,6145,8193,12289,16385,24577];
  const DEXT = [0,0,0,0,1,1,2,2,3,3,4,4,5,5,6,6,7,7,8,8,9,9,10,10,11,11,12,12,13,13];
  const CLORDER = [16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15];

  let fixedL = null, fixedD = null;
  function fixedTrees() {
    if (fixedL) return;
    const l = new Uint8Array(288);
    for (let i = 0; i < 144; i++) l[i] = 8;
    for (let i = 144; i < 256; i++) l[i] = 9;
    for (let i = 256; i < 280; i++) l[i] = 7;
    for (let i = 280; i < 288; i++) l[i] = 8;
    fixedL = buildTree(l, 288);
    const d = new Uint8Array(30).fill(5);
    fixedD = buildTree(d, 30);
  }

  function inflateBlock(lt, dt) {
    for (;;) {
      const sym = decodeSym(lt);
      if (sym < 256) { ensure(1); out[outLen++] = sym; continue; }
      if (sym === 256) return;
      const li = sym - 257;
      const len = LBASE[li] + (LEXT[li] ? bits(LEXT[li]) : 0);
      const di = decodeSym(dt);
      const dist = DBASE[di] + (DEXT[di] ? bits(DEXT[di]) : 0);
      ensure(len);
      let from = outLen - dist;
      for (let k = 0; k < len; k++) out[outLen++] = out[from++];
    }
  }

  let last = 0;
  do {
    last = bits(1);
    const type = bits(2);
    if (type === 0) {
      bitBuf = 0; bitCnt = 0; // バイト境界にそろえる
      const len = src[pos] | (src[pos + 1] << 8);
      pos += 4;
      ensure(len);
      out.set(src.subarray(pos, pos + len), outLen);
      outLen += len;
      pos += len;
    } else if (type === 1) {
      fixedTrees();
      inflateBlock(fixedL, fixedD);
    } else if (type === 2) {
      const hlit = bits(5) + 257, hdist = bits(5) + 1, hclen = bits(4) + 4;
      const cl = new Uint8Array(19);
      for (let i = 0; i < hclen; i++) cl[CLORDER[i]] = bits(3);
      const clt = buildTree(cl, 19);
      const lens = new Uint8Array(hlit + hdist);
      for (let i = 0; i < hlit + hdist;) {
        const sym = decodeSym(clt);
        if (sym < 16) { lens[i++] = sym; continue; }
        let rep = 0, val = 0;
        if (sym === 16) { val = lens[i - 1]; rep = 3 + bits(2); }
        else if (sym === 17) rep = 3 + bits(3);
        else rep = 11 + bits(7);
        while (rep--) lens[i++] = val;
      }
      inflateBlock(buildTree(lens.subarray(0, hlit), hlit), buildTree(lens.subarray(hlit), hdist));
    } else {
      throw new Error('inflate: bad block type');
    }
  } while (!last);

  return out.subarray(0, outLen);
}

// ---- PDFのオブジェクト・ストリームを読む ----
function mtBytesToLatin1_(bytes) {
  const parts = [];
  for (let i = 0; i < bytes.length; i += 8192) {
    parts.push(String.fromCharCode.apply(null, bytes.subarray(i, Math.min(i + 8192, bytes.length))));
  }
  return parts.join('');
}

/** PDF全体から {オブジェクト番号: {dict:文字列, stream:Uint8Array(展開済み)|null}} を作る */
function mtReadPdfObjects_(bytes) {
  const s = mtBytesToLatin1_(bytes);
  const objs = {};
  const re = /(\d+)\s+0\s+obj\b/g;
  let m;
  while ((m = re.exec(s)) !== null) {
    const num = parseInt(m[1], 10);
    const bodyStart = m.index + m[0].length;
    const endObj = s.indexOf('endobj', bodyStart);
    const streamKw = s.indexOf('stream', bodyStart);
    if (streamKw !== -1 && (endObj === -1 || streamKw < endObj)) {
      const dict = s.substring(bodyStart, streamKw);
      let dataStart = streamKw + 6;
      if (s.charCodeAt(dataStart) === 13) dataStart++;
      if (s.charCodeAt(dataStart) === 10) dataStart++;
      let dataEnd;
      const lenM = dict.match(/\/Length\s+(\d+)(?!\d)(?!\s+\d+\s+R)/);
      if (lenM) dataEnd = dataStart + parseInt(lenM[1], 10);
      else dataEnd = s.indexOf('endstream', dataStart);
      let data = bytes.subarray(dataStart, dataEnd);
      if (/\/FlateDecode/.test(dict)) {
        try { data = mtInflate_(data); } catch (e) { data = null; }
      }
      objs[num] = { dict: dict, stream: data };
      re.lastIndex = dataEnd;
    } else {
      objs[num] = { dict: s.substring(bodyStart, endObj === -1 ? s.length : endObj), stream: null };
      if (endObj !== -1) re.lastIndex = endObj;
    }
  }

  // オブジェクトストリーム(/Type/ObjStm)の中に格納されたオブジェクトも展開する
  Object.keys(objs).forEach(k => {
    const o = objs[k];
    if (!o.stream || !/\/Type\s*\/ObjStm/.test(o.dict)) return;
    const n = parseInt((o.dict.match(/\/N\s+(\d+)/) || [])[1], 10);
    const first = parseInt((o.dict.match(/\/First\s+(\d+)/) || [])[1], 10);
    const text = mtBytesToLatin1_(o.stream);
    const header = text.substring(0, first).trim().split(/\s+/).map(Number);
    for (let i = 0; i < n; i++) {
      const objNum = header[i * 2], off = header[i * 2 + 1];
      const nextOff = i + 1 < n ? header[(i + 1) * 2 + 1] : text.length - first;
      if (!objs[objNum]) objs[objNum] = { dict: text.substring(first + off, first + nextOff), stream: null };
    }
  });
  return objs;
}

function mtResolve_(objs, text) {
  const m = String(text).match(/^\s*(\d+)\s+0\s+R/);
  return m ? (objs[parseInt(m[1], 10)] || null) : null;
}

/** 辞書文字列から /Key の値(参照なら参照先の辞書文字列、直接の辞書ならその中身)を取り出す */
function mtDictValue_(objs, dict, key) {
  const idx = dict.search(new RegExp('/' + key + '(?![A-Za-z0-9])'));
  if (idx === -1) return null;
  let rest = dict.substring(idx + key.length + 1).replace(/^\s+/, '');
  const ref = rest.match(/^(\d+)\s+0\s+R/);
  if (ref) {
    const o = objs[parseInt(ref[1], 10)];
    return o ? o.dict : null;
  }
  if (rest.startsWith('<<')) {
    let depth = 0;
    for (let i = 0; i < rest.length - 1; i++) {
      if (rest[i] === '<' && rest[i + 1] === '<') { depth++; i++; }
      else if (rest[i] === '>' && rest[i + 1] === '>') { depth--; i++; if (depth === 0) return rest.substring(2, i - 1); }
    }
  }
  return rest;
}

/** ToUnicode CMapを読み、{コード(数値): 文字列} と コード長(バイト) を返す */
function mtParseCMap_(text) {
  const map = {};
  let codeLen = 1;
  const csr = text.match(/begincodespacerange\s*<([0-9A-Fa-f]+)>/);
  if (csr) codeLen = csr[1].length / 2;
  const hexToStr = h => {
    let r = '';
    for (let i = 0; i + 4 <= h.length; i += 4) r += String.fromCharCode(parseInt(h.substr(i, 4), 16));
    if (h.length === 2) r = String.fromCharCode(parseInt(h, 16));
    return r;
  };
  let m;
  const charBlocks = /beginbfchar([\s\S]*?)endbfchar/g;
  while ((m = charBlocks.exec(text)) !== null) {
    const pairRe = /<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/g;
    let p;
    while ((p = pairRe.exec(m[1])) !== null) map[parseInt(p[1], 16)] = hexToStr(p[2]);
  }
  const rangeBlocks = /beginbfrange([\s\S]*?)endbfrange/g;
  while ((m = rangeBlocks.exec(text)) !== null) {
    const rRe = /<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*(<([0-9A-Fa-f]+)>|\[([^\]]*)\])/g;
    let r;
    while ((r = rRe.exec(m[1])) !== null) {
      const lo = parseInt(r[1], 16), hi = parseInt(r[2], 16);
      if (r[4] !== undefined) {
        const base = parseInt(r[4], 16);
        for (let c = lo; c <= hi; c++) map[c] = String.fromCharCode(base + (c - lo));
      } else {
        const items = (r[5].match(/<([0-9A-Fa-f]+)>/g) || []).map(x => hexToStr(x.slice(1, -1)));
        for (let c = lo; c <= hi && c - lo < items.length; c++) map[c] = items[c - lo];
      }
    }
  }
  return { map: map, codeLen: codeLen };
}

/** 1ページの描画命令からテキストを描画順に取り出す(文字列描画命令ごとに空白で区切る) */
function mtExtractPageText_(content, fonts) {
  const s = mtBytesToLatin1_(content);
  const out = [];
  let font = null;
  let moved = true; // 直前の文字描画から位置移動(Td/Tm等)があったか。移動があれば空白で区切る
  let i = 0;
  function emit(str) {
    if (moved) out.push(' ');
    out.push(str);
    moved = false;
  }
  const n = s.length;
  const operands = [];

  function decode(bytesStr) {
    if (!font) return '';
    let r = '';
    const L = font.codeLen;
    for (let k = 0; k + L <= bytesStr.length; k += L) {
      let code = 0;
      for (let b = 0; b < L; b++) code = (code << 8) | bytesStr.charCodeAt(k + b);
      const u = font.map[code];
      r += (u !== undefined) ? u : (L === 1 ? String.fromCharCode(code) : '');
    }
    return r;
  }

  while (i < n) {
    const c = s[i];
    if (c === ' ' || c === '\n' || c === '\r' || c === '\t' || c === '\f' || c === '\0') { i++; continue; }
    if (c === '%') { while (i < n && s[i] !== '\n' && s[i] !== '\r') i++; continue; }
    if (c === '(') {
      let depth = 1, j = i + 1, str = '';
      while (j < n && depth > 0) {
        const ch = s[j];
        if (ch === '\\') {
          const nx = s[j + 1];
          if (nx >= '0' && nx <= '7') {
            let oct = '', k = j + 1;
            while (k < n && oct.length < 3 && s[k] >= '0' && s[k] <= '7') oct += s[k++];
            str += String.fromCharCode(parseInt(oct, 8));
            j = k;
            continue;
          }
          const esc = { n: '\n', r: '\r', t: '\t', b: '\b', f: '\f' }[nx];
          if (nx === '\r' || nx === '\n') { j += 2; if (nx === '\r' && s[j] === '\n') j++; continue; }
          str += esc !== undefined ? esc : nx;
          j += 2;
          continue;
        }
        if (ch === '(') depth++;
        else if (ch === ')') { depth--; if (depth === 0) break; }
        str += ch;
        j++;
      }
      operands.push({ t: 's', v: str });
      i = j + 1;
      continue;
    }
    if (c === '<' && s[i + 1] !== '<') {
      const j = s.indexOf('>', i);
      let hex = s.substring(i + 1, j).replace(/\s+/g, '');
      if (hex.length % 2) hex += '0';
      let str = '';
      for (let k = 0; k < hex.length; k += 2) str += String.fromCharCode(parseInt(hex.substr(k, 2), 16));
      operands.push({ t: 's', v: str });
      i = j + 1;
      continue;
    }
    if (c === '<' || c === '>') { i += 2; continue; } // << >>(インライン辞書)は無視
    if (c === '[') { operands.push({ t: '[' }); i++; continue; }
    if (c === ']') {
      const arr = [];
      while (operands.length && operands[operands.length - 1].t !== '[') arr.unshift(operands.pop());
      operands.pop();
      operands.push({ t: 'a', v: arr });
      i++;
      continue;
    }
    if (c === '/') {
      let j = i + 1;
      while (j < n && !/[\s\/\[\]<>()%{}]/.test(s[j])) j++;
      operands.push({ t: 'n', v: s.substring(i + 1, j) });
      i = j;
      continue;
    }
    // 数値または演算子
    let j = i;
    while (j < n && !/[\s\/\[\]<>()%{}]/.test(s[j])) j++;
    const tok = s.substring(i, j);
    i = j;
    if (/^[+\-]?(\d+\.?\d*|\.\d+)$/.test(tok)) { operands.push({ t: 'num', v: parseFloat(tok) }); continue; }

    // 演算子
    if (tok === 'Tf') {
      const nameOp = operands.filter(o => o.t === 'n').pop();
      font = nameOp ? (fonts[nameOp.v] || null) : null;
    } else if (tok === 'Td' || tok === 'TD' || tok === 'Tm' || tok === 'T*') {
      moved = true;
    } else if (tok === 'Tj' || tok === "'" || tok === '"') {
      if (tok !== 'Tj') moved = true;
      const sOp = operands.filter(o => o.t === 's').pop();
      if (sOp) emit(decode(sOp.v));
    } else if (tok === 'TJ') {
      const aOp = operands.filter(o => o.t === 'a').pop();
      if (aOp) {
        let piece = '';
        aOp.v.forEach(e => {
          if (e.t === 's') piece += decode(e.v);
          else if (e.t === 'num' && e.v < -300) piece += ' '; // 大きな字間調整は空白とみなす
        });
        emit(piece);
      }
    } else if (tok === 'ET') {
      out.push('\n');
      moved = true;
    } else if (tok === 'BI') {
      const ei = s.indexOf('EI', i);
      i = ei === -1 ? n : ei + 2;
    }
    operands.length = 0;
  }
  return out.join('').replace(/[ \t]*\n[ \t]*/g, '\n');
}

/** ページを文書順に並べる(/Root → /Pages の /Kids をたどる) */
function mtCollectPages_(objs) {
  const pages = [];
  const catalogKey = Object.keys(objs).find(k => /\/Type\s*\/Catalog/.test(objs[k].dict));
  if (!catalogKey) return pages;
  const visit = (dict, depth) => {
    if (!dict || depth > 20) return;
    if (/\/Type\s*\/Page(?![s])/.test(dict)) { pages.push(dict); return; }
    const kidsM = dict.match(/\/Kids\s*\[([^\]]*)\]/);
    if (!kidsM) return;
    const refRe = /(\d+)\s+0\s+R/g;
    let r;
    while ((r = refRe.exec(kidsM[1])) !== null) {
      const o = objs[parseInt(r[1], 10)];
      if (o) visit(o.dict, depth + 1);
    }
  };
  visit(mtDictValue_(objs, objs[catalogKey].dict, 'Pages'), 0);
  return pages;
}

/** mtall.pdfのバイト列から全ページのテキスト(描画順)を返す */
function mtExtractPdfText_(bytes) {
  const objs = mtReadPdfObjects_(bytes);
  const fontCache = {};
  const texts = [];
  mtCollectPages_(objs).forEach(pageDict => {
    const res = mtDictValue_(objs, pageDict, 'Resources') || '';
    const fontDict = mtDictValue_(objs, res, 'Font') || '';
    const fonts = {};
    const fRe = /\/([^\s\/<>\[\]()]+)\s+(\d+)\s+0\s+R/g;
    let f;
    while ((f = fRe.exec(fontDict)) !== null) {
      const fnum = parseInt(f[2], 10);
      if (!fontCache[fnum]) {
        const fo = objs[fnum];
        let cmap = { map: {}, codeLen: 1 };
        if (fo) {
          const tu = fo.dict.match(/\/ToUnicode\s+(\d+)\s+0\s+R/);
          const cmapObj = tu ? objs[parseInt(tu[1], 10)] : null;
          if (cmapObj && cmapObj.stream) cmap = mtParseCMap_(mtBytesToLatin1_(cmapObj.stream));
          else if (/Identity-H/.test(fo.dict)) cmap.codeLen = 2;
        }
        fontCache[fnum] = cmap;
      }
      fonts[f[1]] = fontCache[fnum];
    }
    const contentsM = pageDict.match(/\/Contents\s*(\[[^\]]*\]|\d+\s+0\s+R)/);
    if (!contentsM) return;
    const refs = contentsM[1].match(/\d+\s+0\s+R/g) || [];
    const parts = [];
    refs.forEach(ref => {
      const o = objs[parseInt(ref, 10)];
      if (o && o.stream) parts.push(mtBytesToLatin1_(o.stream));
    });
    const joined = parts.join('\n');
    const content = new Uint8Array(joined.length);
    for (let k = 0; k < joined.length; k++) content[k] = joined.charCodeAt(k);
    texts.push(mtExtractPageText_(content, fonts));
  });
  return texts.join('\n');
}

// ---- テキストから売残・買残を取り出す ----
function mtNum_(s) {
  if (s === undefined || s === null) return null;
  const neg = /▲/.test(s);
  const v = parseInt(String(s).replace(/[▲,\s]/g, ''), 10);
  return isNaN(v) ? null : (neg ? -v : v);
}

// 数値1個(3桁区切りとして正しい形のみ。▲付きもあり)
const MT_NUM_ = '(▲\\s*\\d{1,3}(?:,\\d{3})*|\\d{1,3}(?:,\\d{3})*)';
const MT_DIFF_ = '(?:▲\\s*\\d{1,3}(?:,\\d{3})*|\\d{1,3}(?:,\\d{3})*)';
const MT_RATIO_ = '(?:[\\d.]+%|\\*|-)';
// 「株数 Shs.」の後ろ: 売残 前日比 上場比 買残 前日比 上場比 … の並びが崩れていないことまで確認する
const MT_AFTER_SHS_ = '株数\\s*Shs\\.\\s+' + MT_NUM_ + '\\s+' + MT_DIFF_ + '\\s+' + MT_RATIO_ + '\\s+' + MT_NUM_ + '\\s+' + MT_DIFF_ + '\\s+' + MT_RATIO_ + '(?=\\s|$)';

/**
 * PDF上で1つの数値が2回に分けて描画され、間に空白が入ってしまうことがある(実ファイルで確認)。
 * 「145 ,500」「37,718,7 09」のように、単独では数値として成り立たない形だけを機械的につなぎ直す。
 * (「2 9,700」のように両方とも数値として成り立つものは区別できないため直さない。
 *  必要な列でそれが起きた場合は並びのチェックで不一致になり、取得失敗として扱われる)
 */
function mtRepairSplitNumbers_(text) {
  let t = text;
  for (let k = 0; k < 3; k++) {
    t = t.replace(/(\d) ,(\d{3})(?!\d)/g, '$1,$2');
    t = t.replace(/(\d,\d{1,2}) (\d{1,2})(?=[\s,]|$)/g, (all, a, b) => {
      const merged = a + b;
      return /,\d{3}$/.test(merged) ? merged : all;
    });
  }
  return t;
}

/**
 * 全ページテキストから {銘柄コード(4桁 or 285Aなど): {sell, buy}} と、
 * 総合計の市場別小計 {プライム:{sell,buy,count}, スタンダード:…, グロース:…} を取り出す。
 */
function mtParseMarginText_(text) {
  const t = mtRepairSplitNumbers_(text.replace(/　/g, ' '));
  const stocks = {};
  const stockRe = new RegExp('(?:^|\\s)([0-9][0-9A-Z]{3})0\\s+[A-Z]{2}[0-9A-Z]{10}(?:(?!Val\\.|株数)[^\\n]){0,40}?' + MT_AFTER_SHS_, 'g');
  let m;
  while ((m = stockRe.exec(t)) !== null) {
    if (!stocks[m[1]]) stocks[m[1]] = { sell: mtNum_(m[2]), buy: mtNum_(m[3]) };
  }

  const markets = {};
  const totalIdx = t.search(/総\s*合\s*計/);
  if (totalIdx !== -1) {
    const area = t.substring(totalIdx);
    [['プライム', 'プ\\s*ラ\\s*イ\\s*ム'], ['スタンダード', 'ス\\s*タ\\s*ン\\s*ダ\\s*ー\\s*ド'], ['グロース', 'グ\\s*ロ\\s*ー\\s*ス']].forEach(pair => {
      const re = new RegExp(pair[1] + '\\s*小\\s*計\\s+([\\d,]+)\\s*銘\\s*柄\\s+' + MT_AFTER_SHS_);
      const mm = area.match(re);
      if (mm) markets[pair[0]] = { count: mtNum_(mm[1]), sell: mtNum_(mm[2]), buy: mtNum_(mm[3]) };
    });
  }
  return { stocks: stocks, markets: markets };
}


// ==== 信用残(日次)更新 ====
// データ元: JPX「銘柄別信用取引残高」(https://www.jpx.co.jp/markets/statistics-equities/margin/01.html)
//   ・申込日ごとのPDF(YYYYMMDD_mtall.pdf)。2026年10月分から公表開始のため、それより前のデータはない
//   ・申込日の翌営業日に公表される。一覧ページには直近5日分程度しか載らないので、こまめに取得すること
// 書き込み先: 各月シートの39行目〜69行目(1日〜31日)。終値と同じく A列=日付・B列=曜日・休場マーク
//   市場ブロック(C:E・F:H・I:K)に [売残(百万株), 買残(百万株), 信用倍率]、各銘柄(3列)に [売残(千株), 買残(千株), 信用倍率]
const MARGIN_DATA_START_DATE = new Date(2026, 9, 2); // JPX日次信用残の公表開始(2026/10/2申込分から。月は0始まりなので9=10月)
const MARGIN_LIST_PAGE_URL = 'https://www.jpx.co.jp/markets/statistics-equities/margin/01.html';
const MARGIN_STOCK_UNIT_DIVISOR = 1000;     // 個別銘柄: 株 → 千株(整数)
const MARGIN_MARKET_UNIT_DIVISOR = 1000000; // プライム/スタンダード/グロース: 株 → 百万株
const MARGIN_SMALL_FORMAT = '#,##0.000'; // 市場の値が百万株にして1桁(10未満)になる場合は小数点第3位まで表示
const MARGIN_LARGE_FORMAT = '#,##0';
const MARGIN_AUTO_LOOKBACK_DAYS = 10; // 自動更新(トリガー)で遡って未入力を埋める日数
const MARGIN_WEEKDAY_NAMES_ = ['日曜日', '月曜日', '火曜日', '水曜日', '木曜日', '金曜日', '土曜日'];

function marginPdfUrl_(ymd) {
  return 'https://www.jpx.co.jp/markets/statistics-equities/margin/tvdivq0000001rnl-att/' + ymd + '_mtall.pdf';
}

/** 一覧ページから該当日のPDFリンクを探す(添付フォルダ名が変わった場合のフォールバック) */
function findMarginPdfUrlFromList_(ymd) {
  const response = UrlFetchApp.fetch(MARGIN_LIST_PAGE_URL, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) return null;
  const m = response.getContentText().match(new RegExp('href="([^"]*' + ymd + '_mtall\\.pdf)"'));
  if (!m) return null;
  return m[1].startsWith('http') ? m[1] : 'https://www.jpx.co.jp' + m[1];
}

function fetchPdfBytes_(url) {
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) return null;
  const bytes = new Uint8Array(response.getContent()); // Javaのbyte(-128〜127)は0〜255に変換される
  if (bytes.length < 5 || String.fromCharCode(bytes[0], bytes[1], bytes[2], bytes[3]) !== '%PDF') return null;
  return bytes;
}

/**
 * 指定日(申込日)の銘柄別信用取引残高を取得する。未公表・休場日はnull。
 * 戻り値: { stocks: {コード: [売残, 買残]}, markets: {プライム: [売残, 買残], ...} }(単位: 株)
 * 解析結果は6時間キャッシュする(新規銘柄追加時の補完などで同じ日を何度も読む場合の高速化)。
 */
function fetchJpxMarginForDate_(dateStr) {
  const ymd = dateStr.replace(/-/g, '');
  const cache = CacheService.getScriptCache();
  const cacheKey = 'mtall_' + ymd;
  try {
    const cached = cache.get(cacheKey);
    if (cached) {
      const blob = Utilities.newBlob(Utilities.base64Decode(cached), 'application/x-gzip');
      return JSON.parse(Utilities.ungzip(blob).getDataAsString());
    }
  } catch (e) {}

  let bytes = fetchPdfBytes_(marginPdfUrl_(ymd));
  if (!bytes) {
    const alt = findMarginPdfUrlFromList_(ymd);
    if (alt) bytes = fetchPdfBytes_(alt);
  }
  if (!bytes) return null;

  const parsed = mtParseMarginText_(mtExtractPdfText_(bytes));
  if (Object.keys(parsed.stocks).length === 0) return null;

  const result = { stocks: {}, markets: {} };
  Object.keys(parsed.stocks).forEach(code => { result.stocks[code] = [parsed.stocks[code].sell, parsed.stocks[code].buy]; });
  Object.keys(parsed.markets).forEach(name => { result.markets[name] = [parsed.markets[name].sell, parsed.markets[name].buy]; });

  try {
    const zipped = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(JSON.stringify(result))).getBytes());
    if (zipped.length < 95000) cache.put(cacheKey, zipped, 21600);
  } catch (e) {}
  return result;
}

/**
 * 株数 → シートに書く値。
 *   市場(プライム/スタンダード/グロース): 百万株。10未満(1桁)になる値は小数点第3位まで残し、それ以外は整数
 *     例) 256,314,200株 → 256 / 3,265,500株 → 3.266
 *   個別銘柄: 千株(整数)。例) 846,100株 → 846
 */
function marginUnitValue_(shares, isMarket) {
  if (!isMarket) return Math.round(shares / MARGIN_STOCK_UNIT_DIVISOR);
  const v = shares / MARGIN_MARKET_UNIT_DIVISOR;
  return Math.abs(v) < 10 ? Math.round(v * 1000) / 1000 : Math.round(v);
}

function marginNumberFormat_(value, isMarket) {
  return (isMarket && Math.abs(value) < 10) ? MARGIN_SMALL_FORMAT : MARGIN_LARGE_FORMAT;
}

/** [売残, 買残](株) → シートに書く [売残, 買残, 信用倍率](単位はmarginUnitValue_のとおり) */
function marginCellValues_(pair, isMarket) {
  const sell = pair[0], buy = pair[1];
  const ratio = (sell > 0) ? Math.round((buy / sell) * 100) / 100 : '-';
  return [marginUnitValue_(sell, isMarket), marginUnitValue_(buy, isMarket), ratio];
}

function isMarginBlank_(v) { return v === '' || v === null || v === undefined; }

/**
 * 指定日の信用残を当月(その日の月)シートの信用残行に書き込む。
 * onlyCode を指定すると、その銘柄の列だけを埋める(新規銘柄追加時の補完用)。
 * 戻り値: { message, status: 'done'|'holiday'|'skip'|'unpublished'|'nosheet' }
 */
function updateMarginForSpecificDate(targetDate, onlyCode) {
  const dateStr = Utilities.formatDate(targetDate, 'Asia/Tokyo', 'yyyy-MM-dd');
  if (targetDate < MARGIN_DATA_START_DATE) return { status: 'skip', message: dateStr + ': 2026/10/2より前は日次の信用残データがありません' };
  const sheet = getTargetSheet(targetDate);
  if (!sheet) return { status: 'nosheet', message: dateStr + ': 対応するシートが見つかりません' };

  const row = MARGIN_FIRST_ROW + (targetDate.getDate() - 1);
  const lastCol = sheet.getLastColumn();
  const codeRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const rowValues = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
  if (isMarginBlank_(rowValues[0])) rowValues[0] = targetDate.getDate() + '日';
  if (isMarginBlank_(rowValues[1])) rowValues[1] = MARGIN_WEEKDAY_NAMES_[targetDate.getDay()];

  const blocks = [];
  if (!onlyCode) MARKET_VALUE_ITEMS.forEach(item => blocks.push({ col: item.col, width: MARKET_VALUE_COL_WIDTH, market: item.name }));
  for (let col = STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) {
    const code = codeRow[col - 1];
    if (!code) continue;
    if (onlyCode && String(code) !== String(onlyCode)) continue;
    blocks.push({ col: col, width: STOCK_COL_WIDTH, code: String(code) });
  }

  const dow = targetDate.getDay();
  if ((dow === 0 || dow === 6) || isJapaneseHoliday(targetDate)) {
    blocks.forEach(b => {
      for (let k = 0; k < b.width; k++) rowValues[b.col - 1 + k] = HOLIDAY_LABEL;
      sheet.getRange(row, b.col, 1, b.width).setBackground(HOLIDAY_COLOR);
    });
    sheet.getRange(row, 1, 1, lastCol).setValues([rowValues]);
    return { status: 'holiday', message: dateStr + ': 休場日のため休マークを入力しました' };
  }

  const data = fetchJpxMarginForDate_(dateStr);
  if (!data) {
    sheet.getRange(row, 1, 1, 2).setValues([[rowValues[0], rowValues[1]]]);
    return { status: 'unpublished', message: dateStr + ': 未公表です(申込日の翌営業日に公表されます)' };
  }

  const notFound = [];
  let written = 0;
  const formats = sheet.getRange(row, 1, 1, lastCol).getNumberFormats()[0];
  blocks.forEach(b => {
    const pair = b.market ? data.markets[b.market] : data.stocks[b.code];
    if (!pair) {
      notFound.push(b.market || b.code);
      // 一覧にない銘柄(信用取引の対象外など)は「-」を入れて入力済み扱いにする(毎日取り直さないように)。
      // 市場の小計が読めなかった場合は空欄のままにし、次回の更新で取り直す。
      if (b.code) for (let k = 0; k < 3; k++) rowValues[b.col - 1 + k] = '-';
      return;
    }
    const isMarket = !!b.market;
    const vals = marginCellValues_(pair, isMarket);
    for (let k = 0; k < 3; k++) rowValues[b.col - 1 + k] = vals[k];
    // 売残・買残の表示形式(市場は1桁なら小数点第3位まで、個別銘柄は常に整数)
    formats[b.col - 1] = marginNumberFormat_(vals[0], isMarket);
    formats[b.col] = marginNumberFormat_(vals[1], isMarket);
    written++;
  });
  sheet.getRange(row, 1, 1, lastCol).setValues([rowValues]);
  sheet.getRange(row, 1, 1, lastCol).setNumberFormats([formats]);

  let message = dateStr + ': ' + written + '/' + blocks.length + '項目の信用残を更新しました';
  if (notFound.length > 0) message += ' / 一覧になかった項目(信用取引の対象外など): ' + notFound.join(', ');
  return { status: 'done', message: message };
}

/** その日の信用残行が(市場ブロック・全銘柄とも)すでに埋まっているか */
function isMarginRowFilled_(date) {
  const sheet = getTargetSheet(date);
  if (!sheet) return false;
  const row = MARGIN_FIRST_ROW + (date.getDate() - 1);
  const lastCol = sheet.getLastColumn();
  const codeRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const vals = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
  const cols = MARKET_VALUE_ITEMS.map(item => item.col);
  for (let col = STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) if (codeRow[col - 1]) cols.push(col);
  return cols.every(col => !isMarginBlank_(vals[col - 1]));
}

/** ダイアログからの単日・期間指定更新。入力済みの日はスキップし、時間切れ前に安全に止まる */
function updateMarginForDateRangeInput(startDateStr, endDateStr) {
  const rangeStartMs = Date.now();
  const start = parseDateStr_(startDateStr);
  const end = parseDateStr_(endDateStr || startDateStr);
  if (start > end) return '開始日は終了日より前の日付にしてください';

  const messages = [];
  for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
    if (Date.now() - rangeStartMs > RUN_TIME_LIMIT_MS) {
      messages.push('(実行時間の上限のため、' + Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd') + '以降は未処理です。同じ期間でもう一度実行すると続きから処理されます)');
      break;
    }
    const date = new Date(d);
    if (date >= MARGIN_DATA_START_DATE && isMarginRowFilled_(date)) {
      messages.push(Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd') + ': 入力済みのためスキップしました');
      continue;
    }
    messages.push(updateMarginForSpecificDate(date).message);
  }
  const message = messages.join('\n');
  Logger.log(message);
  return message;
}

/**
 * 時間主導トリガー用(毎日実行): 直近の未入力日(前日まで)の信用残を埋める。
 * 一覧ページには直近5日分程度しか載らないため、毎日動かして取りこぼしを防ぐ。
 */
function updateMarginData() {
  const startMs = Date.now();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const messages = [];
  for (let back = MARGIN_AUTO_LOOKBACK_DAYS; back >= 1; back--) {
    if (Date.now() - startMs > RUN_TIME_LIMIT_MS) { messages.push('(実行時間の上限のため途中で停止しました)'); break; }
    const date = new Date(today);
    date.setDate(today.getDate() - back);
    if (date < MARGIN_DATA_START_DATE) continue;
    if (isMarginRowFilled_(date)) continue;
    messages.push(updateMarginForSpecificDate(date).message);
  }
  const message = messages.length > 0 ? messages.join('\n') : '信用残: 未入力の日はありませんでした';
  Logger.log(message);
  return message;
}

/**
 * 新規銘柄追加時用: 当月の1日〜前日までの、その銘柄の信用残を埋める(入力済みの日・休場日はそのまま)。
 * 時間がかかる場合は途中で止まり、残りは毎日の自動更新または日付入力更新で埋まる。
 */
function fillMarginForCode(code) {
  const startMs = Date.now();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const sheet = getTargetSheet(today);
  if (!sheet) return '当月シートが見つかりません';
  const lastCol = sheet.getLastColumn();
  const codeRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  let targetCol = null;
  for (let col = STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) if (String(codeRow[col - 1]) === String(code)) { targetCol = col; break; }
  if (!targetCol) return '銘柄が見つかりません: ' + code;

  let filled = 0;
  for (let d = 1; d < today.getDate(); d++) {
    if (Date.now() - startMs > RUN_TIME_LIMIT_MS) return code + ': 信用残を' + filled + '日分入力しました(時間切れのため途中まで。残りは自動更新で埋まります)';
    const date = new Date(today.getFullYear(), today.getMonth(), d);
    if (date < MARGIN_DATA_START_DATE) continue;
    const current = sheet.getRange(MARGIN_FIRST_ROW + d - 1, targetCol).getValue();
    if (!isMarginBlank_(current)) continue;
    const r = updateMarginForSpecificDate(date, code);
    if (r.status === 'done') filled++;
  }
  return code + ': 信用残を' + filled + '日分入力しました';
}

/** 動作確認用: 指定日のPDFを読み、解析にかかった時間と一部の値をログに出す(シートには書かない) */
function debugJpxMarginForDate(dateStr) {
  // dateStr = '2026-10-08'; // ← エディタから直接実行する場合はここのコメントを外して日付を指定する
  if (!dateStr) { Logger.log('dateStrが指定されていません(引数なしで実行していませんか?)'); return; }
  CacheService.getScriptCache().remove('mtall_' + dateStr.replace(/-/g, ''));
  const t0 = Date.now();
  const data = fetchJpxMarginForDate_(dateStr);
  Logger.log('所要時間: ' + (Date.now() - t0) + 'ms');
  if (!data) { Logger.log('取得できませんでした(未公表・休場日・URL変更の可能性)'); return; }
  Logger.log('銘柄数: ' + Object.keys(data.stocks).length);
  Logger.log('市場別: ' + JSON.stringify(data.markets));
  Logger.log('3436: ' + JSON.stringify(data.stocks['3436']) + ' / 285A: ' + JSON.stringify(data.stocks['285A']));
}

// ==== データベース書き込み ====
const BACKFILL_PROGRESS_KEY = 'DB_BACKFILL_DONE_CODES';

function getOrCreateDatabaseSheet() {
  const ss = getMainSpreadsheet();
  let dbSheet = ss.getSheetByName(DB_SHEET_NAME);
  if (!dbSheet) {
    dbSheet = ss.insertSheet(DB_SHEET_NAME);
    dbSheet.getRange(1, 1).setValue('日付');
    dbSheet.setFrozenRows(1);
    dbSheet.setFrozenColumns(1);
    PropertiesService.getScriptProperties().deleteProperty(BACKFILL_PROGRESS_KEY);
  }
  return dbSheet;
}

function getOrCreateDatabaseColumn(dbSheet, code) {
  const lastCol = Math.max(dbSheet.getLastColumn(), 1);
  const headerRow = dbSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (let col = 2; col <= lastCol; col++) {
    if (String(headerRow[col - 1]) === String(code)) return col;
  }
  dbSheet.getRange(1, lastCol + 1).setValue(code);
  return lastCol + 1;
}

/**
 * DBシートの中から、beforeDateStr(例 '2026-10-01')より前で最も新しい日付の行を探し、
 * {コード: 終値} のマップを返す(前日比の計算に使う「前営業日の実終値」)。
 * 再実行などで当日分が既にDBに書き込まれていた場合は、その行を飛ばして1つ前の行を使う。
 */
function getPreviousCloseMap(dbSheet, beforeDateStr) {
  const lastRow = dbSheet.getLastRow();
  if (lastRow < 2) return {};
  const lastCol = Math.max(dbSheet.getLastColumn(), 1);
  const headerRow = dbSheet.getRange(1, 1, 1, lastCol).getValues()[0];

  for (let r = lastRow; r >= 2; r--) {
    const d = dbSheet.getRange(r, 1).getValue();
    if (d instanceof Date && Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd') < beforeDateStr) {
      const rowVals = dbSheet.getRange(r, 1, 1, lastCol).getValues()[0];
      const map = {};
      for (let c = 2; c <= lastCol; c++) {
        if (headerRow[c - 1] && typeof rowVals[c - 1] === 'number') map[String(headerRow[c - 1])] = rowVals[c - 1];
      }
      return map;
    }
  }
  return {};
}

function writeClosesToDatabaseForDate(date, priceMap) {
  const codes = Object.keys(priceMap);
  if (codes.length === 0) return;

  const dbSheet = getOrCreateDatabaseSheet();
  const dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');

  let lastCol = Math.max(dbSheet.getLastColumn(), 1);
  const headerRow = dbSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const colMap = {};
  headerRow.forEach((h, i) => { if (h) colMap[String(h)] = i + 1; });

  const newCodes = codes.filter(c => !colMap[c]);
  if (newCodes.length > 0) {
    dbSheet.getRange(1, lastCol + 1, 1, newCodes.length).setValues([newCodes]);
    newCodes.forEach((c, i) => { colMap[c] = lastCol + 1 + i; });
    lastCol += newCodes.length;
  }

  const lastRow = dbSheet.getLastRow();
  const dateColumn = lastRow >= 2 ? dbSheet.getRange(2, 1, lastRow - 1, 1).getValues() : [];
  let targetRow = -1;
  for (let i = 0; i < dateColumn.length; i++) {
    if (dateColumn[i][0] instanceof Date && Utilities.formatDate(dateColumn[i][0], 'Asia/Tokyo', 'yyyy-MM-dd') === dateStr) {
      targetRow = 2 + i; break;
    }
  }
  if (targetRow === -1) {
    targetRow = lastRow + 1;
    dbSheet.getRange(targetRow, 1).setValue(date);
  }

  const rowValues = dbSheet.getRange(targetRow, 1, 1, lastCol).getValues()[0];
  codes.forEach(code => { rowValues[colMap[code] - 1] = priceMap[code]; });
  dbSheet.getRange(targetRow, 1, 1, lastCol).setValues([rowValues]);
}

function sortDatabaseByDate(dbSheet) {
  const lastRow = dbSheet.getLastRow(), lastCol = dbSheet.getLastColumn();
  if (lastRow < 3 || lastCol < 1) return;

  const range = dbSheet.getRange(2, 1, lastRow - 1, lastCol);
  const data = range.getValues();
  data.sort((a, b) => {
    const ta = a[0] instanceof Date ? a[0].getTime() : -Infinity;
    const tb = b[0] instanceof Date ? b[0].getTime() : -Infinity;
    return ta - tb;
  });
  range.setValues(data);
}

function writeClosesToDatabaseBulk(code, dataMap) {
  const dbSheet = getOrCreateDatabaseSheet();
  const col = getOrCreateDatabaseColumn(dbSheet, code);
  const lastRow = dbSheet.getLastRow();
  const existingDates = lastRow >= 2 ? dbSheet.getRange(2, 1, lastRow - 1, 1).getValues() : [];

  const dateToRowMap = {};
  existingDates.forEach((row, idx) => {
    if (row[0] instanceof Date) dateToRowMap[Utilities.formatDate(row[0], 'Asia/Tokyo', 'yyyy-MM-dd')] = 2 + idx;
  });

  const sortedDateStrs = Object.keys(dataMap).sort();
  let nextNewRow = lastRow + 1;
  const newDateRows = [];

  sortedDateStrs.forEach(dateStr => {
    let row = dateToRowMap[dateStr];
    if (!row) {
      row = nextNewRow++;
      newDateRows.push({ row: row, date: new Date(dateStr) });
      dateToRowMap[dateStr] = row;
    }
  });

  if (newDateRows.length > 0) {
    dbSheet.getRange(newDateRows[0].row, 1, newDateRows.length, 1).setValues(newDateRows.map(r => [r.date]));
    sortDatabaseByDate(dbSheet);
  }

  const dbLastRowAfterSort = dbSheet.getLastRow();
  const refreshedDates = dbSheet.getRange(2, 1, dbLastRowAfterSort - 1, 1).getValues();
  const refreshedDateToRow = {};
  refreshedDates.forEach((row, idx) => {
    if (row[0] instanceof Date) refreshedDateToRow[Utilities.formatDate(row[0], 'Asia/Tokyo', 'yyyy-MM-dd')] = 2 + idx;
  });

  const rows = sortedDateStrs.map(ds => refreshedDateToRow[ds]).filter(r => r);
  if (rows.length === 0) return;
  const minRow = Math.min(...rows), maxRow = Math.max(...rows);
  const range = dbSheet.getRange(minRow, col, maxRow - minRow + 1, 1);
  const values = range.getValues();
  sortedDateStrs.forEach(ds => {
    const r = refreshedDateToRow[ds];
    if (r) values[r - minRow][0] = dataMap[ds].close;
  });
  range.setValues(values);
}

function backfillDatabaseForCode(code) {
  const today = new Date();
  const startDate = new Date(today);
  startDate.setFullYear(startDate.getFullYear() - DB_BACKFILL_YEARS);

  const result = fetchHistoricalCloses(code, startDate, today);
  const dataMap = result && result.dateMap;
  if (!dataMap || Object.keys(dataMap).length === 0) return false;
  writeClosesToDatabaseBulk(code, dataMap);
  return true;
}

function initializeDatabaseForAllStocks() {
  const startMs = Date.now();
  const sheet = getTargetSheet(new Date());
  if (!sheet) return '当月シートが見つかりません';
  const codeRow = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];

  const props = PropertiesService.getScriptProperties();
  const doneCodes = new Set(JSON.parse(props.getProperty(BACKFILL_PROGRESS_KEY) || '[]'));

  // プライム/スタンダード/グロースの売買代金ブロックは実在の銘柄ではなく、DBには日次更新で蓄積されるため対象外
  const targetCodes = [];
  for (let col = STOCK_START_COL; col <= sheet.getLastColumn(); col += STOCK_COL_WIDTH) {
    const code = codeRow[col - 1];
    if (code) targetCodes.push(String(code));
  }

  const failedCodes = [];
  let timedOut = false;
  for (const code of targetCodes) {
    if (doneCodes.has(code)) continue;
    if (Date.now() - startMs > RUN_TIME_LIMIT_MS) { timedOut = true; break; }

    const success = backfillDatabaseForCode(code);
    if (success) {
      doneCodes.add(code);
      props.setProperty(BACKFILL_PROGRESS_KEY, JSON.stringify(Array.from(doneCodes)));
    } else {
      failedCodes.push(code);
    }
    Utilities.sleep(300);
  }

  const doneCount = targetCodes.filter(c => doneCodes.has(c)).length;
  let msg;
  if (timedOut) {
    msg = 'データベース初期構築: 実行時間の上限に近づいたため途中で停止しました(完了 ' + doneCount + '/' + targetCodes.length + '銘柄)。もう一度実行すると続きから進みます。';
  } else if (failedCodes.length > 0) {
    // 失敗分だけ再試行できるよう、完了済みの記録は残す
    msg = 'データベース初期構築: ' + doneCount + '/' + targetCodes.length + '銘柄が完了。取得に失敗した銘柄(再実行で再試行されます): ' + failedCodes.join(', ');
  } else {
    props.deleteProperty(BACKFILL_PROGRESS_KEY);
    msg = '全銘柄のデータベース初期構築が完了しました(' + targetCodes.length + '銘柄)';
  }
  Logger.log(msg);
  return msg;
}

// ==== 書式保守 ====
function reapplyFormatFromTemplateToSheet(sheet, templateSheet) {
  const rawLastCol = sheet.getLastColumn(), codeRow = sheet.getRange(2, 1, 1, rawLastCol).getValues()[0];
  const maxRows = sheet.getMaxRows();
  // 書式コピー元: テンプレートの各ブロックの1日目の行(終値: 6行目 / 信用残: 39行目)。
  // 休場マーク(休)の行はグレーの背景を残すため対象外にし、それ以外の日の行にだけ1行分の書式を貼る。
  function applyRows(firstRow, col, width, srcRange) {
    if (firstRow + 30 > maxRows) return;
    const vals = sheet.getRange(firstRow, col, 31, width).getValues();
    let bs = -1;
    for (let r = 0; r <= vals.length; r++) {
      if (r < vals.length && vals[r][0] !== HOLIDAY_LABEL) { if (bs === -1) bs = r; }
      else if (bs !== -1) { srcRange.copyTo(sheet.getRange(firstRow + bs, col, r - bs, width), { formatOnly: true }); bs = -1; }
    }
  }

  const blocks = MARKET_VALUE_ITEMS.map(item => ({ col: item.col, width: MARKET_VALUE_COL_WIDTH, srcCol: item.col }));
  for (let col = STOCK_START_COL; col <= rawLastCol; col += STOCK_COL_WIDTH) {
    if (codeRow[col - 1]) blocks.push({ col: col, width: STOCK_COL_WIDTH, srcCol: STOCK_START_COL });
  }
  blocks.forEach(b => {
    applyRows(CLOSE_FIRST_ROW, b.col, b.width, templateSheet.getRange(CLOSE_FIRST_ROW, b.srcCol, 1, b.width));
    applyRows(MARGIN_FIRST_ROW, b.col, b.width, templateSheet.getRange(MARGIN_FIRST_ROW, b.srcCol, 1, b.width));
  });
}

function reapplyAllFormatsFromTemplate() {
  const ss = getMainSpreadsheet(), template = ss.getSheetByName(TEMPLATE_SHEET_NAME), currentSheet = getTargetSheet(new Date());
  if (!template) return 'テンプレートシートが見つかりません';
  if (currentSheet) reapplyFormatFromTemplateToSheet(currentSheet, template);
  return 'テンプレートの基準行(6行目・39行目)の書式を、当月シートの市場ブロック・全銘柄に再適用しました';
}

// ==== JPX投資部門別売買状況(売買状況シート) ====
// 2026年9月29日掲載分からJPXのファイル形式が変更された:
//   旧形式(〜2026/9/18掲載): 株数ファイルと金額ファイル(stock_val_1_*)が別、市場ごとに金額のPDFをテキスト解析
//   新形式(2026/9/29掲載〜): 1ファイル(stock_1_w_YYYYMMDD_YYYYMMDD)に株数・金額が統合、縦軸=市場横軸=投資部門別の
//     表をExcel(.xlsx)でも配布。PDFのテキスト変換は数値の欠落・順序崩れが実際に確認されたため、
//     新形式はExcel版をGoogleスプレッドシートに変換してセル単位で読み取る(曖昧さがない)。
// 本機能は「終値・前日比を更新(日付入力)」と同様、日付(または期間)を指定すると、その日が含まれる週を
// 自動的に判定して取得・解析・反映する。新形式/旧形式どちらの週も自動判別して処理する
// (「投資部門別売買状況を更新(最新週)」「...の過去分を取得」という旧来の2つの機能を統合したもの)。
const JPX_STATS_SHEET_NAME = '売買状況';
// アーカイブページ(00-00-archives-NN.html)は2026年10月時点でNN=10(2016年)まで開示されていることを確認済み。
// 将来さらに古い年が追加/削除される可能性があるため、2027年以降に実際の開示範囲を確認して必要なら調整する。
const JPX_ARCHIVE_MAX_YEARS_BACK = 10;
const JPX_MARKET_REFORM_DATE = new Date(2022, 3, 4); // 市場区分再編(プライム/スタンダード/グロース開始)。これより前の週は対象外

function convertPdfUrlToText(pdfUrl) {
  const pdfBlob = UrlFetchApp.fetch(pdfUrl, { muteHttpExceptions: true }).getBlob();
  const resource = { name: 'temp_jpx_pdf_convert', mimeType: MimeType.GOOGLE_DOCS };
  const file = Drive.Files.create(resource, pdfBlob, { convert: true });
  const doc = DocumentApp.openById(file.id);
  const text = doc.getBody().getText();
  DriveApp.getFileById(file.id).setTrashed(true);
  return text;
}

const JPX_MARKETS = [
  { key: 'プライム', label: '東証プライム' },
  { key: 'スタンダード', label: '東証スタンダード' },
  { key: 'グロース', label: '東証グロース' }
];

// シート列位置(共通): 総計/委託計/自己計 + 12業種分類
const JPX_ALL_COLS = ['総計', '委託計', '自己計'].map((k, i) => ({ key: k, col: 7 + i * 3 }))
  .concat(['法人', '個人', '海外投資家', '証券会社', '投資信託', '事業法人', 'その他法人等', '金融機関',
    '生保・損保', '都銀・地銀等', '信託銀行', 'その他金融機関'].map((k, i) => ({ key: k, col: 16 + i * 3 })));

function jpxToMillion(thousandYenValue) {
  if (thousandYenValue === null || thousandYenValue === undefined || isNaN(thousandYenValue)) return '';
  return Math.round(thousandYenValue / 1000);
}

// ---- 旧形式(〜2026/9/18掲載分。金額のみのPDF) ----
const JPX_HEADER_CATEGORIES_OLD_ = [
  { key: '自己計', label: '自己計' },
  { key: '委託計', label: '委託計' },
  { key: '総計', label: '総　計' }
];
const JPX_CATEGORIES_OLD_ = [
  { key: '法人', label: '法　人' },
  { key: '個人', label: '個　人' },
  { key: '海外投資家', label: '海外投資家' },
  { key: '証券会社', label: '証券会社' },
  { key: '投資信託', label: '投資信託' },
  { key: '事業法人', label: '事業法人' },
  { key: 'その他法人等', label: 'その他法人等' },
  { key: '金融機関', label: '金融機関' },
  { key: '生保・損保', label: '生保・損保' },
  { key: '都銀・地銀等', label: '都銀・地銀等' },
  { key: '信託銀行', label: '信託銀行' },
  { key: 'その他金融機関', label: 'その他金融機関' }
];

function isJpxAmountToken_(s) { return /^-?[\d,]+$/.test(s); }
function isJpxRatioTokenOld_(s) { return /^\d{1,3}\.\d$/.test(s); }

function parseJpxDirectionBlockOld_(blockText) {
  const tokens = (blockText.match(/-?[\d,]+\.?\d*/g) || []).filter(t => isJpxAmountToken_(t) || isJpxRatioTokenOld_(t));
  const ratioIdx = [];
  tokens.forEach((t, i) => { if (isJpxRatioTokenOld_(t)) ratioIdx.push(i); });
  if (ratioIdx.length < 2) return { amount: null, ratio: null, balance: null };

  const between = tokens.slice(ratioIdx[0] + 1, ratioIdx[1]);
  const after = tokens.slice(ratioIdx[1] + 1);
  const amount2 = between.length > 0 ? between[between.length - 1] : null;
  return {
    amount: amount2 !== null ? parseInt(amount2.replace(/,/g, ''), 10) : null,
    ratio: tokens[ratioIdx[1]] !== undefined ? parseFloat(tokens[ratioIdx[1]]) : null,
    balance: after.length > 0 ? parseInt(after[0].replace(/,/g, ''), 10) : null
  };
}

function parseJpxMarketSectionOld_(sectionText) {
  const result = {};
  const allCats = JPX_HEADER_CATEGORIES_OLD_.concat(JPX_CATEGORIES_OLD_);
  const labels = allCats.map(c => c.label);

  for (let i = 0; i < allCats.length; i++) {
    const cat = allCats[i];
    const startIdx = sectionText.indexOf(cat.label);
    if (startIdx === -1) { result[cat.key] = null; continue; }

    let endIdx = sectionText.length;
    for (let j = 0; j < labels.length; j++) {
      if (labels[j] === cat.label) continue;
      const p = sectionText.indexOf(labels[j], startIdx + cat.label.length);
      if (p !== -1 && p < endIdx) endIdx = p;
    }
    const block = sectionText.substring(startIdx, endIdx);

    const salesIdx = block.indexOf('Sales');
    const purchasesIdx = block.indexOf('Purchases');
    const totalIdx = block.indexOf('合計');
    if (salesIdx === -1 || purchasesIdx === -1 || totalIdx === -1) { result[cat.key] = null; continue; }

    const sellBlock = block.substring(salesIdx, purchasesIdx);
    const buyBlock = block.substring(purchasesIdx, totalIdx);
    const totalBlock = block.substring(totalIdx);

    const sell = parseJpxDirectionBlockOld_(sellBlock);
    const buy = parseJpxDirectionBlockOld_(buyBlock);
    const total = parseJpxDirectionBlockOld_(totalBlock);
    const balance = sell.balance !== null ? sell.balance : buy.balance;

    result[cat.key] = {
      sell: sell.amount, sellRatio: sell.ratio,
      buy: buy.amount, buyRatio: buy.ratio,
      total: total.amount, totalRatio: total.ratio,
      balance: balance
    };
  }
  return result;
}

function parseJpxPdfTextOld_(fullText) {
  const result = {};
  for (let i = 0; i < JPX_MARKETS.length; i++) {
    const m = JPX_MARKETS[i];
    const startIdx = fullText.indexOf(m.label);
    if (startIdx === -1) { result[m.key] = null; continue; }
    let endIdx = fullText.length;
    for (let j = 0; j < JPX_MARKETS.length; j++) {
      if (JPX_MARKETS[j].key === m.key) continue;
      const p = fullText.indexOf(JPX_MARKETS[j].label, startIdx + m.label.length);
      if (p !== -1 && p < endIdx) endIdx = p;
    }
    result[m.key] = parseJpxMarketSectionOld_(fullText.substring(startIdx, endIdx));
  }
  return result;
}

// ---- 新形式(2026/9/29掲載分〜。株数+金額が1ファイルに統合) ----
// PDFをテキスト化して読む方式は、実ファイルで「一部の数値がラベルより前に出現する」「稀に数値が
// 1つ抜け落ちる」といった崩れが確認されたため採用しない(気づかないままシートの値が化けるおそれがある)。
// 新形式はExcel版(.xlsx)も同時に公開されており、announcementのとおり「縦軸=市場、横軸=投資部門別」の
// 単純な表なので、Excel(Googleスプレッドシートに変換)をセル単位で読み取ることで曖昧さなく値を取得する。
// (PDF版のURLの拡張子を.xlsxに置き換えるだけで同じフォルダにExcel版が存在することを確認済み)

/** xlsxのURLをGoogleスプレッドシートに変換して開く(使い終わったらGoogleスプレッドシート側はtrashする) */
function convertXlsxUrlToSheet_(xlsxUrl) {
  const blob = UrlFetchApp.fetch(xlsxUrl, { muteHttpExceptions: true }).getBlob();
  const resource = { name: 'temp_jpx_xlsx_convert', mimeType: MimeType.GOOGLE_SHEETS };
  const file = Drive.Files.create(resource, blob, { convert: true });
  return { ss: SpreadsheetApp.openById(file.id), fileId: file.id };
}

// 列位置(1始まり)。シート構成はJPXの固定テンプレートで、サンプル3週分すべてで同一だったため固定値で扱う。
// 自己(現金取引・信用取引)/委託→個人(現金・信用)/委託→海外投資家(法人・個人)/証券会社/
// 委託→法人内訳(投資信託・事業法人・その他法人等・金融機関内訳(生保損保・都銀地銀等・信託銀行・その他金融機関))
// の順で、各項目は[売,買,差引,合計]の4列ぶんを占める。
const JPX_XLSX_COL_ = {
  jikoGenkin: 4, jikoShinyo: 8,                           // 自己計 = 4+8
  kojinGenkin: 12, kojinShinyo: 16,                       // 個人 = 12+16
  gaikokuHoujin: 20, gaikokuKojin: 24,                    // 海外投資家 = 20+24
  shoken: 28,
  toushin: 32, jigyo: 36, sonotaHoujin: 40,
  seiho: 44, toginChigin: 48, shintaku: 52, sonotaKinyu: 56
};
const JPX_XLSX_CATEGORY_COLS_ = {
  '自己計': [JPX_XLSX_COL_.jikoGenkin, JPX_XLSX_COL_.jikoShinyo],
  '委託計': [JPX_XLSX_COL_.kojinGenkin, JPX_XLSX_COL_.kojinShinyo, JPX_XLSX_COL_.gaikokuHoujin, JPX_XLSX_COL_.gaikokuKojin,
    JPX_XLSX_COL_.shoken, JPX_XLSX_COL_.toushin, JPX_XLSX_COL_.jigyo, JPX_XLSX_COL_.sonotaHoujin,
    JPX_XLSX_COL_.seiho, JPX_XLSX_COL_.toginChigin, JPX_XLSX_COL_.shintaku, JPX_XLSX_COL_.sonotaKinyu],
  '総計': [JPX_XLSX_COL_.jikoGenkin, JPX_XLSX_COL_.jikoShinyo, JPX_XLSX_COL_.kojinGenkin, JPX_XLSX_COL_.kojinShinyo,
    JPX_XLSX_COL_.gaikokuHoujin, JPX_XLSX_COL_.gaikokuKojin, JPX_XLSX_COL_.shoken,
    JPX_XLSX_COL_.toushin, JPX_XLSX_COL_.jigyo, JPX_XLSX_COL_.sonotaHoujin,
    JPX_XLSX_COL_.seiho, JPX_XLSX_COL_.toginChigin, JPX_XLSX_COL_.shintaku, JPX_XLSX_COL_.sonotaKinyu],
  '法人': [JPX_XLSX_COL_.toushin, JPX_XLSX_COL_.jigyo, JPX_XLSX_COL_.sonotaHoujin,
    JPX_XLSX_COL_.seiho, JPX_XLSX_COL_.toginChigin, JPX_XLSX_COL_.shintaku, JPX_XLSX_COL_.sonotaKinyu],
  '個人': [JPX_XLSX_COL_.kojinGenkin, JPX_XLSX_COL_.kojinShinyo],
  '海外投資家': [JPX_XLSX_COL_.gaikokuHoujin, JPX_XLSX_COL_.gaikokuKojin],
  '証券会社': [JPX_XLSX_COL_.shoken],
  '投資信託': [JPX_XLSX_COL_.toushin],
  '事業法人': [JPX_XLSX_COL_.jigyo],
  'その他法人等': [JPX_XLSX_COL_.sonotaHoujin],
  '金融機関': [JPX_XLSX_COL_.seiho, JPX_XLSX_COL_.toginChigin, JPX_XLSX_COL_.shintaku, JPX_XLSX_COL_.sonotaKinyu],
  '生保・損保': [JPX_XLSX_COL_.seiho],
  '都銀・地銀等': [JPX_XLSX_COL_.toginChigin],
  '信託銀行': [JPX_XLSX_COL_.shintaku],
  'その他金融機関': [JPX_XLSX_COL_.sonotaKinyu]
};
const JPX_XLSX_LAST_COL_ = 59; // BG列まで
// 「金額 Value」行の行番号(市場ごとに2行ずつ、奇数行=株数、偶数行=金額。サンプル3週とも同一レイアウトで確認済み)
const JPX_XLSX_VALUE_ROW_ = { 'プライム': 9, 'スタンダード': 11, 'グロース': 13 };

// 比率(%)はxlsxには収録されていないため、PDF脚注にある定義(自己計・委託計の比率は総計に対する構成比、
// それ以外の区分の比率はすべて委託計に対する構成比)どおりに金額から計算して復元する。
// (実際のPDF記載値と照合し、この計算式で一致することを確認済み)
function computeJpxRatios_(cats) {
  const total = cats['総計'];
  const itaku = cats['委託計'];
  const calcRatio = (v, base) => (base ? Math.round((v / base) * 10000) / 100 : null);
  Object.keys(cats).forEach(key => {
    const c = cats[key];
    if (!c) return;
    if (key === '総計') { c.sellRatio = 100; c.buyRatio = 100; return; }
    const base = (key === '自己計' || key === '委託計') ? total : itaku;
    c.sellRatio = base ? calcRatio(c.sell, base.sell) : null;
    c.buyRatio = base ? calcRatio(c.buy, base.buy) : null;
  });
}

function parseJpxXlsxSheet_(ss) {
  const sheet = ss.getSheets()[0];
  const result = {};
  JPX_MARKETS.forEach(market => {
    const row = JPX_XLSX_VALUE_ROW_[market.key];
    const rowValues = sheet.getRange(row, 1, 1, JPX_XLSX_LAST_COL_).getValues()[0];
    const cats = {};
    Object.keys(JPX_XLSX_CATEGORY_COLS_).forEach(catKey => {
      const sums = [0, 0, 0, 0]; // 売,買,差引,合計
      JPX_XLSX_CATEGORY_COLS_[catKey].forEach(startCol => {
        for (let i = 0; i < 4; i++) {
          const v = rowValues[startCol - 1 + i];
          sums[i] += (typeof v === 'number' ? v : 0);
        }
      });
      cats[catKey] = { sell: sums[0], buy: sums[1], balance: sums[2], total: sums[3], sellRatio: null, buyRatio: null, totalRatio: null };
    });
    computeJpxRatios_(cats);
    result[market.key] = cats;
  });
  return result;
}

// ---- 週の判定(ファイル名・アーカイブページから対象日が含まれる週を特定) ----
// アーカイブページ(年ごと、00-00-archives-NN.html / NN=今年からの年数)には新形式・旧形式どちらの週も
// 一覧として載っているため、ファイル名のパターンで新旧を自動判別する。
function getJpxWeekLinksForYear_(yearOffset) {
  const url = 'https://www.jpx.co.jp/markets/statistics-equities/investor-type/00-00-archives-' + String(yearOffset).padStart(2, '0') + '.html';
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) return [];
  const html = response.getContentText();

  const results = [];
  const weekRegex = /(\d{4})年(\d{1,2})月第(\d+)週\s*\((\d{1,2})月(\d{1,2})日\s*[～\-]\s*(\d{1,2})月(\d{1,2})日\)/g;
  let m;
  while ((m = weekRegex.exec(html)) !== null) {
    const year = parseInt(m[1], 10);
    const startDate = new Date(year, parseInt(m[4], 10) - 1, parseInt(m[5], 10));
    let endDate = new Date(year, parseInt(m[6], 10) - 1, parseInt(m[7], 10));
    if (endDate < startDate) endDate = new Date(year + 1, parseInt(m[6], 10) - 1, parseInt(m[7], 10)); // 年またぎ週の保険

    if (endDate < JPX_MARKET_REFORM_DATE) break; // これより古い週はページの並び上これ以降出てこない

    const searchArea = html.substring(m.index, m.index + 3000);
    const pdfMatch = searchArea.match(/href="((?:https:\/\/www\.jpx\.co\.jp)?\/markets\/statistics-equities\/investor-type\/[^"]+(stock_1_w_|stock_val_1_)[^"]+\.pdf)"/);
    if (!pdfMatch) continue;
    const path = pdfMatch[1];
    const pdfUrl = path.startsWith('http') ? path : 'https://www.jpx.co.jp' + path;
    const isNewFormat = pdfMatch[2] === 'stock_1_w_';
    results.push({
      weekKey: m[1] + '年' + m[2] + '月第' + m[3] + '週',
      startDate: startDate,
      endDate: endDate,
      pdfUrl: pdfUrl,
      // 新形式はExcel版を使う(PDF版と同じフォルダに同名で拡張子だけ.xlsxで存在する)
      xlsxUrl: isNewFormat ? pdfUrl.replace(/\.pdf$/, '.xlsx') : null,
      isNewFormat: isNewFormat
    });
  }
  return results;
}

/** 指定日が含まれる週の情報を返す(見つからなければnull)。今年から遡って年ごとのアーカイブページを調べる */
function resolveJpxWeekForDate_(targetDate) {
  if (!targetDate || isNaN(targetDate.getTime())) return null; // 無効な日付(引数なし実行など)
  if (targetDate < JPX_MARKET_REFORM_DATE) return null;
  const thisYear = new Date().getFullYear();
  const targetYear = targetDate.getFullYear();
  const yearOffset = thisYear - targetYear;
  if (yearOffset < 0 || yearOffset > JPX_ARCHIVE_MAX_YEARS_BACK) return null;

  const weeks = getJpxWeekLinksForYear_(yearOffset);
  for (let i = 0; i < weeks.length; i++) {
    if (targetDate >= weeks[i].startDate && targetDate <= weeks[i].endDate) return weeks[i];
  }
  return null;
}

function jpxWeekLabel_(weekInfo) {
  return Utilities.formatDate(weekInfo.startDate, 'Asia/Tokyo', 'yyyy/MM/dd') + '〜' + Utilities.formatDate(weekInfo.endDate, 'Asia/Tokyo', 'yyyy/MM/dd');
}

// ---- シートへの書き込み ----
function clearJpxSparklineMerges(sheet) {
  if (sheet.getLastRow() < 22) return;
  [5, 8, 11, 14, 17, 20].forEach(r => {
    const rng = sheet.getRange(r, 4, 3, 3);
    if (rng.isPartOfMerge()) rng.breakApart();
    rng.clearContent();
  });
}

function setupJpxBalanceSparklines(sheet) {
  const totalCol = JPX_ALL_COLS.find(c => c.key === '総計').col;
  JPX_MARKETS.forEach((market, mi) => {
    const dataBaseRow = 5 + mi * 3;
    const chartBaseRow = 14 + mi * 3;
    const balanceRow = dataBaseRow + 2;
    const lastRow = sheet.getLastRow();

    const values = [];
    for (let r = balanceRow; r <= lastRow && values.length < 24; r += 9) {
      const v = sheet.getRange(r, totalCol).getValue();
      if (typeof v === 'number') values.push(v);
    }
    values.reverse();

    const range = sheet.getRange(chartBaseRow, 4, 3, 3);
    range.merge();
    if (values.length > 0) {
      const maxAbs = Math.max.apply(null, values.map(Math.abs).concat([1]));
      range.setFormula('=SPARKLINE({' + values.join(',') + '}, {"charttype","column";"color","#ea4335";"negcolor","#4285f4";"ymin",-' + maxAbs + ';"ymax",' + maxAbs + '})');
    }
  });
}

function setupJpxTotalLineChart(sheet) {
  const totalCol = JPX_ALL_COLS.find(c => c.key === '総計').col;
  JPX_MARKETS.forEach((market, mi) => {
    const baseRow = 5 + mi * 3;
    const lastRow = sheet.getLastRow();

    const values = [];
    for (let r = baseRow; r <= lastRow && values.length < 24; r += 9) {
      const v = sheet.getRange(r, totalCol).getValue();
      if (typeof v === 'number') values.push(v);
    }
    values.reverse();

    const range = sheet.getRange(baseRow, 4, 3, 3);
    range.merge();
    if (values.length > 0) {
      range.setFormula('=SPARKLINE({' + values.join(',') + '}, {"charttype","line";"color","#4285f4"})');
    }
  });
}

/**
 * A列の週ラベルを9行マージの1セルに「年(改行)開始日〜 終了日」の形で書き込む。
 *   例) A5:A13 = 「2026
 *                  09/24〜 09/25」
 * 改行を表示するため折り返しを有効にし、日付への自動変換を防ぐため書式なしテキスト(@)にしてから書く。
 */
function jpxWeekLabelText_(startDate, endDate) {
  return Utilities.formatDate(startDate, 'Asia/Tokyo', 'yyyy') + '\n' +
    Utilities.formatDate(startDate, 'Asia/Tokyo', 'MM/dd') + '〜 ' +
    Utilities.formatDate(endDate, 'Asia/Tokyo', 'MM/dd');
}

function writeJpxWeekLabelCells_(sheet, baseRow, startDate, endDate) {
  const rng = sheet.getRange(baseRow, 1, 9, 1);
  rng.breakApart();
  rng.setNumberFormat('@');
  rng.setValues([[jpxWeekLabelText_(startDate, endDate)], [''], [''], [''], [''], [''], [''], [''], ['']]);
  rng.merge();
  rng.setWrap(true);
}

/** baseRowから9行ぶん(市場3つ×3行)に週ラベルとparsedDataを書き込む(挿入・上書き共通処理) */
function writeJpxBlockAtRow_(sheet, baseRow, startDate, endDate, parsedData) {
  writeJpxWeekLabelCells_(sheet, baseRow, startDate, endDate);

  JPX_MARKETS.forEach((market, mi) => {
    const r = baseRow + mi * 3;
    const data = parsedData[market.key];

    const marketCell = sheet.getRange(r, 2, 3, 1);
    if (marketCell.isPartOfMerge()) marketCell.breakApart();
    marketCell.merge().setValue(market.key);
    sheet.getRange(r, 3).setValue('金額');
    sheet.getRange(r + 1, 3).setValue('比率');
    sheet.getRange(r + 2, 3).setValue('差引き');
    if (!data) return;

    JPX_ALL_COLS.forEach(c => {
      const d = data[c.key];
      if (!d) return;
      sheet.getRange(r, c.col).setValue(jpxToMillion(d.sell));
      sheet.getRange(r, c.col + 1).setValue(jpxToMillion(d.buy));
      sheet.getRange(r, c.col + 2).setValue(jpxToMillion(d.total));
      sheet.getRange(r + 1, c.col).setValue(d.sellRatio);
      sheet.getRange(r + 1, c.col + 1).setValue(d.buyRatio);
      sheet.getRange(r + 1, c.col + 2).setValue(d.totalRatio);
      const balCell = sheet.getRange(r + 2, c.col, 1, 3);
      if (balCell.isPartOfMerge()) balCell.breakApart();
      balCell.merge().setValue(jpxToMillion(d.balance));
    });
  });
}

/**
 * 売買状況シートの各週ブロック(5行目から9行ごと)の先頭行と、その週の開始日・終了日を上から順に返す。
 * A列の週ラベルは次の形式が混在しうる(①が現行、②〜④は一括変換前の古い形式):
 *   ① 現行: A(r)=「2026(改行)09/28〜 10/02」(9行マージ)
 *   ② 3段形式: A(r)=「2026」、A(r+3)=「09/28〜」、A(r+6)=「10/02」(各3行マージ)
 *   ③ A(r)=「2026/09/28〜2026/10/02」(9行マージ)
 *   ④ 旧実装: A(r)=「09/28〜10/02」(年なし、9行マージ)
 * ④の年は、シートが新しい順に並んでいる前提で、直前(上)のブロックより後の日付にならないよう補う。
 */
function listJpxBlocks_(sheet) {
  const blocks = [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 13) return blocks;
  const labels = sheet.getRange(5, 1, lastRow - 4, 1).getDisplayValues();
  const cell = (row) => String((labels[row - 5] || [''])[0] || '').trim().replace(/\r\n?/g, '\n');
  const today = new Date();
  const md = (year, mon, day) => new Date(year, mon - 1, day);
  // 終了日の年: 開始日と同じ年、終了日が開始日より前なら翌年(年またぎの週)
  const endFrom = (start, mon, day) => {
    let e = md(start.getFullYear(), mon, day);
    if (e < start) e = md(start.getFullYear() + 1, mon, day);
    return e;
  };
  let prevDate = null;

  for (let r = 5; r + 8 <= lastRow; r += 9) {
    const top = cell(r);
    let startDate = null, endDate = null, m;

    if ((m = top.match(/^(\d{4})\s*\n\s*(\d{1,2})\/(\d{1,2})\s*〜\s*(\d{1,2})\/(\d{1,2})/))) {
      // ① 現行レイアウト
      startDate = md(parseInt(m[1], 10), parseInt(m[2], 10), parseInt(m[3], 10));
      endDate = endFrom(startDate, parseInt(m[4], 10), parseInt(m[5], 10));
    } else if (/^\d{4}$/.test(top) && (m = cell(r + 3).match(/^(\d{1,2})\/(\d{1,2})/))) {
      // ② 3段形式
      startDate = md(parseInt(top, 10), parseInt(m[1], 10), parseInt(m[2], 10));
      const e = cell(r + 6).match(/^(\d{1,2})\/(\d{1,2})/);
      if (e) endDate = endFrom(startDate, parseInt(e[1], 10), parseInt(e[2], 10));
    } else if ((m = top.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})(?:\s*〜\s*(\d{4})\/(\d{1,2})\/(\d{1,2}))?/))) {
      // ③ yyyy/MM/dd〜yyyy/MM/dd
      startDate = md(parseInt(m[1], 10), parseInt(m[2], 10), parseInt(m[3], 10));
      if (m[4]) endDate = md(parseInt(m[4], 10), parseInt(m[5], 10), parseInt(m[6], 10));
    } else if ((m = top.match(/^(\d{1,2})\/(\d{1,2})(?:\s*〜\s*(\d{1,2})\/(\d{1,2}))?/))) {
      // ④ 年なし MM/dd〜MM/dd
      const mon = parseInt(m[1], 10), day = parseInt(m[2], 10);
      const year = prevDate ? prevDate.getFullYear() : today.getFullYear();
      startDate = md(year, mon, day);
      if (prevDate ? startDate >= prevDate : startDate > today) startDate = md(year - 1, mon, day);
      if (m[3]) endDate = endFrom(startDate, parseInt(m[3], 10), parseInt(m[4], 10));
    }

    if (startDate && !endDate) { endDate = new Date(startDate); endDate.setDate(endDate.getDate() + 4); }
    blocks.push({ row: r, startDate: startDate, endDate: endDate });
    if (startDate) prevDate = startDate;
  }
  return blocks;
}

function jpxDateKey_(d) {
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
}

/**
 * 週データをシートへ反映する。シートは常に「新しい週が上」の順に保つ。
 * ・同じ開始日の週ブロックが既にあれば、その場で上書き(旧形式の年なしラベルも新形式ラベルに置き換わる)
 * ・なければ、自分より古い最初のブロックの直前に挿入(=日付順の正しい位置。最新週なら先頭)
 * ・既存のどのブロックよりも古ければ末尾に追加
 */
/** 売買状況シートのスパークライン(総計の推移・差引き)を描き直す。一連の書き込みの最後に1回だけ呼ぶ */
function redrawJpxCharts_() {
  const sheet = getMainSpreadsheet().getSheetByName(JPX_STATS_SHEET_NAME);
  if (!sheet) return;
  clearJpxSparklineMerges(sheet);
  setupJpxTotalLineChart(sheet);
  setupJpxBalanceSparklines(sheet);
}

function upsertJpxWeekToSheet_(parsedData, weekStartDate, weekEndDate) {
  const ss = getMainSpreadsheet();
  const sheet = ss.getSheetByName(JPX_STATS_SHEET_NAME);
  if (!sheet) { Logger.log('売買状況シートが見つかりません'); return; }

  const targetKey = jpxDateKey_(weekStartDate);
  const blocks = listJpxBlocks_(sheet);

  const same = blocks.find(b => b.startDate && jpxDateKey_(b.startDate) === targetKey);
  if (same) {
    writeJpxBlockAtRow_(sheet, same.row, weekStartDate, weekEndDate, parsedData);
    return;
  }

  const older = blocks.find(b => b.startDate && b.startDate < weekStartDate);
  clearJpxSparklineMerges(sheet);

  if (older) {
    // olderの直前に9行挿入し、押し下げられたolderブロックの書式をコピー
    const insertRow = older.row;
    sheet.insertRowsBefore(insertRow, 9);
    for (let i = 0; i < 9; i++) {
      sheet.getRange(insertRow + 9 + i, 1, 1, 1).copyTo(sheet.getRange(insertRow + i, 1, 1, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      sheet.getRange(insertRow + 9 + i, 4, 1, 48).copyTo(sheet.getRange(insertRow + i, 4, 1, 48), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    }
    writeJpxBlockAtRow_(sheet, insertRow, weekStartDate, weekEndDate, parsedData);
  } else {
    const baseRow = blocks.length > 0 ? blocks[blocks.length - 1].row + 9 : 5;
    if (blocks.length > 0) {
      for (let i = 0; i < 9; i++) {
        sheet.getRange(5 + i, 1, 1, 1).copyTo(sheet.getRange(baseRow + i, 1, 1, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
        sheet.getRange(5 + i, 4, 1, 48).copyTo(sheet.getRange(baseRow + i, 4, 1, 48), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      }
    }
    writeJpxBlockAtRow_(sheet, baseRow, weekStartDate, weekEndDate, parsedData);
  }
  // グラフ(スパークライン)の再描画は、呼び出し元で全週の書き込みが終わった最後に1回だけ行う
}

/**
 * [一回だけ実行] 売買状況シートの既存の週ラベル(A列)を、すべて新レイアウト
 * (9行マージ1セルに「年(改行)開始日〜 終了日」)に一括変換する。数値データには触れない。
 * 既に新レイアウトのブロックはスキップするので、途中で止まっても再実行すれば続きから進む。
 */
function migrateJpxStatsDateLayout() {
  const startMs = Date.now();
  const sheet = getMainSpreadsheet().getSheetByName(JPX_STATS_SHEET_NAME);
  if (!sheet) return '売買状況シートが見つかりません';
  const blocks = listJpxBlocks_(sheet);
  const labels = sheet.getRange(5, 1, Math.max(sheet.getLastRow() - 4, 1), 1).getDisplayValues();

  let converted = 0, skipped = 0, unknown = 0;
  for (let i = 0; i < blocks.length; i++) {
    const b = blocks[i];
    if (Date.now() - startMs > RUN_TIME_LIMIT_MS) {
      return converted + '週分を変換しました。実行時間の上限のため途中で停止しました。もう一度実行すると続きから進みます。';
    }
    if (!b.startDate) { unknown++; continue; }
    const top = String(labels[b.row - 5][0] || '').trim().replace(/\r\n?/g, '\n');
    if (top === jpxWeekLabelText_(b.startDate, b.endDate)) { skipped++; continue; } // 既に新レイアウト
    writeJpxWeekLabelCells_(sheet, b.row, b.startDate, b.endDate);
    converted++;
  }
  let msg = converted + '週分の日付表示を新レイアウトに変換しました';
  if (skipped > 0) msg += '(変換済み' + skipped + '週はスキップ)';
  if (unknown > 0) msg += ' / 日付を読み取れなかったブロック: ' + unknown + '件';
  return msg;
}

/**
 * [一回だけ実行] 既に下に追加されてしまった週ブロックも含め、売買状況シート全体を
 * 「新しい週が上」の順に並べ直す。各ブロックの値・書式(9行分)をまとめて入れ替える。
 */
function sortJpxStatsSheetByDate() {
  const ss = getMainSpreadsheet();
  const sheet = ss.getSheetByName(JPX_STATS_SHEET_NAME);
  if (!sheet) return '売買状況シートが見つかりません';
  const blocks = listJpxBlocks_(sheet);
  if (blocks.length < 2) return '並べ替える必要はありません';

  const sorted = blocks.slice().sort((a, b) => {
    const ta = a.startDate ? a.startDate.getTime() : 0;
    const tb = b.startDate ? b.startDate.getTime() : 0;
    return tb - ta;
  });
  if (sorted.every((b, i) => b.row === blocks[i].row)) return '既に新しい順に並んでいます';

  clearJpxSparklineMerges(sheet); // 上部のスパークライン(結合セル)が他の週に混ざらないよう先に消す
  const lastCol = sheet.getLastColumn();
  const totalRows = blocks.length * 9;

  // 作業用シートに並べ替え後の順で値・書式・結合ごとコピーし、元の位置へ書き戻す
  const tmpName = 'tmp_jpx_sort_' + Date.now();
  const tmp = ss.insertSheet(tmpName);
  try {
    sorted.forEach((b, i) => {
      sheet.getRange(b.row, 1, 9, lastCol).copyTo(tmp.getRange(1 + i * 9, 1, 9, lastCol));
    });
    const dest = sheet.getRange(5, 1, totalRows, lastCol);
    dest.breakApart();
    tmp.getRange(1, 1, totalRows, lastCol).copyTo(dest);
  } finally {
    ss.deleteSheet(tmp);
  }

  setupJpxTotalLineChart(sheet);
  setupJpxBalanceSparklines(sheet);
  return blocks.length + '週分を新しい順に並べ替えました';
}

/** weekInfoを取得・解析する(新形式はExcel、旧形式はPDFを使う) */
function fetchAndParseJpxWeek_(weekInfo) {
  if (weekInfo.isNewFormat) {
    const converted = convertXlsxUrlToSheet_(weekInfo.xlsxUrl);
    try {
      return parseJpxXlsxSheet_(converted.ss);
    } finally {
      DriveApp.getFileById(converted.fileId).setTrashed(true);
    }
  }
  const text = convertPdfUrlToText(weekInfo.pdfUrl);
  return parseJpxPdfTextOld_(text);
}

function updateJpxStatsForWeek_(weekInfo) {
  const weekLabel = jpxWeekLabel_(weekInfo);
  try {
    const parsed = fetchAndParseJpxWeek_(weekInfo);
    upsertJpxWeekToSheet_(parsed, weekInfo.startDate, weekInfo.endDate);
    return weekLabel + ': 投資部門別売買状況を更新しました';
  } catch (e) {
    return weekLabel + ': 更新に失敗しました。詳細: ' + e.message;
  }
}

/**
 * ダイアログからの単日/期間指定更新。指定日(期間)が含まれる週をすべて判定し、
 * 既にシートに反映済みの週は(forceRefetchがtrueでない限り)再取得せずスキップする。
 * 実行時間の上限が近づいたら安全に止まり、同じ期間で再実行すれば続きの週から処理される。
 */
function updateJpxStatsForDateRangeInput(startDateStr, endDateStr, forceRefetch) {
  const rangeStartMs = Date.now();
  const start = parseDateStr_(startDateStr);
  const end = parseDateStr_(endDateStr || startDateStr);
  if (start > end) return '開始日は終了日より前の日付にしてください';

  const ss = getMainSpreadsheet();
  const sheet = ss.getSheetByName(JPX_STATS_SHEET_NAME);
  const existingWeeks = {}; // 週の開始日(yyyy-MM-dd)で判定(旧形式の年なしラベルにも対応)
  if (sheet) {
    listJpxBlocks_(sheet).forEach(b => { if (b.startDate) existingWeeks[jpxDateKey_(b.startDate)] = true; });
  }

  const messages = [];
  const processedWeekKeys = {};
  let updatedCount = 0;
  for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
    if (Date.now() - rangeStartMs > RUN_TIME_LIMIT_MS) {
      messages.push('(実行時間の上限のため、' + Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd') + '以降は未処理です。同じ期間でもう一度実行すると続きから処理されます)');
      break;
    }
    const weekInfo = resolveJpxWeekForDate_(new Date(d));
    if (!weekInfo) continue; // 市場再編より前、または該当PDFが見つからない日
    if (processedWeekKeys[weekInfo.weekKey]) continue;
    processedWeekKeys[weekInfo.weekKey] = true;

    const weekLabel = jpxWeekLabel_(weekInfo);
    const weekKey = jpxDateKey_(weekInfo.startDate);
    if (!forceRefetch && existingWeeks[weekKey]) continue; // 取得済みはスキップ

    messages.push(updateJpxStatsForWeek_(weekInfo));
    existingWeeks[weekKey] = true;
    updatedCount++;
    Utilities.sleep(300);
  }
  // グラフの描画は全週の書き込みが終わった最後に1回だけ(週ごとに描くと遅くなるため)
  if (updatedCount > 0) {
    try { redrawJpxCharts_(); } catch (e) { messages.push('グラフの描画に失敗しました: ' + e.message); }
  }
  const message = messages.length > 0 ? messages.join('\n') : '対象期間に新規取得が必要な週はありませんでした(既に取得済みです)';
  Logger.log(message);
  return message;
}

/**
 * 指定した年(例: 2026)の、投資部門別売買状況が存在する週の一覧を新しい順で返す。
 * JpxStatsDialogの年・週プルダウンから呼ばれる(ダイアログ側は日付を直接入力しない)。
 * 戻り値はJSONでやり取りできるよう日付は文字列(yyyy-MM-dd)にしてある。
 */
function listJpxWeeksForYear(year) {
  const thisYear = new Date().getFullYear();
  const yearOffset = thisYear - year;
  if (yearOffset < 0 || yearOffset > JPX_ARCHIVE_MAX_YEARS_BACK) return [];

  return getJpxWeekLinksForYear_(yearOffset)
    .filter(w => w.endDate >= JPX_MARKET_REFORM_DATE)
    .map(w => ({
      label: w.weekKey + '(' + Utilities.formatDate(w.startDate, 'Asia/Tokyo', 'MM/dd') + '〜' +
        Utilities.formatDate(w.endDate, 'Asia/Tokyo', 'MM/dd') + ')',
      startDateStr: Utilities.formatDate(w.startDate, 'Asia/Tokyo', 'yyyy-MM-dd')
    }));
}

/** 時間主導トリガー用(引数なし): 本日が含まれる週(=直近週)を取得する。既に取得済みなら何もしない */
function updateJpxStatsForLatestWeek() {
  const todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const message = updateJpxStatsForDateRangeInput(todayStr, todayStr, false);
  Logger.log(message);
  return message;
}

/**
 * 動作確認用: 指定日が含まれる週のURL・形式判定・解析結果をログ出力する(シートへの書き込みは行わない)。
 * エディタの「実行」ボタンから引数なしで直接実行すると、dateStrがundefinedのまま呼ばれてしまい
 * 必ず「該当する週が見つかりませんでした」になるので注意(関数を選んでから実行引数を渡すか、
 * 下の1行のコメントを外して固定の日付でテストすること)。
 */
function debugJpxWeekForDate(dateStr) {
  // dateStr = '2026-10-09'; // ← エディタから直接実行する場合はここのコメントを外して日付を指定する
  if (!dateStr) { Logger.log('dateStrが指定されていません(引数なしで実行していませんか?)'); return; }
  const weekInfo = resolveJpxWeekForDate_(parseDateStr_(dateStr));
  if (!weekInfo) { Logger.log('該当する週が見つかりませんでした'); return; }
  Logger.log('週: ' + weekInfo.weekKey + ' / 形式: ' + (weekInfo.isNewFormat ? '新形式(Excel)' : '旧形式(PDF)') +
    ' / URL: ' + (weekInfo.isNewFormat ? weekInfo.xlsxUrl : weekInfo.pdfUrl));
  Logger.log(JSON.stringify(fetchAndParseJpxWeek_(weekInfo), null, 2));
}

// ==== 空売り集計シート ====
// データ元: https://www.jpx.co.jp/markets/statistics-equities/short-selling/index.html
//   ・空売り集計PDF(市場全体)  : .../short-selling/【添付ID】/YYMMDD-m.pdf
//   ・業種別集計PDF(業種別内訳): .../short-selling/【添付ID】/YYMMDD-g.pdf
// どちらも「実注文/空売り(価格規制あり)/空売り(価格規制なし)/合計」の4区分 × 売買代金・比率で構成される。
const SHORT_SELL_SHEET_NAME = '空売り集計';
const SHORT_SELL_PDF_ATT_ID = 't13vrt0000026iy6-att'; // JPXが添付IDを変更した場合はindex.htmlから探すフォールバックで対応
const SHORT_SELL_CATEGORIES = ['実注文', '空売り（価格規制あり）', '空売り（価格規制なし）', '合計'];

// 業種別集計PDF内での検索に使うラベル(表記ゆれがある項目はsearchで別指定)。シート上の列順もこの並びに従う。
const SHORT_SELL_INDUSTRIES = [
  { label: '水産・農林業' }, { label: '鉱業' }, { label: '建設業' }, { label: '食料品' },
  { label: '繊維製品' }, { label: 'パルプ・紙' }, { label: '化学' }, { label: '医薬品' },
  { label: '石油・石炭製品' }, { label: 'ゴム製品' }, { label: 'ガラス・土石製品' }, { label: '鉄鋼' },
  { label: '非鉄金属' }, { label: '金属製品' }, { label: '機械' }, { label: '電気機器' },
  { label: '輸送用機器' }, { label: '精密機器' }, { label: 'その他製品' }, { label: '電気・ガス業' },
  { label: '陸運業' }, { label: '海運業' }, { label: '空運業' }, { label: '倉庫・運輸関連業' },
  { label: '情報・通信業' }, { label: '卸売業' }, { label: '小売業' }, { label: '銀行業' },
  { label: '証券、商品先物取引業', search: '商品先物取引業' },
  { label: '保険業' }, { label: 'その他金融業' }, { label: '不動産業' }, { label: 'サービス業' },
  { label: 'その他（33業種外）', search: '33業種外' }
];

// シート上の列位置(C列=3起点、1項目=2列幅)。C:D=日次集計(市場全体)、E:F以降=業種別集計。
const SHORT_SELL_COLUMN_BLOCKS = [{ label: '日次集計', col: 3 }].concat(
  SHORT_SELL_INDUSTRIES.map((ind, i) => ({ label: ind.label, col: 5 + i * 2 }))
);

// ==== 銘柄の業種情報(当月シート等2行目J:K)と空売り集計シートの業種別ブロックとの連携 ====

/**
 * 銘柄の業種セルは2列幅(J:K)に収める都合上、長い業種名は省略形で保存している。
 * フル名→保存用の省略形。
 */
const INDUSTRY_STORAGE_ABBREVIATIONS = {
  'ガラス・土石製品': 'ガラス・土石',
  '倉庫・運輸関連業': '倉庫・運輸',
  '証券、商品先物取引業': '証券、商品先物',
  'その他（33業種外）': 'その他'
};

/** 保存用の省略形→空売り集計シートの業種ブロックが使っているフル名(上の逆引き) */
const INDUSTRY_LABEL_ALIASES = Object.keys(INDUSTRY_STORAGE_ABBREVIATIONS).reduce((map, fullLabel) => {
  map[INDUSTRY_STORAGE_ABBREVIATIONS[fullLabel]] = fullLabel;
  return map;
}, {});

/**
 * HTML本文から業種名を抜き出す。
 *
 * 主な方式: Yahoo!ファイナンスの銘柄ページでは、銘柄コードのすぐ右に
 * 業種名が "<a href="https://finance.yahoo.co.jp/search/qi/?ids=XXXX">電気機器</a>"
 * のような、業種検索ページへのリンクとして(括弧なしの)テキストで表示されている。
 * この "search/qi/?ids=" を目印にリンクの表示テキストを取り出す。
 *
 * フォールバック: 上記リンクが見つからないページ構造の場合に備えて、
 * 旧来の括弧([]/【】)ベースの探索も残す。ただしページ冒頭(canonicalタグや
 * meta descriptionなど)には "...(株)【285A】の株価..." のように銘柄コード自体が
 * 全角括弧で囲まれて繰り返し出現するため、①コード自体に一致する括弧内容、
 * ②英数字のみ(ティッカー風)の括弧内容、を候補から除外しながら探索する。
 *
 * 見つからない場合はnull。取得した業種名が長すぎる場合は、保存用の省略形に変換してから返す。
 */
function extractIndustryFromAnchor_(html) {
  const m = html.match(/search\/qi\/?\?ids=/);
  if (!m) return null;
  const idx = m.index;
  const gt = html.indexOf('>', idx);
  if (gt === -1) return null;
  const lt = html.indexOf('<', gt + 1);
  if (lt === -1) return null;
  const text = html.substring(gt + 1, lt).trim();
  return text || null;
}

function parseIndustryFromHtml_(code, html) {
  const anchorIndustry = extractIndustryFromAnchor_(html);
  if (anchorIndustry && anchorIndustry.length <= 20 && !/^[0-9A-Za-z]{1,5}$/.test(anchorIndustry)) {
    let industry = anchorIndustry;
    if (INDUSTRY_STORAGE_ABBREVIATIONS[industry]) industry = INDUSTRY_STORAGE_ABBREVIATIONS[industry];
    return industry;
  }

  // フォールバック: 括弧ベースの探索
  const codeUpper = code.toUpperCase();
  let searchFrom = 0;
  for (let i = 0; i < 20; i++) {
    const codeIdx = html.indexOf(code, searchFrom);
    if (codeIdx === -1) break;
    const nearby = html.substring(codeIdx, codeIdx + 500);
    const bracketRe = /[\[【]([^\[\]【】]{2,12})[\]】]/g;
    let match;
    while ((match = bracketRe.exec(nearby)) !== null) {
      const candidate = match[1].trim();
      if (!candidate) continue;
      if (candidate.toUpperCase() === codeUpper) continue; // コード自身の括弧表記はスキップ
      if (/^[0-9A-Za-z]{1,5}$/.test(candidate)) continue; // ティッカー風の英数字のみもスキップ
      let industry = candidate;
      if (INDUSTRY_STORAGE_ABBREVIATIONS[industry]) industry = INDUSTRY_STORAGE_ABBREVIATIONS[industry];
      return industry;
    }
    searchFrom = codeIdx + code.length;
  }
  return null;
}

/**
 * Yahoo!ファイナンスの銘柄ページから業種名を取得する。
 * 新形式コード(例: 285A)は銘柄ページ本体に業種リンクが出るが、従来形式の数字コード
 * (例: 9433)は本体ページには出ず、/profile(企業情報)ページ側にしか出ないため、
 * 本体ページで見つからない場合は/profileページも試す。
 * 取得できない場合やページ構造が想定と異なる場合はnullを返す(呼び出し側で手動入力にフォールバックする)。
 */
function fetchStockIndustry(code) {
  try {
    const baseUrl = 'https://finance.yahoo.co.jp/quote/' + code + '.T';
    const response = UrlFetchApp.fetch(baseUrl, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
    if (response.getResponseCode() === 200) {
      const industry = parseIndustryFromHtml_(code, response.getContentText());
      if (industry) return industry;
    }
    const profileResponse = UrlFetchApp.fetch(baseUrl + '/profile', { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
    if (profileResponse.getResponseCode() !== 200) return null;
    return parseIndustryFromHtml_(code, profileResponse.getContentText());
  } catch (e) {
    Logger.log('業種取得に失敗しました(' + code + '): ' + e);
    return null;
  }
}

/**
 * 複数銘柄の業種名を並列取得する({コード: 業種名|null})。取得できなかったコードはnullになる
 * (呼び出し側で「取得失敗」として扱い、手動入力を案内する)。
 *
 * 従来形式の数字コードは銘柄ページ本体に業種リンクが出ないため、1回目(本体ページ)で
 * 取得できなかったコードだけ、2回目として/profile(企業情報)ページを試す。
 */
function fetchIndustryPageBatch_(codes, urlSuffix) {
  const results = {};
  const BATCH = 10;
  for (let i = 0; i < codes.length; i += BATCH) {
    const batch = codes.slice(i, i + BATCH);
    const requests = batch.map(code => ({
      url: 'https://finance.yahoo.co.jp/quote/' + code + '.T' + urlSuffix,
      method: 'get', muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' }
    }));
    let responses = [];
    try { responses = UrlFetchApp.fetchAll(requests); } catch (e) { responses = []; }
    batch.forEach((code, k) => {
      const r = responses[k];
      results[code] = (r && r.getResponseCode() === 200) ? parseIndustryFromHtml_(code, r.getContentText()) : null;
    });
    if (i + BATCH < codes.length) Utilities.sleep(300);
  }
  return results;
}

function fetchStockIndustriesBulk(codes) {
  const results = fetchIndustryPageBatch_(codes, '');

  const stillMissing = codes.filter(code => !results[code]);
  if (stillMissing.length > 0) {
    Utilities.sleep(300);
    const profileResults = fetchIndustryPageBatch_(stillMissing, '/profile');
    stillMissing.forEach(code => {
      if (profileResults[code]) results[code] = profileResults[code];
    });
  }
  return results;
}

/** 銘柄の業種セルの文字列(省略形の場合あり)から、空売り集計シートの対応する業種ブロックを探す */
function resolveShortSellBlockForIndustry_(industryLabel) {
  if (!industryLabel) return null;
  const fullLabel = INDUSTRY_LABEL_ALIASES[industryLabel] || industryLabel;
  return SHORT_SELL_COLUMN_BLOCKS.find(b => b.label === fullLabel) || null;
}

/**
 * 銘柄の業種名(省略形可)から、その業種の空売り集計データ(売買代金合計・空売り売買代金合計・比率合計の全履歴)を返す。
 * 戻り値の形はbuildShortSellChartSeriesと同じ({ title, points: [{date, amount, shortAmountSum, ratioSum}, ...] })。
 * 対応する業種ブロックが見つからない場合はnull。
 */
function buildIndustryShortSellSeriesByLabel(industryLabel) {
  const block = resolveShortSellBlockForIndustry_(industryLabel);
  if (!block) return null;
  return buildShortSellChartSeries(block.col);
}

function shortSellPdfUrl_(dateStr, suffix) {
  const parts = dateStr.split('-');
  const ymd = parts[0].slice(-2) + parts[1] + parts[2];
  return 'https://www.jpx.co.jp/markets/statistics-equities/short-selling/' + SHORT_SELL_PDF_ATT_ID + '/' + ymd + '-' + suffix + '.pdf';
}

/**
 * 対象日が含まれる月の一覧ページのURLを返す。
 * 当月分はindex.html、1ヶ月前は00-archives-01.html、2ヶ月前は00-archives-02.html...12ヶ月前は00-archives-12.html
 * (1年より前はアーカイブ対象外のためnull)。
 */
function shortSellArchivePageUrl_(dateStr) {
  const target = parseDateStr_(dateStr);
  const today = new Date();
  const monthsBack = (today.getFullYear() - target.getFullYear()) * 12 + (today.getMonth() - target.getMonth());
  if (monthsBack <= 0) return 'https://www.jpx.co.jp/markets/statistics-equities/short-selling/index.html';
  if (monthsBack > 12) return null;
  return 'https://www.jpx.co.jp/markets/statistics-equities/short-selling/00-archives-' + String(monthsBack).padStart(2, '0') + '.html';
}

function findShortSellPdfUrlFromPage_(pageUrl, dateStr, suffix) {
  const parts = dateStr.split('-');
  const ymd = parts[0].slice(-2) + parts[1] + parts[2];
  const response = UrlFetchApp.fetch(pageUrl, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) return null;
  const html = response.getContentText();
  const re = new RegExp('href="((?:https://www\\.jpx\\.co\\.jp)?/markets/statistics-equities/short-selling/[^"]+' + ymd + '-' + suffix + '\\.pdf)"');
  const m = html.match(re);
  if (!m) return null;
  return m[1].startsWith('http') ? m[1] : 'https://www.jpx.co.jp' + m[1];
}

/**
 * 指定日のPDFテキストを取得する。当月分は添付URLを直接組み立てて試し(速いが添付IDが変わると失敗する)、
 * ダメなら該当月の一覧ページ(当月index.html、過去月は00-archives-NN.html)からリンクを探す。
 * 過去月は直接構築を飛ばし、最初から一覧ページを使う(添付IDが月ごとに異なりうるため)。
 */
function fetchShortSellPdfText_(dateStr, suffix) {
  const pageUrl = shortSellArchivePageUrl_(dateStr);
  const isCurrentMonth = pageUrl && pageUrl.indexOf('index.html') !== -1;

  let text = null;
  if (isCurrentMonth) {
    text = fetchPdfTextFromUrl_(shortSellPdfUrl_(dateStr, suffix));
  }
  if (!text && pageUrl) {
    const url = findShortSellPdfUrlFromPage_(pageUrl, dateStr, suffix);
    if (url) text = fetchPdfTextFromUrl_(url);
  }
  return text;
}

/**
 * PDFの実際のテキスト構造(「業種名/年月日 売買代金(a) 比率(a)/(d) 売買代金(b) 比率(b)/(d) 売買代金(c) 比率(c)/(d) 合計(d)」が
 * 1行に並ぶ表形式。「実注文」等の区分ラベルは見出しとして1回しか出てこないため、区分ごとに文字列検索するのではなく、
 * 与えられたテキスト範囲内の数値トークンを先頭から7個(a金額,a比率,b金額,b比率,c金額,c比率,d金額)として読み取る。
 */
function parseShortSellDataRow_(segmentText) {
  const tokens = segmentText.match(/-?[\d,]+\.?\d*%?/g);
  if (!tokens || tokens.length < 7) return null;

  const toAmount = t => parseInt(t.replace(/,/g, ''), 10);
  const toRatio = t => parseFloat(t.replace('%', ''));

  return {
    '実注文': { amount: toAmount(tokens[0]), ratio: toRatio(tokens[1]) },
    '空売り（価格規制あり）': { amount: toAmount(tokens[2]), ratio: toRatio(tokens[3]) },
    '空売り（価格規制なし）': { amount: toAmount(tokens[4]), ratio: toRatio(tokens[5]) },
    '合計': { amount: toAmount(tokens[6]), ratio: null }
  };
}

/**
 * 空売り集計PDF(市場全体、YYMMDD-m.pdf)のテキストを解析する。
 * 「2026年10月1日」のような和暦区切りの日付の直後に、その日のデータ行(7個の数値)が続く。
 * (ヘッダー付近に出てくる「2026/10/1」形式の日付はデータ行ではないため、和暦形式の日付のみを基準にする)
 */
function parseShortSellMarketPdfText_(fullText) {
  const dateMatch = fullText.match(/\d{4}年\d{1,2}月\d{1,2}日/);
  if (!dateMatch) return null;
  const startIdx = fullText.indexOf(dateMatch[0]) + dateMatch[0].length;
  return parseShortSellDataRow_(fullText.substring(startIdx, startIdx + 500));
}

/** 業種別集計PDF(YYMMDD-g.pdf)のテキストを、業種ごとに1行(7個の数値)として区切って解析する */
function parseShortSellIndustryPdfText_(fullText) {
  const anchors = SHORT_SELL_INDUSTRIES.map(ind => ind.search || ind.label);
  const result = {};
  for (let i = 0; i < SHORT_SELL_INDUSTRIES.length; i++) {
    const ind = SHORT_SELL_INDUSTRIES[i];
    const anchor = anchors[i];
    const startIdx = fullText.indexOf(anchor);
    if (startIdx === -1) { result[ind.label] = null; continue; }

    let endIdx = fullText.length;
    for (let j = 0; j < anchors.length; j++) {
      if (j === i) continue;
      const p = fullText.indexOf(anchors[j], startIdx + anchor.length);
      if (p !== -1 && p < endIdx) endIdx = p;
    }
    result[ind.label] = parseShortSellDataRow_(fullText.substring(startIdx + anchor.length, endIdx));
  }
  return result;
}

/**
 * 指定日の空売り集計(市場全体+業種別)を取得する。
 * 戻り値: { '日次集計': {区分: {amount, ratio}}, '水産・農林業': {...}, ... } または取得失敗時null。
 */
function fetchShortSellDataForDate(dateStr) {
  try {
    const marketText = fetchShortSellPdfText_(dateStr, 'm');
    const industryText = fetchShortSellPdfText_(dateStr, 'g');
    if (!marketText && !industryText) return null;

    const result = {};
    result['日次集計'] = marketText ? parseShortSellMarketPdfText_(marketText) : null;
    const industryData = industryText ? parseShortSellIndustryPdfText_(industryText) : {};
    SHORT_SELL_INDUSTRIES.forEach(ind => { result[ind.label] = industryData[ind.label] || null; });
    return result;
  } catch (e) {
    Logger.log('空売り集計の取得に失敗(' + dateStr + '): ' + e);
    return null;
  }
}

/**
 * 全ブロック(日次集計+全業種)の90日間スパークライン(空売り(価格規制あり)+(価格規制なし)の比率合計)を再計算する。
 * データブロックは1日4行(実注文/価格規制あり/価格規制なし/合計)、新しい日付が上に挿入されていく構造。
 * 日次集計(C列)だけはC4:D5の2行分を使う特別扱い、他の業種はそれぞれの列の4行目1行だけを使う。
 * 1回のgetValuesでまとめて読み込んでから各ブロックを計算する(列数が多いため個別読み出しは避ける)。
 */
function updateShortSellSparklines_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const dataRowCount = lastRow - 6; // 7行目以降のデータ行数
  if (dataRowCount < 4) return;

  const readRows = Math.min(dataRowCount, 90 * 4);
  const allValues = sheet.getRange(7, 1, readRows, lastCol).getValues();

  SHORT_SELL_COLUMN_BLOCKS.forEach(block => {
    if (block.col + 1 > lastCol) return;
    const ratioColIdx = block.col; // 0始まりの配列インデックス(比率列=block.col+1列目 → インデックスはblock.col)

    const series = [];
    for (let r = 0; r + 3 < allValues.length && series.length < 90; r += 4) {
      const ari = allValues[r + 1][ratioColIdx];   // 空売り(価格規制あり)の比率
      const nashi = allValues[r + 2][ratioColIdx]; // 空売り(価格規制なし)の比率
      if (typeof ari === 'number' && typeof nashi === 'number') series.push(ari + nashi);
    }
    series.reverse();

    const isMarketTotal = block.col === 3; // C:D(日次集計)だけ4〜5行目の2行を使う
    const range = isMarketTotal ? sheet.getRange(4, block.col, 2, 2) : sheet.getRange(4, block.col, 1, 2);
    if (!range.isPartOfMerge()) range.merge();
    if (series.length > 0) {
      range.setFormula('=SPARKLINE({' + series.join(',') + '}, {"charttype","line";"color","#ea4335"})');
    }
  });
}

/**
 * 空売り集計シートの見出し(フォーマット崩れ時の復旧用)。
 * 既に業種名が入っている場合はスキップするため、何度実行しても安全。
 */
function setupShortSellSheetHeaders() {
  const ss = getMainSpreadsheet();
  let sheet = ss.getSheetByName(SHORT_SELL_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHORT_SELL_SHEET_NAME);

  if (sheet.getRange(3, 3).getValue() === SHORT_SELL_COLUMN_BLOCKS[0].label) {
    sheet.setFrozenRows(6);
    sheet.setFrozenColumns(2);
    return '空売り集計シート: 見出しは設定済みのためスキップしました';
  }

  sheet.getRange(1, 1, 6, 2 + SHORT_SELL_COLUMN_BLOCKS.length * 2).clearContent();
  sheet.getRange(4, 2).setValue('チャート');
  sheet.getRange(6, 1).setValue('日付');

  SHORT_SELL_COLUMN_BLOCKS.forEach(block => {
    const headerCell = sheet.getRange(3, block.col, 1, 2);
    if (headerCell.isPartOfMerge()) headerCell.breakApart();
    headerCell.merge().setValue(block.label);
    sheet.getRange(6, block.col).setValue('売買代金');
    sheet.getRange(6, block.col + 1).setValue('比率');
  });

  SHORT_SELL_CATEGORIES.forEach((cat, i) => sheet.getRange(7 + i, 2).setValue(cat));

  SHORT_SELL_COLUMN_BLOCKS.forEach(block => {
    const isMarketTotal = block.col === 3;
    const sparklineCell = isMarketTotal ? sheet.getRange(4, block.col, 2, 2) : sheet.getRange(4, block.col, 1, 2);
    if (!sparklineCell.isPartOfMerge()) sparklineCell.merge();
  });

  sheet.setFrozenRows(6);
  sheet.setFrozenColumns(2);
  return '空売り集計シート: 見出しを設定しました';
}

/** 指定日の空売り集計データを、シート上部に新しい4行ブロックとして挿入する */
function writeShortSellDateToSheet(dateStr, data, skipSparkline) {
  const ss = getMainSpreadsheet();
  const sheet = ss.getSheetByName(SHORT_SELL_SHEET_NAME);
  if (!sheet) { Logger.log('空売り集計シートが見つかりません'); return false; }
  if (!data) return false;

  const hadPreviousBlock = sheet.getLastRow() >= 10;
  sheet.insertRowsAfter(6, 4); // 6行目の下に4行追加(7〜10行目が新しい空行になる)
  if (hadPreviousBlock) {
    // 新しい7〜10行目は挿入時に周囲の書式を引き継いでいることがあるため、まず消してから
    // 1つ上(直前まで一番上だった、今は11〜14行目)のブロックの書式だけをコピーする(重複コピー防止)
    const targetRange = sheet.getRange(7, 1, 4, sheet.getLastColumn());
    targetRange.clearFormat();
    sheet.getRange(11, 1, 4, sheet.getLastColumn()).copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  }

  const dateCell = sheet.getRange(7, 1, 4, 1);
  if (dateCell.isPartOfMerge()) dateCell.breakApart();
  dateCell.merge().setValue(dateStr);

  SHORT_SELL_CATEGORIES.forEach((cat, ci) => sheet.getRange(7 + ci, 2).setValue(cat));

  const lastCol = sheet.getLastColumn();
  SHORT_SELL_COLUMN_BLOCKS.forEach(block => {
    if (block.col + 1 > lastCol) return;
    const catData = data[block.label];
    if (!catData) return;
    SHORT_SELL_CATEGORIES.forEach((cat, ci) => {
      const row = 7 + ci;
      const d = catData[cat];
      if (!d) return;
      if (d.amount !== null) sheet.getRange(row, block.col).setValue(d.amount);
      if (d.ratio !== null) sheet.getRange(row, block.col + 1).setValue(d.ratio / 100); // 比率はパーセント表示のセル書式を想定し小数で格納
    });
  });

  if (!skipSparkline) updateShortSellSparklines_(sheet);
  return true;
}

/**
 * 指定日の空売り集計を取得してシートに書き込む(ダイアログ・トリガー共通の単日処理)。
 * skipSparkline=trueの場合はスパークラインの再計算を行わない(期間指定更新で最後に1回だけまとめて行うため)。
 */
function updateShortSellForDateInput(dateStr, skipSparkline) {
  const dow = parseDateStr_(dateStr).getDay();
  if (dow === 0 || dow === 6 || isJapaneseHoliday(parseDateStr_(dateStr))) {
    const msg = dateStr + ': 休場日のためスキップしました';
    Logger.log(msg);
    return msg;
  }

  const data = fetchShortSellDataForDate(dateStr);
  if (!data) {
    const msg = dateStr + ': データを取得できませんでした(未公開、またはJPX側のURL形式が変わった可能性があります)';
    Logger.log(msg);
    return msg;
  }
  const ok = writeShortSellDateToSheet(dateStr, data, skipSparkline);
  const msg = dateStr + ': ' + (ok ? '空売り集計を更新しました' : '書き込みに失敗しました');
  Logger.log(msg);
  return msg;
}

// 期間指定更新は、スパークライン再計算の時間を確保するため5分30秒で日々の書き込みを打ち切る
// (全体の6分制限より短く区切ることで、最後のスパークライン更新がタイムアウトに巻き込まれないようにする)
const SHORT_SELL_RANGE_TIME_LIMIT_MS = 5.5 * 60 * 1000;

/**
 * ダイアログからの期間指定更新。時間切れが近づいたら安全に止まり、続きは再実行で進む。
 * スパークラインの再計算は日ごとに行わず、最後に1回だけまとめて行う。
 */
function updateShortSellForDateRangeInput(startDateStr, endDateStr) {
  const rangeStartMs = Date.now();
  const start = parseDateStr_(startDateStr);
  const end = parseDateStr_(endDateStr);
  if (start > end) return '開始日は終了日より前の日付にしてください';

  const messages = [];
  for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
    if (Date.now() - rangeStartMs > SHORT_SELL_RANGE_TIME_LIMIT_MS) {
      messages.push('(実行時間の上限のため、' + Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd') + '以降は未処理です。同じ期間でもう一度実行すると続きから処理されます)');
      break;
    }
    messages.push(updateShortSellForDateInput(Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd'), true));
  }

  // スパークラインの更新は最後に1回だけ。失敗しても日々のデータ自体は書き込み済みなので、エラーで落とさず記録だけする
  try {
    const sheet = getMainSpreadsheet().getSheetByName(SHORT_SELL_SHEET_NAME);
    if (sheet) updateShortSellSparklines_(sheet);
  } catch (e) {
    Logger.log('スパークラインの更新に失敗しました(データ自体は書き込み済みです): ' + e);
    messages.push('(スパークラインの更新でエラーが発生しましたが、データ自体は書き込み済みです)');
  }

  const message = messages.join('\n');
  Logger.log(message);
  return message;
}

/**
 * 空売り集計シートの全期間チャートボタン用: クリックされた列(売買代金 or 比率どちらでも)から
 * そのブロックを特定し、次の3つを日付昇順の配列で返す(どちらの列を選んでも同じ結果)。
 *  - amount: 「合計」行の売買代金(青線用)
 *  - shortAmountSum: 空売り(価格規制あり)+空売り(価格規制なし)の売買代金合計(黒線用)
 *  - ratioSum: 空売り(価格規制あり)+空売り(価格規制なし)の比率合計(ツールチップ表示用、グラフ線ではない)
 * 戻り値: { title, points: [{date:'yyyy-MM-dd', amount:数値, shortAmountSum:数値|null, ratioSum:数値|null}, ...] } または該当なしの場合null。
 */
function buildShortSellChartSeries(col) {
  const ss = getMainSpreadsheet();
  const sheet = ss.getSheetByName(SHORT_SELL_SHEET_NAME);
  if (!sheet) return null;

  const block = SHORT_SELL_COLUMN_BLOCKS.find(b => col === b.col || col === b.col + 1);
  if (!block) return null;

  const lastRow = sheet.getLastRow();
  const dataRowCount = lastRow - 6;
  if (dataRowCount < 4) return { title: block.label, points: [] };

  const allValues = sheet.getRange(7, 1, dataRowCount, sheet.getLastColumn()).getValues();
  const amountIdx = block.col - 1; // 「売買代金」列の0始まりインデックス
  const ratioIdx = block.col;      // 「比率」列(amountIdx+1)の0始まりインデックス

  const points = [];
  for (let r = 0; r + 3 < allValues.length; r += 4) {
    const dateVal = allValues[r][0]; // A列(日付、4行結合の先頭行に文字列で入っている)
    const dateStr = (typeof dateVal === 'string' && dateVal) ? dateVal
      : (dateVal instanceof Date ? Utilities.formatDate(dateVal, 'Asia/Tokyo', 'yyyy-MM-dd') : null);
    if (!dateStr) continue;

    const amount = allValues[r + 3][amountIdx]; // 「合計」行の売買代金
    if (typeof amount !== 'number') continue;

    const ariAmount = allValues[r + 1][amountIdx];   // 空売り(価格規制あり)の売買代金
    const nashiAmount = allValues[r + 2][amountIdx]; // 空売り(価格規制なし)の売買代金
    const shortAmountSum = (typeof ariAmount === 'number' && typeof nashiAmount === 'number') ? ariAmount + nashiAmount : null;

    const ariRatio = allValues[r + 1][ratioIdx];   // 空売り(価格規制あり)の比率
    const nashiRatio = allValues[r + 2][ratioIdx]; // 空売り(価格規制なし)の比率
    const ratioSum = (typeof ariRatio === 'number' && typeof nashiRatio === 'number') ? ariRatio + nashiRatio : null;

    points.push({ date: dateStr, amount: amount, shortAmountSum: shortAmountSum, ratioSum: ratioSum });
  }
  points.reverse(); // 新しい日付が上にあるため、古い→新しい(昇順)に並べ替える

  return {
    title: block.label,
    points: points
  };
}

/** 時間主導トリガー用: 本日分の空売り集計を更新する */
function updateShortSellForToday() {
  const today = new Date();
  if (today.getDay() === 0 || today.getDay() === 6 || isJapaneseHoliday(today)) {
    Logger.log('本日は休場日のため空売り集計の更新をスキップしました');
    return;
  }
  const message = updateShortSellForDateInput(Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy-MM-dd'));
  Logger.log(message);
  return message;
}