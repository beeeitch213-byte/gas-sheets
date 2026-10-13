/**
 * 日本株 終値自動取得システム(プロジェクト: 株式情報) ― 設定・共通処理
 * ------------------------------------------------------------
 * 2026/10にメイン(株式情報)とDB(株式情報_DB、ライブラリStockLib)を1つのプロジェクトに統合した。
 * ファイル構成(Apps Scriptエディタ上で以下の順に並べる。00が先頭である必要がある):
 *   00_設定・共通      シート構成の定数、シート取得・祝日判定・PDF変換などの共通処理
 *   01_メニュー        「機能」メニュー、各ダイアログの呼び出し口、自動更新トリガーの設定
 *   02_銘柄操作        銘柄の追加・削除・並び替え、色の設定
 *   03_業種            業種情報の取得・設定
 *   04_終値・市場      終値・前日比・売買代金、市場ブロック(出来高・売買代金)、来月シート作成
 *   05_データベース・書式  データベースシート、スパークライン、書式の再適用
 *   06_信用残          JPX日次信用残(mtall.pdf)の解析と書き込み
 *   07_空売り集計      空売り集計シート
 *   08_投資部門別      投資部門別売買状況(売買状況シート)
 *   09_チャート        ボタンチャート(ChartDialog)用のデータ組み立て
 *
 * 【セットアップ】
 *  ・左メニュー「サービス」で Drive API と Google Sheets API を有効にする
 *  ・時間主導トリガーは「機能」→「データ修正」→「自動更新トリガーを設定」で一括設定できる
 *    (updateDailyClose / updateMarketTradingValues / updateShortSellForToday / updateMarginData / updateJpxStatsForLatestWeek)
 */

// ==== 設定項目 ====
const MAIN_SPREADSHEET_ID = '1WE6PgUNDmF-FkuidsgyRSaDKQXdkTngHZArOz1optZg'; // 通常は使わない(getMainSpreadsheetの予備)

const TEMPLATE_SHEET_NAME = 'テンプレート';

const DB_SHEET_NAME = 'データベース';

const STOCK_COL_WIDTH = 3;

const STOCK_START_COL = 12; // 通常銘柄はL列から(C〜K列はプライム/スタンダード/グロースの市場ブロックに固定で使う)

// 行構成: 5行目=見出し / 6〜36行目=日次の終値等(1日〜31日) / 38行目=信用残見出し / 39〜69行目=日次の信用残(1日〜31日)
const CLOSE_FIRST_ROW = 6;

const MARGIN_HEADER_ROW = 38; // 信用残の見出し行

const MARGIN_FIRST_ROW = 39;

const SHEET_DATA_MAX_ROW = 69;       // 当月シート・テンプレートの構造上、意味のあるデータが入りうる最終行

const DB_BACKFILL_YEARS = 3; // データベースに保持する過去データの年数

const RUN_TIME_LIMIT_MS = 4.5 * 60 * 1000; // 6分制限の手前で安全に止めるための実行時間上限

const MA_PERIOD = 25;

const CHART_DISPLAY_MONTHS = 6;      // チャート表示期間(月、通常銘柄のスパークライン用)

const CHART_DIALOG_HEIGHT = 820;     // 終値ボタンチャートのダイアログの高さ(px)

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

const HOLIDAY_LABEL = '休';

const HOLIDAY_COLOR = '#999999';

// 1回の実行の中では同じスプレッドシートを何度も開き直さないよう、最初に開いたものを使い回す
let mainSpreadsheetCache_ = null;

function getMainSpreadsheet() {
  if (!mainSpreadsheetCache_) {
    // このプロジェクトはスプレッドシートに紐付いているので、通常は getActiveSpreadsheet() で取得できる
    // (トリガー実行時も同じ)。念のため取得できない場合だけIDで開く。
    mainSpreadsheetCache_ = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  }
  return mainSpreadsheetCache_;
}

function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * シート名(例: "9月26年")から年と月を抽出する
 */
function parseYearMonthFromSheetName(sheetName) {
  const match = sheetName.match(/^(\d+)月(\d+)年$/);
  if (match) {
    const m = parseInt(match[1], 10);
    const yStr = match[2];
    const y = parseInt("20" + yStr, 10);
    return { year: y, month: m };
  }
  return null;
}

// ==== シート特定・日付ユーティリティ ====
function getTargetSheet(today) {
  const targetYearStr = String(today.getFullYear()).slice(-2);
  const targetMonth = today.getMonth() + 1;
  const targetName = targetMonth + '月' + targetYearStr + '年';
  return getMainSpreadsheet().getSheetByName(targetName);
}

function getDaysInMonth(year, month) {
  return new Date(year, month, 0).getDate();
}

// ==== 日付指定更新(ダイアログから特定日・期間の終値・前日比を更新) ====

function parseDateStr_(dateStr) {
  const parts = String(dateStr).split('-').map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2]);
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

// 祝日カレンダーの問い合わせは遅いため、1回の実行の中では「月ごとに1回だけ」まとめて取得して使い回す
const holidayMonthCache_ = {};

function getHolidaySetForMonth_(year, month0) {
  const key = year + '-' + month0;
  if (holidayMonthCache_[key]) return holidayMonthCache_[key];
  const set = {};
  try {
    const cal = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
    const tz = Session.getScriptTimeZone();
    cal.getEvents(new Date(year, month0, 1), new Date(year, month0 + 1, 1)).forEach(function (e) {
      if (JP_PUBLIC_HOLIDAY_NAMES_.indexOf(e.getTitle()) === -1) return;
      const start = e.isAllDayEvent() ? e.getAllDayStartDate() : e.getStartTime();
      set[Utilities.formatDate(start, tz, 'yyyy-MM-dd')] = true;
    });
  } catch (e) {
    // 取得に失敗した場合は従来どおり「祝日ではない」扱い(年末年始の固定休場日だけは下で判定)
  }
  holidayMonthCache_[key] = set;
  return set;
}

function isJapaneseHoliday(date) {
  const month = date.getMonth() + 1;
  const day = date.getDate();
  if ((month === 1 && day <= 3) || (month === 12 && day === 31)) return true;
  const set = getHolidaySetForMonth_(date.getFullYear(), date.getMonth());
  return !!set[Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd')];
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

function fetchPdfBytes_(url) {
  return pdfBytesFromResponse_(UrlFetchApp.fetch(url, { muteHttpExceptions: true }));
}

/** HTTPの応答がPDFならバイト列を返す(404のページなどPDFでなければnull) */
function pdfBytesFromResponse_(response) {
  if (!response || response.getResponseCode() !== 200) return null;
  const bytes = new Uint8Array(response.getContent()); // Javaのbyte(-128〜127)は0〜255に変換される
  if (bytes.length < 5 || String.fromCharCode(bytes[0], bytes[1], bytes[2], bytes[3]) !== '%PDF') return null;
  return bytes;
}