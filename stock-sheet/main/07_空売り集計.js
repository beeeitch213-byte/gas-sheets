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

  // 新しく挿入した7〜10行目(B列〜最終列)は空なので、値を配列で組み立てて1回で書き込む
  const lastCol = sheet.getLastColumn();
  const block4 = SHORT_SELL_CATEGORIES.map(cat => {
    const row = new Array(lastCol - 1).fill('');
    row[0] = cat; // B列=区分名
    return row;
  });
  SHORT_SELL_COLUMN_BLOCKS.forEach(block => {
    if (block.col + 1 > lastCol) return;
    const catData = data[block.label];
    if (!catData) return;
    SHORT_SELL_CATEGORIES.forEach((cat, ci) => {
      const d = catData[cat];
      if (!d) return;
      if (d.amount !== null) block4[ci][block.col - 2] = d.amount;
      if (d.ratio !== null) block4[ci][block.col - 1] = d.ratio / 100; // 比率はパーセント表示のセル書式を想定し小数で格納
    });
  });
  sheet.getRange(7, 2, 4, lastCol - 1).setValues(block4);

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