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
  const text = fetchPdfTextFromUrl_(pdfUrl); // 変換用の一時ファイルは失敗時も含め必ず削除される
  if (text === null) throw new Error('PDFを取得できませんでした: ' + pdfUrl);
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

/**
 * xlsxの1枚目のシート(xl/worksheets/sheet1.xml)から、指定した行の値を{行番号: [A列, B列, ...]}で返す
 * (数値セルは数値、それ以外は'')。Googleスプレッドシートへの変換(一時ファイルの作成・削除で数秒かかる)を
 * 避けるため、xlsx(zip)を展開してXMLを直接読む。読めなければnull。
 */
function readXlsxRows_(blob, rowNumbers, lastCol) {
  const sheetFile = Utilities.unzip(blob.setContentType('application/zip'))
    .find(f => f.getName() === 'xl/worksheets/sheet1.xml');
  if (!sheetFile) return null;
  const xml = sheetFile.getDataAsString('UTF-8');
  const colIndex = letters => letters.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);

  const rows = {};
  const rowRe = /<row\b[^>]*\br="(\d+)"[^>]*>([\s\S]*?)<\/row>/g;
  let rm;
  while ((rm = rowRe.exec(xml)) !== null) {
    const rowNumber = parseInt(rm[1], 10);
    if (rowNumbers.indexOf(rowNumber) === -1) continue;
    const values = new Array(lastCol).fill('');
    const cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    let cm;
    while ((cm = cellRe.exec(rm[2])) !== null) {
      const ref = cm[1].match(/\br="([A-Z]+)\d+"/);
      const type = cm[1].match(/\bt="([^"]*)"/);
      const v = cm[2] && cm[2].match(/<v>([^<]*)<\/v>/);
      if (!ref || !v || (type && type[1] !== 'n')) continue;
      const col = colIndex(ref[1]);
      if (col <= lastCol && v[1] !== '' && !isNaN(Number(v[1]))) values[col - 1] = Number(v[1]);
    }
    rows[rowNumber] = values;
  }
  return rows;
}

/** xlsxをGoogleスプレッドシートに変換して指定した行を読む(readXlsxRows_で読めなかった場合の保険。変換したファイルはtrashする) */
function readXlsxRowsViaDrive_(blob, rowNumbers, lastCol) {
  const resource = { name: 'temp_jpx_xlsx_convert', mimeType: MimeType.GOOGLE_SHEETS };
  const file = Drive.Files.create(resource, blob, { convert: true });
  try {
    const sheet = SpreadsheetApp.openById(file.id).getSheets()[0];
    const rows = {};
    rowNumbers.forEach(r => { rows[r] = sheet.getRange(r, 1, 1, lastCol).getValues()[0]; });
    return rows;
  } finally {
    DriveApp.getFileById(file.id).setTrashed(true);
  }
}

/** 新形式のxlsxを取得して、各市場の「金額 Value」行を{行番号: 値の配列}で返す */
function fetchJpxXlsxValueRows_(xlsxUrl) {
  const blob = UrlFetchApp.fetch(xlsxUrl, { muteHttpExceptions: true }).getBlob();
  const rowNumbers = JPX_MARKETS.map(m => JPX_XLSX_VALUE_ROW_[m.key]);
  const isReadable = rows => !!rows && rowNumbers.every(r => rows[r] && rows[r].some(v => typeof v === 'number'));
  let rows = null;
  try {
    rows = readXlsxRows_(blob.copyBlob(), rowNumbers, JPX_XLSX_LAST_COL_);
  } catch (e) {
    Logger.log('xlsxの直接読み取りに失敗: ' + e);
  }
  if (isReadable(rows)) return rows;
  Logger.log('xlsxを直接読めなかったため、Googleスプレッドシートに変換して読み直します: ' + xlsxUrl);
  return readXlsxRowsViaDrive_(blob, rowNumbers, JPX_XLSX_LAST_COL_);
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

/** valueRowsはfetchJpxXlsxValueRows_の戻り値({行番号: 値の配列}) */
function parseJpxXlsxRows_(valueRows) {
  const result = {};
  JPX_MARKETS.forEach(market => {
    const rowValues = valueRows[JPX_XLSX_VALUE_ROW_[market.key]];
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
// 年ごとのアーカイブページは1回の実行の中で何度も同じページを読むことになるため、結果を使い回す
const jpxWeekLinksCache_ = {};

function getJpxWeekLinksForYear_(yearOffset) {
  if (jpxWeekLinksCache_[yearOffset]) return jpxWeekLinksCache_[yearOffset];
  const result = fetchJpxWeekLinksForYear_(yearOffset);
  if (result.length > 0) jpxWeekLinksCache_[yearOffset] = result;
  return result;
}

function fetchJpxWeekLinksForYear_(yearOffset) {
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

/** 総計列(5行目〜最終行)を1回で読み込む(スパークライン2種で使い回す) */
function readJpxTotalColumn_(sheet) {
  const totalCol = JPX_ALL_COLS.find(c => c.key === '総計').col;
  const lastRow = sheet.getLastRow();
  return lastRow >= 5 ? sheet.getRange(5, totalCol, lastRow - 4, 1).getValues().map(r => r[0]) : [];
}

/** 総計列の配列から、startRowから9行おきに最大24週分の数値を取り出す(古い順に並べ替えて返す) */
function pickJpxWeeklyValues_(totalColValues, startRow) {
  const values = [];
  for (let r = startRow; r - 5 < totalColValues.length && values.length < 24; r += 9) {
    const v = totalColValues[r - 5];
    if (typeof v === 'number') values.push(v);
  }
  return values.reverse();
}

function setupJpxBalanceSparklines(sheet, totalColValues) {
  const colValues = totalColValues || readJpxTotalColumn_(sheet);
  JPX_MARKETS.forEach((market, mi) => {
    const dataBaseRow = 5 + mi * 3;
    const chartBaseRow = 14 + mi * 3;
    const balanceRow = dataBaseRow + 2;
    const values = pickJpxWeeklyValues_(colValues, balanceRow);

    const range = sheet.getRange(chartBaseRow, 4, 3, 3);
    range.merge();
    if (values.length > 0) {
      const maxAbs = Math.max.apply(null, values.map(Math.abs).concat([1]));
      range.setFormula('=SPARKLINE({' + values.join(',') + '}, {"charttype","column";"color","#ea4335";"negcolor","#4285f4";"ymin",-' + maxAbs + ';"ymax",' + maxAbs + '})');
    }
  });
}

function setupJpxTotalLineChart(sheet, totalColValues) {
  const colValues = totalColValues || readJpxTotalColumn_(sheet);
  JPX_MARKETS.forEach((market, mi) => {
    const baseRow = 5 + mi * 3;
    const values = pickJpxWeeklyValues_(colValues, baseRow);

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
    sheet.getRange(r, 3, 3, 1).setValues([['金額'], ['比率'], ['差引き']]);
    if (!data) return;

    // 数値(7列目〜最終区分の列)は3行分をまとめて読み込み、データのある区分だけ書き換えて1回で書き戻す
    const firstCol = JPX_ALL_COLS[0].col;
    const width = JPX_ALL_COLS[JPX_ALL_COLS.length - 1].col + 2 - firstCol + 1;
    const vals = sheet.getRange(r, firstCol, 3, width).getValues();
    JPX_ALL_COLS.forEach(c => {
      const d = data[c.key];
      if (!d) return;
      const o = c.col - firstCol;
      vals[0][o] = jpxToMillion(d.sell);
      vals[0][o + 1] = jpxToMillion(d.buy);
      vals[0][o + 2] = jpxToMillion(d.total);
      vals[1][o] = d.sellRatio === null || d.sellRatio === undefined ? '' : d.sellRatio;
      vals[1][o + 1] = d.buyRatio === null || d.buyRatio === undefined ? '' : d.buyRatio;
      vals[1][o + 2] = d.totalRatio === null || d.totalRatio === undefined ? '' : d.totalRatio;
      const balCell = sheet.getRange(r + 2, c.col, 1, 3);
      if (balCell.isPartOfMerge()) balCell.breakApart();
      balCell.merge();
      vals[2][o] = jpxToMillion(d.balance);
      vals[2][o + 1] = '';
      vals[2][o + 2] = '';
    });
    sheet.getRange(r, firstCol, 3, width).setValues(vals);
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

/** 売買状況シートのスパークライン(総計の推移・差引き)を描き直す。一連の書き込みの最後に1回だけ呼ぶ */
function redrawJpxCharts_() {
  const sheet = getMainSpreadsheet().getSheetByName(JPX_STATS_SHEET_NAME);
  if (!sheet) return;
  clearJpxSparklineMerges(sheet);
  const totalColValues = readJpxTotalColumn_(sheet);
  setupJpxTotalLineChart(sheet, totalColValues);
  setupJpxBalanceSparklines(sheet, totalColValues);
}

/**
 * 週データをシートへ反映する。シートは常に「新しい週が上」の順に保つ。
 * ・同じ開始日の週ブロックが既にあれば、その場で上書き(旧形式の年なしラベルも新形式ラベルに置き換わる)
 * ・なければ、自分より古い最初のブロックの直前に挿入(=日付順の正しい位置。最新週なら先頭)
 * ・既存のどのブロックよりも古ければ末尾に追加
 * ※グラフは描画しない。書き込み後にredrawJpxCharts_()を呼ぶこと
 */
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
    sheet.getRange(insertRow + 9, 1, 9, 1).copyTo(sheet.getRange(insertRow, 1, 9, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    sheet.getRange(insertRow + 9, 4, 9, 48).copyTo(sheet.getRange(insertRow, 4, 9, 48), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    writeJpxBlockAtRow_(sheet, insertRow, weekStartDate, weekEndDate, parsedData);
  } else {
    const baseRow = blocks.length > 0 ? blocks[blocks.length - 1].row + 9 : 5;
    if (blocks.length > 0) {
      sheet.getRange(5, 1, 9, 1).copyTo(sheet.getRange(baseRow, 1, 9, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      sheet.getRange(5, 4, 9, 48).copyTo(sheet.getRange(baseRow, 4, 9, 48), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    }
    writeJpxBlockAtRow_(sheet, baseRow, weekStartDate, weekEndDate, parsedData);
  }
  // グラフ(スパークライン)の再描画は、呼び出し元で全週の書き込みが終わった最後に1回だけ行う
}

/** weekInfoを取得・解析する(新形式はExcel、旧形式はPDFを使う) */
function fetchAndParseJpxWeek_(weekInfo) {
  if (weekInfo.isNewFormat) return parseJpxXlsxRows_(fetchJpxXlsxValueRows_(weekInfo.xlsxUrl));
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

// 自動更新で、公開済みかどうかを確認する期間(日数)。この期間に含まれる週のうち、シートに無い週を取得する
const JPX_LATEST_LOOKBACK_DAYS = 28;

/**
 * 時間主導トリガー用(引数なし): JPXに公開済みの直近の週のうち、まだシートに無い週を取得する。
 * 投資部門別売買状況は、翌週の第4営業日(木曜ごろ)に前の週の分が公表される。そのため「今日が含まれる週」を
 * 取りにいくと毎回未公表でスキップされ、前の週の分も取りこぼしてしまう(2026年10月第1週で発生)。
 * 直近JPX_LATEST_LOOKBACK_DAYS日に含まれる週をまとめて確認し、取得済みの週はスキップする。
 */
function updateJpxStatsForLatestWeek() {
  const today = new Date();
  const from = new Date(today);
  from.setDate(from.getDate() - JPX_LATEST_LOOKBACK_DAYS);
  const fmt = d => Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  const message = updateJpxStatsForDateRangeInput(fmt(from), fmt(today), false);
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