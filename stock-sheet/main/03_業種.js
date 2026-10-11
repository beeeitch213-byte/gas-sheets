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

// ==== 業種情報 ====

/**
 * 当月シート(なければテンプレート)から、指定銘柄コードの業種セル(銘柄列の右隣2列、省略形の場合あり)を読み取る。
 * 見つからない/未設定の場合は空文字。
 */
function getStockIndustryLabel_(code) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const candidates = [getTargetSheet(new Date()), ss.getSheetByName(TEMPLATE_SHEET_NAME)].filter(s => s);
  for (let i = 0; i < candidates.length; i++) {
    const sheet = candidates[i];
    const lastCol = sheet.getLastColumn();
    if (lastCol < STOCK_START_COL) continue;
    const codeRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
    for (let col = STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) {
      if (String(codeRow[col - 1]) === String(code)) {
        const industry = sheet.getRange(2, col + 1).getValue();
        if (industry) return String(industry);
      }
    }
  }
  return '';
}

/** テンプレート+全月シートの、指定銘柄コードの業種セル(銘柄列の右隣2列)に値を設定する。設定できたシート数を返す。 */
/** テンプレート+全月シートについて、{sheet, colByCode:{コード: 列}} の一覧を作る(複数銘柄をまとめて処理する時の読み込み回数削減用) */
function buildStockColumnMaps_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().filter(s => s.getName() === TEMPLATE_SHEET_NAME || parseYearMonthFromSheetName(s.getName()));
  return sheets.map(sheet => {
    const colByCode = {};
    const lastCol = sheet.getLastColumn();
    if (lastCol >= STOCK_START_COL) {
      const codeRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
      for (let col = STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) {
        const c = codeRow[col - 1];
        if (c && !colByCode[String(c)]) colByCode[String(c)] = col;
      }
    }
    return { sheet: sheet, colByCode: colByCode };
  });
}

/** sheetMaps(省略時はその場で作成)を使い、指定コードの業種セルに値を設定する。設定できたシート数を返す */
function setStockIndustryAcrossAllSheets_(code, industry, sheetMaps) {
  const maps = sheetMaps || buildStockColumnMaps_();
  let count = 0;
  maps.forEach(m => {
    const col = m.colByCode[String(code)];
    if (!col) return;
    const cell = m.sheet.getRange(2, col + 1, 1, 2);
    if (!cell.isPartOfMerge()) cell.merge();
    m.sheet.getRange(2, col + 1).setValue(industry);
    count++;
  });
  return count;
}

/** メニュー「業種情報を一括取得(銘柄選択)」: 当月シートの銘柄一覧をチェックボックスで選び、まとめて業種を自動取得する */
function addIndustryBulkUI() {
  const stocks = getCurrentStockOrder(); // [{code, name}]
  if (stocks.length === 0) {
    SpreadsheetApp.getUi().alert('当月シートに銘柄が見つかりません');
    return;
  }
  const template = HtmlService.createTemplateFromFile('IndustryBulkDialog');
  template.stocksJson = JSON.stringify(stocks);
  const html = template.evaluate().setWidth(420).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '業種情報を一括取得');
}

/**
 * IndustryBulkDialogから呼ばれる。選択された銘柄コードの業種をまとめて自動取得し、
 * 取得できたものは全シート(当月・テンプレート・過去月シート)に設定する。取得できなかったコードは
 * 結果メッセージに列挙するので、該当コードだけ該当シートの2行目業種欄(銘柄列の右隣2列)に直接入力する。
 */
function runIndustryBulkFetch(codes) {
  if (!codes || codes.length === 0) return '銘柄が選択されていません';
  const results = fetchStockIndustriesBulk(codes);
  const succeeded = [], failed = [];
  const sheetMaps = buildStockColumnMaps_(); // 各シートのコード行は1回だけ読む
  codes.forEach(code => {
    const industry = results[code];
    if (industry) {
      setStockIndustryAcrossAllSheets_(code, industry, sheetMaps);
      succeeded.push(code + '(' + industry + ')');
    } else {
      failed.push(code);
    }
  });

  let message = '取得完了: ' + succeeded.length + '/' + codes.length + '件\n';
  if (succeeded.length > 0) message += '成功: ' + succeeded.join(', ') + '\n';
  if (failed.length > 0) {
    message += '失敗(手動入力が必要): ' + failed.join(', ') + '\n'
      + '→該当コードの銘柄列の右隣2列(2行目の業種欄)に直接入力してください。';
  }
  return message;
}

/**
 * 既存シートの2行目(銘柄3列結合=銘柄コード)を、1列目=コード単独セル+残り2列(結合)=業種情報、の
 * レイアウトに移行する。既に移行済みのシート/列はスキップされるので、再実行しても安全。
 */
function migrateCodeCellsForIndustryLayout_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < STOCK_START_COL) return 0;
  const codeRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  let count = 0;
  for (let col = STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) {
    if (!codeRow[col - 1]) continue;
    const fullRange = sheet.getRange(2, col, 1, STOCK_COL_WIDTH);
    if (fullRange.isPartOfMerge()) {
      const merged = fullRange.getMergedRanges()[0];
      if (merged.getNumColumns() === STOCK_COL_WIDTH) {
        fullRange.breakApart();
        sheet.getRange(2, col + 1, 1, 2).merge();
        count++;
        continue;
      }
    }
    const industryRange = sheet.getRange(2, col + 1, 1, 2);
    if (!industryRange.isPartOfMerge()) industryRange.merge();
  }
  return count;
}

function migrateAllSheetsForIndustryLayout() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targets = ss.getSheets().filter(s => s.getName() === TEMPLATE_SHEET_NAME || parseYearMonthFromSheetName(s.getName()));
  const messages = targets.map(s => s.getName() + ': ' + migrateCodeCellsForIndustryLayout_(s) + '件移行');
  return messages.join('\n');
}