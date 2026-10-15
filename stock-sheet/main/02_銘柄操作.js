/**
 * セルから直接呼び出す用:指定銘柄の終値を取得(fetchCloseを利用)
 */
function JQCLOSE(code) {
  const result = fetchClose(code);
  return result === null ? 'データなし' : result.price;
}

function checkStockCodeExists(code) {
  const sheet = getTargetSheet(new Date());
  if (!sheet) return { exists: false };

  const lastCol = sheet.getLastColumn();
  const codeRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];

  for (let col = STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) {
    if (String(codeRow[col - 1]) === String(code)) {
      const name = sheet.getRange(3, col).getValue();
      return { exists: true, name: name };
    }
  }
  return { exists: false };
}

function addChangeConditionalFormat(sheet, changeCol) {
  const colors = getColorSettings();
  const range = sheet.getRange(columnToLetter(changeCol) + '6:' + columnToLetter(changeCol) + '36');
  const rules = sheet.getConditionalFormatRules();

  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.0001).setFontColor(colors.changeUpColor).setRanges([range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThanOrEqualTo(-0.0001).setFontColor(colors.changeDownColor).setRanges([range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(-0.0001, 0.0001).setFontColor(colors.changeFlatColor).setRanges([range]).build());
  sheet.setConditionalFormatRules(rules);
}

function addMarginRatioConditionalFormat(sheet, ratioCol) {
  const colors = getColorSettings();
  const range = sheet.getRange(columnToLetter(ratioCol) + MARGIN_FIRST_ROW + ':' + columnToLetter(ratioCol) + (MARGIN_FIRST_ROW + 30));
  const rules = sheet.getConditionalFormatRules();

  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(1).setFontColor(colors.ratioHighColor).setRanges([range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(1).setFontColor(colors.ratioLowColor).setRanges([range]).build());
  sheet.setConditionalFormatRules(rules);
}

/** 新しい銘柄の前日比列と信倍率列の条件付き書式を、1回の読み書きでまとめて追加する */
function addStockConditionalFormats_(sheet, changeCol, ratioCol) {
  const colors = getColorSettings();
  const changeRange = sheet.getRange(columnToLetter(changeCol) + '6:' + columnToLetter(changeCol) + '36');
  const ratioRange = sheet.getRange(columnToLetter(ratioCol) + MARGIN_FIRST_ROW + ':' + columnToLetter(ratioCol) + (MARGIN_FIRST_ROW + 30));
  const rules = sheet.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.0001).setFontColor(colors.changeUpColor).setRanges([changeRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThanOrEqualTo(-0.0001).setFontColor(colors.changeDownColor).setRanges([changeRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(-0.0001, 0.0001).setFontColor(colors.changeFlatColor).setRanges([changeRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(1).setFontColor(colors.ratioHighColor).setRanges([ratioRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(1).setFontColor(colors.ratioLowColor).setRanges([ratioRange]).build());
  sheet.setConditionalFormatRules(rules);
}

function fillHolidayMarksForNewStock(sheet, col) {
  const parsed = parseYearMonthFromSheetName(sheet.getName());
  if (!parsed) return;

  const year = parsed.year;
  const month = parsed.month;
  const daysInMonth = getDaysInMonth(year, month);
  const existingValues = sheet.getRange(6, col, daysInMonth, STOCK_COL_WIDTH).getValues();
  const newValues = existingValues.map(row => row.slice());
  const holidayRowIndexes = [];

  for (let d = 1; d <= daysInMonth; d++) {
    const date = new Date(year, month - 1, d);
    if ((date.getDay() === 0 || date.getDay() === 6 || isJapaneseHoliday(date)) && existingValues[d - 1][0] === '') {
      newValues[d - 1] = [HOLIDAY_LABEL, HOLIDAY_LABEL, HOLIDAY_LABEL];
      holidayRowIndexes.push(d - 1);
    }
  }

  sheet.getRange(6, col, daysInMonth, STOCK_COL_WIDTH).setValues(newValues);

  // 信用残の日次行(39行目〜)にも同じ日に休マークを入れる
  let marginValues = null;
  if (sheet.getMaxRows() >= MARGIN_FIRST_ROW + daysInMonth - 1) {
    marginValues = sheet.getRange(MARGIN_FIRST_ROW, col, daysInMonth, STOCK_COL_WIDTH).getValues();
    holidayRowIndexes.forEach(i => { if (marginValues[i][0] === '') marginValues[i] = [HOLIDAY_LABEL, HOLIDAY_LABEL, HOLIDAY_LABEL]; });
    sheet.getRange(MARGIN_FIRST_ROW, col, daysInMonth, STOCK_COL_WIDTH).setValues(marginValues);
  }

  // 休場日の背景色は、連続する日をまとめた範囲の一覧にして1回で塗る(範囲ごとにsetBackgroundを呼ばない)
  const a1List = [];
  const colA1 = columnToLetter(col), colEndA1 = columnToLetter(col + STOCK_COL_WIDTH - 1);
  const addRun = (startIdx, endIdx) => {
    a1List.push(colA1 + (6 + startIdx) + ':' + colEndA1 + (6 + endIdx));
    if (marginValues) a1List.push(colA1 + (MARGIN_FIRST_ROW + startIdx) + ':' + colEndA1 + (MARGIN_FIRST_ROW + endIdx));
  };
  let runStart = -1;
  holidayRowIndexes.forEach((idx, i) => {
    if (runStart === -1) runStart = idx;
    const next = holidayRowIndexes[i + 1];
    if (next !== idx + 1) { addRun(runStart, idx); runStart = -1; }
  });
  if (a1List.length > 0) sheet.getRangeList(a1List).setBackground(HOLIDAY_COLOR);
}

/**
 * 当月シートの不足分の終値・前日比・売買代金を埋める(新規銘柄追加時用)
 * 取得自体はfetchHistoricalClosesで行う
 */
function fillMissingClosesForCode(code, prefetched, knownCol) {
  const sheet = getTargetSheet(new Date());
  if (!sheet || !parseYearMonthFromSheetName(sheet.getName())) return;
  const {year, month} = parseYearMonthFromSheetName(sheet.getName());
  let targetCol = knownCol || null; // 銘柄追加の直後は列が分かっているので、2行目を読み直さない
  if (!targetCol) {
    const lastCol = sheet.getLastColumn(), codeRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
    for (let col = STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) if (String(codeRow[col - 1]) === String(code)) { targetCol = col; break; }
  }
  if (!targetCol) return;

  const today = new Date(), startDate = new Date(year, month - 1, 1);
  startDate.setDate(startDate.getDate() - 7);
  // prefetched: 取得済みの日次チャート(3年分など、当月1日の7日前〜今日を含むもの)。無ければここで取得する
  const fetched = prefetched || fetchHistoricalCloses(code, startDate, today);
  if (!fetched) return;
  const dataMap = fetched.dateMap, splits = fetched.splits;

  const daysInMonth = getDaysInMonth(year, month), existingValues = sheet.getRange(6, targetCol, daysInMonth, STOCK_COL_WIDTH).getValues();
  const newValues = existingValues.map(row => row.slice());
  let prevCloseTracker = null;

  for (let d = 1; d <= daysInMonth; d++) {
    const date = new Date(year, month - 1, d);
    if (date > today) break;
    const rowIndex = d - 1, existingClose = existingValues[rowIndex][0], existingChange = existingValues[rowIndex][1];
    let price;

    if (typeof existingClose === 'number') price = existingClose;
    else if (existingClose === HOLIDAY_LABEL) continue;
    else {
      const dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
      const dayData = dataMap[dateStr];
      if (!dayData) continue;
      price = newValues[rowIndex][0] = dayData.close;
      if (dayData.volume !== null && existingValues[rowIndex][2] === '') {
        const splitFactor = computeSplitCorrectionFactor_(dateStr, splits);
        newValues[rowIndex][2] = roundTradingValueToOkuUnit_(dayData.close, dayData.volume, splitFactor);
      }
    }

    if (!existingChange) {
      let prevClose = prevCloseTracker !== null ? prevCloseTracker : (d === 1 ? (Object.keys(dataMap).sort().reverse().find(dStr => dStr < Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd')) ? dataMap[Object.keys(dataMap).sort().reverse().find(dStr => dStr < Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd'))].close : null) : null);
      if (prevClose) newValues[rowIndex][1] = price / prevClose - 1;
    }
    prevCloseTracker = price;
  }
  sheet.getRange(6, targetCol, daysInMonth, STOCK_COL_WIDTH).setValues(newValues);
}

/**
 * 当月シートの不足分の信用残・信用倍率を埋める(新規銘柄追加時用)
 * JPXの日次信用残PDFを日ごとに読むため時間がかかることがある。途中で止まった分は毎日の自動更新で埋まる。
 */
function fillMissingMarginForCode(code, knownCol) {
  try {
    return fillMarginForCode(code, knownCol);
  } catch (e) {
    Logger.log('信用残の補完に失敗(' + code + '): ' + e);
    return '';
  }
}

/**
 * プライム/スタンダード/グロースの市場ブロック(C:E・F:H・I:K)の見出し・結合・スパークライン・条件付き書式を設定する。
 * テンプレートと当月シートの両方に対して実行する。何度実行しても安全(初回設定・崩れた場合の復旧の両方に使える)。
 */
function setupMarketValueLayout() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const template = ss.getSheetByName(TEMPLATE_SHEET_NAME);
  const currentSheet = getTargetSheet(new Date());
  const dbSheet = getOrCreateDatabaseSheet();
  const currentParsedMonth = currentSheet ? parseYearMonthFromSheetName(currentSheet.getName()) : null;
  const results = [];

  [template, currentSheet].filter(s => s).forEach(sheet => {
    const parsedMonth = parseYearMonthFromSheetName(sheet.getName()) || currentParsedMonth;

    MARKET_VALUE_ITEMS.forEach(item => {
      const codeCell = sheet.getRange(2, item.col, 1, MARKET_VALUE_COL_WIDTH);
      if (codeCell.isPartOfMerge()) codeCell.breakApart();
      codeCell.merge().setValue(item.code);

      const nameCell = sheet.getRange(3, item.col, 1, MARKET_VALUE_COL_WIDTH);
      if (nameCell.isPartOfMerge()) nameCell.breakApart();
      nameCell.merge().setValue(item.name);

      const chartCell = sheet.getRange(4, item.col, 1, MARKET_VALUE_COL_WIDTH);
      if (chartCell.isPartOfMerge()) chartCell.breakApart();
      chartCell.merge();
      if (parsedMonth) {
        const dbCol = getOrCreateDatabaseColumn(dbSheet, item.code);
        chartCell.setFormula(buildMarketValueSparklineFormula(columnToLetter(dbCol), parsedMonth.year, parsedMonth.month));
      }

      sheet.getRange(5, item.col, 1, MARKET_VALUE_COL_WIDTH).setValues([['出来高', '売買代金', '前日比']]);
      if (sheet.getMaxRows() >= MARGIN_HEADER_ROW) {
        sheet.getRange(MARGIN_HEADER_ROW, item.col, 1, MARKET_VALUE_COL_WIDTH).setValues([['売残', '買残', '信倍率']]);
      }

    });

    results.push(sheet.getName() + ': 設定しました');
  });

  // 条件付き書式(前日比・信用倍率の色)は、旧レイアウトのルールが残らないよう全体を作り直す
  try { reapplyColorSettingsToAllSheets(); } catch (e) {}
  return 'プライム/スタンダード/グロースのレイアウトを設定しました(' + results.join(' / ') + ')';
}

function addNewStockUI() {
  const html = HtmlService.createHtmlOutputFromFile('StockInputDialog').setWidth(350).setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, '新規銘柄を追加');
}

function addNewStockFromDialog(code, name) {
  addNewStock(code, name);
  return '銘柄コード ' + code + '(' + name + ')を追加しました';
}

/**
 * 銘柄追加に使う外部データ(業種と3年分の日次終値)を同時に取得する。
 * 3年分の日次終値は、当月シートの不足分の入力とデータベースの初期構築の両方に使う(Yahooに2回問い合わせない)。
 * 並列取得で取れなかった分は、従来の取得関数(リトライ・会社概要ページの確認付き)で取り直す。
 */
function prefetchNewStockData_(code) {
  const today = new Date();
  const startDate = new Date(today);
  startDate.setFullYear(startDate.getFullYear() - DB_BACKFILL_YEARS);
  const headers = { 'User-Agent': 'Mozilla/5.0' };
  let responses = [];
  try {
    responses = UrlFetchApp.fetchAll([
      { url: 'https://finance.yahoo.co.jp/quote/' + code + '.T', muteHttpExceptions: true, headers: headers },
      { url: fetchHistoricalChart(code, startDate, today), muteHttpExceptions: true, headers: headers }
    ]);
  } catch (e) {
    responses = [];
  }
  const ok = r => r && r.getResponseCode() === 200;
  const industry = (ok(responses[0]) && parseIndustryFromHtml_(code, responses[0].getContentText())) || fetchStockIndustry(code) || '';
  const history = (ok(responses[1]) && parseHistoricalChart(responses[1].getContentText())) || fetchHistoricalCloses(code, startDate, today);
  return { industry: industry, history: history };
}

function addNewStock(code, name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const template = ss.getSheetByName(TEMPLATE_SHEET_NAME);
  const currentSheet = getTargetSheet(new Date());
  const templateIsActive = template && ss.getActiveSheet().getName() === template.getName();
  // 業種と3年分の日次終値は最初に1回だけ(同時に)取得し、テンプレート・当月シート・データベースで使い回す
  const prefetched = prefetchNewStockData_(code);
  const industry = prefetched.industry;

  function processTemplate() {
    if (!template) return;
    const addedTemplateCol = addStockColumnToSheet(template, code, name, industry);
    if (addedTemplateCol) template.setActiveRange(template.getRange(1, addedTemplateCol + STOCK_COL_WIDTH + 2));
  }

  function processCurrentSheet() {
    if (!currentSheet) return;
    const addedCol = addStockColumnToSheet(currentSheet, code, name, industry);
    if (addedCol) {
      currentSheet.setActiveRange(currentSheet.getRange(1, addedCol + STOCK_COL_WIDTH + 2));
      fillHolidayMarksForNewStock(currentSheet, addedCol);
      fillMissingClosesForCode(code, prefetched.history, addedCol);
      backfillDatabaseForCode(code, prefetched.history);
      fillMissingMarginForCode(code, addedCol);
    }
  }

  if (templateIsActive) { processCurrentSheet(); processTemplate(); }
  else { processTemplate(); processCurrentSheet(); }
}

function addStockColumnToSheet(sheet, code, name, industry) {
  const lastCol = sheet.getLastColumn();
  const codeRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];

  for (let col = STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) {
    if (String(codeRow[col - 1]) === String(code)) return null;
  }

  let lastUsedCol = STOCK_START_COL - STOCK_COL_WIDTH; // 「まだ通常銘柄がない」ことを表す基準値(+STOCK_COL_WIDTHでSTOCK_START_COLになる)
  for (let col = STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) {
    if (codeRow[col - 1]) lastUsedCol = col;
  }
  const newCol = lastUsedCol + STOCK_COL_WIDTH;
  const currentMaxCols = sheet.getMaxColumns();
  const requiredMinCols = newCol + STOCK_COL_WIDTH - 1;

  if (currentMaxCols < requiredMinCols) {
    sheet.insertColumnsAfter(currentMaxCols, requiredMinCols - currentMaxCols + (10 * STOCK_COL_WIDTH));
  }
  sheet.insertColumnsBefore(newCol, STOCK_COL_WIDTH);

  const templateSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEMPLATE_SHEET_NAME);
  if (templateSheet) {
    // 書式コピー元: テンプレートの実際の銘柄列(STOCK_START_COLの1銘柄分)をそのまま使う
    //   2〜5行目(コード/業種/銘柄名/スパークライン/見出し)の4行 ← 値・結合は直後に明示的に再設定するが、
    //     枠線や背景色などの書式はここでコピーしておかないと列挿入時の引き継ぎ任せになってしまうため追加
    //   日次データ+予備行: 6〜37行目(32行) / 信用残見出し+日次データ: 38〜69行目(32行、見出し行の38行目も含める)
    templateSheet.getRange(2, STOCK_START_COL, 4, STOCK_COL_WIDTH).copyTo(sheet.getRange(2, newCol, 4, STOCK_COL_WIDTH), { formatOnly: true });
    templateSheet.getRange(6, STOCK_START_COL, 32, STOCK_COL_WIDTH).copyTo(sheet.getRange(6, newCol, 32, STOCK_COL_WIDTH), { formatOnly: true });
    const marginRows = Math.min(32, sheet.getMaxRows() - (MARGIN_HEADER_ROW - 1), templateSheet.getMaxRows() - (MARGIN_HEADER_ROW - 1));
    if (marginRows > 0) {
      templateSheet.getRange(MARGIN_HEADER_ROW, STOCK_START_COL, marginRows, STOCK_COL_WIDTH).copyTo(sheet.getRange(MARGIN_HEADER_ROW, newCol, marginRows, STOCK_COL_WIDTH), { formatOnly: true });
    }
  }

  // 2行目: 1列目=銘柄コード(単独セル)、2〜3列目(2列結合)=業種情報 / 3行目=銘柄名(3列結合) / 4行目=スパークライン(3列結合)
  // 先に結合だけしておき、2〜5行目の値・数式はまとめて1回で書き込む(結合の左上以外のセルには空文字を入れる)
  sheet.getRange(2, newCol + 1, 1, 2).merge();
  sheet.getRange(3, newCol, 1, STOCK_COL_WIDTH).merge();
  sheet.getRange(4, newCol, 1, STOCK_COL_WIDTH).merge();

  const dbSheet = getOrCreateDatabaseSheet();
  const dbCol = getOrCreateDatabaseColumn(dbSheet, code);
  sheet.getRange(2, newCol, 4, STOCK_COL_WIDTH).setValues([
    [code, industry || '', ''],
    [name, '', ''],
    [buildSparklineFormula(columnToLetter(dbCol)), '', ''],
    ['終値', '前日比', '売買代金']
  ]);

  if (sheet.getMaxRows() >= MARGIN_HEADER_ROW) {
    sheet.getRange(MARGIN_HEADER_ROW, newCol, 1, STOCK_COL_WIDTH).setValues([['売残', '買残', '信倍率']]);
  }

  try { addStockConditionalFormats_(sheet, newCol + 1, newCol + 2); } catch (e) {}

  return newCol;
}

/**
 * [初回のみ]既存シートの5行目見出し「出来高」を「売買代金」に置き換える。
 * (出来高カラムの中身自体は「終値前日比update」実行時に新しい計算式で上書きされるため、
 * ここでは見出し文字列だけを直す。既に「売買代金」になっている列はスキップされるので、再実行しても安全)
 */
function migrateVolumeHeaderToTradingValue_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < STOCK_START_COL) return 0;
  const codeRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  let count = 0;
  for (let col = STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) {
    if (!codeRow[col - 1]) continue;
    const headerCell = sheet.getRange(5, col + 2);
    if (headerCell.getValue() === '出来高') {
      headerCell.setValue('売買代金');
      count++;
    }
  }
  return count;
}

function migrateAllSheetsForTradingValueHeader() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targets = ss.getSheets().filter(s => s.getName() === TEMPLATE_SHEET_NAME || parseYearMonthFromSheetName(s.getName()));
  const messages = targets.map(s => s.getName() + ': ' + migrateVolumeHeaderToTradingValue_(s) + '件移行');
  return messages.join('\n');
}

// ==== 銘柄削除機能 ====
function getStockNameFromSheet(sheet, code) {
  if (!sheet) return null;
  const lastCol = sheet.getLastColumn();
  if (lastCol < 3) return null;
  const codeRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const nameRow = sheet.getRange(3, 1, 1, lastCol).getValues()[0];
  for (let col = STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) {
    if (String(codeRow[col - 1]) === String(code)) return String(nameRow[col - 1]);
  }
  return null;
}

function deleteStockColumnFromSheet(sheet, code) {
  if (!sheet) return false;
  const lastCol = sheet.getLastColumn();
  if (lastCol < 3) return false;
  const codeRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];

  for (let col = STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) {
    if (String(codeRow[col - 1]) === String(code)) {
      sheet.deleteColumns(col, STOCK_COL_WIDTH);
      normalizeBufferColumns(sheet);
      return true;
    }
  }
  return false;
}

function deleteStockColumnFromDatabase(code) {
  const dbSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DB_SHEET_NAME);
  if (!dbSheet) return false;
  const lastCol = dbSheet.getLastColumn();
  if (lastCol < 2) return false;
  const headerRow = dbSheet.getRange(1, 1, 1, lastCol).getValues()[0];

  for (let col = 2; col <= lastCol; col++) {
    if (String(headerRow[col - 1]) === String(code)) {
      dbSheet.deleteColumn(col);
      return true;
    }
  }
  return false;
}

function deleteStock(code) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const template = ss.getSheetByName(TEMPLATE_SHEET_NAME);
  const currentSheet = getTargetSheet(new Date());

  const results = [];
  if (currentSheet) results.push('当月シート: ' + (deleteStockColumnFromSheet(currentSheet, code) ? '削除済み' : '見つかりません'));
  if (template) results.push('テンプレート: ' + (deleteStockColumnFromSheet(template, code) ? '削除済み' : '見つかりません'));
  results.push('データベース: ' + (deleteStockColumnFromDatabase(code) ? '削除済み' : '見つかりません'));

  return '銘柄コード ' + code + ' の削除結果 → ' + results.join(' / ');
}

function deleteStockUI() {
  const ui = SpreadsheetApp.getUi();
  const codeResp = ui.prompt('銘柄を削除', '削除する銘柄コードを入力してください(例: 7203)', ui.ButtonSet.OK_CANCEL);
  if (codeResp.getSelectedButton() !== ui.Button.OK) return;

  const code = codeResp.getResponseText().trim();
  if (!code) { ui.alert('銘柄コードが入力されていません'); return; }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = getStockNameFromSheet(getTargetSheet(new Date()), code)
    || getStockNameFromSheet(ss.getSheetByName(TEMPLATE_SHEET_NAME), code)
    || '';

  const confirmResp = ui.alert(
    '削除の確認',
    '銘柄コード ' + code + (name ? '(' + name + ')' : '') + ' を、当月シート・テンプレート・データベースシートから削除します。\n'
    + 'データベースシートに蓄積された過去の終値データも失われます。元に戻せません。よろしいですか?',
    ui.ButtonSet.YES_NO
  );
  if (confirmResp !== ui.Button.YES) return;

  ui.alert(deleteStock(code));
}

function normalizeBufferColumns(sheet) {
  const lastCol = sheet.getLastColumn();
  const codeRow = sheet.getRange(2, 1, 1, Math.max(lastCol, STOCK_START_COL)).getValues()[0];
  let lastUsedCol = STOCK_START_COL - STOCK_COL_WIDTH;
  for (let col = STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) {
    if (codeRow[col - 1]) lastUsedCol = col;
  }
  const desiredMinCols = lastUsedCol + STOCK_COL_WIDTH * 10;
  const currentMaxCols = sheet.getMaxColumns();
  if (currentMaxCols < desiredMinCols) sheet.insertColumnsAfter(currentMaxCols, desiredMinCols - currentMaxCols);
  else if (currentMaxCols > desiredMinCols + STOCK_COL_WIDTH * 10) sheet.deleteColumns(desiredMinCols + 1, currentMaxCols - desiredMinCols);
}

/**
 * DateオブジェクトをSheets API送信用に'yyyy-MM-dd'文字列化
 */
function sanitizeRowsForApi(rows) {
  return rows.map(row => row.map(cell => (cell instanceof Date) ? Utilities.formatDate(cell, 'Asia/Tokyo', 'yyyy-MM-dd') : cell));
}

/**
 * 並べ替え後に中身が変わる位置(0始まり)の範囲を {first, last} で返す(変わらなければnull)。
 * 1つの銘柄を動かしただけなら、動かした位置の間だけを書き込めばよい。
 * hasData(code)がfalseの銘柄(このシートに無い銘柄)は書き込まれないので、変化なしとして扱う。
 */
function getReorderChangedRange_(currentCodes, orderSet, hasData) {
  let first = -1, last = -1;
  for (let i = 0; i < currentCodes.length; i++) {
    if (hasData(orderSet[i]) && orderSet[i] !== currentCodes[i]) {
      if (first === -1) first = i;
      last = i;
    }
  }
  return first === -1 ? null : { first: first, last: last };
}

/**
 * 並び替えペイロード生成。列位置は固定し中身だけ入れ替える。
 * 書き込むのは中身が変わる銘柄ブロックの範囲だけ(A〜K列や、動かない銘柄の列は送らない)。
 */
function getReorderPayloadForSheet(sheet, orderedCodes) {
  if (!sheet) return null;
  normalizeBufferColumns(sheet);
  const lastCol = sheet.getLastColumn();
  const lastRow = Math.max(sheet.getLastRow(), SHEET_DATA_MAX_ROW);
  if (lastCol < 3) return null;

  const codeRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const blocks = [];
  for (let col = STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) {
    if (codeRow[col - 1]) blocks.push({ code: String(codeRow[col - 1]), startCol: col });
  }

  const currentCodes = blocks.map(b => b.code);
  if (blocks.length < 2 || JSON.stringify(currentCodes) === JSON.stringify(orderedCodes)) {
    return { status: sheet.getName() + ': スキップ(変更なし)' };
  }

  const orderSet = orderedCodes.slice();
  blocks.forEach(b => { if (orderSet.indexOf(b.code) === -1) orderSet.push(b.code); });

  const headerRange = sheet.getRange(2, 1, 2, lastCol);
  const headerValues = headerRange.getValues();
  const bodyRange = sheet.getRange(5, 1, lastRow - 4, lastCol);
  const bodyValues = bodyRange.getValues();

  const headerBlocks = {}, bodyBlocks = {};
  blocks.forEach(b => {
    headerBlocks[b.code] = headerValues.map(row => row.slice(b.startCol - 1, b.startCol - 1 + STOCK_COL_WIDTH));
    bodyBlocks[b.code] = bodyValues.map(row => row.slice(b.startCol - 1, b.startCol - 1 + STOCK_COL_WIDTH));
  });

  for (let i = 0; i < blocks.length; i++) {
    const code = orderSet[i];
    if (!headerBlocks[code]) continue;
    const targetCol = STOCK_START_COL + i * STOCK_COL_WIDTH;
    headerValues.forEach((row, r) => headerBlocks[code][r].forEach((v, k) => { row[targetCol - 1 + k] = v; }));
    bodyValues.forEach((row, r) => bodyBlocks[code][r].forEach((v, k) => { row[targetCol - 1 + k] = v; }));
  }

  const changed = getReorderChangedRange_(currentCodes, orderSet, code => !!headerBlocks[code]);
  if (!changed) return { status: sheet.getName() + ': スキップ(変更なし)' };
  const firstCol = STOCK_START_COL + changed.first * STOCK_COL_WIDTH;
  const endCol = STOCK_START_COL + (changed.last + 1) * STOCK_COL_WIDTH - 1;
  const pickCols = rows => sanitizeRowsForApi(rows.map(row => row.slice(firstCol - 1, endCol)));
  const colRange = (fromRow, toRow) => "'" + sheet.getName() + "'!" + columnToLetter(firstCol) + fromRow + ':' + columnToLetter(endCol) + toRow;

  return {
    status: sheet.getName() + ': 並び替え準備完了',
    payloads: [
      { range: colRange(2, 3), values: pickCols(headerValues) },
      { range: colRange(5, lastRow), values: pickCols(bodyValues) }
    ]
  };
}

/**
 * DB側の列順を表示順に揃える
 */
function getReorderPayloadForDatabase(dbSheet, orderedCodes) {
  if (!dbSheet) return { status: 'データベース: 見つかりません' };
  const lastCol = dbSheet.getLastColumn();
  const lastRow = dbSheet.getLastRow();
  if (lastCol < 3 || lastRow < 1) return { status: 'データベース: スキップ' };

  const dataRange = dbSheet.getRange(1, 1, lastRow, lastCol);
  const data = dataRange.getValues();
  const headerRow = data[0];

  const blocks = [];
  for (let col = 2; col <= lastCol; col++) {
    if (headerRow[col - 1]) blocks.push({ code: String(headerRow[col - 1]), col: col });
  }

  const currentCodes = blocks.map(b => b.code);
  if (blocks.length < 2 || JSON.stringify(currentCodes) === JSON.stringify(orderedCodes)) {
    return { status: 'データベース: スキップ(変更なし)' };
  }

  const orderSet = orderedCodes.slice();
  blocks.forEach(b => { if (orderSet.indexOf(b.code) === -1) orderSet.push(b.code); });

  const colBlocks = {};
  blocks.forEach(b => { colBlocks[b.code] = data.map(row => row[b.col - 1]); });

  for (let i = 0; i < blocks.length; i++) {
    const code = orderSet[i];
    if (!colBlocks[code]) continue;
    const targetCol = 2 + i;
    data.forEach((row, r) => { row[targetCol - 1] = colBlocks[code][r]; });
  }

  // 書き込むのは中身が変わる列の範囲だけ(日付のA列や、動かない銘柄の列は送らない)
  const changed = getReorderChangedRange_(currentCodes, orderSet, code => !!colBlocks[code]);
  if (!changed) return { status: 'データベース: スキップ(変更なし)' };
  const firstCol = 2 + changed.first, endCol = 2 + changed.last;

  return {
    status: 'データベース: 並び替え準備完了',
    payloads: [
      {
        range: "'" + dbSheet.getName() + "'!" + columnToLetter(firstCol) + '1:' + columnToLetter(endCol) + lastRow,
        values: sanitizeRowsForApi(data.map(row => row.slice(firstCol - 1, endCol)))
      }
    ]
  };
}

/**
 * 全シート分のペイロードをSheets APIで一括更新(要:サービスでGoogle Sheets API有効化)
 */
function reorderBothSheets(orderedCodes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = getTargetSheet(new Date());
  const template = ss.getSheetByName(TEMPLATE_SHEET_NAME);
  const dbSheet = ss.getSheetByName(DB_SHEET_NAME);

  const messages = [];
  const allPayloads = [];

  const currResult = getReorderPayloadForSheet(currentSheet, orderedCodes);
  if (currResult) { messages.push(currResult.status); if (currResult.payloads) allPayloads.push(...currResult.payloads); }

  const tmplResult = getReorderPayloadForSheet(template, orderedCodes);
  if (tmplResult) { messages.push(tmplResult.status); if (tmplResult.payloads) allPayloads.push(...tmplResult.payloads); }

  const dbResult = getReorderPayloadForDatabase(dbSheet, orderedCodes);
  messages.push(dbResult.status);
  if (dbResult.payloads) allPayloads.push(...dbResult.payloads);

  if (allPayloads.length === 0) return messages.join(' / ');

  try {
    Sheets.Spreadsheets.Values.batchUpdate({
      valueInputOption: 'USER_ENTERED',
      data: allPayloads
    }, ss.getId());
  } catch (e) {
    return 'APIエラー: Sheets APIが有効になっていない可能性があります(Apps Scriptエディタ→サービス→Google Sheets APIを追加してください)。\n詳細: ' + e.message;
  }

  if (currentSheet && currResult && currResult.payloads) reapplySparklineFormulas(currentSheet);
  if (template && tmplResult && tmplResult.payloads) reapplySparklineFormulas(template);

  return '高速並び替え完了(API一括更新): ' + messages.join(' / ');
}

function reorderStocksUI() {
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('ReorderDialog').setWidth(360).setHeight(800), '銘柄の並び替え');
}

function getCurrentStockOrder() {
  const sheet = getTargetSheet(new Date());
  if (!sheet) return [];
  const lastCol = sheet.getLastColumn();
  const headRows = sheet.getRange(2, 1, 2, lastCol).getValues(); // 2行目=銘柄コード、3行目=銘柄名
  const codeRow = headRows[0], nameRow = headRows[1];
  const list = [];
  for (let col = STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) {
    if (codeRow[col - 1]) list.push({ code: String(codeRow[col - 1]), name: String(nameRow[col - 1]) });
  }
  return list;
}

// ==== 色の設定 ====
function reapplyColorSettingsToAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), colors = getColorSettings();
  [ss.getSheetByName(TEMPLATE_SHEET_NAME), getTargetSheet(new Date())].filter(s => s).forEach(sheet => {
    sheet.setConditionalFormatRules([]);
    const sheetLastCol = sheet.getLastColumn();
    const codeRow = sheet.getRange(2, 1, 1, sheetLastCol).getValues()[0], rules = [];
    // 前日比の列と信用倍率の列: 市場ブロックは[出来高,売買代金,前日比]/[売残,買残,信倍率]、銘柄は[終値,前日比,売買代金]/[売残,買残,信倍率]
    const targets = MARKET_VALUE_ITEMS.map(item => ({ changeCol: item.col + MARKET_CHANGE_OFFSET, ratioCol: item.col + 2 }));
    for (let col = STOCK_START_COL; col <= sheetLastCol; col += STOCK_COL_WIDTH) {
      if (codeRow[col - 1]) targets.push({ changeCol: col + 1, ratioCol: col + 2 });
    }
    const marginLastRow = MARGIN_FIRST_ROW + 30;
    targets.forEach(t => {
      const cRng = sheet.getRange(columnToLetter(t.changeCol) + '6:' + columnToLetter(t.changeCol) + '36'), rRng = sheet.getRange(columnToLetter(t.ratioCol) + MARGIN_FIRST_ROW + ':' + columnToLetter(t.ratioCol) + marginLastRow);
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.0001).setFontColor(colors.changeUpColor).setRanges([cRng]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThanOrEqualTo(-0.0001).setFontColor(colors.changeDownColor).setRanges([cRng]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(-0.0001, 0.0001).setFontColor(colors.changeFlatColor).setRanges([cRng]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(1).setFontColor(colors.ratioHighColor).setRanges([rRng]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(1).setFontColor(colors.ratioLowColor).setRanges([rRng]).build());
    });
    sheet.setConditionalFormatRules(rules);
  });
}

function getColorSettings() {
  const saved = PropertiesService.getScriptProperties().getProperty('COLOR_SETTINGS');
  return saved ? Object.assign({ changeUpColor: '#ff0000', changeDownColor: '#4a90d9', changeFlatColor: '#000000', ratioHighColor: '#ff0000', ratioLowColor: '#6495ed' }, JSON.parse(saved)) : { changeUpColor: '#ff0000', changeDownColor: '#4a90d9', changeFlatColor: '#000000', ratioHighColor: '#ff0000', ratioLowColor: '#6495ed' };
}

function saveColorSettings(s) { PropertiesService.getScriptProperties().setProperty('COLOR_SETTINGS', JSON.stringify(s)); }

function saveAndApplyColorSettings(s) { saveColorSettings(s); return reapplyColorSettingsToAllSheets(); }

function colorSettingsUI() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('ColorSettingsDialog').setWidth(600).setHeight(750), '色の設定'); }