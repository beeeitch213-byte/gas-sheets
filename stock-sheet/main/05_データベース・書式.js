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
 * DBシートの見出し行と日付列を1回で読み込む(1回の更新の中で、前日比の計算とDB書き込みに使い回す)。
 * dateKeys[i] は (i+2)行目の日付を'yyyy-MM-dd'にしたもの(日付でない行はnull)。
 */
function readDatabaseIndex_(dbSheet) {
  const lastRow = dbSheet.getLastRow();
  const lastCol = Math.max(dbSheet.getLastColumn(), 1);
  const headerRow = dbSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const dateKeys = lastRow >= 2
    ? dbSheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r => (r[0] instanceof Date ? Utilities.formatDate(r[0], 'Asia/Tokyo', 'yyyy-MM-dd') : null))
    : [];
  return { lastRow: lastRow, lastCol: lastCol, headerRow: headerRow, dateKeys: dateKeys };
}

/**
 * DBシートの中から、beforeDateStr(例 '2026-10-01')より前で最も新しい日付の行を探し、
 * {コード: 終値} のマップを返す(前日比の計算に使う「前営業日の実終値」)。
 * 再実行などで当日分が既にDBに書き込まれていた場合は、その行を飛ばして1つ前の行を使う。
 * dbIndexはreadDatabaseIndex_の結果(省略時はその場で読み込む)。
 */
function getPreviousCloseMap(dbSheet, beforeDateStr, dbIndex) {
  const idx = dbIndex || readDatabaseIndex_(dbSheet);
  if (idx.lastRow < 2) return {};

  // 対象日より前で一番新しい日付の行を探す
  // (過去日付を後から追加するとDBの並びが日付順でなくなることがあるため、行の位置ではなく日付で判定する)
  let bestRow = -1, bestKey = '';
  idx.dateKeys.forEach((key, i) => {
    if (key && key < beforeDateStr && key > bestKey) { bestKey = key; bestRow = i + 2; }
  });
  if (bestRow === -1) return {};
  const rowVals = dbSheet.getRange(bestRow, 1, 1, idx.lastCol).getValues()[0];
  const map = {};
  for (let c = 2; c <= idx.lastCol; c++) {
    if (idx.headerRow[c - 1] && typeof rowVals[c - 1] === 'number') map[String(idx.headerRow[c - 1])] = rowVals[c - 1];
  }
  return map;
}

/** dbIndexはreadDatabaseIndex_の結果(省略時はその場で読み込む)。読み込み後にDBの行・列が増えていないこと */
function writeClosesToDatabaseForDate(date, priceMap, dbIndex) {
  const codes = Object.keys(priceMap);
  if (codes.length === 0) return;

  const dbSheet = getOrCreateDatabaseSheet();
  const idx = dbIndex || readDatabaseIndex_(dbSheet);
  const dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');

  let lastCol = idx.lastCol;
  const colMap = {};
  idx.headerRow.forEach((h, i) => { if (h) colMap[String(h)] = i + 1; });

  const newCodes = codes.filter(c => !colMap[c]);
  if (newCodes.length > 0) {
    dbSheet.getRange(1, lastCol + 1, 1, newCodes.length).setValues([newCodes]);
    newCodes.forEach((c, i) => { colMap[c] = lastCol + 1 + i; });
    lastCol += newCodes.length;
  }

  const found = idx.dateKeys.indexOf(dateStr);
  let targetRow, rowValues;
  if (found !== -1) {
    targetRow = 2 + found;
    rowValues = dbSheet.getRange(targetRow, 1, 1, lastCol).getValues()[0];
  } else {
    // 新しい日付の行は空なので、読み込まずに日付だけ入れた行を作る
    targetRow = Math.max(idx.lastRow, 1) + 1;
    rowValues = new Array(lastCol).fill('');
    rowValues[0] = date;
  }
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

/**
 * DBの指定列を直接参照する終値スパークライン数式を組み立てる
 */
function buildSparklineFormula(dbColLetter) {
  const dbRef = "'" + DB_SHEET_NAME + "'";
  return '=SPARKLINE(FILTER(' + dbRef + '!$' + dbColLetter + '$2:$' + dbColLetter
    + ', ' + dbRef + '!$A$2:$A>=EDATE(TODAY(),-' + CHART_DISPLAY_MONTHS + ')), '
    + '{"charttype","line";"color","#4285f4";"empty","ignore"})';
}

/**
 * 市場別売買代金(プライム/スタンダード/グロース)用: DBの指定列を、シート自身の月(当月1日〜月末)の範囲だけに絞って参照するスパークライン数式
 */
function buildMarketValueSparklineFormula(dbColLetter, year, month) {
  const dbRef = "'" + DB_SHEET_NAME + "'";
  const startStr = Utilities.formatDate(new Date(year, month - 1, 1), 'Asia/Tokyo', 'yyyy-MM-dd');
  return '=SPARKLINE(FILTER(' + dbRef + '!$' + dbColLetter + '$2:$' + dbColLetter
    + ', ' + dbRef + '!$A$2:$A>=DATE(' + year + ',' + month + ',1), ' + dbRef + '!$A$2:$A<EDATE(DATE(' + year + ',' + month + ',1),1)), '
    + '{"charttype","line";"color","#34a853";"empty","ignore"})';
}

/**
 * 全銘柄の4行目にスパークライン数式を再設定する(並び替え後の作り直し用)。
 * プライム/スタンダード/グロース(C〜K列)のスパークラインも合わせて再設定する。
 */
function reapplySparklineFormulas(sheet) {
  if (!sheet) return 'シートが見つかりません';
  const lastCol = sheet.getLastColumn();
  if (lastCol < 3) return sheet.getName() + ': スキップ';

  const codeRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const dbSheet = getOrCreateDatabaseSheet();

  const dbLastCol = Math.max(dbSheet.getLastColumn(), 1);
  const dbHeaderRow = dbSheet.getRange(1, 1, 1, dbLastCol).getValues()[0];
  const dbColMap = {};
  for (let c = 2; c <= dbLastCol; c++) {
    if (dbHeaderRow[c - 1]) dbColMap[String(dbHeaderRow[c - 1])] = c;
  }

  const formulas = new Array(lastCol).fill('');
  let count = 0;

  const parsedMonth = parseYearMonthFromSheetName(sheet.getName());
  if (parsedMonth) {
    MARKET_VALUE_ITEMS.forEach(item => {
      let dbCol = dbColMap[item.code];
      if (!dbCol) { dbCol = getOrCreateDatabaseColumn(dbSheet, item.code); dbColMap[item.code] = dbCol; }
      formulas[item.col - 1] = buildMarketValueSparklineFormula(columnToLetter(dbCol), parsedMonth.year, parsedMonth.month);
    });
  }

  for (let col = STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) {
    const code = codeRow[col - 1];
    if (!code) continue;
    let dbCol = dbColMap[String(code)];
    if (!dbCol) {
      dbCol = getOrCreateDatabaseColumn(dbSheet, String(code));
      dbColMap[String(code)] = dbCol;
    }
    formulas[col - 1] = buildSparklineFormula(columnToLetter(dbCol));
    count++;
  }

  sheet.getRange(4, 1, 1, lastCol).setFormulas([formulas]);
  return sheet.getName() + ': ' + count + '銘柄のスパークラインを再設定';
}

/**
 * データベースシートを削除→再構築した直後など、テンプレート・全月シートのスパークラインが
 * #REF!になってしまった場合の復旧用。テンプレート+月名のシート全てに再設定する。
 * (削除されたシートへの参照は、同名のシートを新規作成しても自動では直らないため、
 *  数式そのものを書き直す必要がある)
 */
function reapplyAllSparklineFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targets = ss.getSheets().filter(s => s.getName() === TEMPLATE_SHEET_NAME || parseYearMonthFromSheetName(s.getName()));
  const messages = targets.map(s => reapplySparklineFormulas(s));
  return messages.join('\n');
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