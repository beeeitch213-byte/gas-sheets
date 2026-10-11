function onOpen() {
  try {
    const today = new Date();
    const sheet = getTargetSheet(today);
    if (sheet) sheet.activate();
  } catch (e) {
    Logger.log('onOpenでのシート切り替えに失敗: ' + e);
  }
  addStockMenu();
}

// ==== メニュー・銘柄追加 ====
function addStockMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('機能')
    .addItem('新規銘柄を追加', 'addNewStockUI')
    .addItem('銘柄を削除', 'deleteStockUI')
    .addItem('銘柄を並び替え', 'reorderStocksUI')
    .addItem('業種情報を一括取得(銘柄選択)', 'addIndustryBulkUI')
    .addItem('色の設定', 'colorSettingsUI')
    .addSeparator()
    .addSubMenu(ui.createMenu('データ更新')
      .addItem('終値・前日比を更新(日付入力)※市場別売買代金・出来高を含む', 'updateCloseByDateUI')
      .addItem('空売り集計を更新(日付入力)', 'updateShortSellByDateUI')
      .addItem('信用残・信用倍率を更新(日付入力)', 'updateMarginByDateUI')
      .addItem('投資部門別売買状況更新(日付入力)', 'updateJpxStatsByDateUI'))
    .addSubMenu(ui.createMenu('データ修正')
      .addItem('データベースを全銘柄分初期構築', 'runInitializeDatabaseFromMenu')
      .addItem('基準セルの書式を全銘柄に再適用', 'runReapplyFormatsFromMenu')
      .addItem('スパークラインを全シートで再設定(#REF!復旧用)', 'runReapplyAllSparklineFormulasFromMenu')
      .addSeparator()
      .addItem('市場別売買代金レイアウトを設定', 'runSetupMarketValueLayoutFromMenu')
      .addItem('空売り集計シートの見出しを設定', 'runSetupShortSellHeadersFromMenu')
      .addItem('業種情報レイアウトへ移行', 'runMigrateAllSheetsForIndustryLayoutFromMenu')
      .addItem('出来高見出しを売買代金に移行', 'runMigrateAllSheetsForTradingValueHeaderFromMenu')
      .addItem('売買状況シートを新しい週が上の順に並べ替え', 'runSortJpxStatsSheetFromMenu')
      .addItem('売買状況シートの日付表示を新レイアウトに変換', 'runMigrateJpxDateLayoutFromMenu')
      .addSeparator()
      .addItem('自動更新トリガーを設定(統合後に1回)', 'runSetupTriggersFromMenu'))
    .addToUi();
}

// ==== メニューから実行するためのラッパー ====
// 各機能の本体は他のファイルにあり、ここでは呼び出しと完了メッセージの表示だけを行う。
// データベース初期構築・投資部門別売買状況の過去分取得は時間がかかり、実行が長すぎるとタイムアウトすることがある
// (その場合は同じメニューを再度実行すれば続きから進む)。

/**
 * 関数を実行し、返ってきたメッセージ(なければ既定文)をポップアップ表示する。
 * 時間のかかる処理は実行時間の上限前に安全に止まり、再実行で続きから進む作りになっている。
 */
function runLibFromMenu_(fn, defaultMessage, failPrefix) {
  const ui = SpreadsheetApp.getUi();
  try {
    const message = fn();
    ui.alert(message || defaultMessage);
  } catch (e) {
    ui.alert(failPrefix + ': ' + e.message);
  }
}

/** 「終値・前日比を更新(日付入力)」: 単日または期間を指定するダイアログを開く */
function updateCloseByDateUI() {
  const html = HtmlService.createHtmlOutputFromFile('UpdateCloseDialog').setWidth(360).setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, '終値・前日比を更新(日付入力)');
}

/**
 * UpdateCloseDialogから呼ばれる。startDateStr/endDateStrが同じなら単日更新、
 * 異なれば期間更新として扱う(どちらも'yyyy-MM-dd'形式)。
 */
function runUpdateCloseForDateFromDialog(startDateStr, endDateStr) {
  if (!startDateStr || !endDateStr) throw new Error('日付を入力してください');
  if (startDateStr === endDateStr) return updateCloseForDateInput(startDateStr);
  return updateCloseForDateRangeInput(startDateStr, endDateStr);
}

/** 「空売り集計を更新(日付入力)」: 単日または期間を指定するダイアログを開く */
function updateShortSellByDateUI() {
  const html = HtmlService.createHtmlOutputFromFile('UpdateShortSellDialog').setWidth(360).setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, '空売り集計を更新(日付入力)');
}

function runUpdateShortSellForDateFromDialog(startDateStr, endDateStr) {
  if (!startDateStr || !endDateStr) throw new Error('日付を入力してください');
  if (startDateStr === endDateStr) return updateShortSellForDateInput(startDateStr);
  return updateShortSellForDateRangeInput(startDateStr, endDateStr);
}

function runSetupShortSellHeadersFromMenu() {
  runLibFromMenu_(() => setupShortSellSheetHeaders(), '空売り集計シートの見出しを設定しました', '設定に失敗しました');
}

/** 「信用残・信用倍率を更新(日付入力)」: 単日または期間を指定するダイアログを開く */
function updateMarginByDateUI() {
  const html = HtmlService.createHtmlOutputFromFile('UpdateMarginDialog').setWidth(360).setHeight(340);
  SpreadsheetApp.getUi().showModalDialog(html, '信用残・信用倍率を更新(日付入力)');
}

/** UpdateMarginDialogから呼ばれる('yyyy-MM-dd'形式。終了日が空なら開始日1日分) */
function runUpdateMarginForDateFromDialog(startDateStr, endDateStr) {
  if (!startDateStr) throw new Error('日付を入力してください');
  return updateMarginForDateRangeInput(startDateStr, endDateStr || startDateStr);
}

function runInitializeDatabaseFromMenu() {
  runLibFromMenu_(() => initializeDatabaseForAllStocks(), 'データベースの初期構築が完了しました', '処理が途中で停止しました。もう一度実行すると続きから進みます');
}

/** 「投資部門別売買状況更新(日付入力)」: 年・週を選んで取得するダイアログを開く */
function updateJpxStatsByDateUI() {
  const html = HtmlService.createHtmlOutputFromFile('JpxStatsDialog').setWidth(360).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, '投資部門別売買状況更新');
}

function runMigrateJpxDateLayoutFromMenu() {
  runLibFromMenu_(() => migrateJpxStatsDateLayout(), '日付表示の変換が完了しました', '変換に失敗しました');
}

function runSortJpxStatsSheetFromMenu() {
  runLibFromMenu_(() => sortJpxStatsSheetByDate(), '並べ替えが完了しました', '並べ替えに失敗しました');
}

/** JpxStatsDialogの年プルダウンから呼ばれる。その年に存在する週の一覧(新しい順)を返す */
function listJpxWeeksForYearUI(year) {
  return listJpxWeeksForYear(year);
}

/**
 * JpxStatsDialogから呼ばれる。開始週(〜終了週)が含まれる期間をすべて判定して取得する。
 * forceRefetchがtrueなら、既にシートにある週も再取得して上書きする。
 */
function runUpdateJpxStatsForDateFromDialog(startDateStr, endDateStr, forceRefetch) {
  if (!startDateStr) throw new Error('週を選択してください');
  return updateJpxStatsForDateRangeInput(startDateStr, endDateStr || startDateStr, !!forceRefetch);
}

function runReapplyFormatsFromMenu() {
  runLibFromMenu_(() => reapplyAllFormatsFromTemplate(), '基準セルの書式を再適用しました', '処理に失敗しました');
}

function runSetupMarketValueLayoutFromMenu() {
  runLibFromMenu_(() => setupMarketValueLayout(), '市場別売買代金レイアウトを設定しました', '設定に失敗しました');
}

function runReapplyAllSparklineFormulasFromMenu() {
  const ui = SpreadsheetApp.getUi();
  try {
    ui.alert(reapplyAllSparklineFormulas());
  } catch (e) {
    ui.alert('スパークライン再設定でエラーが発生しました: ' + e.message);
  }
}

function runMigrateAllSheetsForIndustryLayoutFromMenu() {
  const ui = SpreadsheetApp.getUi();
  try {
    ui.alert(migrateAllSheetsForIndustryLayout());
  } catch (e) {
    ui.alert('業種レイアウト移行でエラーが発生しました: ' + e.message);
  }
}

function runMigrateAllSheetsForTradingValueHeaderFromMenu() {
  const ui = SpreadsheetApp.getUi();
  try {
    ui.alert(migrateAllSheetsForTradingValueHeader());
  } catch (e) {
    ui.alert('売買代金見出し移行でエラーが発生しました: ' + e.message);
  }
}

// ==== 自動更新トリガー ====
// 統合前は株式情報_DB側に設定していたトリガーを、このプロジェクトに作り直す。
// 同じ関数のトリガーが既にあれば消してから作るので、何度実行しても二重にはならない。
// 指定した「時」の中のどこかで実行される(Googleの仕様。例: 16 → 16:00〜16:59)。
// 各データの公表時刻より後の「時」を指定している。変えたい場合は下の表を編集する。
const AUTO_TRIGGERS_ = [
  { fn: 'updateDailyClose', hour: 16 },              // 大引け(15:30)後: 終値・前日比・売買代金
  { fn: 'updateMarketTradingValues', hour: 17 },     // 商況プリント２(後場、16:10頃)公表後: 市場の出来高・売買代金
  { fn: 'updateShortSellForToday', hour: 18 },       // 空売り集計PDF(17:30頃)公表後
  { fn: 'updateMarginData', hour: 19 },              // 日次信用残(前日までの未入力日を埋める)
  { fn: 'updateJpxStatsForLatestWeek', hour: 20, weekDay: 'FRIDAY' } // 投資部門別(週1回)
];

function setupTriggers() {
  const names = AUTO_TRIGGERS_.map(t => t.fn);
  ScriptApp.getProjectTriggers().forEach(tr => {
    if (names.indexOf(tr.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(tr);
  });
  AUTO_TRIGGERS_.forEach(t => {
    let b = ScriptApp.newTrigger(t.fn).timeBased();
    b = t.weekDay ? b.onWeekDay(ScriptApp.WeekDay[t.weekDay]).atHour(t.hour) : b.everyDays(1).atHour(t.hour);
    b.inTimezone('Asia/Tokyo').create();
  });
  return AUTO_TRIGGERS_.map(t => t.fn + ' : ' + (t.weekDay ? '毎週金曜' : '毎日') + ' ' + t.hour + '時台').join('\n')
    + '\n\n※株式情報_DB側のトリガーは削除してください(残すと二重に動きます)';
}

function runSetupTriggersFromMenu() {
  runLibFromMenu_(() => setupTriggers(), '自動更新トリガーを設定しました', 'トリガーの設定に失敗しました');
}