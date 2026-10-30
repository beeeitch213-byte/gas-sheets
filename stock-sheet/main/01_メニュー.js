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
      .addSeparator()
      .addItem('次月シートを作成', 'runCreateNextMonthSheetFromMenu')
      .addItem('決算予定を更新(今日〜次月末)', 'runUpdateKessanScheduleFromMenu'))
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

/** 「次月シートを作成」: 毎日の更新では月末の6日前から自動で作るが、それより前に作りたいとき用 */
function runCreateNextMonthSheetFromMenu() {
  runLibFromMenu_(() => createNextMonthSheetNow(), '次月シートを作成しました', '次月シートの作成に失敗しました');
}

/** 「決算予定を更新」: 月末の6日前からは毎日自動で更新するが、それより前に見たいとき用 */
function runUpdateKessanScheduleFromMenu() {
  runLibFromMenu_(() => updateKessanSchedule(), '決算予定を更新しました', '決算予定の更新に失敗しました');
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
