// ==== 過去の月次シートを当月に合わせる(バックグラウンド処理) ====
// 当月シートの銘柄(と並び順)に、ほかの月次シート(過去の月・作成済みの次月)を合わせる。
//  ・当月に無い銘柄の列は削除する(その月のデータも消える)。同じ銘柄の列が2つ以上あれば1つにする
//  ・当月にあって無い銘柄は列を足し、終値・前日比・売買代金(Yahoo)と信用残(ドライブの日次データ)を埋める
//  ・終値・前日比が抜けている銘柄(以前の実行で取れなかったなど)も、空いているセルだけ埋め直す
//  ・昔のレイアウトのシートは、週ごとの信用残の空いているセルも埋める(Yahoo!ファイナンスの信用残時系列から)
//  ・当月と同じ順番に並べる
// 銘柄の追加・削除・並び替えのあと(当月・テンプレートの処理が終わったら)と、「データ修正」メニューから始める。
// 処理は時間主導トリガー(1回きり)で進め、時間切れ(RUN_TIME_LIMIT_MS)になったら続きを予約し直す。
// 進み具合はスクリプトプロパティ(PAST_SYNC_STATE_KEY)に置き、小さなダイアログ(PastSyncDialog)が読んで表示する。
// 処理中に別の追加・削除・並び替えがあれば、最初からやり直す(generation が変わったら、古い実行は途中でやめる)。

const PAST_SYNC_STATE_KEY = 'PAST_SHEET_SYNC_STATE';

const PAST_SYNC_HANDLER = 'runPastSheetSync'; // 時間主導トリガーから呼ぶ関数の名前(末尾 _ の関数はトリガーから呼べない)

const PAST_SYNC_FIRST_DELAY_MS = 1000;     // 始めるときの予約(実際には1分ほど後に動く)

const PAST_SYNC_RETRY_DELAY_MS = 60 * 1000; // 時間切れ・ほかの実行中のときの、次の予約までの時間

// 実行の始めに入れておく「保険」の予約。GAS の6分の上限で実行が強制終了されても、続きがこの時間後に動く
// (正常に終われば、終わりに予約し直すか消す)
const PAST_SYNC_WATCHDOG_DELAY_MS = 8 * 60 * 1000;

function readPastSyncState_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(PAST_SYNC_STATE_KEY) || 'null'); } catch (e) { return null; }
}

function writePastSyncState_(state) {
  PropertiesService.getScriptProperties().setProperty(PAST_SYNC_STATE_KEY, JSON.stringify(state));
}

/** 実行中の処理が進み具合を保存する。別の操作で最初からやり直しになっていたら(generation が違えば)保存せず false */
function savePastSyncStateIfCurrent_(state) {
  const saved = readPastSyncState_();
  if (!saved || saved.generation !== state.generation) return false;
  writePastSyncState_(state);
  return true;
}

/**
 * 今のレイアウト(市場ブロックが3列×3でC・F・I列から、銘柄はL列から、信用残は毎日)なら true。
 * 2026/9 以前の月次シートは昔のレイアウト(市場ブロックが2列×3でC・E・G列から、銘柄はI列から、信用残は週ごと)で、
 * 列の位置を前提にする処理(過去のシートを合わせる・スパークラインの再設定・条件付き書式の作り直し)はしない。
 */
function isCurrentLayoutSheet_(sheet) {
  return sheetLayoutOf_(sheet) === 'current';
}

/** 昔のレイアウト(市場ブロックの銘柄コードが C・E・G 列)なら true */
function isOldLayoutSheet_(sheet) {
  return sheetLayoutOf_(sheet) === 'old';
}

/** シートのレイアウト: 'current' / 'old' / ''(どちらでもない)。1回の実行の中では、シートごとに1回だけ調べる */
const sheetLayoutCache_ = {};
function sheetLayoutOf_(sheet) {
  const key = sheet.getSheetId ? sheet.getSheetId() : sheet.getName();
  if (sheetLayoutCache_[key] !== undefined) return sheetLayoutCache_[key];
  const row = sheet.getRange(2, 1, 1, STOCK_START_COL).getValues()[0];
  const has = items => items.every(item => String(row[item.col - 1]) === String(item.code));
  const layout = has(MARKET_VALUE_ITEMS) ? 'current' : (has(OLD_LAYOUT_MARKET_ITEMS) ? 'old' : '');
  sheetLayoutCache_[key] = layout;
  return layout;
}

/** 銘柄の始まりの列(今のレイアウトはL列、昔のレイアウトはI列) */
function stockStartColOf_(sheet) {
  return isOldLayoutSheet_(sheet) ? OLD_LAYOUT_STOCK_START_COL : STOCK_START_COL;
}

/** 当月シート以外の月次シート(今のレイアウトと昔のレイアウト。新しい月から順) */
function listOtherMonthSheets_(ss, currentSheet) {
  return ss.getSheets()
    .map(s => ({ sheet: s, ym: parseYearMonthFromSheetName(s.getName()) }))
    .filter(x => x.ym && (!currentSheet || x.sheet.getName() !== currentSheet.getName()) && sheetLayoutOf_(x.sheet) !== '')
    .sort((a, b) => (b.ym.year * 12 + b.ym.month) - (a.ym.year * 12 + a.ym.month))
    .map(x => x.sheet);
}

/** シートの銘柄 [{code, name, industry, col}](並び順) */
function readStockBlocks_(sheet) {
  const lastCol = sheet.getLastColumn();
  const startCol = stockStartColOf_(sheet); // 昔のレイアウトはI列から
  if (lastCol < startCol) return [];
  const head = sheet.getRange(2, 1, 2, lastCol).getValues(); // 2行目=コード・業種、3行目=銘柄名
  const blocks = [];
  for (let col = startCol; col <= lastCol; col += STOCK_COL_WIDTH) {
    const code = head[0][col - 1];
    if (code) blocks.push({ code: String(code), name: String(head[1][col - 1] || ''), industry: String(head[0][col] || ''), col: col });
  }
  return blocks;
}

function deletePastSyncTriggers_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === PAST_SYNC_HANDLER)
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function schedulePastSheetSync_(delayMs) {
  deletePastSyncTriggers_();
  ScriptApp.newTrigger(PAST_SYNC_HANDLER).timeBased().after(delayMs).create();
}

/**
 * 過去の月次シートを当月に合わせる処理を(最初から)始める。処理中なら、最初からやり直す。
 * currentMessage: 当月の処理の結果(ダイアログの1行目に出す)。対象のシートが無ければ null を返す。
 */
function startPastSheetSync_(currentMessage) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const names = listOtherMonthSheets_(ss, getTargetSheet(new Date())).map(s => s.getName());
  if (names.length === 0) return null;
  const state = {
    generation: Date.now(),
    queue: names,
    total: names.length,
    done: 0,
    current: '',
    pendingFill: [], // 列を足したが中身をまだ埋め終わっていない銘柄 [{sheet, code}]
    running: true,
    currentMessage: currentMessage || '',
    errors: [],
    finishedAt: null
  };
  writePastSyncState_(state);
  schedulePastSheetSync_(PAST_SYNC_FIRST_DELAY_MS);
  return state;
}

/** 当月の処理のあとに呼ぶ: 過去分の処理を始め、進み具合を出す小さなダイアログを開く(開いていたダイアログは閉じる) */
function startPastSheetSyncWithDialog_(currentMessage) {
  const state = startPastSheetSync_(currentMessage);
  if (state) showPastSyncDialog_();
  return state;
}

function showPastSyncDialog_() {
  SpreadsheetApp.getUi().showModelessDialog(
    HtmlService.createHtmlOutputFromFile('PastSyncDialog').setWidth(340).setHeight(110), '過去のシート');
}

/** 「データ修正 → 過去のシートを当月に合わせる」 */
function runPastSheetSyncFromMenu() {
  const state = startPastSheetSyncWithDialog_('過去のシートを当月に合わせます');
  if (!state) SpreadsheetApp.getUi().alert('当月シート以外の月次シートがありません');
}

/** PastSyncDialog から呼ぶ: 進み具合 */
function getPastSheetSyncStatus() {
  const s = readPastSyncState_();
  if (!s) return { running: false, total: 0, done: 0, current: '', currentMessage: '', errors: [] };
  return { running: s.running, total: s.total, done: s.done, current: s.current, currentMessage: s.currentMessage, errors: s.errors || [] };
}

/** 時間主導トリガー(1回きり)から呼ばれる本体。時間切れになったら続きを予約し直す */
function runPastSheetSync() {
  const startMs = Date.now();
  schedulePastSheetSync_(PAST_SYNC_WATCHDOG_DELAY_MS); // 今の予約を消して保険の予約を入れる
  const lock = LockService.getScriptLock(); // この処理どうしの重なりだけを防ぐ(銘柄の列のロックは使わない。当月の操作を待たせないため)
  if (!lock.tryLock(10 * 1000)) { schedulePastSheetSync_(PAST_SYNC_RETRY_DELAY_MS); return; }
  try {
    let state = readPastSyncState_();
    if (!state || !state.running) { deletePastSyncTriggers_(); return; }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const currentSheet = getTargetSheet(new Date());
    if (!currentSheet) { state.running = false; state.errors.push('当月シートが見つかりません'); savePastSyncStateIfCurrent_(state); return; }
    const order = readStockBlocks_(currentSheet);
    const historyCache = {};

    while (state.queue.length > 0) {
      if (Date.now() - startMs > RUN_TIME_LIMIT_MS) break;
      const name = state.queue[0];
      state.current = name;
      if (!savePastSyncStateIfCurrent_(state)) return; // 別の追加・削除・並び替えで、最初からやり直しになった(新しい実行が予約済み)
      const sheet = ss.getSheetByName(name);
      let finished = true;
      if (sheet) {
        try {
          finished = syncSheetToCurrentLayout_(sheet, order, state, startMs, historyCache);
        } catch (e) {
          state.errors.push(name + ': ' + e.message);
          if (!savePastSyncStateIfCurrent_(state)) return;
        }
      }
      if (!finished) break; // 時間切れ(または、やり直しになった)。このシートは次の実行で続きから
      state.queue.shift();
      state.done++;
      if (!savePastSyncStateIfCurrent_(state)) return;
    }

    if (state.queue.length > 0) {
      const delay = state.nextDelayMs || PAST_SYNC_RETRY_DELAY_MS; // 取得に失敗したときは長めに待つ
      delete state.nextDelayMs;
      if (!savePastSyncStateIfCurrent_(state)) return;
      schedulePastSheetSync_(delay);
    } else {
      state.running = false;
      state.current = '';
      state.finishedAt = Date.now();
      if (savePastSyncStateIfCurrent_(state)) deletePastSyncTriggers_(); // 保険の予約も消す
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * 1枚のシートを当月の銘柄・並び順に合わせる。時間切れで途中までなら false(次の実行で続きから。何度実行しても同じ結果になる)。
 * order: 当月シートの銘柄 [{code, name, industry}](並び順)
 */
function syncSheetToCurrentLayout_(sheet, order, state, startMs, historyCache) {
  const layout = sheetLayoutOf_(sheet);
  if (!layout) return true; // どちらのレイアウトでもないシートは触らない
  const sheetName = sheet.getName();
  const wanted = order.map(o => o.code);
  const timeUp = () => Date.now() - startMs > RUN_TIME_LIMIT_MS;

  // 1. 当月に無い銘柄の列と、同じ銘柄が2つ以上ある列(終値が多く入っている方を残す)を削除する
  readStockBlocks_(sheet).filter(b => wanted.indexOf(b.code) === -1).forEach(b => deleteStockColumnFromSheet(sheet, b.code));
  deleteDuplicateStockBlocks_(sheet);

  // 2. 足りない銘柄の列を右端に足す(中身は 3. で埋める)
  const have = readStockBlocks_(sheet).map(b => b.code);
  // 先に「埋める予定」を保存してから列を足す(足した直後に強制終了されても、次の実行で埋められるように)。
  // 条件付き書式は1銘柄ごとには作り直さず、5. で1回だけまとめる(ルールが多いシートでは1回に何十秒もかかるため)
  for (const o of order.filter(o => have.indexOf(o.code) === -1)) {
    if (timeUp()) return false;
    if (!state.pendingFill.some(p => p.sheet === sheetName && p.code === o.code)) state.pendingFill.push({ sheet: sheetName, code: o.code });
    if (!savePastSyncStateIfCurrent_(state)) return false; // やり直しになった
    const col = addStockColumnToSheet(sheet, o.code, o.name, o.industry, { skipConditionalFormats: true });
    if (col) {
      fillHolidayMarksForNewStock(sheet, col);
      // 昔のレイアウトの信用残(38行目〜)は週ごとなので、毎日の休場日の位置で塗ったグレーを消す
      if (layout === 'old') sheet.getRange(MARGIN_FIRST_ROW, col, 31, STOCK_COL_WIDTH).setBackground(null);
    }
  }

  // 3. 足した銘柄と、終値が抜けている銘柄の終値・前日比・売買代金と信用残を埋める
  // (以前の実行で終値を取れなかった銘柄も、ここで埋め直す。埋められるのは空いているセルだけ)
  const targets = state.pendingFill.filter(p => p.sheet === sheetName).map(p => p.code);
  findStocksMissingCloses_(sheet).forEach(code => { if (targets.indexOf(code) === -1) targets.push(code); });
  for (const code of targets) {
    if (timeUp()) return false;
    const block = readStockBlocks_(sheet).find(b => b.code === code);
    if (block) {
      const history = getHistoryForPastSync_(code, historyCache, state);
      if (history === null) return false; // 取得に失敗した。少し待ってから、次の実行で取り直す
      if (history) fillMissingClosesForCode(code, history, block.col, sheet);
      // 昔のレイアウトの信用残は週ごとなので埋めない(ドライブの日次データも、その頃の分は無い)
      if (layout === 'current') {
        const marginMessage = fillMissingMarginForCode(code, block.col, sheet, startMs);
        if (String(marginMessage).indexOf('途中まで') !== -1) return false;
      }
    }
    state.pendingFill = state.pendingFill.filter(x => !(x.sheet === sheetName && x.code === code));
    if (!savePastSyncStateIfCurrent_(state)) return false; // やり直しになった
  }

  // 3b. 昔のレイアウトの、週ごとの信用残(38行目〜)の抜けを埋める
  // (ここで失敗しても、並び替えなどの続きは進める)
  if (layout === 'old') {
    try {
      if (!fillOldLayoutWeeklyMargin_(sheet, state, startMs)) return false; // 時間切れ・取得失敗(次の実行で続きから)
    } catch (e) {
      state.errors.push(sheetName + ': 週の信用残を埋められませんでした(' + e.message + ')');
    }
  }

  // 4. 当月と同じ順番に並べる(並び替えと同じ方法: 列の位置はそのままで中身を入れ替える)
  const result = getReorderPayloadForSheet(sheet, wanted);
  if (result && result.payloads) {
    Sheets.Spreadsheets.Values.batchUpdate({ valueInputOption: 'USER_ENTERED', data: result.payloads }, SpreadsheetApp.getActiveSpreadsheet().getId());
    if (layout === 'old') setOldLayoutSparklines_(sheet);
    else reapplySparklineFormulas(sheet);
  }

  // 5. 前日比・信用倍率の条件付き書式をまとめる
  if (layout === 'old') rebuildOldLayoutColorRules_(sheet);
  else rebuildStockColorRules_(sheet, true); // まとまっていれば何もしない
  return true;
}

const PAST_SYNC_FETCH_MAX_FAILURES = 3; // 1銘柄の終値の取得に、続けてこの回数失敗したらあきらめる

/**
 * 同じ銘柄の列が2つ以上あれば、終値が多く入っている方(同じなら左側)を残して、ほかを削除する。削除した数を返す。
 * (並び替えは銘柄コードで列の中身を入れ替えるため、同じ銘柄が2つあると、片方の中身で上書きされてしまう)
 */
function deleteDuplicateStockBlocks_(sheet) {
  const byCode = {};
  readStockBlocks_(sheet).forEach(b => (byCode[b.code] = byCode[b.code] || []).push(b));
  const dropCols = [];
  Object.keys(byCode).forEach(code => {
    const list = byCode[code];
    if (list.length < 2) return;
    const counts = list.map(b => sheet.getRange(6, b.col, 31, 1).getValues().filter(r => typeof r[0] === 'number').length);
    let keep = 0;
    counts.forEach((n, i) => { if (n > counts[keep]) keep = i; });
    list.forEach((b, i) => { if (i !== keep) dropCols.push(b.col); });
  });
  dropCols.sort((a, b) => b - a).forEach(col => sheet.deleteColumns(col, STOCK_COL_WIDTH)); // 右から消して、左の列の位置をずらさない
  if (dropCols.length > 0) normalizeBufferColumns(sheet);
  return dropCols.length;
}

/**
 * 終値が抜けている銘柄のコードを返す。今日より前の営業日で、終値(か前日比)のセルが空のものがある銘柄。
 * どの銘柄にも終値が入っていない日(シート全体で取れなかった日)は見ない(全銘柄を取りに行かないように)。
 * 上場前の日・上場した日(前日比が無い)も抜けとして数えるが、取りに行っても埋まらないだけ(1回の実行で1銘柄1回しか取りに行かない)。
 */
function findStocksMissingCloses_(sheet) {
  const ym = parseYearMonthFromSheetName(sheet.getName());
  if (!ym) return [];
  const blocks = readStockBlocks_(sheet);
  if (blocks.length === 0) return [];
  const days = getDaysInMonth(ym.year, ym.month);
  const todayStart = new Date();
  todayStart.setHours(0, 0, 0, 0);
  const values = sheet.getRange(6, 1, days, sheet.getLastColumn()).getValues();
  const rows = [];
  for (let d = 1; d <= days; d++) {
    const date = new Date(ym.year, ym.month - 1, d);
    if (date >= todayStart) break;
    if (isJpMarketHoliday_(date)) continue;
    if (blocks.some(b => typeof values[d - 1][b.col - 1] === 'number')) rows.push(d - 1);
  }
  // 終値が空いている日か、終値はあるのに前日比が空いている日(以前の埋め方で、月の最初の営業日の前日比が抜けていた)がある銘柄
  return blocks.filter(b => rows.some(r => values[r][b.col - 1] === '' || (typeof values[r][b.col - 1] === 'number' && values[r][b.col] === '')))
    .map(b => b.code);
}

/**
 * 過去のシートを埋めるための日次チャート(3年分)。1回の実行の中では1銘柄1回だけ取りに行く。
 * 戻り値: 取れたらチャート / 取れず、あとで取り直すなら null / 何度も取れずにあきらめたら false
 */
function getHistoryForPastSync_(code, historyCache, state) {
  if (historyCache[code] !== undefined) return historyCache[code];
  state.fetchFailures = state.fetchFailures || {};
  if ((state.fetchFailures[code] || 0) >= PAST_SYNC_FETCH_MAX_FAILURES) return (historyCache[code] = false);
  const today = new Date(), start = new Date(today);
  start.setFullYear(start.getFullYear() - DB_BACKFILL_YEARS);
  const history = fetchHistoricalCloses(code, start, today);
  if (history) {
    delete state.fetchFailures[code];
    return (historyCache[code] = history);
  }
  state.fetchFailures[code] = (state.fetchFailures[code] || 0) + 1;
  if (state.fetchFailures[code] < PAST_SYNC_FETCH_MAX_FAILURES) return null;
  state.errors.push(code + ': 終値を取得できませんでした(埋められなかった日があります)');
  return (historyCache[code] = false);
}

// ==== 昔のレイアウトの信用残(週ごと)を埋める ====
// 2026/9 以前は JPX の信用残が週1回(金曜時点、第2営業日に公表)だったため、昔のレイアウトのシートは
// 38行目=見出し、39行目〜=1週1行(A列=その週の日付)。その月に公表された週を並べる(例: 9月26年は 8/28〜9/25 の5週)。
// JPX にはもう当時のファイルが無いので、Yahoo!ファイナンスの「信用残時系列」(1ページ20週)から取る。
// 単位は今のレイアウトと同じ千株(整数)、信用倍率は買残÷売残(小数第2位まで、売残が0なら「---」)。空いているセルだけ埋める。

const WEEKLY_MARGIN_URL_PREFIX = 'https://finance.yahoo.co.jp/quote/';

const WEEKLY_MARGIN_MAX_PAGES = 4; // 1ページ20週。2026/4 からなら2ページで足りる

const WEEKLY_MARGIN_CACHE_PREFIX = 'WEEKLY_MARGIN_'; // 取得結果を6時間キャッシュする(ほかのシートでも使い回すため)

const WEEKLY_MARGIN_MAX_FAIL_STREAK = 5; // 続けてこの回数取得に失敗したら(1回失敗するごとに少し待つ)、今回の処理では週の信用残をあきらめる

const WEEKLY_MARGIN_FETCH_INTERVAL_MS = 1000; // 1銘柄取るごとに待つ時間(短い間に続けて取ると、途中から取れなくなるため)

const WEEKLY_MARGIN_BACKOFF_MS = 5 * 60 * 1000; // 取れなかったときに、次の実行まで待つ時間

const OLD_LAYOUT_MARGIN_MAX_WEEKS = 6; // 1か月に入る週の数の上限(39〜44行目)

/** その週の最後の営業日(金曜。休場なら前の営業日) */
function weeklyMarginWeekEnd_(friday) {
  const d = new Date(friday);
  for (let i = 0; i < 5 && isJpMarketHoliday_(d); i++) d.setDate(d.getDate() - 1);
  return d;
}

/** その日より後の、n番目の営業日 */
function nthBusinessDayAfter_(date, n) {
  const d = new Date(date);
  let count = 0;
  while (count < n) {
    d.setDate(d.getDate() + 1);
    if (!isJpMarketHoliday_(d)) count++;
  }
  return d;
}

/** その月(year, month)に公表された週の日付(週の最後の営業日)。古い順 */
function oldLayoutMarginWeeksOf_(year, month) {
  const weeks = [];
  const d = new Date(year, month - 1, 1);
  d.setDate(d.getDate() - 14);
  while (d.getDay() !== 5) d.setDate(d.getDate() + 1); // 金曜まで進める
  for (; d <= new Date(year, month, 0); d.setDate(d.getDate() + 7)) {
    const weekEnd = weeklyMarginWeekEnd_(d);
    const published = nthBusinessDayAfter_(weekEnd, 2);
    if (published.getFullYear() === year && published.getMonth() === month - 1) weeks.push(weekEnd);
  }
  return weeks;
}

/**
 * 昔のレイアウトのシートの週の行を返す [{row, key: 'yyyy-MM-dd'}]。A列に日付が無ければ、公表日から決めて書き込む。
 * あわせて見出し(38行目)を銘柄の列にそろえる(2026/7 以前のシートは見出しがC列から始まっていて、銘柄と2つずれていた)。
 */
function ensureOldLayoutMarginWeeks_(sheet, blocks) {
  const ym = parseYearMonthFromSheetName(sheet.getName());
  if (!ym) return [];
  const range = sheet.getRange(MARGIN_FIRST_ROW, 1, OLD_LAYOUT_MARGIN_MAX_WEEKS, 1);
  let dates = range.getValues().map(r => r[0]).filter(v => v instanceof Date);
  if (dates.length === 0) {
    dates = oldLayoutMarginWeeksOf_(ym.year, ym.month);
    if (dates.length === 0) return [];
    sheet.getRange(MARGIN_FIRST_ROW, 1, dates.length, 1).setValues(dates.map(d => [d])).setNumberFormat('m"月"d"日"');
  }

  if (blocks.length > 0) {
    const lastCol = blocks[blocks.length - 1].col + STOCK_COL_WIDTH - 1;
    const head = sheet.getRange(MARGIN_HEADER_ROW, 3, 1, lastCol - 2).getValues()[0]; // C列から
    const want = head.map((v, i) => (i + 3 < OLD_LAYOUT_STOCK_START_COL) ? '' : ['売残', '買残', '信倍率'][(i + 3 - OLD_LAYOUT_STOCK_START_COL) % STOCK_COL_WIDTH]);
    if (head.some((v, i) => v !== want[i])) sheet.getRange(MARGIN_HEADER_ROW, 3, 1, want.length).setValues([want]);
  }
  return dates.map((d, i) => ({ row: MARGIN_FIRST_ROW + i, key: Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd') }));
}

/**
 * Yahoo!ファイナンスの信用残時系列を取る。{'yyyy-MM-dd': [売残, 買残](株)}。oldestKey の週まで(ページをめくって)取る。
 * 取れなければ null(ページが開けない・表が見つからない)。信用取引の無い銘柄・上場前は {} になる。
 */
function fetchWeeklyMarginHistory_(code, oldestKey) {
  const result = {};
  for (let page = 1; page <= WEEKLY_MARGIN_MAX_PAGES; page++) {
    const url = WEEKLY_MARGIN_URL_PREFIX + encodeURIComponent(code) + '.T/history?styl=margin&page=' + page;
    let html = null;
    for (let attempt = 0; attempt < 2 && html === null; attempt++) {
      if (attempt > 0) Utilities.sleep(1500);
      try {
        const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
        if (res.getResponseCode() === 200) html = res.getContentText('UTF-8');
      } catch (e) {}
    }
    if (html === null) return page === 1 ? null : result;
    if (page === 1 && html.indexOf('信用倍率') === -1) return null; // ページの作りが変わった

    let oldest = null, rows = 0;
    (html.match(/<tr[^>]*>[\s\S]*?<\/tr>/g) || []).forEach(tr => {
      const cells = (tr.match(/<t[hd][^>]*>[\s\S]*?<\/t[hd]>/g) || []).map(c => c.replace(/<[^>]+>/g, '').trim());
      const m = cells.length >= 3 && cells[0].match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
      if (!m) return;
      const key = m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
      const sell = Number(cells[1].replace(/,/g, '')), buy = Number(cells[2].replace(/,/g, ''));
      rows++;
      if (!oldest || key < oldest) oldest = key;
      if (cells[1] !== '' && cells[2] !== '' && isFinite(sell) && isFinite(buy)) result[key] = [sell, buy];
    });
    if (rows === 0 || (oldest && oldest <= oldestKey)) break;
    Utilities.sleep(300);
  }
  return result;
}

/**
 * 週の信用残をどこまでさかのぼって取るか(昔のレイアウトの残りのシートのうち、いちばん古い週)。
 * 1銘柄1回の取得で全部のシートに使えるように、最初のシートからいちばん古い週まで取る。
 */
function pastSyncWeeklyOldestKey_(state) {
  if (state.weeklyOldestKey) return state.weeklyOldestKey;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let oldest = '';
  state.queue.forEach(name => {
    const sh = ss.getSheetByName(name), ym = parseYearMonthFromSheetName(name);
    if (!sh || !ym || !isOldLayoutSheet_(sh)) return;
    const first = oldLayoutMarginWeeksOf_(ym.year, ym.month)[0];
    const key = first ? Utilities.formatDate(first, 'Asia/Tokyo', 'yyyy-MM-dd') : '';
    if (key && (!oldest || key < oldest)) oldest = key;
  });
  return (state.weeklyOldestKey = oldest);
}

/** 週の信用残(キャッシュがあれば使う)。取れなければ null(あとで取り直す)/ 今回はあきらめるなら false */
function getWeeklyMarginForPastSync_(code, oldestKey, state) {
  if (state.weeklyMarginOff) return false;
  const cache = CacheService.getScriptCache(), cacheKey = WEEKLY_MARGIN_CACHE_PREFIX + code;
  try {
    const cached = JSON.parse(cache.get(cacheKey) || 'null');
    if (cached && cached.oldest <= oldestKey) return cached.data;
  } catch (e) {}
  const data = fetchWeeklyMarginHistory_(code, oldestKey);
  Utilities.sleep(WEEKLY_MARGIN_FETCH_INTERVAL_MS);
  if (data) {
    state.weeklyFailStreak = 0;
    try { cache.put(cacheKey, JSON.stringify({ oldest: oldestKey, data: data }), 21600); } catch (e) {}
    return data;
  }
  state.weeklyFailStreak = (state.weeklyFailStreak || 0) + 1;
  if (state.weeklyFailStreak >= WEEKLY_MARGIN_MAX_FAIL_STREAK) {
    state.weeklyMarginOff = true;
    state.errors.push('Yahoo!ファイナンスから週の信用残を取得できませんでした(昔のレイアウトの信用残は埋めていません)');
    return false;
  }
  return null;
}

/**
 * 昔のレイアウトのシートの、週の信用残の空いているセルを埋める。最後まで終われば true、時間切れ・取り直しなら false。
 * 取れなかったら、そこで止めて少し待ってから(WEEKLY_MARGIN_BACKOFF_MS)続きを取る。同じ銘柄が3回取れなければ、その銘柄はあきらめる。
 */
function fillOldLayoutWeeklyMargin_(sheet, state, startMs) {
  const blocks = readStockBlocks_(sheet);
  const weeks = ensureOldLayoutMarginWeeks_(sheet, blocks);
  if (weeks.length === 0 || blocks.length === 0) return true;
  const lastCol = blocks[blocks.length - 1].col + STOCK_COL_WIDTH - 1;
  const values = sheet.getRange(MARGIN_FIRST_ROW, 1, weeks.length, lastCol).getValues();
  const blank = v => v === '' || v === null;
  const needs = blocks.filter(b => weeks.some((w, i) => blank(values[i][b.col - 1]) || blank(values[i][b.col])));
  const wider = pastSyncWeeklyOldestKey_(state);
  const oldestKey = (wider && wider < weeks[0].key) ? wider : weeks[0].key;
  state.weeklyFailures = state.weeklyFailures || {};

  const data = [];
  const flush = () => {
    if (data.length === 0) return;
    Sheets.Spreadsheets.Values.batchUpdate({ valueInputOption: 'RAW', data: data.splice(0) }, SpreadsheetApp.getActiveSpreadsheet().getId());
  };
  for (const b of needs) {
    if (Date.now() - startMs > RUN_TIME_LIMIT_MS) { flush(); return false; }
    if ((state.weeklyFailures[b.code] || 0) >= PAST_SYNC_FETCH_MAX_FAILURES) continue;
    const history = getWeeklyMarginForPastSync_(b.code, oldestKey, state);
    if (history === false) break; // 今回はあきらめる
    if (history === null) {
      state.weeklyFailures[b.code] = (state.weeklyFailures[b.code] || 0) + 1;
      if (state.weeklyFailures[b.code] >= PAST_SYNC_FETCH_MAX_FAILURES) {
        state.errors.push(b.code + ': 週の信用残を取得できませんでした');
        continue;
      }
      flush();
      state.nextDelayMs = WEEKLY_MARGIN_BACKOFF_MS; // 少し待ってから続きを取る
      return false;
    }
    delete state.weeklyFailures[b.code];
    let changed = false;
    const rows = weeks.map((w, i) => {
      const cur = values[i].slice(b.col - 1, b.col + 2);
      const pair = history[w.key];
      if (!pair || !(blank(cur[0]) || blank(cur[1]))) return cur;
      changed = true;
      return [marginUnitValue_(pair[0], false), marginUnitValue_(pair[1], false), pair[0] > 0 ? Math.round(pair[1] / pair[0] * 100) / 100 : '---'];
    });
    if (changed) {
      const a1 = "'" + sheet.getName() + "'!" + columnToLetter(b.col) + weeks[0].row + ':' + columnToLetter(b.col + 2) + weeks[weeks.length - 1].row;
      data.push({ range: a1, values: rows });
      if (data.length >= 20) flush();
    }
  }
  flush();
  if (needs.length > 0) {
    // 数値の表示形式(千株は整数、信用倍率は小数第2位まで)を銘柄の列にそろえる
    const width = lastCol - OLD_LAYOUT_STOCK_START_COL + 1;
    const formatRow = Array.from({ length: width }, (_, i) => (i % STOCK_COL_WIDTH === 2) ? '#,##0.00' : '#,##0');
    sheet.getRange(MARGIN_FIRST_ROW, OLD_LAYOUT_STOCK_START_COL, weeks.length, width).setNumberFormats(weeks.map(() => formatRow));
  }
  return true;
}

// ==== 昔のレイアウトの月次シートの修理(1回だけ使う) ====
// 2026/11/4 に、昔のレイアウトのシートを今のレイアウトとして扱ってしまい、次のように崩れたのを直す。
//  ・9・8・7月26年: L列に、I列と同じ銘柄(285A)の写しの列が増えた(I列からの銘柄をL列からと読み違え、足りないと判断したため)
//  ・9・8・7月26年: 前日比・信用倍率の条件付き書式が、今のレイアウトの列に付いた
//  ・4〜9月26年: 4行目のスパークラインが今のレイアウトの位置に入り、市場ブロック(E・G列)と1銘柄目(I列)がずれた
//  ・9・8・7月26年で足した銘柄: 信用残の欄(昔は週ごと)に、毎日の休場日の位置でグレーが塗られた
// 何度実行しても同じ結果になる。

const OLD_LAYOUT_MARKET_ITEMS = [ // 昔のレイアウトの市場ブロック(2列: 売買代金・前日比)
  { code: '1001', col: 3 }, { code: '1002', col: 5 }, { code: '1003', col: 7 }
];

const OLD_LAYOUT_STOCK_START_COL = 9; // 昔のレイアウトの銘柄はI列から(3列単位は同じ)

/** 昔のレイアウトの銘柄 [{code, col}](並び順) */
function readOldLayoutStockBlocks_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < OLD_LAYOUT_STOCK_START_COL) return [];
  const codeRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const blocks = [];
  for (let col = OLD_LAYOUT_STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) {
    if (codeRow[col - 1]) blocks.push({ code: String(codeRow[col - 1]), col: col });
  }
  return blocks;
}

/** 前日比・信用倍率の色のルールらしいもの(数値の条件で、範囲がすべて日次の行か信用残の行の1列の中) */
function isStockColorLikeRule_(rule) {
  const cond = rule.getBooleanCondition();
  if (!cond) return false;
  const T = SpreadsheetApp.BooleanCriteria;
  const types = [T.NUMBER_GREATER_THAN_OR_EQUAL_TO, T.NUMBER_LESS_THAN_OR_EQUAL_TO, T.NUMBER_BETWEEN, T.NUMBER_LESS_THAN];
  if (types.indexOf(cond.getCriteriaType()) === -1) return false;
  return rule.getRanges().every(rg => rg.getColumn() === rg.getLastColumn() && (
    (rg.getRow() >= CLOSE_FIRST_ROW && rg.getLastRow() <= CLOSE_FIRST_ROW + 30) ||
    (rg.getRow() >= MARGIN_FIRST_ROW && rg.getLastRow() <= MARGIN_FIRST_ROW + 30)));
}

/**
 * 昔のレイアウトのシートの、前日比・信用倍率の色のルールを7ルールに作り直す(ほかのルールは残す)。入れ替えたルールの数を返す。
 * 前日比: 市場ブロック(D・F・H列)と各銘柄+1列 / 信用倍率: 各銘柄+2列(昔の信用残は週ごとで、銘柄だけ)
 */
function rebuildOldLayoutColorRules_(sheet) {
  const blocks = readOldLayoutStockBlocks_(sheet);
  const cols = {
    changeCols: OLD_LAYOUT_MARKET_ITEMS.map(m => m.col + 1).concat(blocks.map(b => b.col + 1)),
    ratioCols: blocks.map(b => b.col + 2)
  };
  const rules = sheet.getConditionalFormatRules();
  const others = rules.filter(r => !isStockColorLikeRule_(r));
  sheet.setConditionalFormatRules(buildStockColorRules_(sheet, cols).concat(others));
  return rules.length - others.length;
}

/**
 * 昔のレイアウトのシートの4行目のスパークラインを入れ直す(市場: C・E・G列はその月の分、銘柄: I列から)。入れた数を返す。
 * A・B列(「チャート」など)は触らず、C列から右を作り直す(ずれて入った数式は消す)。
 */
function setOldLayoutSparklines_(sheet) {
  const ym = parseYearMonthFromSheetName(sheet.getName());
  const blocks = readOldLayoutStockBlocks_(sheet);
  const dbSheet = getOrCreateDatabaseSheet();
  const lastCol = sheet.getLastColumn();
  const formulas = new Array(lastCol - 2).fill('');
  if (ym) {
    OLD_LAYOUT_MARKET_ITEMS.forEach(m => {
      formulas[m.col - 3] = buildMarketValueSparklineFormula(columnToLetter(getOrCreateDatabaseColumn(dbSheet, m.code)), ym.year, ym.month);
    });
  }
  blocks.forEach(b => { formulas[b.col - 3] = buildSparklineFormula(columnToLetter(getOrCreateDatabaseColumn(dbSheet, b.code))); });
  sheet.getRange(4, 3, 1, lastCol - 2).setFormulas([formulas]);
  return blocks.length + (ym ? OLD_LAYOUT_MARKET_ITEMS.length : 0);
}

/** 範囲の値がすべて空なら true */
function isRangeBlank_(range) {
  return range.getValues().every(row => row.every(v => v === '' || v === null));
}

/** 1枚の昔のレイアウトのシートを直す。結果のメッセージを返す */
function repairOldLayoutSheet_(sheet) {
  const notes = [];

  // 1. 1銘柄目(I列)と同じ銘柄がL列にあれば、L列の方を消す。
  //    I列が元からある列(休場日に「休」が入っている)、L列は 11/4 の処理が足した写し(同じ終値をYahooから入れ直したもの)。
  //    それ以外の重なりは、どちらが元か分からないので消さずに知らせる。
  const first = readOldLayoutStockBlocks_(sheet);
  if (first.length >= 2 && first[0].col === OLD_LAYOUT_STOCK_START_COL && first[1].col === OLD_LAYOUT_STOCK_START_COL + STOCK_COL_WIDTH &&
      first[0].code === first[1].code) {
    sheet.deleteColumns(first[1].col, STOCK_COL_WIDTH);
    notes.push(first[1].code + 'の写しの列(L列)を削除');
  }
  const counts = {};
  readOldLayoutStockBlocks_(sheet).forEach(b => { counts[b.code] = (counts[b.code] || 0) + 1; });
  Object.keys(counts).filter(c => counts[c] > 1).forEach(c => notes.push('⚠️' + c + 'が' + counts[c] + 'つある(消さずに残した)'));

  const blocks = readOldLayoutStockBlocks_(sheet);

  // 2. 前日比・信用倍率の条件付き書式を、昔のレイアウトの列で作り直す(ほかのルールは残す)
  notes.push('条件付き書式 ' + rebuildOldLayoutColorRules_(sheet) + '件 → ' + STOCK_COLOR_RULE_COUNT + '件');

  // 3. 4行目のスパークラインを、昔のレイアウトの位置に入れ直す
  notes.push('スパークライン ' + setOldLayoutSparklines_(sheet) + '件');

  // 4. 信用残の欄が空の銘柄は、毎日の休場日の位置で塗られたグレーを消す(昔の信用残は週ごとなので、日ごとの塗りは要らない)
  let cleared = 0;
  blocks.forEach(b => {
    const range = sheet.getRange(MARGIN_FIRST_ROW, b.col, 31, STOCK_COL_WIDTH);
    if (!isRangeBlank_(range)) return;
    const gray = range.getBackgrounds().some(row => row.some(c => String(c).toLowerCase() === HOLIDAY_COLOR.toLowerCase()));
    if (gray) { range.setBackground(null); cleared++; }
  });
  if (cleared > 0) notes.push('信用残の欄のグレーを消した銘柄 ' + cleared);

  return sheet.getName() + ': ' + notes.join('、');
}

/** 「データ修正 → 昔のレイアウトのシートを修理」: 昔のレイアウトの月次シートをすべて直す */
function repairOldLayoutSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // 止めたままになっている「過去のシートを合わせる」処理の状態を片付ける(昔のレイアウトのシートは、もう対象にしない)
  deletePastSyncTriggers_();
  const state = readPastSyncState_();
  if (state && state.running) { state.running = false; state.queue = []; writePastSyncState_(state); }

  const targets = ss.getSheets().filter(s => parseYearMonthFromSheetName(s.getName()) && !isCurrentLayoutSheet_(s));
  if (targets.length === 0) return '昔のレイアウトの月次シートはありません';
  return targets.map(s => repairOldLayoutSheet_(s)).join('\n');
}
