// ==== 過去の月次シートを当月に合わせる(バックグラウンド処理) ====
// 当月シートの銘柄(と並び順)に、ほかの月次シート(過去の月・作成済みの次月)を合わせる。
//  ・当月に無い銘柄の列は削除する(その月のデータも消える)
//  ・当月にあって無い銘柄は列を足し、終値・前日比・売買代金(Yahoo)と信用残(ドライブの日次データ)を埋める
//  ・当月と同じ順番に並べる
// 銘柄の追加・削除・並び替えのあと(当月・テンプレートの処理が終わったら)と、「データ修正」メニューから始める。
// 処理は時間主導トリガー(1回きり)で進め、時間切れ(RUN_TIME_LIMIT_MS)になったら続きを予約し直す。
// 進み具合はスクリプトプロパティ(PAST_SYNC_STATE_KEY)に置き、小さなダイアログ(PastSyncDialog)が読んで表示する。
// 処理中に別の追加・削除・並び替えがあれば、最初からやり直す(generation が変わったら、古い実行は途中でやめる)。

const PAST_SYNC_STATE_KEY = 'PAST_SHEET_SYNC_STATE';

const PAST_SYNC_HANDLER = 'runPastSheetSync'; // 時間主導トリガーから呼ぶ関数の名前(末尾 _ の関数はトリガーから呼べない)

const PAST_SYNC_FIRST_DELAY_MS = 1000;     // 始めるときの予約(実際には1分ほど後に動く)

const PAST_SYNC_RETRY_DELAY_MS = 60 * 1000; // 時間切れ・ほかの実行中のときの、次の予約までの時間

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

/** 当月シート以外の月次シート(新しい月から順) */
function listOtherMonthSheets_(ss, currentSheet) {
  return ss.getSheets()
    .map(s => ({ sheet: s, ym: parseYearMonthFromSheetName(s.getName()) }))
    .filter(x => x.ym && (!currentSheet || x.sheet.getName() !== currentSheet.getName()))
    .sort((a, b) => (b.ym.year * 12 + b.ym.month) - (a.ym.year * 12 + a.ym.month))
    .map(x => x.sheet);
}

/** シートの銘柄 [{code, name, industry, col}](並び順) */
function readStockBlocks_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < STOCK_START_COL) return [];
  const head = sheet.getRange(2, 1, 2, lastCol).getValues(); // 2行目=コード・業種、3行目=銘柄名
  const blocks = [];
  for (let col = STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) {
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
  deletePastSyncTriggers_();
  const lock = LockService.getScriptLock(); // この処理どうしの重なりだけを防ぐ(銘柄の列のロックは使わない。当月の操作を待たせないため)
  if (!lock.tryLock(10 * 1000)) { schedulePastSheetSync_(PAST_SYNC_RETRY_DELAY_MS); return; }
  try {
    let state = readPastSyncState_();
    if (!state || !state.running) return;

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
      if (!savePastSyncStateIfCurrent_(state)) return;
      schedulePastSheetSync_(PAST_SYNC_RETRY_DELAY_MS);
    } else {
      state.running = false;
      state.current = '';
      state.finishedAt = Date.now();
      savePastSyncStateIfCurrent_(state);
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
  const sheetName = sheet.getName();
  const wanted = order.map(o => o.code);
  const timeUp = () => Date.now() - startMs > RUN_TIME_LIMIT_MS;

  // 1. 当月に無い銘柄の列を削除する
  readStockBlocks_(sheet).filter(b => wanted.indexOf(b.code) === -1).forEach(b => deleteStockColumnFromSheet(sheet, b.code));

  // 2. 足りない銘柄の列を右端に足す(中身は 3. で埋める)
  const have = readStockBlocks_(sheet).map(b => b.code);
  order.filter(o => have.indexOf(o.code) === -1).forEach(o => {
    const col = addStockColumnToSheet(sheet, o.code, o.name, o.industry);
    if (!col) return;
    fillHolidayMarksForNewStock(sheet, col);
    state.pendingFill.push({ sheet: sheetName, code: o.code });
    savePastSyncStateIfCurrent_(state);
  });
  if (!savePastSyncStateIfCurrent_(state)) return false; // やり直しになった

  // 3. 足した銘柄の終値・前日比・売買代金と信用残を埋める
  const pending = state.pendingFill.filter(p => p.sheet === sheetName);
  for (const p of pending) {
    if (timeUp()) return false;
    const block = readStockBlocks_(sheet).find(b => b.code === p.code);
    if (block) {
      if (!historyCache[p.code]) {
        const today = new Date(), start = new Date(today);
        start.setFullYear(start.getFullYear() - DB_BACKFILL_YEARS);
        historyCache[p.code] = fetchHistoricalCloses(p.code, start, today) || { dateMap: {}, splits: [] };
      }
      fillMissingClosesForCode(p.code, historyCache[p.code], block.col, sheet);
      const marginMessage = fillMissingMarginForCode(p.code, block.col, sheet, startMs);
      if (String(marginMessage).indexOf('途中まで') !== -1) return false;
    }
    state.pendingFill = state.pendingFill.filter(x => !(x.sheet === p.sheet && x.code === p.code));
    if (!savePastSyncStateIfCurrent_(state)) return false; // やり直しになった
  }

  // 4. 当月と同じ順番に並べる(並び替えと同じ方法: 列の位置はそのままで中身を入れ替える)
  const result = getReorderPayloadForSheet(sheet, wanted);
  if (result && result.payloads) {
    Sheets.Spreadsheets.Values.batchUpdate({ valueInputOption: 'USER_ENTERED', data: result.payloads }, SpreadsheetApp.getActiveSpreadsheet().getId());
    reapplySparklineFormulas(sheet);
  }

  // 5. 前日比・信用倍率の条件付き書式をまとめる(まとまっていれば何もしない)
  rebuildStockColorRules_(sheet, true);
  return true;
}
