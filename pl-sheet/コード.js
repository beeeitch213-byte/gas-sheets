/**
 * 損益管理
 *
 * ・createNextMonthSheet … 毎日自動実行(時間主導トリガー)
 *     当月シートが無ければ作成し、CREATE_NEXT_FROM_DAY 日以降は翌月シートも作成する。
 *     作成途中で止まったシートは、次回の実行で作り直す。
 * ・週間累計 … 各週(月〜日)の最終営業日の下に「週間累計」行を入れる。
 *     月をまたぐ週は、最終営業日がある月のシートに入れ、前月分の日も合計する。
 *     メニュー「機能 → 修正 → 週間累計適用」で、開いている月次シートにも適用できる。
 * ・setupDailyTrigger    … トリガーを「毎日」に設定し直す(手動で1回だけ実行)
 * ・repairCarryOver      … 既存の月次シートの繰越損益を数式に直す(手動で1回だけ実行・任意)
 * ・repairIColumn        … 既存の月次シートのI列を新しい数式に直す(手動で1回だけ実行・任意)
 * ・onOpen               … メニュー登録、当月シートの本日の「実現損益」セルを選択
 */

const TEMPLATE_SHEET_NAME = 'テンプレ';
const CREATE_NEXT_FROM_DAY = 28;   // この日以降に翌月シートを作成する
const TRIGGER_HOUR = 1;            // 毎日の実行時刻(時)
const HOLIDAY_CAL_ID = 'ja.japanese#holiday@group.v.calendar.google.com';
const BUILDING_KEY_PREFIX = 'building_';   // 作成途中の印(ドキュメントプロパティ)
const FIRST_ROW = 4;                       // 1日の行
const BASE_LAST_ROW = 34;                  // 週間累計行が無いときの31日の行
const SCAN_ROWS = 45;                      // A列を走査する行数(4行目から)
const WEEK_LABEL = '週間累計';
const WEEK_BG = '#fff2cc';                 // 週間累計行の背景色
const CLOSED_LABEL = '休場';
const CLOSED_BG = '#f3f3f3';
const DAY_OF_WEEK_KANJI = ['日', '月', '火', '水', '木', '金', '土'];

/** シート名(例: 10月26年) */
function sheetNameOf(year, month) {
  return month + '月' + String(year).slice(-2) + '年';
}

/** シート名から年・月を取り出す(月次シートでなければ null) */
function parseSheetName_(name) {
  const m = String(name).match(/^(\d{1,2})月(\d{2})年$/);
  return m ? { year: 2000 + Number(m[2]), month: Number(m[1]) } : null;
}

// ============================================================
// 月次シートの自動作成
// ============================================================

/**
 * 毎日自動実行。当月・翌月のシートが無い(または作成途中)なら作成する
 */
function createNextMonthSheet() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    console.warn('他の処理が実行中のためスキップしました。');
    return;
  }
  try {
    const now = new Date();
    const year = now.getFullYear();
    const month = now.getMonth() + 1;
    const isBiz = makeBusinessChecker_();

    // 当月(念のため。通常は前月中に作成済み)
    ensureMonthSheet_(year, month, isBiz);

    // 翌月
    if (now.getDate() >= CREATE_NEXT_FROM_DAY) {
      const next = new Date(year, month, 1);
      ensureMonthSheet_(next.getFullYear(), next.getMonth() + 1, isBiz);
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * 指定月のシートを用意する。
 * 完成済みなら何もしない。作成途中の印が残っていれば削除して作り直す。
 */
function ensureMonthSheet_(year, month, isBiz) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();
  const name = sheetNameOf(year, month);
  const key = BUILDING_KEY_PREFIX + name;
  const existing = ss.getSheetByName(name);

  if (existing) {
    if (!props.getProperty(key)) return; // 完成済み(既存シートもこちら)
    console.warn(name + ' は作成途中で止まっていたため作り直します。');
    ss.deleteSheet(existing);
  }

  const template = ss.getSheetByName(TEMPLATE_SHEET_NAME);
  if (!template) {
    throw new Error('「' + TEMPLATE_SHEET_NAME + '」シートが見つかりません。シート名を確認してください。');
  }

  props.setProperty(key, '1'); // 作成開始の印

  const sheet = template.copyTo(ss).setName(name);
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(template.getIndex() + 1); // テンプレのすぐ右隣

  buildMonthSheet_(sheet, year, month, isBiz);
  applyWeeklyRows_(sheet, year, month, isBiz);

  SpreadsheetApp.flush();
  props.deleteProperty(key); // 完成
  console.log(name + ' を作成しました。');
}

/**
 * 月次シートの日付・曜日・休場・H/I列の数式・繰越を設定する(4〜34行目)
 * D列(累計)・合計行・M4・週間累計行は applyWeeklyRows_ で設定する
 */
function buildMonthSheet_(sheet, year, month, isBiz) {
  const ss = sheet.getParent();

  // 繰越損益:前月シートのM5を数式で参照(前月の入力が後から変わっても自動で反映)
  const prev = new Date(year, month - 2, 1);
  const prevName = sheetNameOf(prev.getFullYear(), prev.getMonth() + 1);
  if (ss.getSheetByName(prevName)) {
    sheet.getRange('M3').setFormula(`='${prevName}'!M5`);
  } else {
    sheet.getRange('M3').setValue(0);
    console.warn('前月シート (' + prevName + ') が見つかりません。繰越損益は0からスタートします。');
  }
  sheet.getRange('M5').setFormula('=SUM(M3:M4)');

  const lastDay = new Date(year, month, 0).getDate();

  const dateValues = [];
  const dayValues = [];
  const hFormulas = [];
  const iFormulas = [];
  const closedRanges = []; // 休場日の C:J
  const closedCells = [];  // 休場日の C(「休場」を入れるセル)
  const openRanges = [];   // 営業日・月末以降の C:J

  for (let d = 1; d <= 31; d++) {
    const row = FIRST_ROW - 1 + d;

    if (d > lastDay) {
      dateValues.push(['']);
      dayValues.push(['']);
      hFormulas.push(['']);
      iFormulas.push(['']);
      openRanges.push(`C${row}:J${row}`);
      continue;
    }

    const date = new Date(year, month - 1, d);
    dateValues.push([d]);
    dayValues.push([DAY_OF_WEEK_KANJI[date.getDay()]]);

    if (!isBiz(date)) {
      closedRanges.push(`C${row}:J${row}`);
      closedCells.push(`C${row}`);
      hFormulas.push(['']);
      iFormulas.push(['']);
    } else {
      openRanges.push(`C${row}:J${row}`);
      hFormulas.push([`=E${row}-(F${row}+G${row})`]);
      iFormulas.push([iFormula_(row)]);
    }
  }

  // C〜J列をまとめて初期化
  const block = sheet.getRange(`C${FIRST_ROW}:J${BASE_LAST_ROW}`);
  block.breakApart();
  block.clearContent();

  if (openRanges.length) {
    sheet.getRangeList(openRanges).setBackground(null);
  }
  if (closedRanges.length) {
    closedRanges.forEach(a1 => sheet.getRange(a1).merge());
    sheet.getRangeList(closedRanges)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBackground(CLOSED_BG);
    sheet.getRangeList(closedCells).setValue(CLOSED_LABEL);
  }

  // 日付・曜日・数式を一括セット
  sheet.getRange(`A${FIRST_ROW}:A${BASE_LAST_ROW}`).setValues(dateValues);
  sheet.getRange(`B${FIRST_ROW}:B${BASE_LAST_ROW}`).setValues(dayValues);
  sheet.getRange(`H${FIRST_ROW}:H${BASE_LAST_ROW}`).setFormulas(hFormulas);
  sheet.getRange(`I${FIRST_ROW}:I${BASE_LAST_ROW}`).setFormulas(iFormulas);
}

/** I列の数式: F / (E - G)。E が空欄、または E - G が 0 なら 0 */
function iFormula_(row) {
  return `=IF(OR(E${row}="", E${row}-G${row}=0), 0, F${row}/(E${row}-G${row}))`;
}

// ============================================================
// 週間累計
// ============================================================

/**
 * 月次シートに週間累計行を入れる(既にあれば作り直す)。
 * A〜J列だけを下にずらすので、K列以降(M3〜M5など)は動かない。
 * あわせて D列(累計)・合計行・M4 を設定し直す。
 */
function applyWeeklyRows_(sheet, year, month, isBiz) {
  const ss = sheet.getParent();

  // 1. 既存の週間累計行を削除(下から)
  removeWeeklyRows_(sheet);

  // 2. 週間累計を入れる日(週の最終営業日)を求める
  const weeks = getWeekEndDays_(year, month, isBiz);

  // 3. その日の行の直下にセルを挿入(下から順に入れて行番号がずれないようにする)
  for (let i = weeks.length - 1; i >= 0; i--) {
    const r = FIRST_ROW + weeks[i].day; // 元の行(FIRST_ROW-1+day)の1つ下
    sheet.getRange(`A${r}:J${r}`).insertCells(SpreadsheetApp.Dimension.ROWS);
  }

  const lastRow = BASE_LAST_ROW + weeks.length; // 31日(または週間累計)の最終行
  const totalRow = lastRow + 1;                 // 月合計の行
  const lastDay = new Date(year, month, 0).getDate();
  const weekByDay = {};
  weeks.forEach(w => { weekByDay[w.day] = w; });

  // 4. D列(累計)と週間累計行の位置
  const cValues = sheet.getRange(`C${FIRST_ROW}:C${lastRow}`).getValues();
  const dFormulas = [];
  const weekRows = [];
  let r = FIRST_ROW;
  for (let d = 1; d <= 31; d++) {
    const closedOnSheet = cValues[r - FIRST_ROW][0] === CLOSED_LABEL;
    if (d <= lastDay && !closedOnSheet && isBiz(new Date(year, month - 1, d))) {
      // C列が空欄なら空欄、入力があれば(0含む)当月の累計。週間累計行(A列が文字)は除外
      dFormulas.push([`=IF(C${r}="", "", SUMIF($A$${FIRST_ROW}:A${r}, ">0", $C$${FIRST_ROW}:C${r}))`]);
    } else {
      dFormulas.push(['']);
    }
    r++;

    if (weekByDay[d]) {
      dFormulas.push(['']);
      weekRows.push({ row: r, week: weekByDay[d] });
      r++;
    }
  }
  sheet.getRange(`D${FIRST_ROW}:D${lastRow}`).setFormulas(dFormulas);

  // 5. 週間累計行(A:B「週間累計」、C:D その週の合計)
  weekRows.forEach(({ row, week }) => {
    const rowRange = sheet.getRange(`A${row}:J${row}`);
    rowRange.clearContent();

    sheet.getRange(`A${row}:B${row}`).merge();
    sheet.getRange(`A${row}`).setValue(WEEK_LABEL);

    sheet.getRange(`C${row}:D${row}`).merge();
    const sumCell = sheet.getRange(`C${row}`);
    sumCell.setFormula(weekFormula_(ss, year, month, week, lastRow));
    sumCell.setNumberFormat(sheet.getRange(`C${row - 1}`).getNumberFormat());

    rowRange
      .setBackground(WEEK_BG)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  });

  // 6. 月合計(週間累計行を除く)と当月損益
  sheet.getRange(`C${totalRow}`).setFormula(
    `=SUMIF($A$${FIRST_ROW}:A${lastRow}, ">0", $C$${FIRST_ROW}:C${lastRow})`
  );
  sheet.getRange('M4').setFormula(`=C${totalRow}`);
}

/** 既存の週間累計行を削除する(A〜J列を上に詰める) */
function removeWeeklyRows_(sheet) {
  const aValues = sheet.getRange(FIRST_ROW, 1, SCAN_ROWS, 1).getValues();
  for (let i = aValues.length - 1; i >= 0; i--) {
    if (aValues[i][0] === WEEK_LABEL) {
      const r = FIRST_ROW + i;
      const rg = sheet.getRange(`A${r}:J${r}`);
      rg.breakApart();
      rg.deleteCells(SpreadsheetApp.Dimension.ROWS);
    }
  }
}

/** シート内の週間累計行の数 */
function countWeekRows_(sheet) {
  return sheet.getRange(FIRST_ROW, 1, SCAN_ROWS, 1).getValues()
    .filter(v => v[0] === WEEK_LABEL).length;
}

/**
 * 当月の中で「その週(月〜日)の最終営業日」にあたる日を返す。
 * 週の最終営業日が翌月にある週は含めない(翌月シート側に入る)。
 * 戻り値: [{ day: 日, monday: その週の月曜日(Date) }, ...]
 */
function getWeekEndDays_(year, month, isBiz) {
  const lastDay = new Date(year, month, 0).getDate();
  const result = [];

  for (let d = 1; d <= lastDay; d++) {
    const date = new Date(year, month - 1, d);
    if (!isBiz(date)) continue;

    const monday = new Date(year, month - 1, d - ((date.getDay() + 6) % 7));
    let lastBiz = null;
    for (let i = 0; i < 7; i++) {
      const x = new Date(monday.getFullYear(), monday.getMonth(), monday.getDate() + i);
      if (isBiz(x)) lastBiz = x;
    }

    if (lastBiz &&
        lastBiz.getFullYear() === year &&
        lastBiz.getMonth() === month - 1 &&
        lastBiz.getDate() === d) {
      result.push({ day: d, monday: monday });
    }
  }
  return result;
}

/**
 * 週間累計の数式。週の月曜が前月なら、前月シートの該当日も合計する
 */
function weekFormula_(ss, year, month, week, lastRow) {
  const parts = [];
  const mon = week.monday;
  let startDay = mon.getDate();

  if (mon.getFullYear() !== year || mon.getMonth() !== month - 1) {
    const pName = sheetNameOf(mon.getFullYear(), mon.getMonth() + 1);
    const pSheet = ss.getSheetByName(pName);
    if (pSheet) {
      const pLast = BASE_LAST_ROW + countWeekRows_(pSheet);
      const pEnd = new Date(year, month - 1, 0).getDate();
      const pc = `'${pName}'!$C$${FIRST_ROW}:$C$${pLast}`;
      const pa = `'${pName}'!$A$${FIRST_ROW}:$A$${pLast}`;
      parts.push(`SUMIFS(${pc}, ${pa}, ">=${mon.getDate()}", ${pa}, "<=${pEnd}")`);
    } else {
      console.warn('前月シート (' + pName + ') が見つからないため、週間累計は当月分のみです。');
    }
    startDay = 1;
  }

  const c = `$C$${FIRST_ROW}:$C$${lastRow}`;
  const a = `$A$${FIRST_ROW}:$A$${lastRow}`;
  parts.push(`SUMIFS(${c}, ${a}, ">=${startDay}", ${a}, "<=${week.day}")`);

  return '=' + parts.join(' + ');
}

/**
 * メニュー「機能 → 修正 → 週間累計適用」
 * 開いている月次シートに週間累計行を適用する(既にあれば作り直す)
 */
function menuApplyWeeklyRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getActiveSheet();
  const ym = parseSheetName_(sheet.getName());

  if (!ym) {
    ui.alert('月次シート(例: 10月26年)を開いてから実行してください。');
    return;
  }

  const res = ui.alert(
    '週間累計適用',
    '「' + sheet.getName() + '」に週間累計の行を適用します。\n' +
    '既にある週間累計の行は作り直します。よろしいですか?',
    ui.ButtonSet.OK_CANCEL
  );
  if (res !== ui.Button.OK) return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    ui.alert('他の処理が実行中です。少し待ってから実行してください。');
    return;
  }
  try {
    applyWeeklyRows_(sheet, ym.year, ym.month, makeBusinessChecker_());
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  ui.alert('「' + sheet.getName() + '」に週間累計を適用しました。');
}

// ============================================================
// 営業日判定
// ============================================================

/**
 * 営業日かどうかを判定する関数を返す(祝日は月ごとに1回だけ取得してキャッシュ)
 * 休場:土日・祝日・1/2・1/3・12/31
 */
function makeBusinessChecker_() {
  const cache = {};
  return function (date) {
    const wd = date.getDay();
    if (wd === 0 || wd === 6) return false;

    const y = date.getFullYear();
    const m = date.getMonth() + 1;
    const d = date.getDate();
    if ((m === 1 && (d === 2 || d === 3)) || (m === 12 && d === 31)) return false;

    const key = y + '-' + m;
    if (!cache[key]) cache[key] = getHolidayDays_(y, m);
    return !cache[key].has(d);
  };
}

/**
 * 指定月の祝日(日にち)の集合を返す。
 * 祝日カレンダーには「ひな祭り」「七夕」などの記念日も含まれるため、
 * 説明が「祝日」のものだけを対象にする。
 */
function getHolidayDays_(year, month) {
  const days = new Set();
  const cal = CalendarApp.getCalendarById(HOLIDAY_CAL_ID);
  if (!cal) {
    console.warn('祝日カレンダーを取得できません。土日・年末年始のみ休場にします。');
    return days;
  }

  const events = cal.getEvents(new Date(year, month - 1, 1), new Date(year, month, 1));
  const used = [];
  const skipped = [];

  events.forEach(ev => {
    const title = ev.getTitle();
    const start = ev.isAllDayEvent() ? ev.getAllDayStartDate() : ev.getStartTime();
    if (start.getFullYear() !== year || start.getMonth() !== month - 1) return;

    if (String(ev.getDescription()).indexOf('祝日') !== -1) {
      days.add(start.getDate());
      used.push(start.getDate() + '日 ' + title);
    } else {
      skipped.push(start.getDate() + '日 ' + title);
    }
  });

  console.log(month + '月 休場にした祝日: ' + (used.join(', ') || 'なし'));
  console.log(month + '月 除外した記念日: ' + (skipped.join(', ') || 'なし'));
  return days;
}

// ============================================================
// 保守用(手動で1回だけ実行)
// ============================================================

/**
 * トリガーを「毎日」に設定し直す
 * 既存の createNextMonthSheet のトリガーは削除される
 */
function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'createNextMonthSheet') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('createNextMonthSheet')
    .timeBased()
    .everyDays(1)
    .atHour(TRIGGER_HOUR)
    .create();
  console.log('createNextMonthSheet を毎日 ' + TRIGGER_HOUR + '時台に実行するよう設定しました。');
}

/**
 * 既存の月次シートの繰越損益を数式に直す(任意)
 * 各シートの M4 を合計行への参照、M5 を =SUM(M3:M4)、
 * M3 を直前の月次シートの M5 への参照にする(一番古いシートの M3 はそのまま)
 */
function repairCarryOver() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const monthSheets = ss.getSheets()
    .map(sh => {
      const ym = parseSheetName_(sh.getName());
      return ym ? { sheet: sh, key: ym.year * 100 + ym.month } : null;
    })
    .filter(Boolean)
    .sort((a, b) => a.key - b.key);

  monthSheets.forEach((item, i) => {
    const sh = item.sheet;
    const totalRow = BASE_LAST_ROW + countWeekRows_(sh) + 1;
    sh.getRange('M4').setFormula(`=C${totalRow}`);
    sh.getRange('M5').setFormula('=SUM(M3:M4)');
    if (i > 0) {
      const prevName = monthSheets[i - 1].sheet.getName();
      sh.getRange('M3').setFormula(`='${prevName}'!M5`);
    }
  });

  console.log('繰越損益を修正しました: ' + monthSheets.map(x => x.sheet.getName()).join(' → '));
}

/**
 * 既存の月次シートのI列を新しい数式(iFormula_)に直す(任意)
 * 旧数式 =IF(OR(E="", E=0), 0, F/E) が入っているセルだけを書き換える
 */
function repairIColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const oldPattern = /^=IF\(OR\(E(\d+)="",\s*E\1=0\),\s*0,\s*F\1\/E\1\)$/;
  const fixed = [];

  ss.getSheets().forEach(sh => {
    if (!parseSheetName_(sh.getName())) return;
    const formulas = sh.getRange(FIRST_ROW, 9, SCAN_ROWS, 1).getFormulas();
    let count = 0;
    formulas.forEach((f, i) => {
      if (!oldPattern.test(f[0])) return;
      const row = FIRST_ROW + i;
      sh.getRange(`I${row}`).setFormula(iFormula_(row));
      count++;
    });
    if (count) fixed.push(sh.getName() + '(' + count + ')');
  });

  console.log('I列の数式を修正しました: ' + (fixed.join(', ') || 'なし'));
}

// ============================================================
// 開いたとき
// ============================================================

/**
 * スプレッドシートを開いたときに自動実行。
 * メニューを登録し、今月のシートの本日の「実現損益」セルを選択する
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('機能')
    .addSubMenu(
      ui.createMenu('修正')
        .addItem('週間累計適用', 'menuApplyWeeklyRows')
    )
    .addToUi();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  const currentSheet = ss.getSheetByName(sheetNameOf(now.getFullYear(), now.getMonth() + 1));
  if (!currentSheet) return;

  ss.setActiveSheet(currentSheet);

  // 週間累計行で行がずれるので、A列から本日の日付を探す
  const today = now.getDate();
  const aValues = currentSheet.getRange(FIRST_ROW, 1, SCAN_ROWS, 1).getValues();
  const idx = aValues.findIndex(v => v[0] === today);
  if (idx >= 0) {
    currentSheet.getRange('C' + (FIRST_ROW + idx)).activate();
  }
}