// ==== ポップアップチャート表示(25日移動平均) ====
function showChartForActiveCell() {
  const sheet = SpreadsheetApp.getActiveSheet();

  if (sheet.getName() === TEMPLATE_SHEET_NAME || sheet.getName() === DB_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('このシートでは実行できません。当月シートで実行してください。');
    return;
  }

  const cell = sheet.getActiveCell();
  const col = cell.getColumn();

  let baseCol = null;
  const marketItem = MARKET_VALUE_ITEMS.find(item => col >= item.col && col < item.col + MARKET_VALUE_COL_WIDTH);
  if (marketItem) {
    baseCol = marketItem.col;
  } else if (col >= STOCK_START_COL) {
    baseCol = Math.floor((col - STOCK_START_COL) / STOCK_COL_WIDTH) * STOCK_COL_WIDTH + STOCK_START_COL;
  }

  if (baseCol === null) {
    SpreadsheetApp.getUi().alert('銘柄のデータ列を選択した状態でボタンを押してください。');
    return;
  }

  const codeAndName = sheet.getRange(2, baseCol, 2, 1).getValues(); // 2行目=銘柄コード、3行目=銘柄名
  const code = codeAndName[0][0];

  if (!code) {
    SpreadsheetApp.getUi().alert('選択した列に銘柄コードが見つかりません。');
    return;
  }

  showMovingAverageChart(code, marketItem ? marketItem.name : String(codeAndName[1][0] || ''));
}

/** チャートのタイトル(例: 「トヨタ自動車 7203 のチャート」) */
function chartTitleText_(code, stockName) {
  return (stockName ? stockName + ' ' : '') + code + ' のチャート';
}

/**
 * DateをJST(Asia/Tokyo)の {key: 'yyyy-MM-dd', label: 'MM/dd'} にする。
 * チャートの組み立てでは日付ごとに何千回も呼ぶため、遅いUtilities.formatDateの代わりに使う
 * (JSTは夏時間がないので、+9時間してUTCとして読めばよい)。
 */
function chartDateParts_(d) {
  const t = new Date(d.getTime() + 9 * 3600000);
  const mm = String(t.getUTCMonth() + 1).padStart(2, '0');
  const dd = String(t.getUTCDate()).padStart(2, '0');
  return { key: t.getUTCFullYear() + '-' + mm + '-' + dd, label: mm + '/' + dd };
}

/**
 * 時系列の点({date:Date, value:number, extra?:number|null})配列から、
 * 移動平均(ma)・前日比(change)・前日比率(rate)を付与したチャート表示用の配列を作る共通ヘルパー。
 * extraが入っている点は、そのまま row.extra として引き継ぐ(第3系列用)。
 * rateDecimals: 前日比率の小数点以下桁数(省略時2)。
 */
function buildChartRowsWithMA_(points, period, rateDecimals) {
  if (rateDecimals === undefined || rateDecimals === null) rateDecimals = 2;
  let windowSum = 0; // 直近period個の合計(1つずつ足し引きして、毎回period個を足し直さない)
  return points.map((p, idx) => {
    windowSum += p.value;
    if (idx >= period) windowSum -= points[idx - period].value;
    const ma = idx >= period - 1 ? windowSum / period : null;

    let changeStr = '---';
    let rateStr = '---';
    if (idx > 0) {
      const prevValue = points[idx - 1].value;
      const diff = p.value - prevValue;
      const sign = diff > 0 ? '+' : '';
      changeStr = sign + diff.toLocaleString();
      if (prevValue !== 0) {
        const rate = (diff / prevValue) * 100;
        rateStr = sign + rate.toFixed(rateDecimals) + '%';
      }
    }

    const dateParts = chartDateParts_(p.date);
    const row = {
      dateKey: dateParts.key,
      date: dateParts.label,
      close: p.value,
      ma: ma,
      change: changeStr,
      rate: rateStr
    };
    if (p.extra !== undefined && p.extra !== null) row.extra = p.extra;
    // ratioInfo: グラフの線にはしないが、ツールチップに補足表示したい値(空売り比率など)
    if (p.ratioInfo !== undefined && p.ratioInfo !== null) row.ratioInfo = p.ratioInfo;
    return row;
  });
}

/**
 * 指定した日数(periodDays)だけ表示対象を絞り込む。0または負数の場合は全期間を返す。
 * points(日付昇順)とchartRows(同じ長さ・同じ順)はインデックスが対応している前提。
 */
function filterChartRowsByPeriod_(points, chartRows, periodDays) {
  const days = Number(periodDays);
  if (!days || days <= 0) return chartRows;
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - days);
  cutoff.setHours(0, 0, 0, 0);
  let startIdx = points.findIndex(p => p.date >= cutoff);
  if (startIdx === -1) startIdx = points.length;
  return chartRows.slice(startIdx);
}

/** 指定銘柄の終値の全履歴(DB全期間)を、日付昇順のpoints配列とタイトル文字列で返す。 */
function fetchStockRawPoints_(code) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName(DB_SHEET_NAME);

  if (!dbSheet) {
    throw new Error('データベースシートがまだ存在しません');
  }

  // 当月シートの2行目(銘柄コード・業種)と3行目(銘柄名)を1回で読み、銘柄名・業種・列位置をまとめて取る
  // (業種・列位置はbuildStockChartBundle_で使い回し、同じ行を何度も読まないようにする)
  const sheet = getTargetSheet(new Date());
  let stockName = '', industryLabel = '', sheetCol = null;
  const marketMatch = MARKET_VALUE_ITEMS.find(item => item.code === String(code));
  if (marketMatch) {
    stockName = marketMatch.name;
  } else if (sheet) {
    const lastCol = sheet.getLastColumn();
    const headRows = sheet.getRange(2, 1, 2, lastCol).getValues();
    const codeRow = headRows[0], nameRow = headRows[1];
    for (let col = STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) {
      if (String(codeRow[col - 1]) === String(code)) {
        stockName = String(nameRow[col - 1] || '');
        industryLabel = String(codeRow[col] || ''); // 業種は銘柄コードの右隣のセル
        sheetCol = col;
        break;
      }
    }
  }

  const lastColDb = dbSheet.getLastColumn();
  const headerRow = dbSheet.getRange(1, 1, 1, lastColDb).getValues()[0];

  let targetCol = -1;
  for (let col = 2; col <= lastColDb; col++) {
    if (String(headerRow[col - 1]) === String(code)) {
      targetCol = col;
      break;
    }
  }

  if (targetCol === -1) {
    throw new Error('銘柄コード ' + code + ' のデータが見つかりません');
  }

  const lastRow = dbSheet.getLastRow();
  const readRows = lastRow - 1;
  const dates = dbSheet.getRange(2, 1, readRows, 1).getValues();
  const closes = dbSheet.getRange(2, targetCol, readRows, 1).getValues();

  const points = [];
  for (let i = 0; i < dates.length; i++) {
    const d = dates[i][0];
    const c = closes[i][0];
    if (d instanceof Date && typeof c === 'number') {
      points.push({ date: d, value: c });
    }
  }
  points.sort((a, b) => a.date - b.date);

  return {
    points: points,
    titleText: chartTitleText_(code, stockName),
    industryLabel: industryLabel,
    sheetName: sheet ? sheet.getName() : '',
    sheetCol: sheetCol
  };
}

/**
 * extraの移動平均を、各行の row.extraMa に入れる(全期間で計算してから期間で絞り込むので、表示の先頭から線が出る)。
 * 直近period日のextraに空き(null)があるところは null にする。
 */
function addExtraMovingAverage_(points, chartRows, period) {
  let windowSum = 0;
  let filled = 0; // 直近period日のうち、extraが入っている日数
  points.forEach((p, idx) => {
    const v = typeof p.extra === 'number' ? p.extra : null;
    if (v !== null) { windowSum += v; filled++; }
    if (idx >= period) {
      const old = points[idx - period].extra;
      if (typeof old === 'number') { windowSum -= old; filled--; }
    }
    chartRows[idx].extraMa = idx >= period - 1 && filled === period ? windowSum / period : null;
  });
}

/**
 * 「空売り集計」シートの全期間チャートボタン用のデータを組み立てる。
 * 青=合計売買代金、赤=25日移動平均、黒=空売り(価格規制あり+なし)の売買代金合計、灰=その25日移動平均、の4系列。
 * 空売り比率(あり+なし)はグラフの線にはせず、マウスオーバーのツールチップにのみ小数点第二位まで表示する。
 * periodDays: 表示する期間(日数)。0または負数で全期間。
 */
function buildShortSellChartPayload_(col, periodDays) {
  const series = buildShortSellChartSeries(col);
  if (!series || series.points.length === 0) {
    throw new Error('この列の履歴データが見つかりません。売買代金または比率の列のセルを選択してから実行してください。');
  }

  const points = series.points.map(p => ({
    date: new Date(p.date),
    value: p.amount,
    extra: p.shortAmountSum,
    ratioInfo: p.ratioSum
  }));

  const chartRows = buildChartRowsWithMA_(points, MA_PERIOD, 0); // 前日比率は小数点なし
  addExtraMovingAverage_(points, chartRows, MA_PERIOD);
  const displayChartData = filterChartRowsByPeriod_(points, chartRows, periodDays);

  return {
    chartJson: displayChartData,
    titleText: series.title,
    maPeriod: MA_PERIOD,
    closeLabel: '売買代金合計',
    closeUnit: '百万円',
    extraLabel: '空売り(価格規制あり+なし)の合計売買代金',
    extraMaLabel: '空売り合計の' + MA_PERIOD + '日移動平均' // 凡例に収まるよう短くする
  };
}

/** 単一チャートのpayloadを、ChartDialogが扱う「panelsモード」のbundle形に包む(空売り集計のボタンチャート用)。 */
function wrapSinglePanelBundle_(payload) {
  return {
    mode: 'panels',
    titleText: payload.titleText,
    panels: [{
      chartJson: payload.chartJson,
      titleText: payload.titleText,
      maPeriod: payload.maPeriod,
      closeLabel: payload.closeLabel,
      closeUnit: payload.closeUnit,
      color: '#4285f4',
      maColor: '#ea4335',
      extraLabel: payload.extraLabel || '',
      extraColor: payload.extraLabel ? '#202124' : '',
      extraMaLabel: payload.extraMaLabel || '',
      extraMaColor: '#9e9e9e'
    }]
  };
}

/**
 * 複数系列(終値・業種売買代金合計・空売り売買代金合計など)を日付キーで1つの行配列にマージする。
 * seriesSpecs: [{ points:[{date:Date,value,ratioInfo?}], prefix:'close', rateDecimals:2 }, ...]
 * 各系列ごとに移動平均・前日比・前日比率を計算してから結合する(prefix+Value/Ma/Change/Rate/RatioInfoという
 * キー名で1行にまとめる。日付が存在しない系列のフィールドはundefinedのまま=そのグラフ線だけ途切れる)。
 */
function mergeSeriesForBandsChart_(seriesSpecs) {
  const map = {};
  seriesSpecs.forEach(spec => {
    const rows = buildChartRowsWithMA_(spec.points, MA_PERIOD, spec.rateDecimals === undefined ? 2 : spec.rateDecimals);
    spec.points.forEach((p, idx) => {
      const row = rows[idx];
      const dateKey = row.dateKey;
      if (!map[dateKey]) map[dateKey] = { dateKey: dateKey, date: row.date };
      const m = map[dateKey];
      m[spec.prefix + 'Value'] = row.close;
      m[spec.prefix + 'Change'] = row.change;
      m[spec.prefix + 'Rate'] = row.rate;
      if (row.ma !== null && row.ma !== undefined) m[spec.prefix + 'Ma'] = row.ma;
      if (row.ratioInfo !== undefined && row.ratioInfo !== null) m[spec.prefix + 'RatioInfo'] = row.ratioInfo;
    });
  });
  return Object.keys(map).sort().map(k => map[k]);
}

/** マージ後の行配列(dateKey昇順)を、指定日数(0または負数で全期間)で絞り込む。 */
function filterMergedRowsByPeriod_(rows, periodDays) {
  const days = Number(periodDays);
  if (!days || days <= 0) return rows;
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - days);
  cutoff.setHours(0, 0, 0, 0);
  const cutoffKey = Utilities.formatDate(cutoff, 'Asia/Tokyo', 'yyyy-MM-dd');
  const startIdx = rows.findIndex(r => r.dateKey >= cutoffKey);
  return startIdx === -1 ? [] : rows.slice(startIdx);
}

/**
 * 各月シートの信用残行(39行目〜)から、指定銘柄(または市場コード1001〜1003)の日次信用残の履歴を集める。
 * JPXの日次信用残は2026年10月分からのため、2026年10月以降のシートだけを読む。
 * 戻り値: [{date, sell, buy, ratio}](日付昇順。売残・買残は市場なら百万株、個別銘柄なら千株)
 * knownCols: {シート名: 列}(省略可)。列位置が分かっているシートは銘柄コードの行を読まずに済ませる。
 */
function fetchMarginHistory_(code, knownCols) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const market = MARKET_VALUE_ITEMS.find(item => item.code === String(code));
  const points = [];
  ss.getSheets().forEach(sheet => {
    const ym = parseYearMonthFromSheetName(sheet.getName());
    if (!ym || ym.year * 100 + ym.month < 202610) return;
    const days = getDaysInMonth(ym.year, ym.month);
    if (sheet.getMaxRows() < MARGIN_FIRST_ROW + days - 1) return;

    // knownCols: {シート名: 列} で列位置が分かっているシートは、2行目を読み直さない
    let col = market ? market.col : ((knownCols && knownCols[sheet.getName()]) || null);
    if (!col) {
      const lastCol = sheet.getLastColumn();
      if (lastCol < STOCK_START_COL) return;
      const codeRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
      for (let c = STOCK_START_COL; c <= lastCol; c += STOCK_COL_WIDTH) {
        if (String(codeRow[c - 1]) === String(code)) { col = c; break; }
      }
    }
    if (!col) return;

    const vals = sheet.getRange(MARGIN_FIRST_ROW, col, days, 3).getValues();
    vals.forEach((v, i) => {
      if (typeof v[0] !== 'number' || typeof v[1] !== 'number') return; // 休・未入力・対象外(-)は除く
      const ratio = typeof v[2] === 'number' ? v[2] : (v[0] > 0 ? Math.round((v[1] / v[0]) * 100) / 100 : null);
      points.push({ date: new Date(ym.year, ym.month - 1, i + 1), sell: v[0], buy: v[1], ratio: ratio });
    });
  });
  points.sort((a, b) => a.date - b.date);
  return points;
}

/**
 * 当月シートの終値ボタンチャート用のbundleを組み立てる(「bandsモード」= 1つのチャートに複数グループを統合表示)。
 * 縦の帯を3つに分けて描画する:
 *   上段: 終値(青)・25日移動平均(赤)。右の縦軸に信用倍率(緑)
 *   中段: 業種売買代金合計(水色)・その移動平均(オレンジ)と、業種の空売り(あり+なし)売買代金合計(黒)・その移動平均(グレー)
 *         (どちらも百万円単位なので同じ軸で表示)
 *   下段: 信用売残(青)・信用買残(赤)(市場は百万株、個別銘柄は千株)
 * axes: 縦軸ごとの帯の位置・目盛りの色・表示位置(left=左外側 / right=右外側 / inner=左外側に自前描画 / innerRight=右外側に自前描画)。
 *       tickCount=目盛りの数の目安(省略時4)、centerOnMean=trueなら平均値が帯の縦の中央に来るように範囲を決める
 * series: 各線。axisIndexでどの縦軸に属するかを指定する(帯への割り振りはChartDialog.html側で行う)。
 */
function buildStockChartBundle_(code, periodDays) {
  const stockRaw = fetchStockRawPoints_(code);
  const specs = [{ points: stockRaw.points, prefix: 'close', rateDecimals: 2 }];
  const axes = [{ band: [0.70, 0.98], color: '#4285f4', position: 'left', format: '#,##0' }];
  const seriesDefs = [
    { axisIndex: 0, key: 'closeValue', maKey: 'closeMa', changeKey: 'closeChange', rateKey: 'closeRate',
      color: '#4285f4', maColor: '#ea4335', label: '終値', unit: '円' }
  ];
  const notes = [];

  // 信用残(日次)の履歴
  const knownCols = {};
  if (stockRaw.sheetCol) knownCols[stockRaw.sheetName] = stockRaw.sheetCol;
  const marginPoints = fetchMarginHistory_(code, knownCols);
  const ratioPoints = marginPoints.filter(p => typeof p.ratio === 'number').map(p => ({ date: p.date, value: p.ratio }));
  if (ratioPoints.length > 0) {
    specs.push({ points: ratioPoints, prefix: 'ratio', rateDecimals: 2 });
    axes.push({ band: [0.70, 0.98], color: '#34a853', position: 'right', format: '#,##0.##' });
    seriesDefs.push({ axisIndex: axes.length - 1, key: 'ratioValue', changeKey: 'ratioChange', rateKey: 'ratioRate',
      color: '#34a853', label: '信用倍率', unit: '倍' });
  }

  // 業種売買代金合計+業種の空売り売買代金合計(同じ軸)。
  // 軸の下端を0にして、空売りの線の高さが「売買代金合計を100としたときの空売り比率」と同じ大きさになるようにする
  // (例: 空売り比率40%の日は、空売りの線が売買代金合計の線の4割の高さに来る)。
  // 空売りの線は、目盛りの大きさを売買代金合計の軸と共有する別の軸に載せ、帯の下半分にだけ補助線と目盛り(右側)を足す。
  // 当月シートに業種が無いときだけ、テンプレートも含めて探し直す
  const industryLabel = stockRaw.industryLabel || getStockIndustryLabel_(code);
  const series = industryLabel ? buildIndustryShortSellSeriesByLabel(industryLabel) : null;
  if (!industryLabel) {
    notes.push('この銘柄には業種情報が設定されていません。「機能」メニュー→「業種情報を一括取得(銘柄選択)」から設定してください。');
  } else if (!series || series.points.length === 0) {
    notes.push('業種「' + industryLabel + '」に対応する空売り集計データが見つかりません。');
  } else {
    axes.push({ band: [0.37, 0.64], color: '#4fc3f7', position: 'inner', format: '#,##0', zeroBased: true });
    const industryAxis = axes.length - 1;

    const industryPoints = series.points
      .filter(p => typeof p.amount === 'number')
      .map(p => ({ date: new Date(p.date), value: p.amount }));
    specs.push({ points: industryPoints, prefix: 'industry', rateDecimals: 2 });
    seriesDefs.push({ axisIndex: industryAxis, key: 'industryValue', maKey: 'industryMa', changeKey: 'industryChange', rateKey: 'industryRate',
      color: '#4fc3f7', maColor: '#ff9800', label: '業種(' + industryLabel + ')売買代金合計', unit: '百万円' });

    axes.push({ band: [0.37, 0.64], color: '#202124', position: 'innerRight', format: '#,##0',
      shareScaleWith: industryAxis, tickRange: [0, 0.5] });
    const shortAxis = axes.length - 1;

    const shortPoints = series.points
      .filter(p => typeof p.shortAmountSum === 'number')
      .map(p => ({ date: new Date(p.date), value: p.shortAmountSum, ratioInfo: p.ratioSum }));
    specs.push({ points: shortPoints, prefix: 'short', rateDecimals: 2 });
    seriesDefs.push({ axisIndex: shortAxis, key: 'shortValue', maKey: 'shortMa', changeKey: 'shortChange', rateKey: 'shortRate', ratioKey: 'shortRatioInfo',
      color: '#202124', maColor: '#9e9e9e', label: '業種(' + industryLabel + ')空売り(あり+なし)売買代金合計', unit: '百万円' });
  }

  // 信用売残・買残。桁が大きく違うことが多い(例: 売残5千株・買残4万株)ため、同じ軸にすると
  // 線が帯の上端と下端に分かれて見づらい。同じ帯に重ねたまま売残・買残に別々の軸を持たせ、
  // 目盛りは売残を左・買残を右に表示する。どちらも平均値が帯の縦の中央に来るように範囲を決める(centerOnMean)。
  if (marginPoints.length > 0) {
    const marginUnit = MARKET_VALUE_ITEMS.some(item => item.code === String(code)) ? '百万株' : '千株';
    axes.push({ band: [0.03, 0.30], color: '#1a73e8', position: 'inner', format: '#,##0.###', centerOnMean: true });
    const sellAxis = axes.length - 1;
    axes.push({ band: [0.03, 0.30], color: '#d93025', position: 'innerRight', format: '#,##0.###', centerOnMean: true });
    const buyAxis = axes.length - 1;
    specs.push({ points: marginPoints.map(p => ({ date: p.date, value: p.sell })), prefix: 'sell', rateDecimals: 2 });
    specs.push({ points: marginPoints.map(p => ({ date: p.date, value: p.buy })), prefix: 'buy', rateDecimals: 2 });
    seriesDefs.push({ axisIndex: sellAxis, key: 'sellValue', changeKey: 'sellChange', rateKey: 'sellRate',
      color: '#1a73e8', label: '信用売残', unit: marginUnit });
    seriesDefs.push({ axisIndex: buyAxis, key: 'buyValue', changeKey: 'buyChange', rateKey: 'buyRate',
      color: '#d93025', label: '信用買残', unit: marginUnit });
  } else {
    notes.push('信用残(日次)のデータがまだありません。');
  }

  const mergedAll = mergeSeriesForBandsChart_(specs);
  const rows = filterMergedRowsByPeriod_(mergedAll, periodDays);

  return {
    mode: 'bands',
    titleText: stockRaw.titleText,
    note: notes.join(' '),
    rows: rows,
    axes: axes,
    series: seriesDefs
  };
}

/**
 * チャートのダイアログを先に開く。データはダイアログ側からfetchChartDataForPeriodで取得する
 * (データの組み立てとGoogle Chartsの読み込みを同時に進めて、表示までの待ち時間を短くするため)。
 */
function renderChartDialog_(seriesKind, seriesId, initialPeriodDays, titleText) {
  const template = HtmlService.createTemplateFromFile('ChartDialog');
  template.seriesKind = seriesKind;
  template.seriesId = String(seriesId);
  template.initialPeriodDays = initialPeriodDays;

  // 終値ボタン(bandsモード)は画面いっぱいに近い高さで開く(表示が窮屈な場合はCHART_DIALOG_HEIGHTを調整)
  const height = seriesKind === 'stock' ? CHART_DIALOG_HEIGHT : 530;
  const html = template.evaluate().setWidth(1050).setHeight(height);
  SpreadsheetApp.getUi().showModalDialog(html, titleText);
}

/**
 * ChartDialog内の「参照期間」セレクトから呼ばれる。seriesKindに応じて該当するbundle組み立て関数に振り分ける。
 * 戻り値はrenderChartDialog_が使うbundleと同じ形({ mode, titleText, ... })。
 */
function fetchChartDataForPeriod(seriesKind, seriesId, periodDays) {
  if (seriesKind === 'stock') {
    return buildStockChartBundle_(seriesId, periodDays);
  }
  if (seriesKind === 'shortsell') {
    return wrapSinglePanelBundle_(buildShortSellChartPayload_(Number(seriesId), periodDays));
  }
  throw new Error('不明な種別です: ' + seriesKind);
}

/** stockNameはタイトル用(省略可)。データの取得でエラーになった場合はダイアログ内に表示される */
function showMovingAverageChart(code, stockName) {
  renderChartDialog_('stock', code, CHART_DISPLAY_MONTHS * 30, chartTitleText_(code, stockName));
}

/**
 * 「空売り集計」シートの全期間チャートボタン用。選択中のセルの列(売買代金 or 比率)から
 * 全履歴を取得し、当月シートの終値ボタンチャートと同じポップアップ(ChartDialog)で表示する。
 * シート上に手動で配置したボタンにこの関数を割り当てて使う。
 */
function showShortSellChartForActiveCell() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== '空売り集計') {
    SpreadsheetApp.getUi().alert('このボタンは「空売り集計」シートで使用してください。');
    return;
  }

  const col = sheet.getActiveCell().getColumn();
  const block = SHORT_SELL_COLUMN_BLOCKS.find(b => col === b.col || col === b.col + 1);
  if (!block) {
    SpreadsheetApp.getUi().alert('この列の履歴データが見つかりません。売買代金または比率の列のセルを選択してから実行してください。');
    return;
  }
  renderChartDialog_('shortsell', col, 0, block.label); // 初期表示は全期間
}