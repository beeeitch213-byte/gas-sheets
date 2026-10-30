// ==== 決算発表予定(決算予定シート) ====
// データ元: JPX「決算発表予定日」 https://www.jpx.co.jp/listing/event-schedules/financial-announcement/index.html
//   四半期末・期末の月ごとに Excel(.xlsx)で公開され、会社が発表日を決めるたびに毎営業日更新される。
//   列: A=決算発表予定日(Excelの日付) / B=コード / C=会社名 / E=決算期末 / F=業種名 / H=種別(第２四半期など)
//   発表時刻は載っていないので、時刻の列は空欄にしておく。
// シートには、株式情報に登録している銘柄のうち、今日から次月末までに発表予定のものを並べる。
// 「評価する」のチェックは銘柄ごとにスクリプトプロパティに覚えておき、作り直しても引き継ぐ
// (チェックした銘柄の決算を、発表後に評価して知らせる仕組みで使う)。
const KESSAN_SHEET_NAME = '決算予定';

const KESSAN_PAGE_URL = 'https://www.jpx.co.jp/listing/event-schedules/financial-announcement/index.html';

const KESSAN_EVAL_CODES_KEY = 'KESSAN_EVAL_CODES'; // 「評価する」にチェックした銘柄コードの配列(JSON)

const KESSAN_HEADERS = ['発表予定日', '曜日', '時刻', 'コード', '銘柄名', '種別', '決算期末', '業種', '評価する'];

const KESSAN_FIRST_ROW = 4; // 1行目=タイトル、2行目=説明、3行目=見出し

const KESSAN_CODE_COL = 4;

const KESSAN_EVAL_COL = 9;

/** XMLの文字参照(&amp; &#12354; など)を元の文字に戻す */
function decodeXmlEntities_(s) {
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (m, e) => {
    if (e[0] === '#') return String.fromCodePoint(e[1] === 'x' ? parseInt(e.substring(2), 16) : parseInt(e.substring(1), 10));
    return { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }[e];
  });
}

/**
 * xlsxの1枚目のシートの全行を、行ごとの配列(1列目から。数値セルは数値、文字セルは文字列、空は'')で返す。
 * readXlsxRows_(08_投資部門別)は数値セルしか読まないため、会社名などの文字も読めるようにした版。
 */
function readXlsxAllRows_(blob) {
  const files = Utilities.unzip(blob.setContentType('application/zip'));
  const fileOf = name => files.find(f => f.getName() === name);
  const sheetFile = fileOf('xl/worksheets/sheet1.xml');
  if (!sheetFile) return [];

  const shared = [];
  const sharedFile = fileOf('xl/sharedStrings.xml');
  if (sharedFile) {
    const sxml = sharedFile.getDataAsString('UTF-8');
    const siRe = /<si>([\s\S]*?)<\/si>/g;
    let si;
    while ((si = siRe.exec(sxml)) !== null) {
      let text = '';
      const tRe = /<t[^>]*>([\s\S]*?)<\/t>/g;
      let t;
      while ((t = tRe.exec(si[1])) !== null) text += t[1];
      shared.push(decodeXmlEntities_(text));
    }
  }

  const colIndex = letters => letters.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const xml = sheetFile.getDataAsString('UTF-8');
  const rows = [];
  const rowRe = /<row\b[^>]*>([\s\S]*?)<\/row>/g;
  let rm;
  while ((rm = rowRe.exec(xml)) !== null) {
    const values = [];
    const cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    let cm;
    while ((cm = cellRe.exec(rm[1])) !== null) {
      const ref = cm[1].match(/\br="([A-Z]+)\d+"/);
      if (!ref) continue;
      const type = (cm[1].match(/\bt="([^"]*)"/) || [])[1];
      const inner = cm[2] || '';
      const v = inner.match(/<v>([^<]*)<\/v>/);
      let value = '';
      if (type === 's' && v) value = shared[parseInt(v[1], 10)] || '';
      else if (type === 'inlineStr') value = decodeXmlEntities_((inner.match(/<t[^>]*>([\s\S]*?)<\/t>/) || ['', ''])[1]);
      else if (type === 'str' && v) value = decodeXmlEntities_(v[1]);
      else if (v && v[1] !== '' && !isNaN(Number(v[1]))) value = Number(v[1]);
      const col = colIndex(ref[1]);
      while (values.length < col - 1) values.push('');
      values[col - 1] = value;
    }
    rows.push(values);
  }
  return rows;
}

/** Excelの日付(1900年方式のシリアル値)をDateにする */
function excelSerialToDate_(serial) {
  return new Date(1899, 11, 30 + Math.floor(serial));
}

/**
 * JPXの決算発表予定日の全ファイルから、指定した銘柄の、from〜to(両端を含む)に発表予定の行を集める。
 * 戻り値: [{date, code, name, kind, fiscalEnd, industry}](日付→コード順)
 */
function fetchKessanSchedule_(codeSet, from, to) {
  const page = UrlFetchApp.fetch(KESSAN_PAGE_URL, { muteHttpExceptions: true });
  if (page.getResponseCode() !== 200) throw new Error('JPXの決算発表予定日のページを読めませんでした');
  const links = [];
  const linkRe = /href="([^"]+\.xlsx)"/g;
  let m;
  while ((m = linkRe.exec(page.getContentText())) !== null) {
    const url = m[1].startsWith('http') ? m[1] : 'https://www.jpx.co.jp' + m[1];
    if (links.indexOf(url) === -1) links.push(url);
  }
  if (links.length === 0) throw new Error('JPXの決算発表予定日のExcelが見つかりませんでした');

  const responses = UrlFetchApp.fetchAll(links.map(url => ({ url: url, muteHttpExceptions: true })));
  const byKey = {}; // 同じ銘柄・同じ日が複数のファイルに載っていても1行にする
  responses.forEach(r => {
    if (r.getResponseCode() !== 200) return;
    readXlsxAllRows_(r.getBlob()).forEach(row => {
      if (typeof row[0] !== 'number') return; // 見出しなど、日付で始まらない行
      const code = String(row[1]).trim().toUpperCase();
      if (!codeSet[code]) return;
      const date = excelSerialToDate_(row[0]);
      if (date < from || date > to) return;
      const key = code + '_' + row[0];
      byKey[key] = {
        date: date,
        code: code,
        name: String(row[2] || ''),
        kind: String(row[7] || ''),
        fiscalEnd: typeof row[4] === 'number' ? excelSerialToDate_(row[4]) : '',
        industry: String(row[5] || '')
      };
    });
  });
  return Object.keys(byKey).map(k => byKey[k]).sort((a, b) => (a.date - b.date) || (a.code < b.code ? -1 : a.code > b.code ? 1 : 0));
}

/** 「評価する」にチェックした銘柄コードの集合 {コード: true} */
function getKessanEvalCodes_() {
  const saved = JSON.parse(PropertiesService.getScriptProperties().getProperty(KESSAN_EVAL_CODES_KEY) || '[]');
  const set = {};
  saved.forEach(c => { set[String(c).toUpperCase()] = true; });
  return set;
}

/** 今の決算予定シートのチェックを読み、銘柄ごとの「評価する」の記憶に反映して返す(シートに無い銘柄は前の記憶のまま) */
function syncKessanEvalCodesFromSheet_(sheet) {
  const set = getKessanEvalCodes_();
  if (sheet && sheet.getLastRow() >= KESSAN_FIRST_ROW) {
    const rows = sheet.getRange(KESSAN_FIRST_ROW, 1, sheet.getLastRow() - KESSAN_FIRST_ROW + 1, KESSAN_EVAL_COL).getValues();
    rows.forEach(r => {
      const code = String(r[KESSAN_CODE_COL - 1]).trim().toUpperCase();
      if (!code) return;
      if (r[KESSAN_EVAL_COL - 1] === true) set[code] = true;
      else delete set[code];
    });
  }
  PropertiesService.getScriptProperties().setProperty(KESSAN_EVAL_CODES_KEY, JSON.stringify(Object.keys(set)));
  return set;
}

/**
 * 決算予定シートを、今日から次月末までの発表予定で作り直す(登録銘柄のみ)。結果のメッセージを返す。
 * 月末の6日前からは毎日の更新(createNextMonthSheetIfNeeded)で自動的に呼ばれる。メニューからも実行できる。
 */
function updateKessanSchedule() {
  const ss = getMainSpreadsheet();
  const today = new Date();
  const from = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const to = new Date(today.getFullYear(), today.getMonth() + 2, 0); // 次月の末日

  const codeSet = {};
  const nameByCode = {};
  getCurrentStockOrder().forEach(s => {
    const code = String(s.code).trim().toUpperCase();
    codeSet[code] = true;
    nameByCode[code] = s.name;
  });
  if (Object.keys(codeSet).length === 0) return '当月シートに銘柄が登録されていません';

  const schedule = fetchKessanSchedule_(codeSet, from, to);

  let sheet = ss.getSheetByName(KESSAN_SHEET_NAME);
  const evalCodes = syncKessanEvalCodesFromSheet_(sheet); // 作り直す前に、今のチェックを記憶に反映しておく
  if (!sheet) sheet = ss.insertSheet(KESSAN_SHEET_NAME);
  sheet.clear();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations(); // 前回のチェックボックスも消す

  const fmt = (d, f) => Utilities.formatDate(d, 'Asia/Tokyo', f);
  sheet.getRange(1, 1).setValue('決算発表予定(' + fmt(from, 'yyyy/MM/dd') + '〜' + fmt(to, 'yyyy/MM/dd') + ')').setFontWeight('bold').setFontSize(12);
  sheet.getRange(2, 1).setValue('JPX「決算発表予定日」より登録銘柄を抜き出し(' + fmt(new Date(), 'yyyy/MM/dd HH:mm') + '更新)。'
    + '「評価する」にチェックした銘柄は、決算発表後に評価してお知らせします(チェックは銘柄ごとに次回以降も引き継ぎます)。')
    .setFontColor('#666666');
  sheet.getRange(3, 1, 1, KESSAN_HEADERS.length).setValues([KESSAN_HEADERS]).setFontWeight('bold').setBackground('#f1f3f4');
  sheet.setFrozenRows(3);

  if (schedule.length > 0) {
    const values = schedule.map(s => [
      s.date, MARGIN_WEEKDAY_NAMES_[s.date.getDay()].charAt(0), '', s.code, nameByCode[s.code] || s.name,
      s.kind, s.fiscalEnd, s.industry, !!evalCodes[s.code]
    ]);
    // 表示形式は値より先に設定する(銘柄コードを文字列にしておかないと、7203などが数値になる)
    sheet.getRange(KESSAN_FIRST_ROW, 1, values.length, 1).setNumberFormat('yyyy/MM/dd');
    sheet.getRange(KESSAN_FIRST_ROW, 7, values.length, 1).setNumberFormat('yyyy/MM');
    sheet.getRange(KESSAN_FIRST_ROW, KESSAN_CODE_COL, values.length, 1).setNumberFormat('@');
    sheet.getRange(KESSAN_FIRST_ROW, KESSAN_EVAL_COL, values.length, 1).insertCheckboxes();
    sheet.getRange(KESSAN_FIRST_ROW, 1, values.length, KESSAN_HEADERS.length).setValues(values);
  } else {
    sheet.getRange(KESSAN_FIRST_ROW, 1).setValue('この期間に発表予定の登録銘柄はありません(JPXの一覧に載っていない会社は、発表日が決まると追加されます)');
  }
  sheet.autoResizeColumns(1, KESSAN_HEADERS.length);

  return '決算予定を更新しました(' + fmt(from, 'MM/dd') + '〜' + fmt(to, 'MM/dd') + ': ' + schedule.length + '件)';
}
