// ==== プログラム売買・裁定取引(プログラムシート) ====
// JPX「プログラム売買・裁定取引」の2つのページから取り、プログラムシートに新しい日を上にして並べる。
//  ・裁定取引(日ごと) https://www.jpx.co.jp/markets/statistics-equities/program/index.html
//      毎営業日16時ごろに前々営業日の分が載る。ページには直近10日分しか無い(それより前のファイルはたどれない)
//  ・プログラム売買(週ごと) https://www.jpx.co.jp/markets/statistics-equities/program/01.html
//      毎週第3営業日16時ごろに前の週の分が載る。バックナンバー(01-archives-NN.html)は2023年から
// どちらも Excel(.xls、BIFF8)を自前で読む(PDF は自前の解析では数字が抜けるため使わない)。
//
// シートの形(1日2行。5行目から下へ、新しい日が上):
//   A列=日付(2行を結合。週ごとのデータしか無い日は「(週)」を付けて表示する)、B列=買付け/売付け
//   C〜G列=裁定取引に係る現物ポジション(当限・翌限以降・合計・前日比(%)・週間売買金額)
//     買付けの行は買いポジション、売付けの行は売りポジション。前日比は合計の変化率(週ごとの行は前週末比)
//     週間売買金額はプログラム売買全体の金額で、その週の最後の営業日の行に入れる
//   H列から2列ずつ=証券会社(3行目=名前、4行目=枚数/比率)。買付けの行は買付株数・比率、売付けの行は売付株数・比率
//     知らない証券会社が出てきたら右に2列足す。その日に名前が無い証券会社は空欄
// シートの中身をそのまま保存先として使う(読み直して、足して、全部書き直す)。何度実行しても同じ結果になる。

const PROGRAM_SHEET_NAME = 'プログラム';

const PROGRAM_JPX_ORIGIN = 'https://www.jpx.co.jp';

const PROGRAM_DAILY_PAGE_URL = 'https://www.jpx.co.jp/markets/statistics-equities/program/index.html';

const PROGRAM_WEEKLY_PAGE_URL = 'https://www.jpx.co.jp/markets/statistics-equities/program/01.html';

const PROGRAM_BROKER_NAME_ROW = 3; // 証券会社の名前(2列を結合)

const PROGRAM_HEADER_ROW = 4;      // 当限・翌限以降・…・枚数・比率

const PROGRAM_FIRST_ROW = 5;       // データの始まり(1日2行: 買付け・売付け)

const PROGRAM_BROKER_START_COL = 8; // H列

const PROGRAM_DATA_COLS = 5;        // C〜G列(当限・翌限以降・合計・前日比・週間売買金額)

const PROGRAM_DAY_FORMAT = 'yyyy/m/d';

const PROGRAM_WEEK_FORMAT = 'yyyy/m/d"(週)"';

// JPX の証券会社名(全角を半角にし、空白を除いたもの) → シートの名前
const PROGRAM_BROKER_ALIASES = {
  '三菱UFJモルガン・スタンレー': '三菱UFJ MS',
  'みずほ': 'みずほ',
  'ジェフリーズ日本': 'ジェフリーズJ',
  'ジェフリーズ': 'ジェフリーズJ',
  '野村': '野村',
  'UBS': 'UBS',
  'ソシエテ・ジェネラル': 'ソシエテ・G'
};

// ---- Excel(.xls)の読み取り ----

/**
 * .xls(BIFF8)の最初のワークシートを、セルの2次元配列(行・列は0始まり、空のセルは '')で返す。
 * 読めなければ例外。bytes は Uint8Array(0〜255)。
 */
function xlsReadFirstSheet_(bytes) {
  const wb = xlsReadCfbStream_(bytes, ['Workbook', 'Book']);
  const dv = new DataView(wb.buffer, wb.byteOffset, wb.byteLength);
  const cells = [];
  const put = (r, c, v) => { (cells[r] = cells[r] || [])[c] = v; };
  let sst = [];
  let pos = 0, inSheet = false, sheetDone = false;
  while (pos + 4 <= wb.length && !sheetDone) {
    const type = dv.getUint16(pos, true), len = dv.getUint16(pos + 2, true);
    const start = pos + 4;
    pos = start + len;
    if (type === 0x00FC) { // SST(続きは CONTINUE レコードに分かれる)
      const parts = [wb.subarray(start, start + len)];
      while (pos + 4 <= wb.length && dv.getUint16(pos, true) === 0x003C) {
        const l = dv.getUint16(pos + 2, true);
        parts.push(wb.subarray(pos + 4, pos + 4 + l));
        pos += 4 + l;
      }
      sst = xlsReadSst_(parts);
    } else if (type === 0x0809) { // BOF
      const kind = dv.getUint16(start + 2, true);
      if (kind === 0x0010) inSheet = true; // ワークシート
    } else if (type === 0x000A) { // EOF
      if (inSheet) sheetDone = true;
    } else if (inSheet) {
      if (type === 0x00FD) put(dv.getUint16(start, true), dv.getUint16(start + 2, true), sst[dv.getUint32(start + 6, true)] || '');
      else if (type === 0x0203) put(dv.getUint16(start, true), dv.getUint16(start + 2, true), dv.getFloat64(start + 6, true));
      else if (type === 0x027E) put(dv.getUint16(start, true), dv.getUint16(start + 2, true), xlsRk_(dv, start + 6));
      else if (type === 0x00BD) { // MULRK
        const r = dv.getUint16(start, true), c0 = dv.getUint16(start + 2, true);
        const n = (len - 6) / 6;
        for (let i = 0; i < n; i++) put(r, c0 + i, xlsRk_(dv, start + 4 + i * 6 + 2));
      } else if (type === 0x0006) { // FORMULA(計算結果が数値のときだけ使う)
        if (dv.getUint16(start + 12, true) !== 0xFFFF) put(dv.getUint16(start, true), dv.getUint16(start + 2, true), dv.getFloat64(start + 6, true));
      }
    }
  }
  if (!inSheet) throw new Error('ワークシートが見つかりません');
  const width = cells.reduce((m, row) => Math.max(m, row ? row.length : 0), 0);
  return Array.from({ length: cells.length }, (_, r) => Array.from({ length: width }, (_, c) => {
    const v = cells[r] && cells[r][c];
    return v === undefined ? '' : v;
  }));
}

/** RK 値(4バイト)を数値にする */
function xlsRk_(dv, offset) {
  const rk = dv.getUint32(offset, true);
  let v;
  if (rk & 2) v = rk >> 2; // 整数(符号付き30ビット)
  else {
    const buf = new DataView(new ArrayBuffer(8));
    buf.setUint32(4, rk & 0xFFFFFFFC, true);
    v = buf.getFloat64(0, true);
  }
  return (rk & 1) ? v / 100 : v;
}

/** SST(共有文字列の表)を読む。parts は SST と続きの CONTINUE の中身 */
function xlsReadSst_(parts) {
  let pi = 0, p = 0, cur = parts[0];
  const next = () => { pi++; p = 0; cur = parts[pi]; };
  const u8 = () => { if (p >= cur.length) next(); return cur[p++]; };
  const u16 = () => u8() | (u8() << 8);
  const u32 = () => (u16() | (u16() << 16)) >>> 0;
  u32(); // 全体の数
  const count = u32();
  const out = [];
  for (let i = 0; i < count && pi < parts.length; i++) {
    if (p >= cur.length) next();
    const n = u16();
    let flags = u8();
    const rich = (flags & 0x08) ? u16() : 0;
    const ext = (flags & 0x04) ? u32() : 0;
    let s = '';
    for (let k = 0; k < n; k++) {
      if (p >= cur.length) { next(); flags = u8(); } // CONTINUE の先頭で、文字の幅を示すバイトが入り直す
      s += String.fromCharCode((flags & 0x01) ? u16() : u8());
    }
    for (let k = 0; k < rich * 4 + ext; k++) u8();
    out.push(s);
  }
  return out;
}

/** 複合ファイル(OLE2)から、名前の合うストリームの中身を取り出す */
function xlsReadCfbStream_(bytes, names) {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (dv.getUint32(0, true) !== 0xE011CFD0) throw new Error('Excel(.xls)のファイルではありません');
  const secSize = 1 << dv.getUint16(30, true), miniSecSize = 1 << dv.getUint16(32, true);
  const nFat = dv.getUint32(44, true), dirStart = dv.getUint32(48, true);
  const miniCutoff = dv.getUint32(56, true), miniFatStart = dv.getUint32(60, true);
  let difatStart = dv.getUint32(68, true);
  const secOff = s => (s + 1) * secSize;
  const fatSecs = [];
  for (let i = 0; i < 109 && fatSecs.length < nFat; i++) fatSecs.push(dv.getUint32(76 + i * 4, true));
  while (fatSecs.length < nFat && difatStart < 0xFFFFFFFA) {
    const o = secOff(difatStart), per = secSize / 4 - 1;
    for (let i = 0; i < per && fatSecs.length < nFat; i++) fatSecs.push(dv.getUint32(o + i * 4, true));
    difatStart = dv.getUint32(o + per * 4, true);
  }
  const fat = [];
  fatSecs.forEach(s => { const o = secOff(s); for (let i = 0; i < secSize / 4; i++) fat.push(dv.getUint32(o + i * 4, true)); });
  const chain = (start, table) => {
    const out = [];
    for (let s = start; s < 0xFFFFFFFA && out.length <= table.length; s = table[s]) out.push(s);
    return out;
  };
  const readChain = (start, size) => {
    const out = new Uint8Array(size);
    let w = 0;
    chain(start, fat).forEach(s => {
      const n = Math.min(secSize, size - w);
      if (n > 0) out.set(bytes.subarray(secOff(s), secOff(s) + n), w);
      w += n;
    });
    return out;
  };
  const dirSecs = chain(dirStart, fat);
  const entries = [];
  dirSecs.forEach(s => {
    for (let o = secOff(s); o < secOff(s) + secSize; o += 128) {
      const nameLen = dv.getUint16(o + 64, true);
      let name = '';
      for (let i = 0; i < nameLen / 2 - 1; i++) name += String.fromCharCode(dv.getUint16(o + i * 2, true));
      entries.push({ name: name, type: dv.getUint8(o + 66), start: dv.getUint32(o + 116, true), size: dv.getUint32(o + 120, true) });
    }
  });
  const entry = entries.find(e => e.type === 2 && names.indexOf(e.name) !== -1);
  if (!entry) throw new Error('ブックのデータが見つかりません');
  if (entry.size >= miniCutoff) return readChain(entry.start, entry.size);
  // 小さいストリームはミニストリームに入っている
  const root = entries.find(e => e.type === 5);
  const mini = readChain(root.start, root.size);
  const miniFatBytes = readChain(miniFatStart, chain(miniFatStart, fat).length * secSize);
  const mdv = new DataView(miniFatBytes.buffer);
  const miniFat = [];
  for (let i = 0; i < miniFatBytes.length / 4; i++) miniFat.push(mdv.getUint32(i * 4, true));
  const out = new Uint8Array(entry.size);
  let w = 0;
  chain(entry.start, miniFat).forEach(s => {
    const n = Math.min(miniSecSize, entry.size - w);
    if (n > 0) out.set(mini.subarray(s * miniSecSize, s * miniSecSize + n), w);
    w += n;
  });
  return out;
}

// ---- JPX のファイルの読み取り ----

/** 空白(全角も)を除き、全角英数字を半角にする */
function programNormalize_(s) {
  return String(s).replace(/[！-～]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0)).replace(/[\s　]/g, '');
}

function programFindRow_(grid, from, pred) {
  for (let r = from; r < grid.length; r++) if (pred(grid[r].map(programNormalize_))) return r;
  return -1;
}

function programNum_(v) {
  return typeof v === 'number' ? v : null;
}

/** 変化率(増減 ÷ 前の値)。前の値が0なら '' */
function programChangeRate_(total, diff) {
  if (total === null || diff === null) return '';
  const before = total - diff;
  return before ? diff / before : '';
}

/**
 * 「２．裁定取引に係る現物ポジション」を読む。{buy: [当限, 翌限以降, 合計, 変化率], sell: [...]}
 * diffLabel: 増減の行の名前(日ごと: 前日比、週ごと: 前週末比)
 */
function programReadPositions_(grid, diffLabel) {
  const sec = programFindRow_(grid, 0, row => row.some(v => v.indexOf('裁定取引に係る現物ポジション') !== -1));
  const head = programFindRow_(grid, sec + 1, row => row.indexOf('当限') !== -1);
  const shares = programFindRow_(grid, head + 1, row => row.indexOf('株数') !== -1);
  const diff = programFindRow_(grid, shares + 1, row => row.indexOf(diffLabel) !== -1);
  if (sec < 0 || head < 0 || shares < 0 || diff < 0) throw new Error('裁定取引に係る現物ポジションの表が見つかりません');
  const h = grid[head].map(programNormalize_);
  const colsOf = label => h.map((v, i) => v === label ? i : -1).filter(i => i >= 0);
  const cur = colsOf('当限'), next = colsOf('翌限以降'), total = colsOf('合計');
  const side = k => {
    const cols = [cur[k], next[k], total[k]];
    const vals = cols.map(c => programNum_(grid[shares][c]));
    return vals.concat([programChangeRate_(vals[2], programNum_(grid[diff][total[k]]))]);
  };
  return { sell: side(0), buy: side(1) }; // 左が売りポジション、右が買いポジション
}

/** 「取引参加者別裁定取引の状況」を読む。{シートの名前: [買付株数, 買付比率, 売付株数, 売付比率]}(比率は 0〜1) */
function programReadBrokers_(grid) {
  const head = programFindRow_(grid, 0, row => row.indexOf('証券会社名') !== -1);
  if (head < 0) throw new Error('取引参加者別の表が見つかりません');
  const h = grid[head].map(programNormalize_), h2 = grid[head + 1].map(programNormalize_);
  const nameCol = h.indexOf('証券会社名'), sellCol = h.indexOf('売付株数'), buyCol = h.indexOf('買付株数');
  const ratioAfter = c => h2.findIndex((v, i) => i > c && v === '比率');
  const sellRatioCol = ratioAfter(sellCol), buyRatioCol = ratioAfter(buyCol);
  const brokers = {};
  for (let r = head + 2; r < grid.length; r++) {
    const name = programNormalize_(grid[r][nameCol]);
    if (!name || /^(上位|全社)/.test(name)) break;
    if (name === '-') continue;
    const ratio = v => typeof v === 'number' ? v / 100 : '';
    brokers[PROGRAM_BROKER_ALIASES[name] || name] = [
      programNum_(grid[r][buyCol]), ratio(grid[r][buyRatioCol]), programNum_(grid[r][sellCol]), ratio(grid[r][sellRatioCol])
    ];
  }
  return brokers;
}

/** 「１．プログラム売買に係る現物株式の売買」のプログラム売買全体の金額 [買付け, 売付け](百万円) */
function programReadWeeklyAmount_(grid) {
  const sec = programFindRow_(grid, 0, row => row.some(v => v.indexOf('プログラム売買に係る現物株式の売買') !== -1));
  const head = programFindRow_(grid, sec + 1, row => row.indexOf('売付け') !== -1);
  const amount = programFindRow_(grid, head + 1, row => row.indexOf('金額') !== -1);
  if (sec < 0 || head < 0 || amount < 0) throw new Error('プログラム売買の金額の表が見つかりません');
  const h = grid[head].map(programNormalize_);
  return [programNum_(grid[amount][h.indexOf('買付け')]), programNum_(grid[amount][h.indexOf('売付け')])]; // 最初の組がプログラム売買全体
}

/** 日ごとのファイル → 記録 */
function programParseDaily_(grid, key) {
  const pos = programReadPositions_(grid, '前日比');
  return { key: key, kind: 'day', buy: pos.buy, sell: pos.sell, amount: null, brokers: programReadBrokers_(grid) };
}

/** 週ごとのファイル → 記録(その週の最後の営業日の日付) */
function programParseWeekly_(grid, key) {
  const pos = programReadPositions_(grid, '前週末比');
  return { key: key, kind: 'week', buy: pos.buy, sell: pos.sell, amount: programReadWeeklyAmount_(grid), brokers: programReadBrokers_(grid) };
}

// ---- JPX のページ(ファイルの一覧) ----

function programFetchText_(url) {
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error(url + ' を開けませんでした(HTTP ' + res.getResponseCode() + ')');
  return res.getContentText('UTF-8');
}

/** 日ごとのファイルの一覧 [{key: 'yyyy-MM-dd', url}] */
function programDailyFiles_() {
  const html = programFetchText_(PROGRAM_DAILY_PAGE_URL);
  const out = [];
  (html.match(/<tr[\s\S]*?<\/tr>/g) || []).forEach(tr => {
    const d = tr.match(/(\d{4})年(\d{1,2})月(\d{1,2})日売買分/), x = tr.match(/href="([^"]+\.xls)"/);
    if (d && x) out.push({ key: d[1] + '-' + ('0' + d[2]).slice(-2) + '-' + ('0' + d[3]).slice(-2), url: PROGRAM_JPX_ORIGIN + x[1] });
  });
  return out;
}

/**
 * 週ごとのファイルの一覧 [{key: 週の最後の営業日 'yyyy-MM-dd', url}]。新しい順。
 * withArchives: バックナンバー(2023年から)のページも見る。false なら最新のページ(直近5週)だけ
 */
function programWeeklyFiles_(withArchives) {
  const first = programFetchText_(PROGRAM_WEEKLY_PAGE_URL);
  const pages = [first];
  if (withArchives) (first.match(/01-archives-\d+\.html/g) || []).filter((v, i, a) => a.indexOf(v) === i).forEach(name => {
    pages.push(programFetchText_(PROGRAM_WEEKLY_PAGE_URL.replace('01.html', name)));
  });
  const seen = {}, out = [];
  pages.forEach(html => (html.match(/href="[^"]+\/(\d{8})\.xls"/g) || []).forEach(a => {
    const m = a.match(/href="([^"]+\/(\d{4})(\d{2})(\d{2})\.xls)"/);
    const key = m[2] + '-' + m[3] + '-' + m[4];
    if (!seen[key]) { seen[key] = true; out.push({ key: key, url: PROGRAM_JPX_ORIGIN + m[1] }); }
  }));
  return out.sort((a, b) => a.key < b.key ? 1 : -1);
}

/** ファイルをまとめて取りに行き、表にする。[{file, grid}](取れなかったものは grid が null) */
function programFetchGrids_(files) {
  const responses = UrlFetchApp.fetchAll(files.map(f => ({ url: f.url, muteHttpExceptions: true })));
  return files.map((f, i) => {
    const res = responses[i];
    if (res.getResponseCode() !== 200) return { file: f, grid: null, error: 'HTTP ' + res.getResponseCode() };
    try {
      return { file: f, grid: xlsReadFirstSheet_(new Uint8Array(res.getContent())) };
    } catch (e) {
      return { file: f, grid: null, error: e.message };
    }
  });
}

// ---- シートの読み書き ----

/** シートの中身を記録にする。{records: {key: 記録}, brokers: [シートの証券会社名(左から)]} */
function programReadSheet_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), PROGRAM_BROKER_START_COL - 1);
  const names = sheet.getRange(PROGRAM_BROKER_NAME_ROW, 1, 1, lastCol).getValues()[0];
  const brokers = [];
  for (let c = PROGRAM_BROKER_START_COL; c <= lastCol; c += 2) if (names[c - 1]) brokers.push({ name: String(names[c - 1]), col: c });

  const records = {};
  const lastRow = sheet.getLastRow();
  if (lastRow >= PROGRAM_FIRST_ROW) {
    const n = lastRow - PROGRAM_FIRST_ROW + 1;
    const values = sheet.getRange(PROGRAM_FIRST_ROW, 1, n, lastCol).getValues();
    const formats = sheet.getRange(PROGRAM_FIRST_ROW, 1, n, 1).getNumberFormats();
    for (let i = 0; i < n; i++) {
      const date = values[i][0];
      if (!(date instanceof Date)) continue;
      const b = values[i], s = values[i + 1] || b.map(() => '');
      const key = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
      const rec = {
        key: key,
        kind: String(formats[i][0]).indexOf('(週)') !== -1 ? 'week' : 'day',
        buy: b.slice(2, 6), sell: s.slice(2, 6),
        amount: (b[6] !== '' || s[6] !== '') ? [b[6], s[6]] : null,
        brokers: {}
      };
      brokers.forEach(br => {
        const v = [b[br.col - 1], b[br.col], s[br.col - 1], s[br.col]];
        if (v.some(x => x !== '')) rec.brokers[br.name] = v;
      });
      records[key] = rec;
      i++; // 売付けの行を飛ばす
    }
  }
  return { records: records, brokers: brokers.map(b => b.name) };
}

/** 記録をシートに書き直す(新しい日が上)。足りない証券会社の列・行は足す */
function programWriteSheet_(sheet, records, brokerNames) {
  const oldWidth = PROGRAM_BROKER_START_COL - 1 + brokerNames.length * 2;
  const list = Object.keys(records).sort().reverse().map(k => records[k]);
  Object.keys(records).forEach(k => Object.keys(records[k].brokers).forEach(name => {
    if (brokerNames.indexOf(name) === -1) brokerNames.push(name);
  }));
  const width = PROGRAM_BROKER_START_COL - 1 + brokerNames.length * 2;
  const rows = list.length * 2;
  if (rows === 0) return;
  const lastRow = PROGRAM_FIRST_ROW + rows - 1;
  if (sheet.getMaxRows() < lastRow) sheet.insertRowsAfter(sheet.getMaxRows(), lastRow - sheet.getMaxRows());
  if (sheet.getMaxColumns() < width) sheet.insertColumnsAfter(sheet.getMaxColumns(), width - sheet.getMaxColumns());

  // 見出し(証券会社)。新しく足した証券会社の列は、H・I列(最初の証券会社)の書式をまねる
  const nameRow = [], headRow = [];
  brokerNames.forEach(name => { nameRow.push(name, ''); headRow.push('枚数', '比率'); });
  const ssId = sheet.getParent().getId(), sheetId = sheet.getSheetId();
  const grid = (r1, r2, c1, c2) => ({ sheetId: sheetId, startRowIndex: r1 - 1, endRowIndex: r2, startColumnIndex: c1 - 1, endColumnIndex: c2 });
  const requests = [];
  if (width > oldWidth && oldWidth > PROGRAM_BROKER_START_COL) {
    requests.push({ copyPaste: { source: grid(PROGRAM_BROKER_NAME_ROW, lastRow, PROGRAM_BROKER_START_COL, PROGRAM_BROKER_START_COL + 1), destination: grid(PROGRAM_BROKER_NAME_ROW, lastRow, oldWidth + 1, width), pasteType: 'PASTE_FORMAT' } });
  }
  // データの行の書式は、いちばん上の2行(5・6行目)をまねる
  if (rows > 2) {
    requests.push({ copyPaste: { source: grid(PROGRAM_FIRST_ROW, PROGRAM_FIRST_ROW + 1, 1, width), destination: grid(PROGRAM_FIRST_ROW + 2, lastRow, 1, width), pasteType: 'PASTE_FORMAT' } });
  }
  requests.push({ unmergeCells: { range: grid(PROGRAM_FIRST_ROW, lastRow, 1, 1) } });
  for (let r = PROGRAM_FIRST_ROW; r < lastRow; r += 2) requests.push({ mergeCells: { range: grid(r, r + 1, 1, 1), mergeType: 'MERGE_ALL' } });
  requests.push({ unmergeCells: { range: grid(PROGRAM_BROKER_NAME_ROW, PROGRAM_BROKER_NAME_ROW, PROGRAM_BROKER_START_COL, width) } });
  for (let c = PROGRAM_BROKER_START_COL; c < width; c += 2) requests.push({ mergeCells: { range: grid(PROGRAM_BROKER_NAME_ROW, PROGRAM_BROKER_NAME_ROW, c, c + 1), mergeType: 'MERGE_ALL' } });
  Sheets.Spreadsheets.batchUpdate({ requests: requests }, ssId);

  sheet.getRange(PROGRAM_BROKER_NAME_ROW, PROGRAM_BROKER_START_COL, 1, nameRow.length).setValues([nameRow]);
  sheet.getRange(PROGRAM_HEADER_ROW, PROGRAM_BROKER_START_COL, 1, headRow.length).setValues([headRow]);

  const blank = v => (v === null || v === undefined) ? '' : v;
  const values = [], dateFormats = [];
  list.forEach(rec => {
    const amount = rec.amount || ['', ''];
    const buyRow = [new Date(rec.key.replace(/-/g, '/')), '買付け'].concat(rec.buy.map(blank), [blank(amount[0])]);
    const sellRow = ['', '売付け'].concat(rec.sell.map(blank), [blank(amount[1])]);
    brokerNames.forEach(name => {
      const v = (rec.brokers[name] || ['', '', '', '']).map(blank);
      buyRow.push(v[0], v[1]);
      sellRow.push(v[2], v[3]);
    });
    values.push(buyRow, sellRow);
    dateFormats.push([rec.kind === 'week' ? PROGRAM_WEEK_FORMAT : PROGRAM_DAY_FORMAT], [PROGRAM_DAY_FORMAT]);
  });
  sheet.getRange(PROGRAM_FIRST_ROW, 1, rows, width).setValues(values);
  sheet.getRange(PROGRAM_FIRST_ROW, 1, rows, 1).setNumberFormats(dateFormats);
  // 前のデータより行が減ったときは、下に残った分を消す(ふつうは増える一方)
  const oldLast = sheet.getLastRow();
  if (oldLast > lastRow) sheet.getRange(lastRow + 1, 1, oldLast - lastRow, width).clearContent();

  programSpreadChangeColors_(sheet, rows);
}

/**
 * 前日比(F列)の条件付き書式(増えたら赤・減ったら青など)を、データのある行の全体に広げる。
 * 書式をまねたときに同じルールが増えることがあるので、同じ条件のルールは1つにまとめる。
 */
function programSpreadChangeColors_(sheet, rows) {
  const fRange = sheet.getRange(PROGRAM_FIRST_ROW, 6, rows, 1);
  const seen = {};
  const rules = [];
  sheet.getConditionalFormatRules().forEach(rule => {
    const rs = rule.getRanges();
    const onF = rs.length > 0 && rs.every(r => r.getColumn() === 6 && r.getNumColumns() === 1 && r.getRow() >= PROGRAM_FIRST_ROW);
    if (!onF) { rules.push(rule); return; }
    const cond = rule.getBooleanCondition();
    let sig = '';
    try {
      sig = cond ? JSON.stringify([String(cond.getCriteriaType()), cond.getCriteriaValues(), String(cond.getFontColorObject() && cond.getFontColorObject().asRgbColor().asHexString())]) : String(Math.random());
    } catch (e) { sig = JSON.stringify([String(cond.getCriteriaType()), cond.getCriteriaValues()]); }
    if (seen[sig]) return;
    seen[sig] = true;
    rules.push(rule.copy().setRanges([fRange]).build());
  });
  sheet.setConditionalFormatRules(rules);
}

// ---- 更新 ----

const PROGRAM_FETCH_BATCH = 10; // 一度に取りに行くファイルの数

/** 時間主導トリガー(毎日)から呼ぶ。週ごとは最新のページ(直近5週)だけを見る */
function updateProgramTrading() {
  return programUpdate_(false);
}

/** 「データ更新 → プログラム売買・裁定取引を更新(過去分も)」。週ごとはバックナンバー(2023年から)も見る */
function runUpdateProgramTradingFromMenu() {
  runLibFromMenu_(() => programUpdate_(true), 'プログラム売買を更新しました', 'プログラム売買の更新に失敗しました');
}

/**
 * プログラムシートを更新する。
 * 日ごと: ページにあって、まだシートに無い日を足す。週ごと: 週間売買金額がまだ入っていない週を足す
 * (その週の最後の営業日の日ごとの行があればそこへ金額だけ入れ、無ければ週ごとの行を作る)。
 * 時間切れ(RUN_TIME_LIMIT_MS)になったら、そこまでを書いて止まる(もう一度実行すると続きから)。
 */
function programUpdate_(withArchives) {
  const startMs = Date.now();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) return 'プログラム売買: ほかの更新が実行中のため、あとでもう一度実行してください';
  try {
    const sheet = getMainSpreadsheet().getSheetByName(PROGRAM_SHEET_NAME);
    if (!sheet) return 'プログラム売買: 「' + PROGRAM_SHEET_NAME + '」シートが見つかりません';
    const loaded = programReadSheet_(sheet);
    const records = loaded.records;
    const errors = [];
    let days = 0, weeks = 0, timeUp = false;

    const dailyTodo = programDailyFiles_().filter(f => !(records[f.key] && records[f.key].kind === 'day'));
    programFetchGrids_(dailyTodo).forEach(g => {
      if (!g.grid) { errors.push(g.file.key + '(日ごと): ' + g.error); return; }
      try {
        const rec = programParseDaily_(g.grid, g.file.key);
        if (records[g.file.key]) rec.amount = records[g.file.key].amount; // 週ごとの行だった日は、金額を引き継ぐ
        records[g.file.key] = rec;
        days++;
      } catch (e) { errors.push(g.file.key + '(日ごと): ' + e.message); }
    });

    const weeklyTodo = programWeeklyFiles_(withArchives).filter(f => !(records[f.key] && records[f.key].amount));
    for (let i = 0; i < weeklyTodo.length; i += PROGRAM_FETCH_BATCH) {
      if (Date.now() - startMs > RUN_TIME_LIMIT_MS) { timeUp = true; break; }
      programFetchGrids_(weeklyTodo.slice(i, i + PROGRAM_FETCH_BATCH)).forEach(g => {
        if (!g.grid) { errors.push(g.file.key + '(週ごと): ' + g.error); return; }
        try {
          const rec = programParseWeekly_(g.grid, g.file.key);
          if (records[g.file.key] && records[g.file.key].kind === 'day') records[g.file.key].amount = rec.amount;
          else records[g.file.key] = rec;
          weeks++;
        } catch (e) { errors.push(g.file.key + '(週ごと): ' + e.message); }
      });
    }

    if (days > 0 || weeks > 0) programWriteSheet_(sheet, records, loaded.brokers);
    let message = 'プログラム売買: 日ごと' + days + '日分・週ごと' + weeks + '週分を入れました';
    if (days === 0 && weeks === 0) message = 'プログラム売買: 新しいデータはありませんでした';
    if (timeUp) message += '(時間切れのため途中まで。もう一度実行すると続きから進みます)';
    if (errors.length > 0) message += '\n取れなかったもの: ' + errors.slice(0, 10).join(' / ') + (errors.length > 10 ? ' ほか' + (errors.length - 10) + '件' : '');
    Logger.log(message);
    return message;
  } finally {
    lock.releaseLock();
  }
}
