// ==== JPX 銘柄別信用取引残高(mtall.pdf)の解析 ====
// 109ページ程度の大きなPDFのため、Googleドキュメント変換(ページ数制限・時間)に頼らず、
// PDFの中身(圧縮されたテキスト描画命令)を直接展開して文字を取り出す。
// PDFは「描画順」に文字が並ぶため、1銘柄が
//   「コード ISIN 株数 Shs. 売残 前日比 上場比 買残 前日比 上場比 …」
// の順で連続する(実ファイルで確認済み)。この並びから売残・買残を読む。

// ---- Deflate(Flate)展開(zlib形式のストリームを展開する最小実装) ----
function mtInflate_(src) {
  // zlibヘッダ(2バイト)を飛ばす。ヘッダがない(生deflate)場合はそのまま
  let pos = 0;
  if (src.length > 2 && (src[0] & 0x0f) === 8 && ((src[0] << 8) | src[1]) % 31 === 0) pos = 2;

  let out = new Uint8Array(Math.max(src.length * 4, 1024));
  let outLen = 0;
  function ensure(n) {
    if (outLen + n <= out.length) return;
    let size = out.length * 2;
    while (size < outLen + n) size *= 2;
    const b = new Uint8Array(size);
    b.set(out.subarray(0, outLen));
    out = b;
  }

  let bitBuf = 0, bitCnt = 0;
  function bits(n) {
    while (bitCnt < n) {
      if (pos >= src.length) throw new Error('inflate: unexpected end');
      bitBuf |= src[pos++] << bitCnt;
      bitCnt += 8;
    }
    const v = bitBuf & ((1 << n) - 1);
    bitBuf >>>= n;
    bitCnt -= n;
    return v;
  }

  // Huffman表: counts[len], symbols[]
  function buildTree(lengths, n) {
    const counts = new Uint16Array(16), offs = new Uint16Array(16), symbols = new Uint16Array(n);
    for (let i = 0; i < n; i++) counts[lengths[i]]++;
    counts[0] = 0;
    for (let i = 1; i < 16; i++) offs[i] = offs[i - 1] + counts[i - 1];
    for (let i = 0; i < n; i++) if (lengths[i]) symbols[offs[lengths[i]]++] = i;
    return { counts: counts, symbols: symbols };
  }
  function decodeSym(t) {
    let code = 0, first = 0, index = 0;
    for (let len = 1; len < 16; len++) {
      code |= bits(1);
      const count = t.counts[len];
      if (code - count < first) return t.symbols[index + (code - first)];
      index += count;
      first += count;
      first <<= 1;
      code <<= 1;
    }
    throw new Error('inflate: bad code');
  }

  const LBASE = [3,4,5,6,7,8,9,10,11,13,15,17,19,23,27,31,35,43,51,59,67,83,99,115,131,163,195,227,258];
  const LEXT = [0,0,0,0,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,5,5,5,5,0];
  const DBASE = [1,2,3,4,5,7,9,13,17,25,33,49,65,97,129,193,257,385,513,769,1025,1537,2049,3073,4097,6145,8193,12289,16385,24577];
  const DEXT = [0,0,0,0,1,1,2,2,3,3,4,4,5,5,6,6,7,7,8,8,9,9,10,10,11,11,12,12,13,13];
  const CLORDER = [16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15];

  let fixedL = null, fixedD = null;
  function fixedTrees() {
    if (fixedL) return;
    const l = new Uint8Array(288);
    for (let i = 0; i < 144; i++) l[i] = 8;
    for (let i = 144; i < 256; i++) l[i] = 9;
    for (let i = 256; i < 280; i++) l[i] = 7;
    for (let i = 280; i < 288; i++) l[i] = 8;
    fixedL = buildTree(l, 288);
    const d = new Uint8Array(30).fill(5);
    fixedD = buildTree(d, 30);
  }

  function inflateBlock(lt, dt) {
    for (;;) {
      const sym = decodeSym(lt);
      if (sym < 256) { ensure(1); out[outLen++] = sym; continue; }
      if (sym === 256) return;
      const li = sym - 257;
      const len = LBASE[li] + (LEXT[li] ? bits(LEXT[li]) : 0);
      const di = decodeSym(dt);
      const dist = DBASE[di] + (DEXT[di] ? bits(DEXT[di]) : 0);
      ensure(len);
      let from = outLen - dist;
      for (let k = 0; k < len; k++) out[outLen++] = out[from++];
    }
  }

  let last = 0;
  do {
    last = bits(1);
    const type = bits(2);
    if (type === 0) {
      bitBuf = 0; bitCnt = 0; // バイト境界にそろえる
      const len = src[pos] | (src[pos + 1] << 8);
      pos += 4;
      ensure(len);
      out.set(src.subarray(pos, pos + len), outLen);
      outLen += len;
      pos += len;
    } else if (type === 1) {
      fixedTrees();
      inflateBlock(fixedL, fixedD);
    } else if (type === 2) {
      const hlit = bits(5) + 257, hdist = bits(5) + 1, hclen = bits(4) + 4;
      const cl = new Uint8Array(19);
      for (let i = 0; i < hclen; i++) cl[CLORDER[i]] = bits(3);
      const clt = buildTree(cl, 19);
      const lens = new Uint8Array(hlit + hdist);
      for (let i = 0; i < hlit + hdist;) {
        const sym = decodeSym(clt);
        if (sym < 16) { lens[i++] = sym; continue; }
        let rep = 0, val = 0;
        if (sym === 16) { val = lens[i - 1]; rep = 3 + bits(2); }
        else if (sym === 17) rep = 3 + bits(3);
        else rep = 11 + bits(7);
        while (rep--) lens[i++] = val;
      }
      inflateBlock(buildTree(lens.subarray(0, hlit), hlit), buildTree(lens.subarray(hlit), hdist));
    } else {
      throw new Error('inflate: bad block type');
    }
  } while (!last);

  return out.subarray(0, outLen);
}

// ---- PDFのオブジェクト・ストリームを読む ----
function mtBytesToLatin1_(bytes) {
  const parts = [];
  for (let i = 0; i < bytes.length; i += 8192) {
    parts.push(String.fromCharCode.apply(null, bytes.subarray(i, Math.min(i + 8192, bytes.length))));
  }
  return parts.join('');
}

/** PDF全体から {オブジェクト番号: {dict:文字列, stream:Uint8Array(展開済み)|null}} を作る */
function mtReadPdfObjects_(bytes) {
  const s = mtBytesToLatin1_(bytes);
  const objs = {};
  const re = /(\d+)\s+0\s+obj\b/g;
  let m;
  while ((m = re.exec(s)) !== null) {
    const num = parseInt(m[1], 10);
    const bodyStart = m.index + m[0].length;
    const endObj = s.indexOf('endobj', bodyStart);
    const streamKw = s.indexOf('stream', bodyStart);
    if (streamKw !== -1 && (endObj === -1 || streamKw < endObj)) {
      const dict = s.substring(bodyStart, streamKw);
      let dataStart = streamKw + 6;
      if (s.charCodeAt(dataStart) === 13) dataStart++;
      if (s.charCodeAt(dataStart) === 10) dataStart++;
      let dataEnd;
      const lenM = dict.match(/\/Length\s+(\d+)(?!\d)(?!\s+\d+\s+R)/);
      if (lenM) dataEnd = dataStart + parseInt(lenM[1], 10);
      else dataEnd = s.indexOf('endstream', dataStart);
      let data = bytes.subarray(dataStart, dataEnd);
      if (/\/FlateDecode/.test(dict)) {
        try { data = mtInflate_(data); } catch (e) { data = null; }
      }
      objs[num] = { dict: dict, stream: data };
      re.lastIndex = dataEnd;
    } else {
      objs[num] = { dict: s.substring(bodyStart, endObj === -1 ? s.length : endObj), stream: null };
      if (endObj !== -1) re.lastIndex = endObj;
    }
  }

  // オブジェクトストリーム(/Type/ObjStm)の中に格納されたオブジェクトも展開する
  Object.keys(objs).forEach(k => {
    const o = objs[k];
    if (!o.stream || !/\/Type\s*\/ObjStm/.test(o.dict)) return;
    const n = parseInt((o.dict.match(/\/N\s+(\d+)/) || [])[1], 10);
    const first = parseInt((o.dict.match(/\/First\s+(\d+)/) || [])[1], 10);
    const text = mtBytesToLatin1_(o.stream);
    const header = text.substring(0, first).trim().split(/\s+/).map(Number);
    for (let i = 0; i < n; i++) {
      const objNum = header[i * 2], off = header[i * 2 + 1];
      const nextOff = i + 1 < n ? header[(i + 1) * 2 + 1] : text.length - first;
      if (!objs[objNum]) objs[objNum] = { dict: text.substring(first + off, first + nextOff), stream: null };
    }
  });
  return objs;
}

/** 辞書文字列から /Key の値(参照なら参照先の辞書文字列、直接の辞書ならその中身)を取り出す */
function mtDictValue_(objs, dict, key) {
  const idx = dict.search(new RegExp('/' + key + '(?![A-Za-z0-9])'));
  if (idx === -1) return null;
  let rest = dict.substring(idx + key.length + 1).replace(/^\s+/, '');
  const ref = rest.match(/^(\d+)\s+0\s+R/);
  if (ref) {
    const o = objs[parseInt(ref[1], 10)];
    return o ? o.dict : null;
  }
  if (rest.startsWith('<<')) {
    let depth = 0;
    for (let i = 0; i < rest.length - 1; i++) {
      if (rest[i] === '<' && rest[i + 1] === '<') { depth++; i++; }
      else if (rest[i] === '>' && rest[i + 1] === '>') { depth--; i++; if (depth === 0) return rest.substring(2, i - 1); }
    }
  }
  return rest;
}

/** ToUnicode CMapを読み、{コード(数値): 文字列} と コード長(バイト) を返す */
function mtParseCMap_(text) {
  const map = {};
  let codeLen = 1;
  const csr = text.match(/begincodespacerange\s*<([0-9A-Fa-f]+)>/);
  if (csr) codeLen = csr[1].length / 2;
  const hexToStr = h => {
    let r = '';
    for (let i = 0; i + 4 <= h.length; i += 4) r += String.fromCharCode(parseInt(h.substr(i, 4), 16));
    if (h.length === 2) r = String.fromCharCode(parseInt(h, 16));
    return r;
  };
  let m;
  const charBlocks = /beginbfchar([\s\S]*?)endbfchar/g;
  while ((m = charBlocks.exec(text)) !== null) {
    const pairRe = /<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/g;
    let p;
    while ((p = pairRe.exec(m[1])) !== null) map[parseInt(p[1], 16)] = hexToStr(p[2]);
  }
  const rangeBlocks = /beginbfrange([\s\S]*?)endbfrange/g;
  while ((m = rangeBlocks.exec(text)) !== null) {
    const rRe = /<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*(<([0-9A-Fa-f]+)>|\[([^\]]*)\])/g;
    let r;
    while ((r = rRe.exec(m[1])) !== null) {
      const lo = parseInt(r[1], 16), hi = parseInt(r[2], 16);
      if (r[4] !== undefined) {
        const base = parseInt(r[4], 16);
        for (let c = lo; c <= hi; c++) map[c] = String.fromCharCode(base + (c - lo));
      } else {
        const items = (r[5].match(/<([0-9A-Fa-f]+)>/g) || []).map(x => hexToStr(x.slice(1, -1)));
        for (let c = lo; c <= hi && c - lo < items.length; c++) map[c] = items[c - lo];
      }
    }
  }
  return { map: map, codeLen: codeLen };
}

/** 1ページの描画命令からテキストを描画順に取り出す(文字列描画命令ごとに空白で区切る) */
function mtExtractPageText_(content, fonts) {
  const s = mtBytesToLatin1_(content);
  const out = [];
  let font = null;
  let moved = true; // 直前の文字描画から位置移動(Td/Tm等)があったか。移動があれば空白で区切る
  let i = 0;
  function emit(str) {
    if (moved) out.push(' ');
    out.push(str);
    moved = false;
  }
  const n = s.length;
  const operands = [];

  function decode(bytesStr) {
    if (!font) return '';
    let r = '';
    const L = font.codeLen;
    for (let k = 0; k + L <= bytesStr.length; k += L) {
      let code = 0;
      for (let b = 0; b < L; b++) code = (code << 8) | bytesStr.charCodeAt(k + b);
      const u = font.map[code];
      r += (u !== undefined) ? u : (L === 1 ? String.fromCharCode(code) : '');
    }
    return r;
  }

  while (i < n) {
    const c = s[i];
    if (c === ' ' || c === '\n' || c === '\r' || c === '\t' || c === '\f' || c === '\0') { i++; continue; }
    if (c === '%') { while (i < n && s[i] !== '\n' && s[i] !== '\r') i++; continue; }
    if (c === '(') {
      let depth = 1, j = i + 1, str = '';
      while (j < n && depth > 0) {
        const ch = s[j];
        if (ch === '\\') {
          const nx = s[j + 1];
          if (nx >= '0' && nx <= '7') {
            let oct = '', k = j + 1;
            while (k < n && oct.length < 3 && s[k] >= '0' && s[k] <= '7') oct += s[k++];
            str += String.fromCharCode(parseInt(oct, 8));
            j = k;
            continue;
          }
          const esc = { n: '\n', r: '\r', t: '\t', b: '\b', f: '\f' }[nx];
          if (nx === '\r' || nx === '\n') { j += 2; if (nx === '\r' && s[j] === '\n') j++; continue; }
          str += esc !== undefined ? esc : nx;
          j += 2;
          continue;
        }
        if (ch === '(') depth++;
        else if (ch === ')') { depth--; if (depth === 0) break; }
        str += ch;
        j++;
      }
      operands.push({ t: 's', v: str });
      i = j + 1;
      continue;
    }
    if (c === '<' && s[i + 1] !== '<') {
      const j = s.indexOf('>', i);
      let hex = s.substring(i + 1, j).replace(/\s+/g, '');
      if (hex.length % 2) hex += '0';
      let str = '';
      for (let k = 0; k < hex.length; k += 2) str += String.fromCharCode(parseInt(hex.substr(k, 2), 16));
      operands.push({ t: 's', v: str });
      i = j + 1;
      continue;
    }
    if (c === '<' || c === '>') { i += 2; continue; } // << >>(インライン辞書)は無視
    if (c === '[') { operands.push({ t: '[' }); i++; continue; }
    if (c === ']') {
      const arr = [];
      while (operands.length && operands[operands.length - 1].t !== '[') arr.unshift(operands.pop());
      operands.pop();
      operands.push({ t: 'a', v: arr });
      i++;
      continue;
    }
    if (c === '/') {
      let j = i + 1;
      while (j < n && !/[\s\/\[\]<>()%{}]/.test(s[j])) j++;
      operands.push({ t: 'n', v: s.substring(i + 1, j) });
      i = j;
      continue;
    }
    // 数値または演算子
    let j = i;
    while (j < n && !/[\s\/\[\]<>()%{}]/.test(s[j])) j++;
    const tok = s.substring(i, j);
    i = j;
    if (/^[+\-]?(\d+\.?\d*|\.\d+)$/.test(tok)) { operands.push({ t: 'num', v: parseFloat(tok) }); continue; }

    // 演算子
    if (tok === 'Tf') {
      const nameOp = operands.filter(o => o.t === 'n').pop();
      font = nameOp ? (fonts[nameOp.v] || null) : null;
    } else if (tok === 'Td' || tok === 'TD' || tok === 'Tm' || tok === 'T*') {
      moved = true;
    } else if (tok === 'Tj' || tok === "'" || tok === '"') {
      if (tok !== 'Tj') moved = true;
      const sOp = operands.filter(o => o.t === 's').pop();
      if (sOp) emit(decode(sOp.v));
    } else if (tok === 'TJ') {
      const aOp = operands.filter(o => o.t === 'a').pop();
      if (aOp) {
        let piece = '';
        aOp.v.forEach(e => {
          if (e.t === 's') piece += decode(e.v);
          else if (e.t === 'num' && e.v < -300) piece += ' '; // 大きな字間調整は空白とみなす
        });
        emit(piece);
      }
    } else if (tok === 'ET') {
      out.push('\n');
      moved = true;
    } else if (tok === 'BI') {
      const ei = s.indexOf('EI', i);
      i = ei === -1 ? n : ei + 2;
    }
    operands.length = 0;
  }
  return out.join('').replace(/[ \t]*\n[ \t]*/g, '\n');
}

/** ページを文書順に並べる(/Root → /Pages の /Kids をたどる) */
function mtCollectPages_(objs) {
  const pages = [];
  const catalogKey = Object.keys(objs).find(k => /\/Type\s*\/Catalog/.test(objs[k].dict));
  if (!catalogKey) return pages;
  const visit = (dict, depth) => {
    if (!dict || depth > 20) return;
    if (/\/Type\s*\/Page(?![s])/.test(dict)) { pages.push(dict); return; }
    const kidsM = dict.match(/\/Kids\s*\[([^\]]*)\]/);
    if (!kidsM) return;
    const refRe = /(\d+)\s+0\s+R/g;
    let r;
    while ((r = refRe.exec(kidsM[1])) !== null) {
      const o = objs[parseInt(r[1], 10)];
      if (o) visit(o.dict, depth + 1);
    }
  };
  visit(mtDictValue_(objs, objs[catalogKey].dict, 'Pages'), 0);
  return pages;
}

/** mtall.pdfのバイト列から全ページのテキスト(描画順)を返す */
function mtExtractPdfText_(bytes) {
  const objs = mtReadPdfObjects_(bytes);
  const fontCache = {};
  const texts = [];
  mtCollectPages_(objs).forEach(pageDict => {
    const res = mtDictValue_(objs, pageDict, 'Resources') || '';
    const fontDict = mtDictValue_(objs, res, 'Font') || '';
    const fonts = {};
    const fRe = /\/([^\s\/<>\[\]()]+)\s+(\d+)\s+0\s+R/g;
    let f;
    while ((f = fRe.exec(fontDict)) !== null) {
      const fnum = parseInt(f[2], 10);
      if (!fontCache[fnum]) {
        const fo = objs[fnum];
        let cmap = { map: {}, codeLen: 1 };
        if (fo) {
          const tu = fo.dict.match(/\/ToUnicode\s+(\d+)\s+0\s+R/);
          const cmapObj = tu ? objs[parseInt(tu[1], 10)] : null;
          if (cmapObj && cmapObj.stream) cmap = mtParseCMap_(mtBytesToLatin1_(cmapObj.stream));
          else if (/Identity-H/.test(fo.dict)) cmap.codeLen = 2;
        }
        fontCache[fnum] = cmap;
      }
      fonts[f[1]] = fontCache[fnum];
    }
    const contentsM = pageDict.match(/\/Contents\s*(\[[^\]]*\]|\d+\s+0\s+R)/);
    if (!contentsM) return;
    const refs = contentsM[1].match(/\d+\s+0\s+R/g) || [];
    const parts = [];
    refs.forEach(ref => {
      const o = objs[parseInt(ref, 10)];
      if (o && o.stream) parts.push(mtBytesToLatin1_(o.stream));
    });
    const joined = parts.join('\n');
    const content = new Uint8Array(joined.length);
    for (let k = 0; k < joined.length; k++) content[k] = joined.charCodeAt(k);
    texts.push(mtExtractPageText_(content, fonts));
  });
  return texts.join('\n');
}

// ---- テキストから売残・買残を取り出す ----
function mtNum_(s) {
  if (s === undefined || s === null) return null;
  const neg = /▲/.test(s);
  const v = parseInt(String(s).replace(/[▲,\s]/g, ''), 10);
  return isNaN(v) ? null : (neg ? -v : v);
}

// 数値1個(3桁区切りとして正しい形のみ。▲付きもあり)
const MT_NUM_ = '(▲\\s*\\d{1,3}(?:,\\d{3})*|\\d{1,3}(?:,\\d{3})*)';

const MT_DIFF_ = '(?:▲\\s*\\d{1,3}(?:,\\d{3})*|\\d{1,3}(?:,\\d{3})*)';

const MT_RATIO_ = '(?:[\\d.]+%|\\*|-)';

// 「株数 Shs.」の後ろ: 売残 前日比 上場比 買残 前日比 上場比 … の並びが崩れていないことまで確認する
const MT_AFTER_SHS_ = '株数\\s*Shs\\.\\s+' + MT_NUM_ + '\\s+' + MT_DIFF_ + '\\s+' + MT_RATIO_ + '\\s+' + MT_NUM_ + '\\s+' + MT_DIFF_ + '\\s+' + MT_RATIO_ + '(?=\\s|$)';

/**
 * PDF上で1つの数値が2回に分けて描画され、間に空白が入ってしまうことがある(実ファイルで確認)。
 * 「145 ,500」「37,718,7 09」のように、単独では数値として成り立たない形だけを機械的につなぎ直す。
 * (「2 9,700」のように両方とも数値として成り立つものは区別できないため直さない。
 *  必要な列でそれが起きた場合は並びのチェックで不一致になり、取得失敗として扱われる)
 */
function mtRepairSplitNumbers_(text) {
  let t = text;
  for (let k = 0; k < 3; k++) {
    t = t.replace(/(\d) ,(\d{3})(?!\d)/g, '$1,$2');
    t = t.replace(/(\d,\d{1,2}) (\d{1,2})(?=[\s,]|$)/g, (all, a, b) => {
      const merged = a + b;
      return /,\d{3}$/.test(merged) ? merged : all;
    });
  }
  return t;
}

/**
 * 全ページテキストから {銘柄コード(4桁 or 285Aなど): {sell, buy}} と、
 * 総合計の市場別小計 {プライム:{sell,buy,count}, スタンダード:…, グロース:…} を取り出す。
 */
function mtParseMarginText_(text) {
  const t = mtRepairSplitNumbers_(text.replace(/　/g, ' '));
  const stocks = {};
  const stockRe = new RegExp('(?:^|\\s)([0-9][0-9A-Z]{3})0\\s+[A-Z]{2}[0-9A-Z]{10}(?:(?!Val\\.|株数)[^\\n]){0,40}?' + MT_AFTER_SHS_, 'g');
  let m;
  while ((m = stockRe.exec(t)) !== null) {
    if (!stocks[m[1]]) stocks[m[1]] = { sell: mtNum_(m[2]), buy: mtNum_(m[3]) };
  }

  const markets = {};
  const totalIdx = t.search(/総\s*合\s*計/);
  if (totalIdx !== -1) {
    const area = t.substring(totalIdx);
    [['プライム', 'プ\\s*ラ\\s*イ\\s*ム'], ['スタンダード', 'ス\\s*タ\\s*ン\\s*ダ\\s*ー\\s*ド'], ['グロース', 'グ\\s*ロ\\s*ー\\s*ス']].forEach(pair => {
      const re = new RegExp(pair[1] + '\\s*小\\s*計\\s+([\\d,]+)\\s*銘\\s*柄\\s+' + MT_AFTER_SHS_);
      const mm = area.match(re);
      if (mm) markets[pair[0]] = { count: mtNum_(mm[1]), sell: mtNum_(mm[2]), buy: mtNum_(mm[3]) };
    });
  }
  return { stocks: stocks, markets: markets };
}

// ==== 信用残(日次)更新 ====
// データ元: JPX「銘柄別信用取引残高」(https://www.jpx.co.jp/markets/statistics-equities/margin/01.html)
//   ・申込日ごとのPDF(YYYYMMDD_mtall.pdf)。2026年10月分から公表開始のため、それより前のデータはない
//   ・申込日の翌営業日に公表される。一覧ページには直近5日分程度しか載らないので、こまめに取得すること
// 書き込み先: 各月シートの39行目〜69行目(1日〜31日)。終値と同じく A列=日付・B列=曜日・休場マーク
//   市場ブロック(C:E・F:H・I:K)に [売残(百万株), 買残(百万株), 信用倍率]、各銘柄(3列)に [売残(千株), 買残(千株), 信用倍率]
const MARGIN_DATA_START_DATE = new Date(2026, 9, 2); // JPX日次信用残の公表開始(2026/10/2申込分から。月は0始まりなので9=10月)

const MARGIN_LIST_PAGE_URL = 'https://www.jpx.co.jp/markets/statistics-equities/margin/01.html';

const MARGIN_STOCK_UNIT_DIVISOR = 1000;     // 個別銘柄: 株 → 千株(整数)

const MARGIN_MARKET_UNIT_DIVISOR = 1000000; // プライム/スタンダード/グロース: 株 → 百万株

const MARGIN_SMALL_FORMAT = '#,##0.000'; // 市場の値が百万株にして1桁(10未満)になる場合は小数点第3位まで表示

const MARGIN_LARGE_FORMAT = '#,##0';

const MARGIN_AUTO_LOOKBACK_DAYS = 10; // 自動更新(トリガー)で遡って未入力を埋める日数

const MARGIN_WEEKDAY_NAMES_ = ['日曜日', '月曜日', '火曜日', '水曜日', '木曜日', '金曜日', '土曜日'];

function marginPdfUrl_(ymd) {
  return 'https://www.jpx.co.jp/markets/statistics-equities/margin/tvdivq0000001rnl-att/' + ymd + '_mtall.pdf';
}

/** 一覧ページから該当日のPDFリンクを探す(添付フォルダ名が変わった場合のフォールバック) */
function findMarginPdfUrlFromList_(ymd) {
  const response = UrlFetchApp.fetch(MARGIN_LIST_PAGE_URL, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) return null;
  const m = response.getContentText().match(new RegExp('href="([^"]*' + ymd + '_mtall\\.pdf)"'));
  if (!m) return null;
  return m[1].startsWith('http') ? m[1] : 'https://www.jpx.co.jp' + m[1];
}

/**
 * 指定日(申込日)の銘柄別信用取引残高を取得する。未公表・休場日はnull。
 * 戻り値: { stocks: {コード: [売残, 買残]}, markets: {プライム: [売残, 買残], ...} }(単位: 株)
 * 解析結果は6時間キャッシュする(新規銘柄追加時の補完などで同じ日を何度も読む場合の高速化)。
 */
function fetchJpxMarginForDate_(dateStr) {
  const ymd = dateStr.replace(/-/g, '');
  const cache = CacheService.getScriptCache();
  const cacheKey = 'mtall_' + ymd;
  try {
    const cached = cache.get(cacheKey);
    if (cached) {
      const blob = Utilities.newBlob(Utilities.base64Decode(cached), 'application/x-gzip');
      return JSON.parse(Utilities.ungzip(blob).getDataAsString());
    }
  } catch (e) {}

  let bytes = fetchPdfBytes_(marginPdfUrl_(ymd));
  if (!bytes) {
    const alt = findMarginPdfUrlFromList_(ymd);
    if (alt) bytes = fetchPdfBytes_(alt);
  }
  if (!bytes) return null;

  const parsed = mtParseMarginText_(mtExtractPdfText_(bytes));
  if (Object.keys(parsed.stocks).length === 0) return null;

  const result = { stocks: {}, markets: {} };
  Object.keys(parsed.stocks).forEach(code => { result.stocks[code] = [parsed.stocks[code].sell, parsed.stocks[code].buy]; });
  Object.keys(parsed.markets).forEach(name => { result.markets[name] = [parsed.markets[name].sell, parsed.markets[name].buy]; });

  try {
    const zipped = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(JSON.stringify(result))).getBytes());
    if (zipped.length < 95000) cache.put(cacheKey, zipped, 21600);
  } catch (e) {}
  return result;
}

/**
 * 株数 → シートに書く値。
 *   市場(プライム/スタンダード/グロース): 百万株。10未満(1桁)になる値は小数点第3位まで残し、それ以外は整数
 *     例) 256,314,200株 → 256 / 3,265,500株 → 3.266
 *   個別銘柄: 千株(整数)。例) 846,100株 → 846
 */
function marginUnitValue_(shares, isMarket) {
  if (!isMarket) return Math.round(shares / MARGIN_STOCK_UNIT_DIVISOR);
  const v = shares / MARGIN_MARKET_UNIT_DIVISOR;
  return Math.abs(v) < 10 ? Math.round(v * 1000) / 1000 : Math.round(v);
}

function marginNumberFormat_(value, isMarket) {
  return (isMarket && Math.abs(value) < 10) ? MARGIN_SMALL_FORMAT : MARGIN_LARGE_FORMAT;
}

/** [売残, 買残](株) → シートに書く [売残, 買残, 信用倍率](単位はmarginUnitValue_のとおり) */
function marginCellValues_(pair, isMarket) {
  const sell = pair[0], buy = pair[1];
  const ratio = (sell > 0) ? Math.round((buy / sell) * 100) / 100 : '-';
  return [marginUnitValue_(sell, isMarket), marginUnitValue_(buy, isMarket), ratio];
}

function isMarginBlank_(v) { return v === '' || v === null || v === undefined; }

/**
 * 指定日の信用残を当月(その日の月)シートの信用残行に書き込む。
 * onlyCode を指定すると、その銘柄の列だけを埋める(新規銘柄追加時の補完用)。
 * 戻り値: { message, status: 'done'|'holiday'|'skip'|'unpublished'|'nosheet' }
 */
function updateMarginForSpecificDate(targetDate, onlyCode) {
  const dateStr = Utilities.formatDate(targetDate, 'Asia/Tokyo', 'yyyy-MM-dd');
  if (targetDate < MARGIN_DATA_START_DATE) return { status: 'skip', message: dateStr + ': 2026/10/2より前は日次の信用残データがありません' };
  const sheet = getTargetSheet(targetDate);
  if (!sheet) return { status: 'nosheet', message: dateStr + ': 対応するシートが見つかりません' };

  const row = MARGIN_FIRST_ROW + (targetDate.getDate() - 1);
  const lastCol = sheet.getLastColumn();
  const codeRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const rowValues = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
  if (isMarginBlank_(rowValues[0])) rowValues[0] = targetDate.getDate() + '日';
  if (isMarginBlank_(rowValues[1])) rowValues[1] = MARGIN_WEEKDAY_NAMES_[targetDate.getDay()];

  const blocks = [];
  if (!onlyCode) MARKET_VALUE_ITEMS.forEach(item => blocks.push({ col: item.col, width: MARKET_VALUE_COL_WIDTH, market: item.name }));
  for (let col = STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) {
    const code = codeRow[col - 1];
    if (!code) continue;
    if (onlyCode && String(code) !== String(onlyCode)) continue;
    blocks.push({ col: col, width: STOCK_COL_WIDTH, code: String(code) });
  }

  const dow = targetDate.getDay();
  if ((dow === 0 || dow === 6) || isJapaneseHoliday(targetDate)) {
    blocks.forEach(b => {
      for (let k = 0; k < b.width; k++) rowValues[b.col - 1 + k] = HOLIDAY_LABEL;
    });
    if (blocks.length > 0) {
      sheet.getRangeList(blocks.map(b => columnToLetter(b.col) + row + ':' + columnToLetter(b.col + b.width - 1) + row)).setBackground(HOLIDAY_COLOR);
    }
    sheet.getRange(row, 1, 1, lastCol).setValues([rowValues]);
    return { status: 'holiday', message: dateStr + ': 休場日のため休マークを入力しました' };
  }

  const data = fetchJpxMarginForDate_(dateStr);
  if (!data) {
    sheet.getRange(row, 1, 1, 2).setValues([[rowValues[0], rowValues[1]]]);
    return { status: 'unpublished', message: dateStr + ': 未公表です(申込日の翌営業日に公表されます)' };
  }

  const notFound = [];
  let written = 0;
  const formats = sheet.getRange(row, 1, 1, lastCol).getNumberFormats()[0];
  blocks.forEach(b => {
    const pair = b.market ? data.markets[b.market] : data.stocks[b.code];
    if (!pair) {
      notFound.push(b.market || b.code);
      // 一覧にない銘柄(信用取引の対象外など)は「-」を入れて入力済み扱いにする(毎日取り直さないように)。
      // 市場の小計が読めなかった場合は空欄のままにし、次回の更新で取り直す。
      if (b.code) for (let k = 0; k < 3; k++) rowValues[b.col - 1 + k] = '-';
      return;
    }
    const isMarket = !!b.market;
    const vals = marginCellValues_(pair, isMarket);
    for (let k = 0; k < 3; k++) rowValues[b.col - 1 + k] = vals[k];
    // 売残・買残の表示形式(市場は1桁なら小数点第3位まで、個別銘柄は常に整数)
    formats[b.col - 1] = marginNumberFormat_(vals[0], isMarket);
    formats[b.col] = marginNumberFormat_(vals[1], isMarket);
    written++;
  });
  sheet.getRange(row, 1, 1, lastCol).setValues([rowValues]);
  sheet.getRange(row, 1, 1, lastCol).setNumberFormats([formats]);

  let message = dateStr + ': ' + written + '/' + blocks.length + '項目の信用残を更新しました';
  if (notFound.length > 0) message += ' / 一覧になかった項目(信用取引の対象外など): ' + notFound.join(', ');
  return { status: 'done', message: message };
}

/** その日の信用残行が(市場ブロック・全銘柄とも)すでに埋まっているか */
function isMarginRowFilled_(date) {
  const sheet = getTargetSheet(date);
  if (!sheet) return false;
  const row = MARGIN_FIRST_ROW + (date.getDate() - 1);
  const lastCol = sheet.getLastColumn();
  const codeRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const vals = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
  const cols = MARKET_VALUE_ITEMS.map(item => item.col);
  for (let col = STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) if (codeRow[col - 1]) cols.push(col);
  return cols.every(col => !isMarginBlank_(vals[col - 1]));
}

/** ダイアログからの単日・期間指定更新。入力済みの日はスキップし、時間切れ前に安全に止まる */
function updateMarginForDateRangeInput(startDateStr, endDateStr) {
  const rangeStartMs = Date.now();
  const start = parseDateStr_(startDateStr);
  const end = parseDateStr_(endDateStr || startDateStr);
  if (start > end) return '開始日は終了日より前の日付にしてください';

  const messages = [];
  for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
    if (Date.now() - rangeStartMs > RUN_TIME_LIMIT_MS) {
      messages.push('(実行時間の上限のため、' + Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd') + '以降は未処理です。同じ期間でもう一度実行すると続きから処理されます)');
      break;
    }
    const date = new Date(d);
    if (date >= MARGIN_DATA_START_DATE && isMarginRowFilled_(date)) {
      messages.push(Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd') + ': 入力済みのためスキップしました');
      continue;
    }
    messages.push(updateMarginForSpecificDate(date).message);
  }
  const message = messages.join('\n');
  Logger.log(message);
  return message;
}

/**
 * 時間主導トリガー用(毎日実行): 直近の未入力日(前日まで)の信用残を埋める。
 * 一覧ページには直近5日分程度しか載らないため、毎日動かして取りこぼしを防ぐ。
 */
function updateMarginData() {
  const startMs = Date.now();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const messages = [];
  for (let back = MARGIN_AUTO_LOOKBACK_DAYS; back >= 1; back--) {
    if (Date.now() - startMs > RUN_TIME_LIMIT_MS) { messages.push('(実行時間の上限のため途中で停止しました)'); break; }
    const date = new Date(today);
    date.setDate(today.getDate() - back);
    if (date < MARGIN_DATA_START_DATE) continue;
    if (isMarginRowFilled_(date)) continue;
    messages.push(updateMarginForSpecificDate(date).message);
  }
  const message = messages.length > 0 ? messages.join('\n') : '信用残: 未入力の日はありませんでした';
  Logger.log(message);
  return message;
}

/**
 * 新規銘柄追加時用: 当月の1日〜前日までの、その銘柄の信用残を埋める(入力済みの日・休場日はそのまま)。
 * 時間がかかる場合は途中で止まり、残りは毎日の自動更新または日付入力更新で埋まる。
 */
function fillMarginForCode(code) {
  const startMs = Date.now();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const sheet = getTargetSheet(today);
  if (!sheet) return '当月シートが見つかりません';
  const lastCol = sheet.getLastColumn();
  const codeRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  let targetCol = null;
  for (let col = STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) if (String(codeRow[col - 1]) === String(code)) { targetCol = col; break; }
  if (!targetCol) return '銘柄が見つかりません: ' + code;

  let filled = 0;
  const existing = today.getDate() > 1 ? sheet.getRange(MARGIN_FIRST_ROW, targetCol, today.getDate() - 1, 1).getValues() : [];
  for (let d = 1; d < today.getDate(); d++) {
    if (Date.now() - startMs > RUN_TIME_LIMIT_MS) return code + ': 信用残を' + filled + '日分入力しました(時間切れのため途中まで。残りは自動更新で埋まります)';
    const date = new Date(today.getFullYear(), today.getMonth(), d);
    if (date < MARGIN_DATA_START_DATE) continue;
    if (!isMarginBlank_(existing[d - 1][0])) continue;
    const r = updateMarginForSpecificDate(date, code);
    if (r.status === 'done') filled++;
  }
  return code + ': 信用残を' + filled + '日分入力しました';
}

/** 動作確認用: 指定日のPDFを読み、解析にかかった時間と一部の値をログに出す(シートには書かない) */
function debugJpxMarginForDate(dateStr) {
  // dateStr = '2026-10-08'; // ← エディタから直接実行する場合はここのコメントを外して日付を指定する
  if (!dateStr) { Logger.log('dateStrが指定されていません(引数なしで実行していませんか?)'); return; }
  CacheService.getScriptCache().remove('mtall_' + dateStr.replace(/-/g, ''));
  const t0 = Date.now();
  const data = fetchJpxMarginForDate_(dateStr);
  Logger.log('所要時間: ' + (Date.now() - t0) + 'ms');
  if (!data) { Logger.log('取得できませんでした(未公表・休場日・URL変更の可能性)'); return; }
  Logger.log('銘柄数: ' + Object.keys(data.stocks).length);
  Logger.log('市場別: ' + JSON.stringify(data.markets));
  Logger.log('3436: ' + JSON.stringify(data.stocks['3436']) + ' / 285A: ' + JSON.stringify(data.stocks['285A']));
}