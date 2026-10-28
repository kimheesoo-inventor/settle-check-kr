/*
 * 정산 확인기 (settle-check-kr) - core functions. No DOM, no network.
 * Loaded as a classic <script> in the page (window.SettleCheckCore)
 * and with require() in Node tests (module.exports). MIT License, (c) 2026 Heesoo Kim.
 *
 * What it does: compares an order/delivery export and a settlement export (xlsx or csv)
 * and lists orders whose delivery date + waiting days + grace days has passed
 * but whose order number is not in the settlement export.
 */
(function (root, factory) {
  var api = factory();
  if (typeof module === 'object' && module && module.exports) module.exports = api;
  else root.SettleCheckCore = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  // ---------- errors (plain Korean) ----------

  const MESSAGES = {
    EMPTY: '파일에 읽을 내용이 없습니다. 내려받은 엑셀이 비어 있지 않은지 확인해 주세요.',
    OLE: '이 파일은 옛 엑셀(.xls) 형식이거나 암호가 걸린 엑셀입니다. 엑셀에서 연 뒤 "다른 이름으로 저장"에서 "Excel 통합 문서(.xlsx)" 또는 "CSV UTF-8"로 저장해 다시 넣어 주세요. 암호가 걸려 있다면 암호를 푼 뒤 저장해 주세요.',
    ZIP_BAD: '엑셀(.xlsx) 파일을 읽지 못했습니다. 파일이 손상됐거나 엑셀 파일이 아닐 수 있습니다. 엑셀에서 열어 다시 저장한 뒤 넣어 주세요.',
    ZIP_METHOD: '이 엑셀 파일은 지원하지 않는 방식으로 압축돼 있습니다. 엑셀에서 열어 .xlsx 또는 CSV로 다시 저장해 주세요.',
    NO_DECOMPRESS: '이 브라우저는 엑셀(.xlsx) 압축 풀기를 지원하지 않습니다. 최신 크롬·엣지·사파리·파이어폭스를 쓰거나, 엑셀에서 CSV로 저장해 넣어 주세요.',
    NO_SHEET: '엑셀 안에서 시트를 찾지 못했습니다. 엑셀에서 열어 다시 저장해 주세요.',
    BAD_ASOF: '기준일을 읽지 못했습니다. 예: 2026-10-29',
    BAD_DAYS: '기다리는 날과 여유일은 0~365 사이의 정수로 넣어 주세요.',
    MAP_ORDERS: '주문 엑셀에서 "주문번호"와 "배송완료일" 열을 골라 주세요.',
    MAP_SETTLEMENT: '정산 엑셀에서 "주문번호" 열을 골라 주세요.'
  };

  function fail(code) {
    const e = new Error(MESSAGES[code] || code);
    e.code = code;
    return e;
  }

  // ---------- small helpers ----------

  function str(v) { return v == null ? '' : String(v); }

  function isEmptyRow(r) {
    if (!r || !r.length) return true;
    for (let i = 0; i < r.length; i++) if (str(r[i]).trim() !== '') return false;
    return true;
  }

  // 'A' -> 0, 'Z' -> 25, 'AA' -> 26, 'AA10' -> 26
  function columnIndex(ref) {
    const m = /^([A-Za-z]+)/.exec(str(ref));
    if (!m) return -1;
    const s = m[1].toUpperCase();
    let n = 0;
    for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
    return n - 1;
  }

  // 0 -> 'A', 26 -> 'AA'
  function columnName(i) {
    let n = i + 1, s = '';
    while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  function decodeXml(s) {
    return str(s).replace(/&(#[xX][0-9a-fA-F]+|#\d+|lt|gt|amp|quot|apos|nbsp);/g, function (m, e) {
      if (e[0] === '#') {
        const cp = (e[1] === 'x' || e[1] === 'X') ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
        try { return String.fromCodePoint(cp); } catch (_) { return m; }
      }
      return { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'", nbsp: '\u00a0' }[e];
    });
  }

  // OOXML writes some characters as _xHHHH_ (for example _x000D_ for a carriage return).
  function decodeOoxmlEscapes(s) {
    return str(s).replace(/_x([0-9A-Fa-f]{4})_/g, function (m, h) { return String.fromCharCode(parseInt(h, 16)); });
  }

  // ---------- text decoding (UTF-8 first, then EUC-KR) ----------

  function toU8(buf) {
    if (buf instanceof Uint8Array) return buf;
    if (buf instanceof ArrayBuffer) return new Uint8Array(buf);
    if (buf && buf.buffer instanceof ArrayBuffer) return new Uint8Array(buf.buffer, buf.byteOffset || 0, buf.byteLength);
    return new Uint8Array(0);
  }

  function decodeBytes(input) {
    const u8 = toU8(input);
    function dec(label, bytes, opts) { return new TextDecoder(label, opts).decode(bytes); }
    try {
      if (u8[0] === 0xEF && u8[1] === 0xBB && u8[2] === 0xBF) return { text: dec('utf-8', u8.subarray(3)), encoding: 'utf-8' };
      if (u8[0] === 0xFF && u8[1] === 0xFE) return { text: dec('utf-16le', u8.subarray(2)), encoding: 'utf-16le' };
      if (u8[0] === 0xFE && u8[1] === 0xFF) return { text: dec('utf-16be', u8.subarray(2)), encoding: 'utf-16be' };
    } catch (_) { /* fall through */ }
    try { return { text: dec('utf-8', u8, { fatal: true }), encoding: 'utf-8' }; } catch (_) { /* not utf-8 */ }
    try { return { text: dec('euc-kr', u8), encoding: 'euc-kr' }; } catch (_) { /* no euc-kr support */ }
    return { text: dec('utf-8', u8), encoding: 'unknown' };
  }

  // ---------- CSV ----------

  function stripBOM(t) { t = str(t); return t.charCodeAt(0) === 0xfeff ? t.slice(1) : t; }

  function detectDelimiter(text) {
    text = stripBOM(text);
    const counts = { ',': 0, ';': 0, '\t': 0 };
    let inQ = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (c === '"') inQ = !inQ;
      else if (!inQ && (c === '\n' || c === '\r')) { if (counts[','] + counts[';'] + counts['\t'] > 0) break; }
      else if (!inQ && Object.prototype.hasOwnProperty.call(counts, c)) counts[c]++;
    }
    let best = ',';
    if (counts['\t'] > counts[best]) best = '\t';
    if (counts[';'] > counts[best]) best = ';';
    return best;
  }

  // RFC 4180-style: quoted fields, "" escapes, delimiters and newlines inside quotes, CRLF/LF/CR, BOM.
  function parseCSV(text, delimiter) {
    text = stripBOM(text);
    const d = delimiter || detectDelimiter(text);
    const rows = [];
    let row = [], field = '', inQ = false, i = 0;
    const n = text.length;
    while (i < n) {
      const c = text[i];
      if (inQ) {
        if (c === '"') {
          if (text[i + 1] === '"') { field += '"'; i += 2; continue; }
          inQ = false; i++; continue;
        }
        field += c; i++; continue;
      }
      if (c === '"' && field === '') { inQ = true; i++; continue; }
      if (c === d) { row.push(field); field = ''; i++; continue; }
      if (c === '\r' || c === '\n') {
        row.push(field); rows.push(row); row = []; field = '';
        if (c === '\r' && text[i + 1] === '\n') i++;
        i++; continue;
      }
      field += c; i++;
    }
    if (field !== '' || row.length) { row.push(field); rows.push(row); }
    return rows;
  }

  // ---------- HTML table (some admin sites save an HTML table with an .xls name) ----------

  function parseHtmlTable(html) {
    const rows = [];
    const trRe = /<tr\b[^>]*>([\s\S]*?)<\/tr>/gi;
    let tm;
    while ((tm = trRe.exec(str(html)))) {
      const row = [];
      const tdRe = /<t([dh])\b[^>]*>([\s\S]*?)<\/t\1>/gi;
      let cm;
      while ((cm = tdRe.exec(tm[1]))) {
        const t = decodeXml(cm[2].replace(/<br\s*\/?>/gi, ' ').replace(/<[^>]+>/g, '')).replace(/[\s\u00a0]+/g, ' ').trim();
        row.push(t);
      }
      rows.push(row);
    }
    return rows;
  }

  // ---------- XLSX (ZIP + DecompressionStream, no library) ----------

  function utf8(bytes) { return new TextDecoder('utf-8').decode(bytes); }

  function unzipEntries(u8) {
    if (u8.length < 22) throw fail('ZIP_BAD');
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let eocd = -1;
    const stop = Math.max(0, u8.length - 22 - 65535);
    for (let i = u8.length - 22; i >= stop; i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw fail('ZIP_BAD');
    const count = dv.getUint16(eocd + 10, true);
    const cdOffset = dv.getUint32(eocd + 16, true);
    if (count === 0xffff || cdOffset === 0xffffffff || cdOffset >= u8.length) throw fail('ZIP_BAD');
    const entries = {};
    let p = cdOffset;
    for (let k = 0; k < count; k++) {
      if (p + 46 > u8.length || dv.getUint32(p, true) !== 0x02014b50) throw fail('ZIP_BAD');
      const method = dv.getUint16(p + 10, true);
      const csize = dv.getUint32(p + 20, true);
      const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
      const lho = dv.getUint32(p + 42, true);
      const name = utf8(u8.subarray(p + 46, p + 46 + nlen));
      entries[name.toLowerCase()] = { name: name, method: method, csize: csize, lho: lho };
      p += 46 + nlen + xlen + clen;
    }
    return entries;
  }

  async function inflateRaw(bytes) {
    if (typeof DecompressionStream === 'undefined') throw fail('NO_DECOMPRESS');
    let ds;
    try { ds = new DecompressionStream('deflate-raw'); } catch (_) { throw fail('NO_DECOMPRESS'); }
    try {
      const stream = new Blob([bytes]).stream().pipeThrough(ds);
      return new Uint8Array(await new Response(stream).arrayBuffer());
    } catch (_) { throw fail('ZIP_BAD'); }
  }

  async function readEntry(u8, entries, name) {
    const e = entries[String(name).toLowerCase()];
    if (!e) return null;
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    if (e.lho + 30 > u8.length || dv.getUint32(e.lho, true) !== 0x04034b50) throw fail('ZIP_BAD');
    const nlen = dv.getUint16(e.lho + 26, true), xlen = dv.getUint16(e.lho + 28, true);
    const start = e.lho + 30 + nlen + xlen;
    const data = u8.subarray(start, start + e.csize);
    if (e.method === 0) return data;
    if (e.method === 8) return inflateRaw(data);
    throw fail('ZIP_METHOD');
  }

  async function readEntryText(u8, entries, name) {
    const b = await readEntry(u8, entries, name);
    return b == null ? null : utf8(b);
  }

  // Remove namespace prefixes from element names (<x:c> -> <c>). Attribute names keep theirs.
  function stripNs(xml) { return str(xml).replace(/<(\/?)[A-Za-z_][\w.-]*:(?=[A-Za-z_])/g, '<$1'); }

  function attrs(s) {
    const o = {};
    const re = /([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
    let m;
    while ((m = re.exec(str(s)))) o[m[1]] = decodeXml(m[2] !== undefined ? m[2] : m[3]);
    return o;
  }

  // Text of <t> runs, skipping phonetic hints (<rPh>).
  function textOfRuns(xml) {
    const clean = str(xml).replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
    let out = '';
    const re = /<t\b[^>]*?(?:\/>|>([\s\S]*?)<\/t>)/g;
    let m;
    while ((m = re.exec(clean))) out += m[1] || '';
    return decodeOoxmlEscapes(decodeXml(out));
  }

  function parseSharedStrings(xml) {
    const out = [];
    if (!xml) return out;
    const re = /<si\b[^>]*?(?:\/>|>([\s\S]*?)<\/si>)/g;
    let m;
    const x = stripNs(xml);
    while ((m = re.exec(x))) out.push(textOfRuns(m[1] || ''));
    return out;
  }

  function parseSheetXml(xml, shared) {
    const x = stripNs(xml);
    const sd = /<sheetData\b[^>]*?(?:\/>|>([\s\S]*?)<\/sheetData>)/.exec(x);
    if (!sd) return [];
    const body = sd[1] || '';
    const rows = [];
    const rowRe = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g;
    let rm, nextRow = 0;
    while ((rm = rowRe.exec(body))) {
      const ra = attrs(rm[1]);
      let ri = ra.r ? parseInt(ra.r, 10) - 1 : nextRow;
      if (!(ri >= 0)) ri = nextRow;
      nextRow = ri + 1;
      const row = [];
      const cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
      let cm, nextCol = 0;
      const inner = rm[2] || '';
      while ((cm = cellRe.exec(inner))) {
        const ca = attrs(cm[1]);
        let ci = ca.r ? columnIndex(ca.r) : nextCol;
        if (!(ci >= 0)) ci = nextCol;
        nextCol = ci + 1;
        const cin = cm[2] || '';
        const t = ca.t || 'n';
        let val = '';
        if (t === 'inlineStr') {
          const is = /<is\b[^>]*>([\s\S]*?)<\/is>/.exec(cin);
          val = is ? textOfRuns(is[1]) : '';
        } else {
          const vm = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(cin);
          const v = vm ? decodeOoxmlEscapes(decodeXml(vm[1])) : '';
          if (t === 's') {
            const k = parseInt(v, 10);
            val = (shared && k >= 0 && k < shared.length) ? shared[k] : '';
          } else if (t === 'b') val = v === '1' ? 'TRUE' : (v === '0' ? 'FALSE' : v);
          else val = v;
        }
        while (row.length < ci) row.push('');
        row[ci] = val;
      }
      while (rows.length < ri) rows.push([]);
      rows[ri] = row;
    }
    return rows;
  }

  function dirOf(path) { const i = path.lastIndexOf('/'); return i >= 0 ? path.slice(0, i + 1) : ''; }

  function resolvePath(baseDir, target) {
    target = str(target).replace(/\\/g, '/');
    const full = target[0] === '/' ? target.slice(1) : baseDir + target;
    const out = [];
    full.split('/').forEach(function (seg) {
      if (seg === '..') out.pop();
      else if (seg !== '.' && seg !== '') out.push(seg);
    });
    return out.join('/');
  }

  function parseRels(xml) {
    const list = [];
    const re = /<Relationship\b([^>]*?)\/?>/g;
    let m;
    const x = stripNs(xml || '');
    while ((m = re.exec(x))) list.push(attrs(m[1]));
    return list;
  }

  function relId(a) {
    if (a['r:id']) return a['r:id'];
    for (const k in a) if (/:id$/.test(k)) return a[k];
    return a.id || '';
  }

  // Returns { sheets: [names], sheetIndex, sheetName, rows: [[...]] }. Empty cells are kept as ''.
  async function readXlsx(buf, opts) {
    const u8 = toU8(buf);
    if (u8[0] === 0xD0 && u8[1] === 0xCF && u8[2] === 0x11 && u8[3] === 0xE0) throw fail('OLE');
    const entries = unzipEntries(u8);
    let wbPath = 'xl/workbook.xml';
    const rootRels = parseRels(await readEntryText(u8, entries, '_rels/.rels'));
    for (const r of rootRels) if (/\/officeDocument$/.test(r.Type || '')) { wbPath = resolvePath('', r.Target); break; }
    if (!entries[wbPath.toLowerCase()] && entries['xl/workbook.xml']) wbPath = 'xl/workbook.xml';
    const baseDir = dirOf(wbPath);
    const wbXml = await readEntryText(u8, entries, wbPath);
    const relsPath = baseDir + '_rels/' + wbPath.slice(baseDir.length) + '.rels';
    const rels = parseRels(await readEntryText(u8, entries, relsPath));
    const relById = {};
    rels.forEach(function (r) { relById[r.Id] = r; });

    const sheets = [];
    if (wbXml) {
      const re = /<sheet\b([^>]*?)\/?>/g;
      let m;
      const x = stripNs(wbXml);
      while ((m = re.exec(x))) {
        const a = attrs(m[1]);
        const rel = relById[relId(a)];
        sheets.push({ name: a.name || ('Sheet' + (sheets.length + 1)), path: rel ? resolvePath(baseDir, rel.Target) : null });
      }
    }
    if (!sheets.length) {
      if (entries['xl/worksheets/sheet1.xml']) sheets.push({ name: 'Sheet1', path: 'xl/worksheets/sheet1.xml' });
      else throw fail('NO_SHEET');
    }
    let idx = opts && Number.isInteger(opts.sheetIndex) ? opts.sheetIndex : 0;
    if (idx < 0 || idx >= sheets.length) idx = 0;
    let path = sheets[idx].path;
    if (!path || !entries[path.toLowerCase()]) {
      const guess = baseDir + 'worksheets/sheet' + (idx + 1) + '.xml';
      if (entries[guess.toLowerCase()]) path = guess; else throw fail('NO_SHEET');
    }
    let ssPath = baseDir + 'sharedStrings.xml';
    for (const r of rels) if (/\/sharedStrings$/.test(r.Type || '')) { ssPath = resolvePath(baseDir, r.Target); break; }
    const shared = parseSharedStrings(await readEntryText(u8, entries, ssPath));
    const sheetXml = await readEntryText(u8, entries, path);
    return {
      sheets: sheets.map(function (s) { return s.name; }),
      sheetIndex: idx,
      sheetName: sheets[idx].name,
      rows: parseSheetXml(sheetXml, shared)
    };
  }

  function sniffFormat(u8) {
    if (u8.length >= 4 && u8[0] === 0x50 && u8[1] === 0x4B && (u8[2] === 0x03 || u8[2] === 0x05) && (u8[3] === 0x04 || u8[3] === 0x06)) return 'xlsx';
    if (u8.length >= 4 && u8[0] === 0xD0 && u8[1] === 0xCF && u8[2] === 0x11 && u8[3] === 0xE0) return 'ole';
    return 'text';
  }

  // ---------- fields, header row, column mapping ----------

  const ORDER_NO_EXCLUDE = ['묶음배송', '배송번호', '운송장', '송장', '출고'];

  const FIELDS = {
    orders: [
      { key: 'orderNo', label: '주문번호', required: true, keywords: ['주문번호'], exclude: ORDER_NO_EXCLUDE },
      { key: 'deliveredAt', label: '배송완료일', required: true, keywords: ['배송완료일', '배송완료일시', '배송완료'], exclude: ['예정'] },
      { key: 'status', label: '주문 상태', required: false, keywords: ['주문상태', '배송상태', '상태'] },
      { key: 'product', label: '상품명', required: false, keywords: ['상품명', '노출상품명', '등록상품명', '옵션명'] },
      { key: 'amount', label: '금액', required: false, keywords: ['결제액', '결제금액', '판매가', '주문금액', '매출금액'] },
      { key: 'qty', label: '수량', required: false, keywords: ['수량', '구매수'] }
    ],
    settlement: [
      { key: 'orderNo', label: '주문번호', required: true, keywords: ['주문번호'], exclude: ORDER_NO_EXCLUDE },
      { key: 'amount', label: '정산 금액', required: false, keywords: ['정산금액', '정산대상금액', '매출금액', '판매액'] },
      { key: 'recognizedAt', label: '매출인식일·정산일', required: false, keywords: ['매출인식일', '구매확정일', '정산일', '지급예정일'] }
    ]
  };

  // Remove spaces and brackets, lower-case.
  function normText(s) { return str(s).replace(/[\s()\[\]{}<>（）［］]/g, '').toLowerCase(); }

  function keywordsFor(kind) {
    const defs = kind && FIELDS[kind] ? FIELDS[kind] : FIELDS.orders.concat(FIELDS.settlement);
    const set = [];
    defs.forEach(function (f) { f.keywords.forEach(function (k) { const n = normText(k); if (set.indexOf(n) < 0) set.push(n); }); });
    return set;
  }

  // Export files may have a title line or a blank line above the real header.
  // Look at the first 15 rows and pick the row with the most keyword matches (earliest on ties).
  function findHeaderRow(rows, kind) {
    const kws = keywordsFor(kind);
    let best = 0, bestScore = 0;
    const lim = Math.min((rows || []).length, 15);
    for (let i = 0; i < lim; i++) {
      let s = 0;
      (rows[i] || []).forEach(function (cell) {
        const h = normText(cell);
        if (!h || h.length > 30) return;
        if (kws.some(function (k) { return h.indexOf(k) >= 0; })) s++;
      });
      if (s > bestScore) { best = i; bestScore = s; }
    }
    return best;
  }

  function tableFromRows(rows, kind) {
    rows = rows || [];
    const h = findHeaderRow(rows, kind);
    const headers = (rows[h] || []).map(function (x) { return str(x).trim(); });
    const data = [];
    for (let i = h + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      if (!isEmptyRow(r)) data.push(r.map(str));
    }
    let maxLen = headers.length;
    data.forEach(function (r) { if (r.length > maxLen) maxLen = r.length; });
    while (headers.length < maxLen) headers.push('');
    return { headerRowIndex: h, headers: headers, rows: data };
  }

  // Returns { fieldKey: columnIndex or -1 }. Exact match first (keyword order), then partial
  // match (keyword order, shortest header wins). A column is used for at most one field.
  function autoMapColumns(headers, kind) {
    const defs = FIELDS[kind];
    if (!defs) throw new Error('kind must be orders or settlement');
    const norm = (headers || []).map(normText);
    const used = {};
    const map = {};
    defs.forEach(function (f) {
      const ok = function (h, i) {
        return h && !used[i] && !(f.exclude || []).some(function (x) { return h.indexOf(normText(x)) >= 0; });
      };
      let found = -1;
      for (let k = 0; k < f.keywords.length && found < 0; k++) {
        const kw = normText(f.keywords[k]);
        for (let i = 0; i < norm.length; i++) if (ok(norm[i], i) && norm[i] === kw) { found = i; break; }
      }
      for (let k = 0; k < f.keywords.length && found < 0; k++) {
        const kw = normText(f.keywords[k]);
        let best = -1;
        for (let i = 0; i < norm.length; i++) {
          if (ok(norm[i], i) && norm[i].indexOf(kw) >= 0 && (best < 0 || norm[i].length < norm[best].length)) best = i;
        }
        found = best;
      }
      map[f.key] = found;
      if (found >= 0) used[found] = true;
    });
    return map;
  }

  // ---------- values ----------

  // Order numbers: remove spaces and hyphens. Flags values that may have lost digits in Excel.
  function normalizeOrderNo(v) {
    let warning = null;
    if (typeof v === 'number') {
      if (!Number.isSafeInteger(v)) warning = 'precision';
      v = (Number.isInteger(v) && Math.abs(v) < 1e21) ? BigInt(v).toString() : String(v);
    }
    const raw = str(v).trim();
    let s = raw.replace(/^'/, '');
    const eq = /^="(.*)"$/.exec(s);
    if (eq) s = eq[1];
    s = s.replace(/[\s\u00a0\-\u2010-\u2015]/g, '');
    if (/^\d+\.0+$/.test(s)) s = s.replace(/\.0+$/, '');
    const ex = /^(\d+)(?:\.(\d+))?[eE]\+?(\d+)$/.exec(s);
    if (ex) {
      warning = 'exponent';
      const ip = ex[1], fr = ex[2] || '', e = parseInt(ex[3], 10);
      if (fr.length <= e && e < 40) s = (ip + fr + '0'.repeat(e - fr.length)).replace(/^0+(?=\d)/, '');
    } else if (!warning && /^\d{17,}$/.test(s)) warning = 'long';
    return { value: s, raw: raw, warning: warning };
  }

  function pad2(n) { return (n < 10 ? '0' : '') + n; }

  function mkDate(y, m, d) {
    if (!(y >= 1900 && y <= 2200) || !(m >= 1 && m <= 12) || !(d >= 1 && d <= 31)) return null;
    const t = Date.UTC(y, m - 1, d);
    const dt = new Date(t);
    if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return null;
    return { y: y, m: m, d: d, day: Math.round(t / 86400000), iso: y + '-' + pad2(m) + '-' + pad2(d) };
  }

  // Excel 1900 date system. Serial 60 is the non-existent 1900-02-29.
  function fromSerial(n) {
    if (!(n >= 1 && n < 2958466)) return null;
    const days = Math.floor(n);
    if (days === 60) return null;
    const base = days < 60 ? Date.UTC(1899, 11, 31) : Date.UTC(1899, 11, 30);
    const dt = new Date(base + days * 86400000);
    return mkDate(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate());
  }

  // Returns { y, m, d, day, iso } or null.
  function parseDate(v) {
    if (v == null) return null;
    if (v instanceof Date) return isNaN(v.getTime()) ? null : mkDate(v.getFullYear(), v.getMonth() + 1, v.getDate());
    if (typeof v === 'number') return fromSerial(v);
    const s = String(v).trim();
    if (!s) return null;
    let m;
    if ((m = /^(\d{4})(\d{2})(\d{2})(?:\d{4}|\d{6})?$/.exec(s))) return mkDate(+m[1], +m[2], +m[3]);
    if ((m = /^(\d{4})\s*[-./]\s*(\d{1,2})\s*[-./]\s*(\d{1,2})\s*\.?(?:$|[\sT])/.exec(s))) return mkDate(+m[1], +m[2], +m[3]);
    if ((m = /^(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일/.exec(s))) return mkDate(+m[1], +m[2], +m[3]);
    if (/^\d{1,7}(\.\d+)?$/.test(s)) return fromSerial(parseFloat(s));
    return null;
  }

  function parseAmount(v) {
    if (v == null) return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    let s = String(v).trim();
    if (!s) return null;
    let neg = false;
    if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
    s = s.replace(/[,\s\u00a0원₩]|KRW/gi, '');
    if (s[0] === '-' || s[0] === '\u2212') { neg = !neg; s = s.slice(1); }
    if (!/^\d+(\.\d+)?$/.test(s)) return null;
    const n = parseFloat(s);
    return neg ? -n : n;
  }

  const EXCLUDED_STATUS = /취소|반품|교환/;
  function isExcludedStatus(s) { return EXCLUDED_STATUS.test(str(s).replace(/\s/g, '')); }

  function cell(row, idx) { return (idx >= 0 && row) ? (row[idx] == null ? '' : row[idx]) : ''; }

  const WARNING_TEXT = {
    exponent: '주문번호가 "1.23457E+15"처럼 지수 형태로 저장된 줄이 있습니다. 엑셀에서 숫자로 바뀌며 끝자리가 사라졌을 수 있어 주문번호 대조가 틀릴 수 있습니다. WING에서 파일을 다시 내려받아, 엑셀에서 열고 저장하지 말고 그대로 이 화면에 넣어 주세요.',
    long: '17자리 이상 숫자로 된 주문번호가 있습니다. 엑셀에서 숫자로 바뀌며 끝자리가 달라졌을 수 있으니 WING 화면의 주문번호와 비교해 주세요.',
    precision: '숫자로 저장된 주문번호 중 정밀도를 넘는 값이 있습니다. 끝자리가 바뀌었을 수 있습니다.',
    badDate: '배송완료일을 날짜로 읽지 못한 줄이 있습니다. 이 주문은 "제외"로 분류했습니다.',
    noOrderNo: '주문번호가 비어 있는 줄은 건너뛰었습니다(합계 줄 등).',
    noStatus: '주문 상태 열을 고르지 않아 취소·반품·교환 주문을 거르지 못했습니다. 상태 열이 있으면 골라 주세요.'
  };

  // Main comparison. orders/settlement = { rows: [[...]], map: { field: columnIndex } }.
  function findUnsettled(opts) {
    opts = opts || {};
    const orders = opts.orders || {}, settlement = opts.settlement || {};
    const waitDays = opts.waitDays == null || opts.waitDays === '' ? 7 : Number(opts.waitDays);
    const graceDays = opts.graceDays == null || opts.graceDays === '' ? 3 : Number(opts.graceDays);
    if (!Number.isInteger(waitDays) || waitDays < 0 || waitDays > 365 || !Number.isInteger(graceDays) || graceDays < 0 || graceDays > 365) throw fail('BAD_DAYS');
    const asOf = parseDate(opts.asOf);
    if (!asOf) throw fail('BAD_ASOF');
    const om = orders.map || {}, sm = settlement.map || {};
    if (!(om.orderNo >= 0) || !(om.deliveredAt >= 0)) throw fail('MAP_ORDERS');
    if (!(sm.orderNo >= 0)) throw fail('MAP_SETTLEMENT');
    const threshold = waitDays + graceDays;
    const wc = { exponent: 0, long: 0, precision: 0, badDate: 0, noOrderNo: 0, noStatus: 0 };

    // Settlement order numbers
    const settled = {};
    let sMin = null, sMax = null;
    (settlement.rows || []).forEach(function (r) {
      if (isEmptyRow(r)) return;
      const no = normalizeOrderNo(cell(r, sm.orderNo));
      if (!no.value) return;
      if (no.warning) wc[no.warning]++;
      const rec = settled[no.value] || (settled[no.value] = { orderNo: no.raw, amount: null, lines: 0 });
      rec.lines++;
      const a = sm.amount >= 0 ? parseAmount(cell(r, sm.amount)) : null;
      if (a != null) rec.amount = (rec.amount || 0) + a;
      if (sm.recognizedAt >= 0) {
        const d = parseDate(cell(r, sm.recognizedAt));
        if (d) { if (!sMin || d.day < sMin.day) sMin = d; if (!sMax || d.day > sMax.day) sMax = d; }
      }
    });

    // Group order lines by order number
    const groups = {}, order = [];
    (orders.rows || []).forEach(function (r) {
      if (isEmptyRow(r)) return;
      const no = normalizeOrderNo(cell(r, om.orderNo));
      if (!no.value) { wc.noOrderNo++; return; }
      if (no.warning) wc[no.warning]++;
      let g = groups[no.value];
      if (!g) { g = groups[no.value] = { key: no.value, orderNo: no.raw, lines: [] }; order.push(g); }
      const rawDate = str(cell(r, om.deliveredAt)).trim();
      const d = parseDate(rawDate);
      if (!d && rawDate) wc.badDate++;
      g.lines.push({
        delivered: d, rawDate: rawDate,
        status: om.status >= 0 ? str(cell(r, om.status)).trim() : '',
        product: om.product >= 0 ? str(cell(r, om.product)).trim() : '',
        amount: om.amount >= 0 ? parseAmount(cell(r, om.amount)) : null,
        qty: om.qty >= 0 ? parseAmount(cell(r, om.qty)) : null
      });
    });
    if (!(om.status >= 0)) wc.noStatus = 1;

    const out = { needsCheck: [], waiting: [], settled: [], excluded: [] };
    order.forEach(function (g) {
      const active = g.lines.filter(function (l) { return !isExcludedStatus(l.status); });
      const pool = active.length ? active : g.lines;
      let first = '';
      for (const l of pool) if (l.product) { first = l.product; break; }
      const product = first + (pool.length > 1 ? (first ? ' ' : '') + '외 ' + (pool.length - 1) : '');
      let amount = null, qty = null;
      pool.forEach(function (l) {
        if (l.amount != null) amount = (amount || 0) + l.amount;
        if (l.qty != null) qty = (qty || 0) + l.qty;
      });
      const base = { orderNo: g.orderNo, key: g.key, lines: g.lines.length, product: product, amount: amount, qty: qty };

      if (!active.length) {
        out.excluded.push(Object.assign(base, { reason: '취소·반품·교환', status: g.lines[0].status }));
        return;
      }
      const dates = active.map(function (l) { return l.delivered; }).filter(Boolean);
      const last = dates.length ? dates.reduce(function (a, b) { return b.day > a.day ? b : a; }) : null;
      if (settled[g.key]) {
        out.settled.push(Object.assign(base, { deliveredAt: last ? last.iso : '' }));
        return;
      }
      if (!last) {
        const bad = active.some(function (l) { return l.rawDate; });
        out.excluded.push(Object.assign(base, { reason: bad ? '배송완료일을 읽지 못함' : '배송완료일 없음(미배송 등)', status: active[0].status }));
        return;
      }
      const notes = [];
      if (dates.length < active.length) notes.push('일부 상품은 배송완료일 없음');
      if (active.length < g.lines.length) notes.push('일부 상품 취소·반품·교환 제외');
      const elapsed = asOf.day - last.day;
      const rec = Object.assign(base, { deliveredAt: last.iso, elapsedDays: elapsed });
      if (elapsed > threshold) {
        if (sMin && sMax) {
          const expected = last.day + waitDays;
          if (expected < sMin.day || expected > sMax.day) notes.push('예상 구매확정일이 정산 엑셀 날짜 범위 밖');
        }
        rec.note = notes.join(', ');
        out.needsCheck.push(rec);
      } else {
        rec.note = notes.join(', ');
        out.waiting.push(rec);
      }
    });

    out.needsCheck.sort(function (a, b) { return b.elapsedDays - a.elapsedDays || (a.key < b.key ? -1 : a.key > b.key ? 1 : 0); });
    out.waiting.sort(function (a, b) { return b.elapsedDays - a.elapsedDays; });

    const settlementOnly = [];
    Object.keys(settled).forEach(function (k) { if (!groups[k]) settlementOnly.push(settled[k].orderNo); });

    let needsCheckAmount = null;
    out.needsCheck.forEach(function (r) { if (r.amount != null) needsCheckAmount = (needsCheckAmount || 0) + r.amount; });

    const warnings = [];
    Object.keys(wc).forEach(function (k) { if (wc[k]) warnings.push({ code: k, count: wc[k], message: WARNING_TEXT[k] }); });

    return {
      asOf: asOf.iso, waitDays: waitDays, graceDays: graceDays, threshold: threshold,
      needsCheck: out.needsCheck, waiting: out.waiting, settled: out.settled, excluded: out.excluded,
      settlementOnly: settlementOnly,
      settlementRange: sMin ? { min: sMin.iso, max: sMax.iso } : null,
      needsCheckAmount: needsCheckAmount,
      counts: {
        orders: order.length, needsCheck: out.needsCheck.length, waiting: out.waiting.length,
        settled: out.settled.length, excluded: out.excluded.length, settlementOnly: settlementOnly.length
      },
      warnings: warnings
    };
  }

  // ---------- CSV export ('확인 필요' list) ----------

  function csvCell(v) {
    let s = str(v);
    if (/^[=+\-@\t\r]/.test(s)) s = "'" + s;            // spreadsheet formula guard
    if (/[",\r\n]/.test(s)) s = '"' + s.replace(/"/g, '""') + '"';
    return s;
  }

  // Long digit-only order numbers are written as ="..." so Excel keeps every digit.
  function orderNoCell(v) {
    const s = str(v);
    return /^\d{12,}$/.test(s) ? '="' + s + '"' : csvCell(s);
  }

  function toCSV(result) {
    const lines = [['주문번호', '배송완료일', '경과일', '상품명', '주문 엑셀 기준 금액(정산 예정액 아님)', '메모'].map(csvCell).join(',')];
    (result && result.needsCheck || []).forEach(function (r) {
      lines.push([
        orderNoCell(r.orderNo), csvCell(r.deliveredAt), String(r.elapsedDays),
        csvCell(r.product), r.amount == null ? '' : String(r.amount), csvCell(r.note || '')
      ].join(','));
    });
    return '\ufeff' + lines.join('\r\n') + '\r\n';
  }

  // ---------- one call for a file ----------

  async function loadTable(buf, fileName, kind, opts) {
    const u8 = toU8(buf);
    if (!u8.length) throw fail('EMPTY');
    const fmt = sniffFormat(u8);
    if (fmt === 'ole') throw fail('OLE');
    const meta = { name: fileName || '', format: fmt };
    let rows;
    if (fmt === 'xlsx') {
      const x = await readXlsx(u8, opts);
      rows = x.rows; meta.sheets = x.sheets; meta.sheetIndex = x.sheetIndex; meta.sheetName = x.sheetName;
    } else {
      const dec = decodeBytes(u8);
      meta.encoding = dec.encoding;
      if (/^\s*</.test(dec.text) && /<table\b/i.test(dec.text)) { rows = parseHtmlTable(dec.text); meta.format = 'html'; }
      else { rows = parseCSV(dec.text); meta.format = 'csv'; }
    }
    if (!rows.some(function (r) { return !isEmptyRow(r); })) throw fail('EMPTY');
    const t = tableFromRows(rows, kind);
    return Object.assign(t, meta);
  }

  return {
    MESSAGES: MESSAGES, FIELDS: FIELDS,
    columnIndex: columnIndex, columnName: columnName,
    decodeBytes: decodeBytes, parseCSV: parseCSV, detectDelimiter: detectDelimiter, parseHtmlTable: parseHtmlTable,
    readXlsx: readXlsx, sniffFormat: sniffFormat,
    findHeaderRow: findHeaderRow, tableFromRows: tableFromRows, autoMapColumns: autoMapColumns,
    normalizeOrderNo: normalizeOrderNo, parseDate: parseDate, parseAmount: parseAmount, isExcludedStatus: isExcludedStatus,
    findUnsettled: findUnsettled, toCSV: toCSV, loadTable: loadTable
  };
});
