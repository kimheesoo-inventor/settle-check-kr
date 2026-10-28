// 정산 확인기 tests. Run: node test.mjs   (Node 18+ built-ins only, no packages)
// SETTLE_CORE=<path> lets a mutation check load a modified copy of core.js.
import { createRequire } from 'node:module';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join, resolve } from 'node:path';
import assert from 'node:assert/strict';

const here = dirname(fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);
const C = require(process.env.SETTLE_CORE ? resolve(process.env.SETTLE_CORE) : './core.js');
const S = require('./samples.js');

let passed = 0, failed = 0;
const queue = [];
function t(name, fn) { queue.push([name, fn]); }

// ---------- helpers ----------

const enc = new TextEncoder();

function crc32(u8) {
  let c, crc = 0xffffffff;
  for (let i = 0; i < u8.length; i++) {
    c = (crc ^ u8[i]) & 0xff;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crc = (crc >>> 8) ^ c;
  }
  return (crc ^ 0xffffffff) >>> 0;
}

async function deflateRaw(u8) {
  const s = new Blob([u8]).stream().pipeThrough(new CompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(s).arrayBuffer());
}

// files: [{ name, text, method: 0 | 8 }]
async function buildZip(files) {
  const parts = [], central = [];
  let offset = 0;
  for (const f of files) {
    const name = enc.encode(f.name);
    const raw = enc.encode(f.text);
    const data = f.method === 8 ? await deflateRaw(raw) : raw;
    const crc = crc32(raw);
    const lh = new DataView(new ArrayBuffer(30));
    lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true);
    lh.setUint16(8, f.method, true); lh.setUint32(14, crc, true);
    lh.setUint32(18, data.length, true); lh.setUint32(22, raw.length, true);
    lh.setUint16(26, name.length, true); lh.setUint16(28, 0, true);
    parts.push(new Uint8Array(lh.buffer), name, data);
    const ch = new DataView(new ArrayBuffer(46));
    ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true);
    ch.setUint16(10, f.method, true); ch.setUint32(16, crc, true);
    ch.setUint32(20, data.length, true); ch.setUint32(24, raw.length, true);
    ch.setUint16(28, name.length, true); ch.setUint32(42, offset, true);
    central.push(new Uint8Array(ch.buffer), name);
    offset += 30 + name.length + data.length;
  }
  const cdSize = central.reduce((a, b) => a + b.length, 0);
  const eocd = new DataView(new ArrayBuffer(22));
  eocd.setUint32(0, 0x06054b50, true);
  eocd.setUint16(8, files.length, true); eocd.setUint16(10, files.length, true);
  eocd.setUint32(12, cdSize, true); eocd.setUint32(16, offset, true);
  const all = [...parts, ...central, new Uint8Array(eocd.buffer)];
  const out = new Uint8Array(all.reduce((a, b) => a + b.length, 0));
  let p = 0;
  for (const a of all) { out.set(a, p); p += a.length; }
  return out;
}

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function xlsxFiles(method) {
  return [
    { name: '[Content_Types].xml', method, text: '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>' },
    { name: '_rels/.rels', method, text: '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' },
    { name: 'xl/workbook.xml', method, text: `<?xml version="1.0"?><workbook ${NS}><sheets><sheet name="주문목록" sheetId="1" r:id="rId1"/><sheet name="요약" sheetId="2" r:id="rId2"/></sheets></workbook>` },
    { name: 'xl/_rels/workbook.xml.rels', method, text: '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/data.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/sheet2.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>' },
    { name: 'xl/sharedStrings.xml', method, text: `<?xml version="1.0"?><sst ${NS} count="6" uniqueCount="6">` +
      '<si><t>주문번호</t></si>' +
      '<si><r><rPr><b/></rPr><t>배송</t></r><r><t xml:space="preserve">완료일</t></r></si>' +
      '<si><t>주문상태</t></si>' +
      '<si><t>상품 &amp; 옵션</t><rPh sb="0" eb="1"><t>ショウヒン</t></rPh></si>' +
      '<si><t>결제액</t></si>' +
      '<si><t>배송완료</t></si></sst>' },
    { name: 'xl/worksheets/data.xml', method, text: `<?xml version="1.0"?><worksheet ${NS}><dimension ref="A1:F5"/><cols><col min="1" max="1" width="20"/></cols><sheetData>` +
      '<row r="1"><c r="A1" t="inlineStr"><is><t>주문 목록 (테스트)</t></is></c></row>' +
      '<row r="3"><c r="A3" t="s"><v>0</v></c><c r="B3" t="s"><v>1</v></c><c r="C3" t="s"><v>2</v></c><c r="D3" t="s"><v>3</v></c><c r="E3" t="s"><v>4</v></c><c r="F3" t="inlineStr"><is><t>확인</t></is></c></row>' +
      '<row r="4"><c r="A4"><v>3000000000123</v></c><c r="B4" s="2"><v>46296</v></c><c r="C4" t="s"><v>5</v></c><c r="E4"><v>12900</v></c><c r="F4" t="b"><v>1</v></c></row>' +
      '<row r="5"><c r="A5" t="str"><f>A4+1</f><v>3000000000124</v></c><c r="B5" t="inlineStr"><is><t>2026-10-25</t></is></c><c r="C5" t="s"><v>5</v></c><c r="D5" t="inlineStr"><is><t>줄_x000D_바꿈</t></is></c><c r="E5" s="1"/><c r="F5" t="b"><v>0</v></c></row>' +
      '</sheetData></worksheet>' },
    { name: 'xl/worksheets/sheet2.xml', method, text: '<?xml version="1.0"?><x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData>' +
      '<x:row r="1"><x:c r="A1" t="inlineStr"><x:is><x:t>요약</x:t></x:is></x:c><x:c r="C1"><x:v>7</x:v></x:c></x:row></x:sheetData></x:worksheet>' }
  ];
}

// EUC-KR encoder built from the platform decoder (Hangul block of KS X 1001).
let eucMap = null;
function eucKr(text) {
  if (!eucMap) {
    eucMap = new Map();
    const d = new TextDecoder('euc-kr');
    for (let a = 0xB0; a <= 0xC8; a++) for (let b = 0xA1; b <= 0xFE; b++) {
      const ch = d.decode(new Uint8Array([a, b]));
      if (ch.length === 1 && !eucMap.has(ch)) eucMap.set(ch, [a, b]);
    }
  }
  const out = [];
  for (const ch of text) {
    const code = ch.charCodeAt(0);
    if (code < 0x80) out.push(code);
    else if (eucMap.has(ch)) out.push(...eucMap.get(ch));
    else throw new Error('no euc-kr byte for ' + ch);
  }
  return new Uint8Array(out);
}

function sampleTables() {
  const o = C.tableFromRows(S.orderRows, 'orders'); o.map = C.autoMapColumns(o.headers, 'orders');
  const s = C.tableFromRows(S.settlementRows, 'settlement'); s.map = C.autoMapColumns(s.headers, 'settlement');
  return { o, s };
}

// Small builder: orders with columns [주문번호, 배송완료일, 주문상태, 상품명, 결제액]
const OMAP = { orderNo: 0, deliveredAt: 1, status: 2, product: 3, amount: 4, qty: -1 };
const SMAP = { orderNo: 0, amount: -1, recognizedAt: 1 };
function run(orderRows, settleRows, extra) {
  return C.findUnsettled(Object.assign({
    orders: { rows: orderRows, map: OMAP }, settlement: { rows: settleRows, map: SMAP }, asOf: '2026-10-29'
  }, extra || {}));
}

// ---------- CSV / decoding ----------

t('parseCSV: 따옴표 안의 쉼표와 줄바꿈', () => {
  const rows = C.parseCSV('주문번호,상품명\n1,"가상, 상품\n둘째 줄"\n');
  assert.deepEqual(rows, [['주문번호', '상품명'], ['1', '가상, 상품\n둘째 줄']]);
});
t('parseCSV: BOM과 CRLF', () => {
  const rows = C.parseCSV('\ufeffa,b\r\n1,2\r\n');
  assert.deepEqual(rows, [['a', 'b'], ['1', '2']]);
});
t('parseCSV: "" 이스케이프와 빈 칸', () => {
  assert.deepEqual(C.parseCSV('x,"말 ""그대로""",,z'), [['x', '말 "그대로"', '', 'z']]);
});
t('parseCSV: 탭 구분 자동 인식', () => {
  assert.deepEqual(C.parseCSV('주문번호\t배송완료일\n1\t2026-10-01'), [['주문번호', '배송완료일'], ['1', '2026-10-01']]);
});
t('decodeBytes: UTF-8', () => {
  const r = C.decodeBytes(enc.encode('주문번호,배송완료일'));
  assert.equal(r.encoding, 'utf-8'); assert.equal(r.text, '주문번호,배송완료일');
});
t('decodeBytes: EUC-KR 재시도', () => {
  const r = C.decodeBytes(eucKr('주문번호,배송완료일\n1,2026-10-01'));
  assert.equal(r.encoding, 'euc-kr'); assert.equal(r.text, '주문번호,배송완료일\n1,2026-10-01');
});
t('loadTable: EUC-KR CSV + 제목줄 건너뛰기 + 자동 매칭', async () => {
  const bytes = eucKr('주문 목록\n\n주문번호,배송완료일,주문상태\n111,2026-10-01,배송완료\n');
  const tb = await C.loadTable(bytes, 'a.csv', 'orders');
  assert.equal(tb.format, 'csv'); assert.equal(tb.encoding, 'euc-kr'); assert.equal(tb.headerRowIndex, 2);
  assert.deepEqual(tb.rows, [['111', '2026-10-01', '배송완료']]);
  const m = C.autoMapColumns(tb.headers, 'orders');
  assert.equal(m.orderNo, 0); assert.equal(m.deliveredAt, 1); assert.equal(m.status, 2);
});
t('parseHtmlTable: .xls 이름의 HTML 표', async () => {
  const html = '<html><body><table><tr><th>주문번호</th><th>배송완료일</th></tr><tr><td> 9&amp;1 </td><td>2026-10-01<br></td></tr></table></body></html>';
  assert.deepEqual(C.parseHtmlTable(html), [['주문번호', '배송완료일'], ['9&1', '2026-10-01']]);
  const tb = await C.loadTable(enc.encode(html), 'x.xls', 'orders');
  assert.equal(tb.format, 'html'); assert.equal(tb.rows.length, 1);
});

// ---------- dates, numbers, order numbers ----------

t('parseDate: 2026-10-01', () => assert.equal(C.parseDate('2026-10-01').iso, '2026-10-01'));
t('parseDate: 2026.10.01 과 "2026. 10. 1."', () => {
  assert.equal(C.parseDate('2026.10.01').iso, '2026-10-01');
  assert.equal(C.parseDate('2026. 10. 1.').iso, '2026-10-01');
});
t('parseDate: 2026/10/01 과 시각 포함', () => {
  assert.equal(C.parseDate('2026/10/01').iso, '2026-10-01');
  assert.equal(C.parseDate('2026-10-01 13:22:11').iso, '2026-10-01');
  assert.equal(C.parseDate('2026-10-01T13:22').iso, '2026-10-01');
});
t('parseDate: 20261001 과 20261001132211', () => {
  assert.equal(C.parseDate('20261001').iso, '2026-10-01');
  assert.equal(C.parseDate('20261001132211').iso, '2026-10-01');
});
t('parseDate: 2026년 10월 1일', () => {
  assert.equal(C.parseDate('2026년 10월 1일').iso, '2026-10-01');
  assert.equal(C.parseDate('2026년 10월 1일 오후 3:00').iso, '2026-10-01');
});
t('parseDate: 엑셀 일련번호(1900 체계)', () => {
  assert.equal(C.parseDate('46296').iso, '2026-10-01');
  assert.equal(C.parseDate(46296.75).iso, '2026-10-01');
  assert.equal(C.parseDate(45658).iso, '2025-01-01');
  assert.equal(C.parseDate(1).iso, '1900-01-01');
  assert.equal(C.parseDate(59).iso, '1900-02-28');
  assert.equal(C.parseDate(61).iso, '1900-03-01');
  assert.equal(C.parseDate(60), null);
});
t('parseDate: 없는 날짜와 글자는 null', () => {
  assert.equal(C.parseDate('2026-02-30'), null);
  assert.equal(C.parseDate('배송중'), null);
  assert.equal(C.parseDate(''), null);
  assert.equal(C.parseDate('2026-13-01'), null);
});
t('parseAmount: 쉼표·원·음수·괄호', () => {
  assert.equal(C.parseAmount('12,900'), 12900);
  assert.equal(C.parseAmount('12,900원'), 12900);
  assert.equal(C.parseAmount('-3,000'), -3000);
  assert.equal(C.parseAmount('(3,000)'), -3000);
  assert.equal(C.parseAmount('₩ 1,000'), 1000);
  assert.equal(C.parseAmount('abc'), null);
  assert.equal(C.parseAmount(''), null);
});
t('normalizeOrderNo: 공백·하이픈 제거', () => {
  const r = C.normalizeOrderNo(' 2026-1001 23 ');
  assert.equal(r.value, '2026100123'); assert.equal(r.warning, null); assert.equal(r.raw, '2026-1001 23');
});
t('normalizeOrderNo: 지수표기 경고', () => {
  const r = C.normalizeOrderNo('1.23457E+15');
  assert.equal(r.warning, 'exponent'); assert.equal(r.value, '1234570000000000');
});
t('normalizeOrderNo: 17자리 이상·안전 범위 밖 숫자·.0·="..."', () => {
  assert.equal(C.normalizeOrderNo('12345678901234567').warning, 'long');
  assert.equal(C.normalizeOrderNo('1234567890123456').warning, null);
  assert.equal(C.normalizeOrderNo(12345678901234568).warning, 'precision');
  assert.equal(C.normalizeOrderNo('12345.0').value, '12345');
  assert.equal(C.normalizeOrderNo('="300000000012"').value, '300000000012');
  assert.equal(C.normalizeOrderNo("'300000000012").value, '300000000012');
});
t('columnIndex / columnName', () => {
  assert.equal(C.columnIndex('A1'), 0); assert.equal(C.columnIndex('Z9'), 25);
  assert.equal(C.columnIndex('AA10'), 26); assert.equal(C.columnIndex('AZ1'), 51);
  assert.equal(C.columnName(0), 'A'); assert.equal(C.columnName(26), 'AA'); assert.equal(C.columnName(51), 'AZ');
});

// ---------- header row and column mapping ----------

t('findHeaderRow: 제목줄·빈 줄 위에 있어도 머리글 찾기', () => {
  assert.equal(C.findHeaderRow(S.orderRows, 'orders'), 2);
  assert.equal(C.findHeaderRow(S.settlementRows, 'settlement'), 2);
  assert.equal(C.findHeaderRow([['주문번호', '배송완료일'], ['1', '2026-10-01']], 'orders'), 0);
});
t('tableFromRows: 빈 줄 버리고 머리글 길이 맞춤', () => {
  const tb = C.tableFromRows([['제목'], ['주문번호', '배송완료일'], ['1', '2026-10-01', '덤'], [], ['', ' ']], 'orders');
  assert.equal(tb.headerRowIndex, 1); assert.equal(tb.rows.length, 1); assert.equal(tb.headers.length, 3);
});
t('autoMapColumns: 주문 엑셀 일반 머리글', () => {
  const m = C.autoMapColumns(S.orderRows[2], 'orders');
  assert.deepEqual(m, { orderNo: 2, deliveredAt: 4, status: 5, product: 6, amount: 9, qty: 8 });
});
t('autoMapColumns: 묶음배송번호 함정', () => {
  const m = C.autoMapColumns(['묶음배송번호', '주문 번호', '배송완료일'], 'orders');
  assert.equal(m.orderNo, 1);
  const only = C.autoMapColumns(['묶음배송번호', '배송완료일'], 'orders');
  assert.equal(only.orderNo, -1);
  // 두 말이 한 제목에 같이 있어도 묶음배송 쪽은 고르지 않음
  assert.equal(C.autoMapColumns(['묶음배송 주문번호', '쿠팡 주문번호 (주문별)'], 'orders').orderNo, 1);
  assert.equal(C.autoMapColumns(['운송장 주문번호'], 'settlement').orderNo, -1);
});
t('autoMapColumns: 정확히 같은 이름 우선·예정일 제외·괄호 제거', () => {
  const m = C.autoMapColumns(['배송완료예정일', '배송완료일시', '결제액(원)', '노출상품명(옵션명)'], 'orders');
  assert.equal(m.deliveredAt, 1); assert.equal(m.amount, 2); assert.equal(m.product, 3);
  const m2 = C.autoMapColumns(['배송완료일시', '배송완료일'], 'orders');
  assert.equal(m2.deliveredAt, 1);
});
t('autoMapColumns: 정산 엑셀', () => {
  const m = C.autoMapColumns(S.settlementRows[2], 'settlement');
  assert.deepEqual(m, { orderNo: 1, amount: 4, recognizedAt: 0 });
  const m2 = C.autoMapColumns(['주문번호', '구매확정일', '매출금액'], 'settlement');
  assert.deepEqual(m2, { orderNo: 0, amount: 2, recognizedAt: 1 });
});

// ---------- findUnsettled ----------

t('findUnsettled: 가상 예시 데이터 5분류', () => {
  const { o, s } = sampleTables();
  const r = C.findUnsettled({ orders: o, settlement: s, asOf: S.asOf });
  const e = S.expected;
  assert.equal(r.counts.needsCheck, e.needsCheck); assert.equal(r.counts.waiting, e.waiting);
  assert.equal(r.counts.settled, e.settled); assert.equal(r.counts.excluded, e.excluded);
  assert.equal(r.counts.settlementOnly, e.settlementOnly); assert.equal(r.needsCheckAmount, e.needsCheckAmount);
  assert.deepEqual(r.needsCheck.map(x => x.orderNo), ['DEMO-0006', 'DEMO-0007', 'DEMO-0008']);
  assert.deepEqual(r.settlementOnly, ['DEMO-0099']);
  assert.deepEqual(r.settlementRange, { min: '2026-10-02', max: '2026-10-27' });
  assert.equal(r.warnings.length, 0);
});
t('findUnsettled: 경계일(정확히 7+3일째는 기다리는 중, 하루 지나면 확인 필요)', () => {
  const r = run([['A', '2026-10-19', '배송완료', 'p', '1'], ['B', '2026-10-18', '배송완료', 'p', '1']], []);
  assert.deepEqual(r.waiting.map(x => x.orderNo), ['A']); assert.equal(r.waiting[0].elapsedDays, 10);
  assert.deepEqual(r.needsCheck.map(x => x.orderNo), ['B']); assert.equal(r.needsCheck[0].elapsedDays, 11);
});
t('findUnsettled: 기다리는 날·여유일 바꾸기', () => {
  const rows = [['A', '2026-10-19', '배송완료', 'p', '1']];
  assert.equal(run(rows, [], { graceDays: 0 }).counts.needsCheck, 1);
  assert.equal(run(rows, [], { waitDays: 10, graceDays: 0 }).counts.waiting, 1);
  assert.equal(run(rows, [], { waitDays: 9, graceDays: 0 }).counts.needsCheck, 1);
});
t('findUnsettled: 취소·반품·교환은 오래돼도 제외', () => {
  const r = run([['A', '2026-09-01', '취소완료', 'p', '1'], ['B', '2026-09-01', '반품 접수', 'p', '1'], ['C', '2026-09-01', '교환완료', 'p', '1']], []);
  assert.equal(r.counts.needsCheck, 0); assert.equal(r.counts.excluded, 3);
  assert.ok(r.excluded.every(x => x.reason === '취소·반품·교환'));
});
t('findUnsettled: 여러 줄 주문은 한 주문으로 금액 합·"외 N"', () => {
  const r = run([['A', '2026-10-01', '배송완료', '상품1', '1,000'], ['A', '2026-10-02', '배송완료', '상품2', '2,500'], ['A', '2026-10-02', '배송완료', '상품3', '500']], []);
  assert.equal(r.counts.orders, 1); assert.equal(r.needsCheck.length, 1);
  const x = r.needsCheck[0];
  assert.equal(x.amount, 4000); assert.equal(x.product, '상품1 외 2'); assert.equal(x.lines, 3);
  assert.equal(x.deliveredAt, '2026-10-02'); assert.equal(x.elapsedDays, 27);
});
t('findUnsettled: 일부 취소된 주문은 남은 줄만 합산하고 메모', () => {
  const r = run([['A', '2026-10-01', '배송완료', '상품1', '1,000'], ['A', '', '취소완료', '상품2', '9,000']], []);
  assert.equal(r.needsCheck[0].amount, 1000); assert.match(r.needsCheck[0].note, /일부 상품 취소/);
});
t('findUnsettled: 배송완료일 없음·못 읽음은 제외, 경고', () => {
  const r = run([['A', '', '배송중', 'p', '1'], ['B', '어제쯤', '배송완료', 'p', '1']], []);
  assert.equal(r.counts.excluded, 2);
  assert.equal(r.excluded.find(x => x.orderNo === 'A').reason, '배송완료일 없음(미배송 등)');
  assert.equal(r.excluded.find(x => x.orderNo === 'B').reason, '배송완료일을 읽지 못함');
  assert.ok(r.warnings.some(w => w.code === 'badDate' && w.count === 1));
});
t('findUnsettled: 정산 엑셀에 있으면 오래돼도 정산 확인됨(하이픈 차이 무시)', () => {
  const r = run([['2026-0001', '2026-09-01', '배송완료', 'p', '1']], [['20260001', '2026-09-08']]);
  assert.equal(r.counts.settled, 1); assert.equal(r.counts.needsCheck, 0); assert.equal(r.counts.settlementOnly, 0);
});
t('findUnsettled: 정산 엑셀에만 있는 주문 수', () => {
  const r = run([['A', '2026-10-01', '배송완료', 'p', '1']], [['A', '2026-10-08'], ['X', '2026-10-08'], ['Y', '2026-10-09'], ['X', '2026-10-10']]);
  assert.equal(r.counts.settlementOnly, 2); assert.deepEqual(r.settlementOnly, ['X', 'Y']);
});
t('findUnsettled: 예상 구매확정일이 정산 엑셀 날짜 범위 밖이면 메모', () => {
  const r = run([['A', '2026-10-01', '배송완료', 'p', '1'], ['B', '2026-08-20', '배송완료', 'p', '1'], ['C', '2026-09-10', '배송완료', 'p', '1']], [['Z', '2026-09-05'], ['Q', '2026-10-05']]);
  const a = r.needsCheck.find(x => x.orderNo === 'A'), b = r.needsCheck.find(x => x.orderNo === 'B');
  assert.match(a.note, /날짜 범위 밖/);
  assert.match(b.note, /날짜 범위 밖/);
  assert.equal(r.needsCheck.find(x => x.orderNo === 'C').note, '');
  const r2 = run([['A', '2026-10-01', '배송완료', 'p', '1']], [['Z', '2026-10-01'], ['Q', '2026-10-20']]);
  assert.equal(r2.needsCheck[0].note, '');
});
t('findUnsettled: 지수표기 주문번호 경고가 결과에 남음', () => {
  const r = run([['1.23457E+15', '2026-10-01', '배송완료', 'p', '1']], []);
  assert.ok(r.warnings.some(w => w.code === 'exponent' && w.count === 1));
});
t('findUnsettled: 상태 열이 없으면 경고, 금액 열이 없으면 합계 null', () => {
  const r = C.findUnsettled({ orders: { rows: [['A', '2026-10-01']], map: { orderNo: 0, deliveredAt: 1 } }, settlement: { rows: [], map: { orderNo: 0 } }, asOf: '2026-10-29' });
  assert.ok(r.warnings.some(w => w.code === 'noStatus'));
  assert.equal(r.needsCheckAmount, null); assert.equal(r.counts.needsCheck, 1);
});
t('findUnsettled: 필수 열·기준일·날짜 수 오류', () => {
  assert.throws(() => C.findUnsettled({ orders: { rows: [], map: { orderNo: 0, deliveredAt: -1 } }, settlement: { rows: [], map: { orderNo: 0 } }, asOf: '2026-10-29' }), e => e.code === 'MAP_ORDERS');
  assert.throws(() => C.findUnsettled({ orders: { rows: [], map: { orderNo: 0, deliveredAt: 1 } }, settlement: { rows: [], map: { orderNo: -1 } }, asOf: '2026-10-29' }), e => e.code === 'MAP_SETTLEMENT');
  assert.throws(() => run([], [], { asOf: '언젠가' }), e => e.code === 'BAD_ASOF');
  assert.throws(() => run([], [], { graceDays: -1 }), e => e.code === 'BAD_DAYS');
  assert.throws(() => run([], [], { waitDays: 1.5 }), e => e.code === 'BAD_DAYS');
});

// ---------- CSV export ----------

t('toCSV: BOM·머리글·긴 주문번호 보존·따옴표·수식 방지', () => {
  const csv = C.toCSV({ needsCheck: [
    { orderNo: '30000000001234', deliveredAt: '2026-10-01', elapsedDays: 28, product: '가상 "큰" 상품, 1개', amount: 12900, note: '' },
    { orderNo: 'DEMO-1', deliveredAt: '2026-10-02', elapsedDays: 27, product: '=HYPERLINK("x")', amount: null, note: '메모' }
  ] });
  assert.equal(csv.charCodeAt(0), 0xfeff);
  const lines = csv.slice(1).split('\r\n');
  assert.equal(lines[0], '주문번호,배송완료일,경과일,상품명,주문 엑셀 기준 금액(정산 예정액 아님),메모');
  assert.equal(lines[1], '="30000000001234",2026-10-01,28,"가상 ""큰"" 상품, 1개",12900,');
  assert.equal(lines[2], 'DEMO-1,2026-10-02,27,"\'=HYPERLINK(""x"")",,메모');
  assert.equal(lines.length, 4);
});

// ---------- XLSX ----------

async function checkXlsxRead(method) {
  const zip = await buildZip(xlsxFiles(method));
  const x = await C.readXlsx(zip.buffer);
  assert.deepEqual(x.sheets, ['주문목록', '요약']); assert.equal(x.sheetName, '주문목록');
  assert.deepEqual(x.rows[0], ['주문 목록 (테스트)']);
  assert.deepEqual(x.rows[1], []);                                     // 빈 줄 유지
  assert.deepEqual(x.rows[2], ['주문번호', '배송완료일', '주문상태', '상품 & 옵션', '결제액', '확인']); // 서식 run 합치기, rPh 빼기
  assert.deepEqual(x.rows[3], ['3000000000123', '46296', '배송완료', '', '12900', 'TRUE']);   // 빈 칸(D4) 유지
  assert.deepEqual(x.rows[4], ['3000000000124', '2026-10-25', '배송완료', '줄\r바꿈', '', 'FALSE']);
}
t('readXlsx: 무압축(store) ZIP - 공유문자열·inlineStr·숫자·빈 칸', () => checkXlsxRead(0));
t('readXlsx: deflate 압축 ZIP - 같은 결과', () => checkXlsxRead(8));
t('readXlsx: 두 번째 시트(x: 접두사, 절대경로 대상)', async () => {
  const zip = await buildZip(xlsxFiles(8));
  const x = await C.readXlsx(zip, { sheetIndex: 1 });
  assert.equal(x.sheetName, '요약'); assert.deepEqual(x.rows, [['요약', '', '7']]);
});
t('loadTable: xlsx → 머리글 찾기 → 비교까지', async () => {
  const zip = await buildZip(xlsxFiles(8));
  const o = await C.loadTable(zip, '주문.xlsx', 'orders');
  assert.equal(o.format, 'xlsx'); assert.equal(o.headerRowIndex, 2); assert.equal(o.rows.length, 2);
  o.map = C.autoMapColumns(o.headers, 'orders');
  assert.equal(o.map.orderNo, 0); assert.equal(o.map.deliveredAt, 1); assert.equal(o.map.status, 2);
  const r = C.findUnsettled({ orders: o, settlement: { rows: [['3000000000124']], map: { orderNo: 0 } }, asOf: '2026-10-29' });
  assert.deepEqual(r.needsCheck.map(x => x.orderNo), ['3000000000123']);
  assert.equal(r.needsCheck[0].deliveredAt, '2026-10-01'); assert.equal(r.counts.settled, 1);
});
t('loadTable: 옛 xls·암호 걸린 파일(OLE)은 쉬운 오류', async () => {
  const ole = new Uint8Array([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0, 0, 0, 0]);
  await assert.rejects(C.loadTable(ole, 'old.xls', 'orders'), e => e.code === 'OLE' && /\.xlsx/.test(e.message));
});
t('loadTable: 깨진 ZIP·빈 파일 오류', async () => {
  await assert.rejects(C.loadTable(new Uint8Array([0x50, 0x4B, 0x03, 0x04, 1, 2, 3]), 'bad.xlsx', 'orders'), e => e.code === 'ZIP_BAD');
  await assert.rejects(C.loadTable(new Uint8Array(0), 'empty.csv', 'orders'), e => e.code === 'EMPTY');
  await assert.rejects(C.loadTable(enc.encode('\n\n , \n'), 'blank.csv', 'orders'), e => e.code === 'EMPTY');
});

// ---------- page checks ----------

const html = readFileSync(join(here, 'index.html'), 'utf8');
t('index.html: 외부 스크립트는 방문 통계 1개뿐, 스타일·폰트 없음', () => {
  const srcs = [...html.matchAll(/<script\b[^>]*\bsrc="([^"]*)"/g)].map(m => m[1]);
  assert.deepEqual(srcs, ['core.js', 'samples.js', 'app.js', 'https://static.cloudflareinsights.com/beacon.min.js']);
  const csp = (html.match(/http-equiv="Content-Security-Policy" content="([^"]*)"/) || [])[1] || '';
  assert.ok(csp.includes("connect-src https://cloudflareinsights.com;"), 'connect-src는 방문 통계 주소만');
  assert.ok(csp.includes("script-src 'self' https://static.cloudflareinsights.com;"), 'script-src 제한');
  assert.ok(html.includes('고른 파일의 내용은 들어가지 않습니다'), '방문 통계 고지');
  assert.equal(/<link\b[^>]*rel="stylesheet"/i.test(html), false);
  assert.equal(/@import|url\(\s*['"]?https?:/i.test(html), false);
});
t('index.html: 필수 고지 문구', () => {
  for (const s of ['서버로 보내지 않습니다', '브라우저 안에서만', '쿠팡과 관련 없는 개인 프로젝트', 'AI 도구의 도움', '결과는 참고용', '반드시 쿠팡 WING에서 확인',
    'https://github.com/kimheesoo-inventor/settle-check-kr', '가상 예시 데이터로 해 보기', 'lang="ko"', 'width=device-width']) {
    assert.ok(html.includes(s), '없음: ' + s);
  }
});
t('index.html: 효과 약속·과장 문구 없음', () => {
  const text = html.replace(/<style[\s\S]*?<\/style>/g, '').replace(/<script[\s\S]*?<\/script>/g, '');
  for (const s of ['찾아 드립니다', '찾아드립니다', '돌려받', '보장', '100%', '손실 방지', '나도 모르는']) assert.equal(text.includes(s), false, '있음: ' + s);
});
t('index.html: 요소 id가 app.js와 맞음', () => {
  const app = readFileSync(join(here, 'app.js'), 'utf8');
  const ids = new Set([...app.matchAll(/\$\('([A-Za-z]+)'\)/g)].map(m => m[1]));
  for (const id of ids) assert.ok(html.includes('id="' + id + '"') || /Info$|Mapping$/.test(id), 'id 없음: ' + id);
  for (const id of ['ordersInfo', 'settlementInfo', 'ordersMapping', 'settlementMapping']) assert.ok(html.includes('id="' + id + '"'), id);
});

// ---------- run ----------

for (const [name, fn] of queue) {
  try { await fn(); passed++; console.log('ok   ' + name); }
  catch (e) { failed++; console.log('FAIL ' + name + '\n     ' + (e && e.message)); }
}
console.log(`\n${passed} passed, ${failed} failed (${queue.length} tests)`);
process.exitCode = failed ? 1 : 0;
