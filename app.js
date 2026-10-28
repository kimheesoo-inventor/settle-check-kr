/* 정산 확인기 - page logic. Uses SettleCheckCore (core.js) and SettleCheckSamples (samples.js).
 * No network calls. Files are read in the browser. All text is inserted with textContent. MIT License. */
(function () {
  'use strict';
  var C = window.SettleCheckCore;
  var S = window.SettleCheckSamples;

  var KINDS = ['orders', 'settlement'];
  var TITLE = { orders: '주문 엑셀', settlement: '정산 엑셀' };
  var state = { orders: null, settlement: null, files: {}, result: null };

  function $(id) { return document.getElementById(id); }

  function el(tag, attrs, children) {
    var e = document.createElement(tag);
    if (attrs) Object.keys(attrs).forEach(function (k) {
      if (k === 'text') e.textContent = attrs[k];
      else if (k === 'class') e.className = attrs[k];
      else e.setAttribute(k, attrs[k]);
    });
    (children || []).forEach(function (c) {
      if (c == null) return;
      e.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
    });
    return e;
  }

  function clear(node) { while (node.firstChild) node.removeChild(node.firstChild); }

  function won(n) { return n == null ? '-' : Math.round(n).toLocaleString('ko-KR') + '원'; }

  function todayISO() {
    var d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }

  function friendly(e) {
    if (e && e.code && C.MESSAGES[e.code]) return C.MESSAGES[e.code];
    return '파일을 읽는 중 문제가 생겼습니다. 엑셀에서 열어 .xlsx 또는 CSV로 다시 저장한 뒤 넣어 주세요.';
  }

  function hideResults() { state.result = null; $('resultsSection').hidden = true; }

  function readArrayBuffer(file) {
    if (file.arrayBuffer) return file.arrayBuffer();
    return new Promise(function (resolve, reject) {
      var r = new FileReader();
      r.onload = function () { resolve(r.result); };
      r.onerror = function () { reject(r.error); };
      r.readAsArrayBuffer(file);
    });
  }

  function formatLabel(t) {
    if (t.format === 'sample') return '가상 예시 데이터';
    if (t.format === 'xlsx') return '엑셀(.xlsx)' + (t.sheetName ? ', 시트 "' + t.sheetName + '"' : '');
    if (t.format === 'html') return 'HTML 표';
    return 'CSV' + (t.encoding === 'euc-kr' ? ' (EUC-KR 한글)' : '');
  }

  function renderInfo(kind) {
    var box = $(kind + 'Info');
    clear(box);
    var t = state[kind];
    if (!t) return;
    var line = el('p', null, [t.name + ': ' + formatLabel(t) + ' · 데이터 ' + t.rows.length + '줄 · 열 제목은 ' + (t.headerRowIndex + 1) + '번째 줄']);
    if (t.isSample) line.appendChild(el('span', { class: 'badge', text: ' 가상 데이터' }));
    box.appendChild(line);
    if (t.sheets && t.sheets.length > 1 && !t.isSample) {
      var id = kind + 'Sheet';
      var sel = el('select', { id: id, class: 'sheet-select' });
      t.sheets.forEach(function (name, i) {
        var o = el('option', { value: String(i), text: (i + 1) + '. ' + name });
        if (i === t.sheetIndex) o.selected = true;
        sel.appendChild(o);
      });
      sel.addEventListener('change', function () { loadFile(kind, parseInt(sel.value, 10)); });
      box.appendChild(el('label', { for: id, text: '시트가 ' + t.sheets.length + '개 있습니다. 읽을 시트:' }));
      box.appendChild(sel);
    }
  }

  function setTable(kind, t) {
    t.map = C.autoMapColumns(t.headers, kind);
    state[kind] = t;
    hideResults();
    renderInfo(kind);
    renderMappings();
  }

  function loadFile(kind, sheetIndex) {
    var f = state.files[kind];
    if (!f) return;
    var box = $(kind + 'Info');
    clear(box);
    box.appendChild(el('p', { text: '읽는 중…' }));
    readArrayBuffer(f).then(function (buf) {
      return C.loadTable(buf, f.name, kind, { sheetIndex: sheetIndex });
    }).then(function (t) {
      setTable(kind, t);
    }).catch(function (e) {
      state[kind] = null;
      hideResults();
      clear(box);
      box.appendChild(el('p', { class: 'error', text: f.name + ': ' + friendly(e) }));
      renderMappings();
    });
  }

  function onFile(kind, input) {
    var f = input.files && input.files[0];
    if (!f) return;
    state.files[kind] = f;
    loadFile(kind, 0);
  }

  function columnLabel(h, i) { return C.columnName(i) + '열: ' + (h || '(제목 없음)'); }

  function renderMapping(kind) {
    var box = $(kind + 'Mapping');
    clear(box);
    var t = state[kind];
    box.appendChild(el('h3', { text: TITLE[kind] }));
    if (!t) { box.appendChild(el('p', { class: 'muted', text: '아직 파일을 고르지 않았습니다.' })); return; }
    C.FIELDS[kind].forEach(function (f) {
      var id = 'map-' + kind + '-' + f.key;
      var sel = el('select', { id: id });
      sel.appendChild(el('option', { value: '-1', text: '(없음)' }));
      t.headers.forEach(function (h, i) {
        var o = el('option', { value: String(i), text: columnLabel(h, i) });
        if (t.map[f.key] === i) o.selected = true;
        sel.appendChild(o);
      });
      sel.addEventListener('change', function () {
        t.map[f.key] = parseInt(sel.value, 10);
        hideResults();
        renderProblems();
      });
      var tag = t.map[f.key] >= 0 ? el('span', { class: 'tag', text: '자동 선택' }) : el('span', { class: 'tag', text: '못 찾음' });
      var label = el('label', { for: id }, [f.label, f.required ? el('span', { class: 'req', text: ' *' }) : null, tag]);
      box.appendChild(el('div', { class: 'map-row' }, [label, sel]));
    });
  }

  function renderProblems() {
    var box = $('mappingProblems');
    clear(box);
    var probs = [];
    KINDS.forEach(function (kind) {
      var t = state[kind];
      if (!t) { probs.push(TITLE[kind] + '을 골라 주세요.'); return; }
      C.FIELDS[kind].forEach(function (f) {
        if (f.required && !(t.map[f.key] >= 0)) probs.push(TITLE[kind] + ': "' + f.label + '" 열을 골라 주세요.');
      });
    });
    var o = state.orders, s = state.settlement;
    if (o && s && o.map.orderNo >= 0 && s.map.orderNo >= 0 && o.isSample !== s.isSample) {
      box.appendChild(el('p', { class: 'warn', text: '한쪽은 가상 예시 데이터, 다른 쪽은 실제 파일입니다. 두 파일을 같은 종류로 맞춰 주세요.' }));
    }
    probs.forEach(function (p) { box.appendChild(el('p', { class: 'error', text: p })); });
    $('runBtn').disabled = probs.length > 0;
    return probs;
  }

  function renderMappings() {
    $('mappingSection').hidden = false;
    renderMapping('orders');
    renderMapping('settlement');
    renderProblems();
  }

  function summaryBox(cls, label, num, sub) {
    return el('div', { class: 'box' + (cls ? ' ' + cls : '') }, [
      el('div', { class: 'lbl', text: label }),
      el('div', { class: 'num', text: num }),
      sub ? el('div', { class: 'sub', text: sub }) : null
    ]);
  }

  function cardTable(headers, rows, numCols) {
    var thead = el('thead', null, [el('tr', null, headers.map(function (h, i) {
      return el('th', { class: numCols.indexOf(i) >= 0 ? 'num' : '', scope: 'col', text: h });
    }))]);
    var tbody = el('tbody', null, rows.map(function (r) {
      return el('tr', null, r.map(function (v, i) {
        return el('td', { class: numCols.indexOf(i) >= 0 ? 'num' : '', 'data-label': headers[i], text: v == null ? '' : String(v) });
      }));
    }));
    return el('div', { class: 'table-wrap' }, [el('table', { class: 'cards' }, [thead, tbody])]);
  }

  function listDetails(title, rows, headers, numCols) {
    if (!rows.length) return null;
    return el('details', null, [el('summary', { text: title + ' (' + rows.length + '건)' }), cardTable(headers, rows, numCols)]);
  }

  function renderResults(r) {
    var box = $('results');
    clear(box);
    var sample = state.orders && state.orders.isSample;
    if (sample) box.appendChild(el('p', { class: 'warn', text: '지금 보는 결과는 가상 예시 데이터입니다. 실제 판매자·주문이 아닙니다.' }));

    box.appendChild(el('p', { text: '기준일 ' + r.asOf + ' 기준으로, 배송완료일 + 기다리는 날 ' + r.waitDays + '일 + 여유일 ' + r.graceDays +
      '일(합계 ' + r.threshold + '일)이 지났는데 정산 엑셀에 주문번호가 없는 주문을 "확인 필요"로 표시했습니다.' }));
    if (r.settlementRange) box.appendChild(el('p', { class: 'muted', text: '정산 엑셀 날짜 범위: ' + r.settlementRange.min + ' ~ ' + r.settlementRange.max }));

    box.appendChild(el('div', { class: 'summary' }, [
      summaryBox('main', '확인 필요', r.counts.needsCheck + '건', '정산 엑셀에서 못 찾음(누락 확정 아님) · 주문 엑셀 기준 금액 합 ' + won(r.needsCheckAmount) + ' (정산 예정액 아님)'),
      summaryBox('', '기다리는 중', r.counts.waiting + '건', '배송완료 후 아직 기간 안'),
      summaryBox('', '정산 확인됨', r.counts.settled + '건', '정산 엑셀에 주문번호 있음'),
      summaryBox('', '제외', r.counts.excluded + '건', '취소·반품·교환, 배송완료일 없음'),
      summaryBox('', '정산 엑셀에만 있음', r.counts.settlementOnly + '건', '다른 기간 주문일 수 있음')
    ]));

    r.warnings.forEach(function (w) { box.appendChild(el('p', { class: 'warn', text: w.message + (w.code === 'noStatus' ? '' : ' (' + w.count + '줄)') })); });

    box.appendChild(el('h3', { text: '확인 필요 목록', style: 'margin-top:18px' }));
    if (r.needsCheck.length) {
      box.appendChild(el('p', { class: 'muted', text: '정산 엑셀에서 찾지 못한 주문입니다. 정산이 안 됐다는 뜻은 아닙니다. 구매확정이 늦어졌거나 정산이 보류된 경우, 정산 엑셀 기간이 모자란 경우도 여기에 나옵니다. WING에서 주문번호로 먼저 확인해 주세요.' }));
      box.appendChild(cardTable(['주문번호', '배송완료일', '경과일', '상품명', '주문 엑셀 금액', '메모'],
        r.needsCheck.map(function (x) { return [x.orderNo, x.deliveredAt, x.elapsedDays + '일', x.product || '-', won(x.amount), x.note || '']; }), [2, 4]));
      $('downloadBtn').hidden = false;
    } else {
      box.appendChild(el('p', { class: 'ok', text: '기간이 지났는데 정산 엑셀에 없는 주문을 찾지 못했습니다.' }));
      $('downloadBtn').hidden = true;
    }

    var more = [
      listDetails('기다리는 중', r.waiting.map(function (x) { return [x.orderNo, x.deliveredAt, x.elapsedDays + '일', x.product || '-']; }), ['주문번호', '배송완료일', '경과일', '상품명'], [2]),
      listDetails('제외', r.excluded.map(function (x) { return [x.orderNo, x.reason, x.status || '-', x.product || '-']; }), ['주문번호', '이유', '상태', '상품명'], []),
      listDetails('정산 확인됨', r.settled.map(function (x) { return [x.orderNo, x.deliveredAt || '-', x.product || '-']; }), ['주문번호', '배송완료일', '상품명'], []),
      listDetails('정산 엑셀에만 있는 주문', r.settlementOnly.slice(0, 200).map(function (no) { return [no]; }), ['주문번호'], [])
    ];
    more.forEach(function (d) { if (d) box.appendChild(d); });
    $('resultsSection').hidden = false;
  }

  function run() {
    if (renderProblems().length) return;
    var box = $('mappingProblems');
    try {
      state.result = C.findUnsettled({
        orders: state.orders, settlement: state.settlement,
        asOf: $('asOf').value,
        waitDays: $('waitDays').value === '' ? null : Number($('waitDays').value),
        graceDays: $('graceDays').value === '' ? null : Number($('graceDays').value)
      });
      renderResults(state.result);
      $('step3').scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (e) {
      box.appendChild(el('p', { class: 'error', text: (e && e.code && C.MESSAGES[e.code]) ? C.MESSAGES[e.code] : '비교 중 문제가 생겼습니다. 열 선택을 다시 확인해 주세요.' }));
    }
  }

  function download() {
    if (!state.result) return;
    var blob = new Blob([C.toCSV(state.result)], { type: 'text/csv;charset=utf-8' });
    var url = URL.createObjectURL(blob);
    var a = el('a', { href: url, download: '정산확인필요_' + state.result.asOf + '.csv' });
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }

  function loadSample() {
    $('ordersFile').value = ''; $('settlementFile').value = '';
    state.files = {};
    var o = C.tableFromRows(S.orderRows, 'orders');
    o.name = S.orderFileName; o.format = 'sample'; o.isSample = true;
    var s = C.tableFromRows(S.settlementRows, 'settlement');
    s.name = S.settlementFileName; s.format = 'sample'; s.isSample = true;
    $('asOf').value = S.asOf;
    $('waitDays').value = '7'; $('graceDays').value = '3';
    setTable('orders', o);
    setTable('settlement', s);
    var note = $('mappingProblems');
    note.insertBefore(el('p', { class: 'muted', text: '가상 데이터에 맞춰 기준일을 ' + S.asOf + '로 바꿨습니다. 실제 파일을 넣을 때는 기준일을 오늘로 바꿔 주세요.' }), note.firstChild);
  }

  $('asOf').value = todayISO();
  $('ordersFile').addEventListener('change', function () { onFile('orders', this); });
  $('settlementFile').addEventListener('change', function () { onFile('settlement', this); });
  $('loadSample').addEventListener('click', loadSample);
  $('runBtn').addEventListener('click', run);
  $('downloadBtn').addEventListener('click', download);
})();
