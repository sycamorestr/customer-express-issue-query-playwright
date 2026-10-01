'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '..', 'skills',
  'customer-express-issue-query-playwright', 'scripts', 'query.js'), 'utf8');
const sensitive = 'PRIVATE_TOKEN_OR_RESPONSE_BODY';
const envelope = rows => ({ IsSuccess: true, ReturnValue: { datas: rows } });
const response = body => ({ status: 200, text: async () => typeof body === 'string' ? body : JSON.stringify(body) });

function harness(responses = [], overrides = {}) {
  const calls = [];
  const document = { getElementById: id => id === '__VIEWSTATE' ? { value: sensitive } : null };
  const window = {
    location: { pathname: '/app/order/order/list.aspx', origin: 'https://www.erp321.com' },
    fetch: async (url, options) => {
      const callback = JSON.parse(new URLSearchParams(options.body).get('__CALLBACKPARAM'));
      calls.push({ url, options, callback });
      assert.ok(responses.length > 0, 'Unexpected request');
      const next = responses.shift();
      return typeof next === 'function' ? next(url, options) : next;
    }
  };
  const query = vm.runInNewContext(source, { window, document, URLSearchParams,
    AbortController, setTimeout, clearTimeout, ...overrides });
  return { calls, run: async input => JSON.parse(JSON.stringify(await query(input))) };
}

test('uses only the read-only logistics callback and preserves all four fields and order', async () => {
  const rows = [
    { o_id: 1, l_id: '@abc123', confirm_date: ' 2026-10-01\n  08:00 ', paid_amount: 0, remark: ' 商品A\n  辅料 B ' },
    { o_id: 2, plat_l_id: 'ABC123', confirm_date: null, paid_amount: null, remark: null }
  ];
  const { run, calls } = harness([response(envelope([...rows, rows[0]]))]);
  const result = await run({ nos: ['ABC123', '@XYZ456', 'MISS123'] });
  assert.equal(result.ok, true);
  assert.deepEqual(result.results, [
    { no: 'ABC123', matched: 2, data: [['ABC123', '2026-10-01 08:00', '0', '商品A 辅料 B'], ['ABC123', '', '', '']] },
    { no: '@XYZ456', matched: 0, data: [['@XYZ456', '', '', '']] },
    { no: 'MISS123', matched: 0, data: [['MISS123', '', '', '']] }
  ]);
  assert.deepEqual(result.logs, [{ requested: 3, rows: 3, matched: 1, fallback: 0 }]);
  assert.equal(calls.length, 1);
  assert.equal(calls[0].options.method, 'POST');
  assert.equal(calls[0].options.credentials, 'same-origin');
  assert.equal(calls[0].options.mode, 'same-origin');
  assert.equal(calls[0].options.redirect, 'error');
  assert.equal(calls[0].callback.Method, 'LoadDataToJSON');
  assert.deepEqual(JSON.parse(calls[0].callback.Args[1]), [
    { k: 'l_id', v: 'ABC123,@ABC123,@XYZ456,XYZ456,MISS123,@MISS123', c: '@=' }
  ]);
  assert.equal(new URLSearchParams(calls[0].options.body).get('_jt_page_size'), '500');
  assert.ok(calls[0].url.startsWith('https://www.erp321.com/app/order/order/list.aspx?'));
  assert.equal(JSON.stringify(result).includes(sensitive), false);
});

test('keeps literal pipes in JSON remarks and accepts numeric callback prefixes', async () => {
  for (const prefix of ['', '0|', '123|', ' \n123|']) {
    const payload = { IsSuccess: true, ReturnValue: JSON.stringify({ data: [
      { o_id: 1, l_id: 'ABC123', remark: '商品A | 商品B' }
    ] }) };
    const { run } = harness([response(prefix + JSON.stringify(payload))]);
    const result = await run({ nos: ['ABC123'] });
    assert.equal(result.ok, true, prefix);
    assert.equal(result.results[0].data[0][3], '商品A | 商品B');
  }
});

test('does not strip arbitrary prefixes or treat invalid responses as missing orders', async () => {
  for (const body of [sensitive + '|' + JSON.stringify(envelope([])), '<html>' + sensitive + '</html>',
    { IsSuccess: false, Message: sensitive }, { IsSuccess: true, ReturnValue: sensitive },
    { IsSuccess: true, ReturnValue: { datas: [null] } }, { IsSuccess: true, ReturnValue: {} }]) {
    const { run } = harness([response(body)]);
    const result = await run({ nos: ['ABC123'] });
    assert.equal(result.ok, false);
    assert.equal('results' in result, false);
    assert.equal(JSON.stringify(result).includes(sensitive), false);
  }
});

test('matches every supported logistics field with separators and optional @', async () => {
  const fields = ['l_id', 'plat_l_id', 'purchase_lid', 'multiWaybillLid',
    'logistics_no', 'express_no', 'waybillNo', 'waybill_no', 'lc_id'];
  const rows = fields.map((field, index) => ({ o_id: index, [field]: 'OTHER，@abc123、EXTRA;MORE|LAST', paid_amount: index }));
  const { run } = harness([response(envelope(rows))]);
  const result = await run({ nos: ['ABC123'] });
  assert.equal(result.ok, true);
  assert.equal(result.results[0].matched, fields.length);
  assert.deepEqual(result.results[0].data.map(row => row[2]), fields.map((_, index) => String(index)));
  assert.equal(result.logs[0].fallback, 0);
});

test('fallback has alphanumeric boundaries, escapes regexp characters, and records use', async () => {
  const rows = [
    { o_id: 1, remark: '查询 @ABC123 正常' },
    { o_id: 2, remark: 'long XABC123 or ABC1239 must not match' },
    { o_id: 3, remark: 'Exact AB.123 only' },
    { o_id: 4, remark: 'Pattern ABX123 must not match' }
  ];
  const { run } = harness([response(envelope(rows))]);
  const result = await run({ nos: ['ABC123', 'AB.123'] });
  assert.equal(result.ok, true);
  assert.deepEqual(result.results.map(row => row.matched), [1, 1]);
  assert.equal(result.logs[0].fallback, 2);
});

test('fallback is only used on rows without a logistics-field match', async () => {
  const { run } = harness([response(envelope([{ o_id: 1, l_id: 'ABC123', remark: 'OTHER123' }]))]);
  const result = await run({ nos: ['ABC123', 'OTHER123'] });
  assert.deepEqual(result.results.map(row => row.matched), [1, 0]);
  assert.equal(result.logs[0].fallback, 0);
});

test('500-row limit causes sequential subdivision and capped rows are discarded', async () => {
  const { run, calls } = harness([
    response(envelope(Array.from({ length: 500 }, (_, index) => ({ o_id: index, l_id: 'ABC123', remark: 'capped' })))),
    response(envelope([{ o_id: 1, l_id: 'ABC123', remark: 'complete A' }])),
    response(envelope([{ o_id: 2, l_id: '@XYZ456', remark: 'complete B' }]))
  ]);
  const result = await run({ nos: ['ABC123', 'XYZ456'] });
  assert.equal(result.ok, true);
  assert.deepEqual(calls.map(call => JSON.parse(call.callback.Args[1])[0].v),
    ['ABC123,@ABC123,XYZ456,@XYZ456', 'ABC123,@ABC123', 'XYZ456,@XYZ456']);
  assert.deepEqual(result.results.map(row => row.data[0][3]), ['complete A', 'complete B']);
  assert.equal(result.logs[0].split, true);
  assert.equal(result.logs.length, 3);
});

test('failure after a successful subdivision exposes no partial results', async () => {
  const { run } = harness([
    response(envelope(Array.from({ length: 500 }, () => ({ l_id: 'ABC123' })))),
    response(envelope([{ o_id: 1, l_id: 'ABC123', remark: sensitive }])),
    { status: 403, text: async () => sensitive }
  ]);
  const result = await run({ nos: ['ABC123', 'XYZ456'] });
  assert.equal(result.ok, false);
  assert.match(result.reason, /HTTP 403/);
  assert.equal('results' in result, false);
  assert.equal(JSON.stringify(result).includes(sensitive), false);
  assert.equal(result.logs.length, 2);
});

test('a single number at the row cap cannot be reported complete', async () => {
  const { run, calls } = harness([response(envelope(Array.from({ length: 500 }, () => ({ l_id: 'ABC123' }))))]);
  const result = await run({ nos: ['ABC123'] });
  assert.equal(result.ok, false);
  assert.match(result.reason, /cannot certify completeness/);
  assert.equal('results' in result, false);
  assert.equal(calls.length, 1);
});

test('aborts a stalled fetch at the configured timeout', async () => {
  let aborted = false;
  const { run } = harness([(_, options) => new Promise((resolve, reject) => {
    options.signal.addEventListener('abort', () => {
      aborted = true;
      reject(new Error(sensitive));
    });
  })]);
  const result = await run({ nos: ['ABC123'], timeoutMs: 10 });
  assert.equal(aborted, true);
  assert.equal(result.ok, false);
  assert.match(result.reason, /timed out/);
  assert.equal('results' in result, false);
  assert.equal(JSON.stringify(result).includes(sensitive), false);
});

test('timeout also covers stalled response bodies', async () => {
  const { run } = harness([(_, options) => ({ status: 200,
    text: () => new Promise((resolve, reject) => options.signal.addEventListener('abort', () => reject(new Error(sensitive))))
  })]);
  const result = await run({ nos: ['ABC123'], timeoutMs: 10 });
  assert.equal(result.ok, false);
  assert.match(result.reason, /timed out/);
});

test('unexpected network errors are sanitized', async () => {
  const { run } = harness([() => { throw new Error(sensitive); }]);
  const result = await run({ nos: ['ABC123'] });
  assert.equal(result.ok, false);
  assert.equal('results' in result, false);
  assert.equal(JSON.stringify(result).includes(sensitive), false);
});

test('invalid batches, timeouts, and frames fail before making a request', async () => {
  for (const input of [{ nos: [] }, { nos: [''] }, { nos: ['@'] }, { nos: [42] },
    { nos: Array(61).fill('ABC123') }, { nos: ['ABC123'], timeoutMs: 0 },
    { nos: ['ABC123'], timeoutMs: NaN }, { nos: ['ABC123'], timeoutMs: 300001 }]) {
    const { run, calls } = harness();
    assert.equal((await run(input)).ok, false);
    assert.equal(calls.length, 0);
  }
  for (const overrides of [
    { document: { getElementById: () => null } },
    { window: { location: { pathname: '/login' } } }
  ]) {
    const { run, calls } = harness([], overrides);
    assert.equal((await run({ nos: ['ABC123'] })).ok, false);
    assert.equal(calls.length, 0);
  }
});
