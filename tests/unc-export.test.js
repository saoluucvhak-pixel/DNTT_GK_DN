// Kiểm thử createUpUncFile với Google Apps Script giả lập
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const read = f => fs.readFileSync(path.join(__dirname, '..', f), 'utf8');
const src = read('Code.gs').split('function doGet')[0] + read('Banks.gs') + read('Unc.gs') +
  ';return { createUpUncFile, UNC_CFG };';

function run(ghiId) {
  const mk = (lan, ten, bank, stk, tien) => ['I' + ten, '', 'x', ten, bank, stk, tien, 'TT', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, lan];
  const rows = [mk(1, 'A', 'Vietcombank', '111111', 1000), mk(1, 'B', 'BIDV', '222222', 2000), mk(2, 'C', 'Agribank', '333333', 3000)];
  const cap = { up: null, upClear: null, th: null };
  const rng = (n, a) => ({
    setNumberFormat() { return this; },
    setValues(v) { if (n === 'UP' && a[0] === 4) cap.up = v; if (n === 'TH') cap.th = v; return this; },
    clearContent() { if (n === 'UP') cap.upClear = a; return this; },
    getValues() { return rows.map(r => { r = r.slice(); while (r.length < 22) r.push(''); return r; }); }
  });
  const sheet = n => ({ getLastRow: () => 6, getMaxRows: () => 1000, insertRowsAfter() {}, getRange: (...a) => rng(n, a) });
  const S = { UP_UNC: sheet('UP'), TH_UNC: sheet('TH'), DNTT_GK_DN_112: sheet('112') };
  const g = {
    SpreadsheetApp: { getActive: () => ({ getSheetByName: n => S[n] }), openById: () => ({ getSheetByName: n => S[n] }), flush() {} },
    DriveApp: { getFolderById: () => ({ createFile: () => ({ getUrl: () => 'u', getId: () => 'i' }) }) },
    Utilities: { getUuid: () => require('crypto').randomUUID(), formatDate: (d, tz, f) => ({ yyyy: '2026', MM: '11', dd: '17' }[f] || 'x'), base64Encode: () => 'B' },
    UrlFetchApp: { fetch: () => ({ getResponseCode: () => 200, getBlob: () => ({ setName() { return this; }, getBytes: () => [1] }) }) },
    ScriptApp: { getOAuthToken: () => 't' },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) }
  };
  const m = new Function(...Object.keys(g), src)(...Object.values(g));
  if (ghiId !== undefined) m.UNC_CFG.GHI_ID_UNC = ghiId;
  const res = m.createUpUncFile(['IA', 'IB', 'IC'], { tkTrichNo: '8619299999', tkThuPhi: '5620727934' });
  return { res, cap };
}

test('mặc định: UP_UNC chỉ gồm 16 cột (A-P), không có ID_UNC / ID_UNCCT', () => {
  const { res, cap } = run();
  assert.strictEqual(res.success, true);
  assert.strictEqual(cap.up.length, 3);
  cap.up.forEach(r => assert.strictEqual(r.length, 16));
  assert.strictEqual(cap.upClear[3], 18, 'phải xóa cả 2 cột ID của lần xuất trước');
});

test('các cột theo mẫu vẫn đúng', () => {
  const { cap } = run();
  const byName = Object.fromEntries(cap.up.map(r => [r[3], r]));
  assert.strictEqual(byName.A[1], 'O - Chuyển ngoài BIDV');
  assert.strictEqual(byName.B[1], 'I - Chuyển trong BIDV');
  assert.strictEqual(byName.A[2], '8619299999');
  assert.strictEqual(byName.A[14], '5620727934');
  assert.strictEqual(byName.A[6], 1000);
});

test('TH_UNC vẫn được tạo (tách lô theo Lần + phương thức)', () => {
  const { cap } = run();
  assert.strictEqual(cap.th.length, 3);
});

test('bật GHI_ID_UNC: có thêm 2 cột ID', () => {
  const { cap } = run(true);
  cap.up.forEach(r => { assert.strictEqual(r.length, 18); assert.ok(r[16] && r[17]); });
});
