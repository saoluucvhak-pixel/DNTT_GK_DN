// Kiểm thử đơn vị các hàm thuần (không cần Google Apps Script)
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
const read = f => fs.readFileSync(path.join(root, f), 'utf8');

// Chỉ lấy phần trước doGet của Code.gs (CFG, utils, sysLock) + Banks.gs + Unc.gs
const src = read('Code.gs').split('function doGet')[0] + read('Banks.gs') + read('Unc.gs') +
  ';return { CFG, utils, findBankFull, _cleanAccount, getUncDefaults };';
const stubs = { Utilities: { formatDate: () => '' }, LockService: {} };
const { CFG, utils, findBankFull, _cleanAccount, getUncDefaults } =
  new Function(...Object.keys(stubs), src)(...Object.values(stubs));

test('parseNum: số kiểu Việt Nam và kiểu US', () => {
  assert.strictEqual(utils.parseNum('1.234.567'), 1234567);
  assert.strictEqual(utils.parseNum('1.234,5'), 1234.5);
  assert.strictEqual(utils.parseNum('1,234,567.89'), 1234567.89);
  assert.strictEqual(utils.parseNum('-2.500,75'), -2500.75);
  assert.strictEqual(utils.parseNum('12.5'), 12.5);
  assert.strictEqual(utils.parseNum(1500), 1500);
  assert.strictEqual(utils.parseNum('abc'), 0);
  assert.strictEqual(utils.parseNum(''), 0);
});

test('standardize: bỏ khoảng trắng và viết hoa', () => {
  assert.strictEqual(utils.standardize(' ab c '), 'ABC');
  assert.strictEqual(utils.standardize(null), '');
});

test('findBankFull: không phân biệt hoa thường, dấu, viết tắt', () => {
  const tail = s => findBankFull(s).split(' - ').pop();
  assert.strictEqual(tail('vietcombank'), 'Vietcombank');
  assert.strictEqual(tail('VIETCOMBANK'), 'Vietcombank');
  assert.strictEqual(tail('VcB'), 'Vietcombank');
  assert.strictEqual(tail('AGRIBANK'), 'Agribank');
  assert.strictEqual(tail('ngân hàng nông nghiệp'), 'Agribank');
  assert.strictEqual(tail('NGÂN HÀNG KỸ THƯƠNG'), 'Techcombank');
  assert.strictEqual(tail('Mb bank'), 'MB');
  assert.strictEqual(tail('Techcombank - CN Đà Nẵng'), 'Techcombank');
  assert.strictEqual(findBankFull('01203001 - Ngan hang TMCP Ngoai thuong Viet Nam - Vietcombank').split(' - ').pop(), 'Vietcombank');
  assert.strictEqual(findBankFull('không có ngân hàng này'), '');
  assert.strictEqual(findBankFull(''), '');
});

test('_cleanAccount: làm sạch, mặc định và báo lỗi', () => {
  assert.strictEqual(_cleanAccount(' 1111 222333 ', 'd', 'x'), '1111222333');
  assert.strictEqual(_cleanAccount('', 'd', 'x'), 'd');
  assert.strictEqual(_cleanAccount(null, 'd', 'x'), 'd');
  assert.throws(() => _cleanAccount('abc', 'd', 'x'));
  assert.throws(() => _cleanAccount('123', 'd', 'x'));
});

test('cấu hình tài khoản: có trích nợ và thu phí', () => {
  assert.match(CFG.TK_TRICH_NO, /^\d{6,25}$/);
  assert.match(CFG.TK_THU_PHI, /^\d{6,25}$/);
  assert.deepStrictEqual(getUncDefaults(), { tkTrichNo: CFG.TK_TRICH_NO, tkThuPhi: CFG.TK_THU_PHI });
});
