// Kiểm thử giao diện Index.html bằng trình duyệt headless (giả lập google.script.run)
const assert = require('assert');
const path = require('path');
const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage();
  const errors = [];
  const dialogs = [];
  page.on('pageerror', e => errors.push(e.message));
  page.on('dialog', async d => { dialogs.push(d.message()); await d.accept(); });

  await page.addInitScript(() => {
    const rows = [{ idHeThong: 'A1', ngayDN: '07/11/2026', ngayISO: '2026-11-07', soLan: '1', chuRung: '<i id="xss">X</i>', nguoiNhan: 'Y', stk: '1', nganHang: 'VCB', klTan: 1, soTien: 1000, noiDungCK: 'n', ghiChu: '' }];
    window.__calls = [];
    function runner() {
      let ok, px;
      const o = { withSuccessHandler(f) { ok = f; return px; }, withFailureHandler() { return px; } };
      px = new Proxy(o, { get(t, k) {
        if (k in t) return t[k];
        return (...a) => { window.__calls.push([k, a]); setTimeout(() => {
          if (k === 'get112ViewData' || k === 'getMainTableData') ok(rows);
          else if (k === 'getUncDefaults') ok({ tkTrichNo: '8619299999', tkThuPhi: '8619299999' });
          else if (k === 'createUpUncFile') ok({ success: true, fileName: 'UP_UNC.xlsx', fileUrl: 'https://drive.google.com/file/d/x/view', downloadUrl: 'https://drive.google.com/uc?export=download&id=x', base64: 'UEsDBA==', count: 1, total: 1000, warnings: ['w1'] });
          else if (k === 'updateBankFile') ok({ success: true, fileName: 'B.xlsx', fileUrl: 'u', downloadUrl: 'd', base64: 'UEsDBA==', count: 2 });
          else ok({ success: true });
        }, 10); };
      } });
      return px;
    }
    window.google = { script: { get run() { return runner(); } } };
  });

  await page.goto('file://' + path.join(__dirname, '..', 'Index.html'));
  await page.waitForTimeout(300);

  // XSS: dữ liệu có thẻ HTML phải hiển thị như chữ, không được thành thẻ thật
  await page.click('text=2. LẬP ĐỀ NGHỊ');
  await page.waitForTimeout(300);
  assert.strictEqual(await page.locator('#xss').count(), 0, 'dữ liệu từ sheet phải được escape');

  // Các nút theo đúng thứ tự trong ảnh tham khảo
  const labels = await page.locator('.btn-hak').allInnerTexts();
  const joined = labels.join('|');
  for (const l of ['3. XUẤT BÁO CÁO', '4. FILE NGÂN HÀNG', '5. TẠO UNC', '6. CHỐT & KHÓA']) assert(joined.includes(l), 'thiếu nút ' + l);

  // Tạo UNC: form -> xác nhận -> link tải
  await page.click('text=5. TẠO UNC');
  await page.waitForTimeout(300);
  assert(await page.isVisible('#uncForm'));
  assert.strictEqual(await page.inputValue('#uncTkTrichNo'), '8619299999');
  await page.fill('#uncTkThuPhi', '5620727934');
  await page.click('text=Xác nhận tạo UNC');
  await page.waitForTimeout(1200);
  assert(await page.isVisible('#uncResult'));
  const html = await page.innerHTML('#uncResult');
  assert(html.includes('drive.google.com/uc?export=download'), 'phải có link tải');
  const call = await page.evaluate(() => window.__calls.find(c => c[0] === 'createUpUncFile'));
  assert.deepStrictEqual(call[1][1], { tkTrichNo: '8619299999', tkThuPhi: '5620727934' });

  // File ngân hàng
  await page.click('text=4. FILE NGÂN HÀNG');
  await page.waitForTimeout(1200);
  assert((await page.innerText('#uncResult')).includes('B.xlsx'));

  assert.deepStrictEqual(errors, [], 'không được có lỗi JavaScript: ' + errors.join('; '));
  assert(dialogs.some(d => d.includes('Đã tạo file')), 'phải có hộp thoại báo kết quả');
  console.log('UI tests passed');
  await browser.close();
})().catch(e => { console.error(e); process.exit(1); });
