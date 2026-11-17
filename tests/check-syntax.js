// Kiểm tra cú pháp các file Apps Script (.gs) và script trong Index.html
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
let failed = 0;

function check(name, code) {
  try {
    new Function(code);
    console.log('ok   ' + name);
  } catch (e) {
    failed++;
    console.error('FAIL ' + name + ': ' + e.message);
  }
}

for (const f of fs.readdirSync(root).filter(f => f.endsWith('.gs'))) {
  check(f, fs.readFileSync(path.join(root, f), 'utf8'));
}

const html = fs.readFileSync(path.join(root, 'Index.html'), 'utf8');
const m = html.match(/<script>([\s\S]*?)<\/script>/);
if (!m) { failed++; console.error('FAIL Index.html: không tìm thấy <script>'); }
else check('Index.html (script)', m[1]);

// Mọi hàm gọi từ onclick phải tồn tại
const handlers = [...new Set([...html.matchAll(/onclick="([A-Za-z0-9_]+)\(/g)].map(x => x[1]))];
for (const h of handlers) {
  if (!new RegExp('function\\s+' + h + '\\s*\\(').test(html)) { failed++; console.error('FAIL onclick gọi hàm không tồn tại: ' + h); }
}
console.log('onclick handlers: ' + handlers.length);

process.exit(failed ? 1 : 0);
