/**
 * XUẤT FILE UP_UNC (Ủy nhiệm chi) THEO MẪU UNC_NGANHANG.xlsx ĐỂ UP LÊN NGÂN HÀNG
 * - Sheet UP_UNC : mỗi hồ sơ ở bảng 112 = 1 dòng chuyển khoản (header 3 dòng, dữ liệu từ dòng 4 như mẫu)
 * - Sheet TH_UNC : mỗi "Lần" = 1 lô (ID_UNC)
 */
const UNC_CFG = {
  SS_ID: "1v6MlQaMF4N8BoTUqaraInxJFPUKcA7u3z_zVle8cXw8", // Sheet mẫu UNC_NGANHANG_HAK (UP_UNC, TH_UNC, DM_TKNH)
  PHUONG_THUC: "O - Chuyển ngoài BIDV",
  LOAI_TIEN: "VND",
  BEN_CHIU_PHI: "O - Bên chuyển",
  HEADERS: ["STT", "PhuongThucChuyenTien", "TkTrichNo", "TenNguoiHuong", "SoTKNguoiHuong", "NganHangHuong", "SoTien", "LoaiTien", "NoiDung", "A", "B", "C", "D", "BenChiuPhi", "TaiKhoanThuPhi", "NgayHieuLuc", "ID_UNC", "ID_UNCCT"],
  TH_HEADERS: ["ID_UNC", "Ngày", "PhuongThucChuyenTien", "TaiKhoanTrichNo", "LoaiTien", "TKThuPhi", "LanTT", "BenChiuPhi"]
};

// Tên viết tắt/cách gọi thường gặp -> tên ngắn trong danh mục
const BANK_ALIASES = [
  ["vcb", "Vietcombank"], ["ngoai thuong", "Vietcombank"],
  ["agri", "Agribank"], ["nong nghiep", "Agribank"],
  ["ctg", "Vietinbank"], ["cong thuong", "Vietinbank"],
  ["dau tu va phat trien", "BIDV"],
  ["tcb", "Techcombank"], ["ky thuong", "Techcombank"],
  ["quan doi", "MB"], ["mbbank", "MB"],
  ["stb", "Sacombank"], ["sai gon thuong tin", "Sacombank"],
  ["a chau", "ACB"], ["viet nam thinh vuong", "VPBank"],
  ["tien phong", "TPBank"], ["xuat nhap khau", "Eximbank"], ["phuong dong", "OCB"]
];

function _uncNorm(s) {
  return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/đ/gi, "d")
    .toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}
function _uncHasWord(text, tok) {
  if (!tok) return false;
  return (" " + text + " ").indexOf(" " + tok + " ") >= 0;
}

let _bankIndex = null;
function _getBankIndex() {
  if (_bankIndex) return _bankIndex;
  _bankIndex = BANK_LIST.map(b => {
    const full = b[0], shortC = b[1];
    const lastSeg = full.split(" - ").pop();
    const toks = [...new Set([_uncNorm(shortC), _uncNorm(lastSeg)].filter(t => t.length >= 2))];
    return { full: full, code: full.split(" - ")[0].trim(), toks: toks, short: _uncNorm(shortC) };
  });
  return _bankIndex;
}

/** Trả về tên đầy đủ trong danh mục ngân hàng, hoặc "" nếu không nhận diện được */
function findBankFull(input) {
  const raw = String(input || "").trim();
  if (!raw) return "";
  const idx = _getBankIndex();
  const m = raw.match(/^0?\d{6,8}\b/);
  if (m) {
    const code = m[0].replace(/^0+/, "");
    const hit = idx.find(b => b.code.replace(/^0+/, "") === code);
    if (hit) return hit.full;
  }
  const n = _uncNorm(raw);
  let best = null, bestLen = 0;
  idx.forEach(b => b.toks.forEach(t => {
    if (_uncHasWord(n, t) && t.length > bestLen) { best = b; bestLen = t.length; }
  }));
  if (best) return best.full;
  for (const a of BANK_ALIASES) {
    if (_uncHasWord(n, a[0]) || (a[0].indexOf(" ") > 0 && n.indexOf(a[0]) >= 0)) {
      const hit = idx.find(b => b.short === _uncNorm(a[1]));
      if (hit) return hit.full;
    }
  }
  return "";
}

/** Giá trị mặc định (lấy từ CFG trong Code.gs) để điền sẵn vào form tạo UNC */
function getUncDefaults() {
  return { tkTrichNo: CFG.TK_TRICH_NO, tkThuPhi: CFG.TK_THU_PHI };
}

function _cleanAccount(v, def, label) {
  const s = String(v == null ? "" : v).replace(/[\s'.\-]/g, "");
  if (!s) return def;
  if (!/^\d{6,25}$/.test(s)) throw new Error(label + " không hợp lệ (chỉ gồm 6-25 chữ số): " + v);
  return s;
}

/**
 * @param {string[]} selectedIds  ID hệ thống (cột A sheet 112) của các hồ sơ cần xuất
 * Ghi dữ liệu vào sheet mẫu UNC_NGANHANG_HAK (xóa dữ liệu cũ từ dòng 4 của UP_UNC và dòng 2 của TH_UNC),
 * xuất toàn bộ file ra .xlsx, lưu vào thư mục báo cáo và trả về base64 để trình duyệt tải về ngay.
 * @return {{success:boolean,fileName?:string,fileUrl?:string,base64?:string,message?:string,warnings?:string[],count?:number,total?:number}}
 */
function createUpUncFile(selectedIds, opts) {
  const lock = sysLock.acquire(); // sheet mẫu dùng chung, tránh 2 người ghi cùng lúc
  try {
    const tkTrichNo = _cleanAccount(opts && opts.tkTrichNo, CFG.TK_TRICH_NO, "Tài khoản trích nợ");
    const tkThuPhi = _cleanAccount(opts && opts.tkThuPhi, CFG.TK_THU_PHI, "Tài khoản thu phí");
    const ids = new Set((selectedIds || []).map(x => String(x).trim()));
    if (!ids.size) return { success: false, message: "Chưa có hồ sơ nào để xuất UNC." };

    const sh = SpreadsheetApp.getActive().getSheetByName(CFG.DNTT_112);
    if (!sh || sh.getLastRow() < 2) return { success: false, message: "Sheet 112 trống." };
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, 22).getValues();

    // Đọc lại từ server (không tin số liệu do trình duyệt gửi lên)
    const rows = data.filter(r => ids.has(String(r[0]).trim()));
    if (!rows.length) return { success: false, message: "Không tìm thấy hồ sơ trong sheet 112." };

    const tmpl = SpreadsheetApp.openById(UNC_CFG.SS_ID);

    const warnings = [];
    const today = new Date();
    const effDate = new Date(Number(Utilities.formatDate(today, CFG.TZ, "yyyy")), Number(Utilities.formatDate(today, CFG.TZ, "MM")) - 1, Number(Utilities.formatDate(today, CFG.TZ, "dd")));

    // Mỗi "Lần" (cột S) = 1 lô UNC
    const batchId = new Map();
    const newId = () => Utilities.getUuid().replace(/-/g, "").substring(0, 8);
    const lanOf = r => String(r[18] || "").trim() || "1";
    rows.sort((a, b) => (utils.parseNum(lanOf(a)) - utils.parseNum(lanOf(b))));
    rows.forEach(r => { const l = lanOf(r); if (!batchId.has(l)) batchId.set(l, newId()); });

    const body = [];
    let total = 0, unmatchedBank = 0;
    rows.forEach(r => {
      const ten = String(r[3] || "").trim().toUpperCase();
      const stk = String(r[5] || "").replace(/[\s'.\-]/g, "");
      const tien = utils.parseNum(r[6]);
      const nd = String(r[7] || "").trim();
      const label = ten || String(r[0]);
      if (!stk) { warnings.push("Thiếu số tài khoản: " + label); return; }
      if (!(tien > 0)) { warnings.push("Số tiền không hợp lệ: " + label); return; }
      const bank = findBankFull(r[4]);
      if (!bank) { unmatchedBank++; warnings.push("Không nhận diện được ngân hàng \"" + (r[4] || "") + "\": " + label + " (để trống cột NganHangHuong)"); }
      if (bank.indexOf("BIDV") >= 0) warnings.push("Người nhận tại BIDV: " + label + " - mẫu đang để \"" + UNC_CFG.PHUONG_THUC + "\", hãy kiểm tra phương thức chuyển.");
      total += tien;
      body.push([
        String(body.length + 1), UNC_CFG.PHUONG_THUC, tkTrichNo, ten, stk, bank, tien, UNC_CFG.LOAI_TIEN,
        nd, "", "", "", "", UNC_CFG.BEN_CHIU_PHI, tkThuPhi, effDate, batchId.get(lanOf(r)), newId()
      ]);
    });
    if (!body.length) return { success: false, message: "Không có dòng hợp lệ để xuất.", warnings: warnings };

    const fileName = "UP_UNC (" + Utilities.formatDate(today, CFG.TZ, "dd-MM-yyyy HHmm") + ")";
    const sh1 = tmpl.getSheetByName("UP_UNC");
    const sh2 = tmpl.getSheetByName("TH_UNC");
    if (!sh1 || !sh2) throw new Error("Sheet mẫu thiếu UP_UNC hoặc TH_UNC.");
    const ensureRows = (sh, need) => { const mr = sh.getMaxRows(); if (mr < need) sh.insertRowsAfter(mr, need - mr); };

    // --- UP_UNC: giữ 3 dòng tiêu đề của mẫu, ghi dữ liệu từ dòng 4 ---
    const nCol = UNC_CFG.HEADERS.length;
    if (sh1.getLastRow() >= 4) sh1.getRange(4, 1, sh1.getLastRow() - 3, nCol).clearContent();
    ensureRows(sh1, 3 + body.length);
    const textCols = [1, 2, 3, 4, 5, 6, 8, 9, 14, 15, 17, 18];
    textCols.forEach(c => sh1.getRange(4, c, body.length, 1).setNumberFormat("@"));
    sh1.getRange(4, 1, body.length, nCol).setValues(body);
    sh1.getRange(4, 7, body.length, 1).setNumberFormat("#,##0");
    sh1.getRange(4, 16, body.length, 1).setNumberFormat('d"/"m"/"yyyy');

    // --- TH_UNC: mỗi Lần = 1 lô ---
    const th = [];
    batchId.forEach((id, lan) => {
      th.push([id, effDate, UNC_CFG.PHUONG_THUC, tkTrichNo, UNC_CFG.LOAI_TIEN, tkThuPhi, utils.parseNum(lan), UNC_CFG.BEN_CHIU_PHI]);
    });
    if (sh2.getLastRow() >= 2) sh2.getRange(2, 1, sh2.getLastRow() - 1, 8).clearContent();
    ensureRows(sh2, 1 + th.length);
    [1, 3, 4, 6].forEach(c => sh2.getRange(2, c, th.length, 1).setNumberFormat("@"));
    sh2.getRange(2, 2, th.length, 1).setNumberFormat('d"/"m"/"yyyy');
    sh2.getRange(2, 1, th.length, 8).setValues(th);

    SpreadsheetApp.flush();

    // Xuất toàn bộ file mẫu (UP_UNC, TH_UNC, DM_TKNH) sang .xlsx
    const resp = UrlFetchApp.fetch("https://docs.google.com/spreadsheets/d/" + UNC_CFG.SS_ID + "/export?format=xlsx", {
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) throw new Error("Không xuất được file Excel (HTTP " + resp.getResponseCode() + ").");
    const xlsxName = fileName + ".xlsx";
    const blob = resp.getBlob().setName(xlsxName);
    const xlsxFile = DriveApp.getFolderById(CFG.FOLDER_REPORT_ID).createFile(blob);

    return {
      success: true,
      fileName: xlsxName,
      fileUrl: xlsxFile.getUrl(),
      base64: Utilities.base64Encode(blob.getBytes()),
      count: body.length,
      total: total,
      warnings: warnings
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally { lock.releaseLock(); }
}
