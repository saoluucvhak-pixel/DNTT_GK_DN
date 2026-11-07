/**
 * XUẤT FILE UP_UNC (Ủy nhiệm chi) THEO MẪU UNC_NGANHANG.xlsx ĐỂ UP LÊN NGÂN HÀNG
 * - Sheet UP_UNC : mỗi hồ sơ ở bảng 112 = 1 dòng chuyển khoản (header 3 dòng, dữ liệu từ dòng 4 như mẫu)
 * - Sheet TH_UNC : mỗi "Lần" = 1 lô (ID_UNC)
 */
const UNC_CFG = {
  TK_TRICH_NO: "5622289731",              // Tài khoản trích nợ (theo file mẫu) - sửa nếu đổi tài khoản
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

/**
 * @param {string[]} selectedIds  ID hệ thống (cột A sheet 112) của các hồ sơ cần xuất
 * @return {{success:boolean,url?:string,xlsxUrl?:string,message?:string,warnings?:string[],count?:number,total?:number}}
 */
function createUpUncFile(selectedIds) {
  let newSS = null;
  try {
    const ids = new Set((selectedIds || []).map(x => String(x).trim()));
    if (!ids.size) return { success: false, message: "Chưa có hồ sơ nào để xuất UNC." };

    const sh = SpreadsheetApp.getActive().getSheetByName(CFG.DNTT_112);
    if (!sh || sh.getLastRow() < 2) return { success: false, message: "Sheet 112 trống." };
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, 22).getValues();

    // Đọc lại từ server (không tin số liệu do trình duyệt gửi lên)
    const rows = data.filter(r => ids.has(String(r[0]).trim()));
    if (!rows.length) return { success: false, message: "Không tìm thấy hồ sơ trong sheet 112." };

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
        String(body.length + 1), UNC_CFG.PHUONG_THUC, UNC_CFG.TK_TRICH_NO, ten, stk, bank, tien, UNC_CFG.LOAI_TIEN,
        nd, "", "", "", "", UNC_CFG.BEN_CHIU_PHI, UNC_CFG.TK_TRICH_NO, effDate, batchId.get(lanOf(r)), newId()
      ]);
    });
    if (!body.length) return { success: false, message: "Không có dòng hợp lệ để xuất.", warnings: warnings };

    const fileName = "UP_UNC (" + Utilities.formatDate(today, CFG.TZ, "dd-MM-yyyy HHmm") + ")";
    newSS = SpreadsheetApp.create(fileName);
    DriveApp.getFileById(newSS.getId()).moveTo(DriveApp.getFolderById(CFG.FOLDER_REPORT_ID));

    // --- UP_UNC ---
    const sh1 = newSS.getSheets()[0];
    sh1.setName("UP_UNC");
    const nCol = UNC_CFG.HEADERS.length;
    sh1.getRange(1, 1, 3, nCol).setNumberFormat("@").setValues([UNC_CFG.HEADERS, UNC_CFG.HEADERS, UNC_CFG.HEADERS]);
    sh1.getRange(1, 1, 3, nCol).setFontWeight("bold");
    const textCols = [1, 2, 3, 4, 5, 6, 8, 9, 14, 15, 17, 18];
    textCols.forEach(c => sh1.getRange(4, c, body.length, 1).setNumberFormat("@"));
    sh1.getRange(4, 1, body.length, nCol).setValues(body);
    sh1.getRange(4, 7, body.length, 1).setNumberFormat("#,##0");
    sh1.getRange(4, 16, body.length, 1).setNumberFormat('d"/"m"/"yyyy');
    sh1.setFrozenRows(3);
    sh1.autoResizeColumns(1, nCol);

    // --- TH_UNC ---
    const sh2 = newSS.insertSheet("TH_UNC");
    sh2.getRange(1, 1, 1, UNC_CFG.TH_HEADERS.length).setValues([UNC_CFG.TH_HEADERS]).setFontWeight("bold");
    const th = [];
    batchId.forEach((id, lan) => {
      th.push([id, effDate, UNC_CFG.PHUONG_THUC, UNC_CFG.TK_TRICH_NO, UNC_CFG.LOAI_TIEN, UNC_CFG.TK_TRICH_NO, utils.parseNum(lan), UNC_CFG.BEN_CHIU_PHI]);
    });
    [1, 3, 4, 6].forEach(c => sh2.getRange(2, c, th.length, 1).setNumberFormat("@"));
    sh2.getRange(2, 2, th.length, 1).setNumberFormat('d"/"m"/"yyyy');
    sh2.getRange(2, 1, th.length, 8).setValues(th);
    sh2.autoResizeColumns(1, 8);

    SpreadsheetApp.flush();
    const id = newSS.getId();
    return {
      success: true,
      url: newSS.getUrl(),
      xlsxUrl: "https://docs.google.com/spreadsheets/d/" + id + "/export?format=xlsx",
      count: body.length,
      total: total,
      warnings: warnings
    };
  } catch (e) {
    if (newSS) { try { DriveApp.getFileById(newSS.getId()).setTrashed(true); } catch (x) {} }
    return { success: false, message: e.toString() };
  }
}
