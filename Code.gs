/** * HỆ THỐNG QUẢN LÝ THANH TOÁN HAK - PHIÊN BẢN TỐI ƯU HOÀN HẢO 2026
 */

const CFG = {
  PC_SS_ID: "1vqMVxccBA7zlAMHrGsVBydGFwZJ6QuDZW10zJ74V29g",
  HD_SS_ID: "1cv11ORWuAF3Sit4f-kA0xrP6-ab4SF-7LEdkCvGi_gI",
  UPDATE_NH_SS_ID: "1id7m092q2OVY4dJgcavzPM5HNO14_QIUburVN7y7szg",
  FOLDER_REPORT_ID: "13_ed2jCP8CgL-UJk0eOn-vIH6LqQElxc",
  DNTT_SRC: "DNTT_GK_DN",
  DNTT_CT: "DNTT_GK_DN_CT",
  DNTT_112: "DNTT_GK_DN_112",
  PC_SHEET: "PhieuCan_DN",
  SPLIT_COL: "Số phiếu cân",
  LAN_TT: "Lanthanhtoan",
  TZ: "Asia/Ho_Chi_Minh",
  TK_TRICH_NO: "8619299999", // Tài khoản trích nợ (dùng cho file ngân hàng và UNC)
  TK_THU_PHI: "8619299999"   // Tài khoản thu phí (UNC) - sửa nếu khác tài khoản trích nợ
};


const utils = {
  standardize: (v) => v ? String(v).replace(/\s+/g, '').toUpperCase().trim() : "",
  formatDate: (d) => (d instanceof Date && d.getFullYear() > 1900) ? Utilities.formatDate(d, CFG.TZ, "dd/MM/yyyy") : (d ? String(d) : ""),
  // Hiểu cả số kiểu VN (1.234.567,89) lẫn kiểu US (1,234,567.89)
  parseNum: (v) => {
    if (typeof v === 'number') return isNaN(v) ? 0 : v;
    if (!v) return 0;
    let str = String(v).replace(/[^0-9.,-]+/g, "");
    const neg = str.charAt(0) === "-";
    str = str.replace(/-/g, "");
    const lastDot = str.lastIndexOf("."), lastComma = str.lastIndexOf(",");
    const cntDot = (str.match(/\./g) || []).length, cntComma = (str.match(/,/g) || []).length;
    let decSep = "";
    if (lastDot >= 0 && lastComma >= 0) decSep = lastDot > lastComma ? "." : ",";
    else if (cntDot > 1 || cntComma > 1) decSep = "";            // nhiều dấu giống nhau => ngăn cách nghìn
    else if (lastDot >= 0) decSep = (str.length - lastDot - 1 === 3) ? "" : ".";
    else if (lastComma >= 0) decSep = (str.length - lastComma - 1 === 3) ? "" : ",";
    if (decSep) {
      const idx = str.lastIndexOf(decSep);
      str = str.slice(0, idx).replace(/[.,]/g, "") + "." + str.slice(idx + 1);
    } else {
      str = str.replace(/[.,]/g, "");
    }
    let n = parseFloat(str);
    if (isNaN(n)) return 0;
    return neg ? -n : n;
  }
};

const sysLock = {
  acquire: () => {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(30000); 
      return lock;
    } catch (e) {
      throw new Error("Hệ thống đang bận do có người khác đang chốt dữ liệu. Vui lòng thử lại sau 1 phút.");
    }
  }
};

function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    let action = e.parameter.action;
    let result = "";
    try {
      if (action === "tach_phieu") result = runProcessDetail();
      else if (action === "lap_de_nghi") result = runCreate112();
      return ContentService.createTextOutput(JSON.stringify({status: "success", message: result})).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({status: "error", message: err.toString()})).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return HtmlService.createTemplateFromFile('Index').evaluate()
      .setTitle("QUẢN LÝ THANH TOÁN HAK")
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- DATA VIEW FOR WEB ---
function get112ViewData(fDate, tDate) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.DNTT_112);
  if (!sh || sh.getLastRow() < 2) return [];
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 22).getValues();
  const filtered = data.filter(r => {
    let d = r[16]; // Cột Q
    if (!(d instanceof Date)) return false;
    let iso = Utilities.formatDate(d, CFG.TZ, "yyyy-MM-dd");
    return (!fDate || iso >= fDate) && (!tDate || iso <= tDate);
  });
  return filtered.map(r => ({
    idHeThong: String(r[0] || ""),
    ngayDN: utils.formatDate(r[16]),
    ngayISO: Utilities.formatDate(r[16], CFG.TZ, "yyyy-MM-dd"),
    soLan: String(r[18] || ""),
    chuRung: String(r[2] || ""),
    nguoiNhan: String(r[3] || ""),
    stk: String(r[5] || ""),
    nganHang: String(r[4] || ""),
    klTan: utils.parseNum(r[11]) / 1000,
    soTien: utils.parseNum(r[6]),
    noiDungCK: String(r[7] || ""),
    ghiChu: String(r[15] || "")
  })).reverse();
}

function getMainTableData(fDate, tDate) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.DNTT_SRC);
  if (!sh || sh.getLastRow() < 2) return [];
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 18).getValues();
  const filtered = data.filter(r => {
    let d = r[11]; // Cột L
    if (!(d instanceof Date)) return false;
    let iso = Utilities.formatDate(d, CFG.TZ, "yyyy-MM-dd");
    return (!fDate || iso >= fDate) && (!tDate || iso <= tDate);
  });
  return filtered.map(r => ({
    idHeThong: String(r[0] || ""),
    ngayDN: utils.formatDate(r[11]),
    ngayISO: Utilities.formatDate(r[11], CFG.TZ, "yyyy-MM-dd"),
    nguoiDN: String(r[5] || ""),
    chuRung: String(r[3] || ""),
    nguoiNhan: String(r[7] || ""),
    stk: String(r[8] || ""),
    kl: utils.parseNum(r[9]).toFixed(2),
    soHD: String(r[13] || ""),
    lapDNTT: String(r[14] || ""),
    trangThaiTT: String(r[17] || "")
  })).reverse();
}

// ============================================================
// NÚT 6: TẠO BÁO CÁO & SHEET NGÂN HÀNG
// ============================================================
function createFinalReportFromFilteredData(filteredRows, dateRange) {
  let newSS = null;
  try {
    const folder = DriveApp.getFolderById(CFG.FOLDER_REPORT_ID);
    const fileName = "BÁO CÁO ĐNTT (" + dateRange + ") - " + Utilities.formatDate(new Date(), CFG.TZ, "HHmm");
    newSS = SpreadsheetApp.create(fileName);
    DriveApp.getFileById(newSS.getId()).moveTo(folder);

    const sheet1 = newSS.getSheets()[0];
    sheet1.setName("DeNghiThanhToanCK");
    renderSheet1Full(sheet1, filteredRows, dateRange);

    const sheet2 = newSS.insertSheet("BangKeChiTietCK");
    renderSheet2Detail(sheet2, filteredRows, dateRange);

    return { success: true, url: newSS.getUrl() };
  } catch (e) {
    if (newSS) { try { DriveApp.getFileById(newSS.getId()).setTrashed(true); } catch (x) {} }
    return { success: false, message: e.toString() };
  }
}

// ============================================================
// SHEET 1: TỐI ƯU GIAO DIỆN IN ẤN
// ============================================================
function renderSheet1Full(sheet, filteredRows, dateRange) {
  sheet.clear();
  sheet.getRange("A1").setValue("CÔNG TY TNHH HOÀNG ANH KHÔI").setFontSize(10).setFontWeight("bold");
  sheet.getRange("A2:K2").merge().setValue("BẢNG ĐỀ XUẤT THANH TOÁN CHUYỂN KHOẢN CÔNG NỢ")
        .setFontSize(15).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("A3:K3").merge().setValue(dateRange).setFontStyle("italic").setHorizontalAlignment("center");

  const headers = ["STT", "Ngày đề nghị", "Lần", "Họ tên Chủ rừng", "Người nhận tiền", "Số tài khoản", "Tên ngân hàng", "KL (Tấn)", "Số tiền", "Nội dung chuyển khoản", "Ghi chú"];
  sheet.getRange(4, 1, 1, 11).setValues([headers]).setBackground("#d9ead3").setFontWeight("bold").setBorder(true, true, true, true, true, true).setHorizontalAlignment("center");

  const body = filteredRows.map((r, i) => [
    i + 1, r.ngayDN, r.soLan, r.chuRung, r.nguoiNhan, "'" + r.stk, r.nganHang, 
    utils.parseNum(r.klTan), utils.parseNum(r.soTien), r.noiDungCK, r.ghiChu
  ]);

  if (body.length > 0) {
    sheet.getRange(5, 1, body.length, 11).setValues(body).setBorder(true, true, true, true, true, true).setVerticalAlignment("middle");
    
    // ĐỊNH DẠNG CỘT
    sheet.setColumnWidth(1, 30);  // STT cực nhỏ
    sheet.setColumnWidth(11, 150); // Ghi chú thu gọn
    sheet.getRange(5, 10, body.length, 2).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); // Tự động xuống dòng
    
    sheet.getRange(5, 8, body.length, 1).setNumberFormat("#,##0.00");
    sheet.getRange(5, 9, body.length, 1).setNumberFormat("#,##0");
    
    const lastR = 5 + body.length;
    sheet.getRange(lastR, 1, 1, 7).merge().setValue("TỔNG CỘNG").setFontWeight("bold").setHorizontalAlignment("right");
    sheet.getRange(lastR, 8, 1, 2).setFormulaR1C1(`=SUM(R5C:R[-1]C)`).setFontWeight("bold");
    sheet.getRange(lastR, 1, 1, 11).setBackground("#FFF2CC").setBorder(true, true, true, true, true, true);
    
    sheet.autoResizeRows(5, body.length); // Mở rộng hàng theo chiều dọc
  }
}

// ============================================================
// SHEET 2: CHI TIẾT VÀ ĐỊNH DẠNG TIME/NUMBER
// ============================================================
/** Gom dữ liệu chi tiết (phiếu cân + hợp đồng) của các hồ sơ đã lọc - dùng chung cho Sheet 2 và File ngân hàng */
function collectDetailRows_(filteredRows) {
  const ss = SpreadsheetApp.getActive();
  const idSet = new Set(filteredRows.map(r => r.idHeThong));
  const mapParent = new Map(filteredRows.map(r => [r.idHeThong, r]));
  const ctData = ss.getSheetByName(CFG.DNTT_CT).getDataRange().getValues();
  const pcData = SpreadsheetApp.openById(CFG.PC_SS_ID).getSheetByName(CFG.PC_SHEET).getDataRange().getValues();
  const hdData = SpreadsheetApp.openById(CFG.HD_SS_ID).getSheetByName("HD_NCC").getDataRange().getValues();
  const mapPC = new Map(pcData.slice(1).map(r => [utils.standardize(r[22]), r]));
  const mapHD = new Map(hdData.slice(1).map(r => [utils.standardize(r[2]), r]));

  const tempRows = [];
  ctData.slice(1).forEach(ctRow => {
    let parentId = String(ctRow[1]).trim();
    if (!idSet.has(parentId)) return;
    let pc = mapPC.get(utils.standardize(ctRow[11])) || new Array(30).fill("");
    let hd = mapHD.get(utils.standardize(ctRow[19])) || new Array(30).fill("");
    let parentInfo = mapParent.get(parentId);
    tempRows.push({
      data: [
        0, parentInfo?.soLan, "'" + ctRow[11], utils.formatDate(ctRow[20]), utils.formatDate(pc[1]),
        pc[13], "'" + ctRow[19], pc[5], ctRow[3], "'" + (hd[6]||""), hd[10], "'" + (hd[11]||""),
        hd[5], hd[18], pc[2], pc[4], utils.parseNum(pc[7]), utils.parseNum(pc[8]),
        utils.parseNum(pc[9]), utils.parseNum(pc[9])/1000, utils.parseNum(pc[23]),
        utils.parseNum(pc[25]), pc[14], hd[16], "'" + (hd[15]||""), ctRow[2]
      ],
      bank: { ngayCK: utils.formatDate(ctRow[20]), soPhieuCan: ctRow[11], noiDung: parentInfo?.noiDungCK, thanhTien: ctRow[16], soHD: ctRow[19], hd: hd }
    });
  });
  return tempRows;
}

function renderSheet2Detail(sheet, filteredRows, dateRange) {
  const tempRows = collectDetailRows_(filteredRows);

  sheet.clear();
  sheet.getRange("A1:Z1").merge().setValue("BẢNG KÊ CHI TIẾT CHUYỂN KHOẢN").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("A2:Z2").merge().setValue("Thời gian: " + dateRange).setFontStyle("italic").setHorizontalAlignment("center");

  const headers = ["STT", "Lần TT", "Số phiếu cân", "Ngày CK", "Ngày nhập", "Đại lý", "Số hợp đồng", "Số xe", "Họ tên chủ rừng", "CCCD Chủ rừng", "Tên người thụ hưởng", "CCCD Người thụ hưởng", "Địa chỉ Chủ rừng", "Địa chỉ rừng", "Giờ vào", "Giờ ra", "KL hàng + xe", "Bì", "KL_KG", "KL_Tấn", "Đơn giá", "Thành tiền", "Nguồn gốc", "Ngân hàng", "Số tài khoản", "Thời gian nhập liệu"];
  sheet.getRange(4, 1, 1, headers.length).setValues([headers]).setBackground("#cfe2f3").setFontWeight("bold").setBorder(true, true, true, true, true, true).setHorizontalAlignment("center");

  if (tempRows.length > 0) {
    let finalRows = tempRows.map((item, index) => { item.data[0] = index + 1; return item.data; });
    const lastR = 5 + finalRows.length - 1;
    sheet.getRange(5, 1, finalRows.length, headers.length).setValues(finalRows).setBorder(true, true, true, true, true, true).setVerticalAlignment("middle");

    // ĐỊNH DẠNG SHEET 2
    sheet.getRange(5, 15, finalRows.length, 2).setNumberFormat("HH:mm"); // Giờ vào/ra
    sheet.getRange(5, 17, finalRows.length, 3).setNumberFormat("#,##0"); // KL KG
    sheet.getRange(5, 20, finalRows.length, 1).setNumberFormat("#,##0.000"); // Tấn
    sheet.getRange(5, 21, finalRows.length, 2).setNumberFormat("#,##0"); // Tiền

    // TỔNG CỘNG SHEET 2
    const totalRow = lastR + 1;
    sheet.getRange(totalRow, 1, 1, 16).merge().setValue("TỔNG CỘNG").setFontWeight("bold").setHorizontalAlignment("right");
    sheet.getRange(totalRow, 17, 1, 3).setFormulaR1C1(`=SUM(R5C:R[-1]C)`);
    sheet.getRange(totalRow, 20).setFormulaR1C1(`=SUM(R5C:R[-1]C)`);
    sheet.getRange(totalRow, 22).setFormulaR1C1(`=SUM(R5C:R[-1]C)`);
    sheet.getRange(totalRow, 1, 1, headers.length).setBackground("#fff2cc").setBorder(true, true, true, true, true, true);
  }
}

/** Xuất đúng 1 sheet ra blob .xlsx (qua bảng tính tạm, xóa sau khi xuất) */
function exportSheetToXlsx_(sheet, fileName) {
  const tmp = SpreadsheetApp.create(fileName);
  try {
    const copied = sheet.copyTo(tmp);
    copied.setName(sheet.getName());
    tmp.deleteSheet(tmp.getSheets()[0]); // bỏ sheet mặc định
    SpreadsheetApp.flush();
    const resp = UrlFetchApp.fetch("https://docs.google.com/spreadsheets/d/" + tmp.getId() + "/export?format=xlsx", {
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) throw new Error("Không xuất được file Excel (HTTP " + resp.getResponseCode() + ").");
    return resp.getBlob().setName(fileName);
  } finally {
    try { DriveApp.getFileById(tmp.getId()).setTrashed(true); } catch (x) {}
  }
}

// ============================================================
// NÚT 4: FILE NGÂN HÀNG (XÓA VÀ GHI MỚI sheet Update_NganHang_DN)
// ============================================================
function updateBankFile(filteredRows) {
  const lock = sysLock.acquire();
  try {
    const tempRows = collectDetailRows_(filteredRows || []);
    if (!tempRows.length) return { success: false, message: "Không có dữ liệu chi tiết để ghi file ngân hàng." };

    const ssNH = SpreadsheetApp.openById(CFG.UPDATE_NH_SS_ID);
    const shUpdateNH = ssNH.getSheetByName("Update_NganHang_DN");
    if (!shUpdateNH) return { success: false, message: "Không tìm thấy sheet Update_NganHang_DN." };

    // Xóa toàn bộ dữ liệu cũ từ dòng 2
    if (shUpdateNH.getLastRow() > 1) {
      shUpdateNH.getRange(2, 1, shUpdateNH.getLastRow() - 1, 33).clearContent();
    }

    const bankUpdateRows = tempRows.map(item => {
      let b = item.bank;
      let row = new Array(33).fill("");
      row[0] = ""; // Cột A bỏ trống
      row[1] = 0;       // Cột B: Giá trị mặc định là 0 (Số nguyên)
      row[2] = b.ngayCK; // Ngày hạch toán (Lấy từ ngày chốt)
      row[3] = b.ngayCK; // Ngày chứng từ
      row[4] = "'" + b.soPhieuCan;
      row[5] = "'" + CFG.TK_TRICH_NO; row[6] = "Ngân hàng TMCP Đầu tư và Phát triển Việt Nam - Quảng Nam - VND"; row[7] = 43;
      row[8] = b.noiDung || ("Thanh toán phiếu cân " + b.soPhieuCan); // Cột I diễn giải
      row[9] = "'" + String(b.hd[6] || ""); row[10] = b.hd[4]; row[11] = b.hd[5];
      row[12] = "'" + String(b.hd[15] || ""); row[13] = b.hd[16]; row[14] = b.hd[10];
      row[15] = "'" + String(b.hd[11] || ""); row[19] = "VND"; row[21] = b.noiDung;
      row[22] = "33111"; row[23] = "1121"; row[24] = b.thanhTien; row[32] = "'" + b.soHD;
      return row;
    });
    shUpdateNH.getRange(2, 1, bankUpdateRows.length, 33).setValues(bankUpdateRows);
    SpreadsheetApp.flush();

    // Xuất riêng sheet Update_NganHang_DN ra file Excel, lưu vào thư mục báo cáo
    const fileName = "UPDATE_NGANHANG (" + Utilities.formatDate(new Date(), CFG.TZ, "dd-MM-yyyy HHmm") + ").xlsx";
    const blob = exportSheetToXlsx_(shUpdateNH, fileName);
    const xlsxFile = DriveApp.getFolderById(CFG.FOLDER_REPORT_ID).createFile(blob);
    return {
      success: true,
      count: bankUpdateRows.length,
      fileName: fileName,
      fileUrl: xlsxFile.getUrl(),
      downloadUrl: "https://drive.google.com/uc?export=download&id=" + xlsxFile.getId(),
      base64: Utilities.base64Encode(blob.getBytes())
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally { lock.releaseLock(); }
}

// ============================================================
// NÚT 8: CHỐT THANH TOÁN (KHÓA DỮ LIỆU ĐA TẦNG)
// ============================================================
function runConfirmPayment(selectedIds, payDateStr) {
  const lock = sysLock.acquire(); 
  try {
    const ss = SpreadsheetApp.getActive();
    const shSrc = ss.getSheetByName(CFG.DNTT_SRC);
    const shCT = ss.getSheetByName(CFG.DNTT_CT);
    const sh112 = ss.getSheetByName(CFG.DNTT_112);
    const shPC = SpreadsheetApp.openById(CFG.PC_SS_ID).getSheetByName(CFG.PC_SHEET);
    
    const m = String(payDateStr || "").trim().match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
    if (!m) throw new Error("Ngày thanh toán không hợp lệ. Định dạng đúng: dd/mm/yyyy.");
    const dd = m[1].padStart(2, '0'), mm = m[2].padStart(2, '0'), yyyy = m[3];
    const chk = new Date(Number(yyyy), Number(mm) - 1, Number(dd));
    if (chk.getFullYear() != yyyy || chk.getMonth() != Number(mm) - 1 || chk.getDate() != Number(dd)) {
      throw new Error("Ngày thanh toán không tồn tại: " + payDateStr);
    }
    const prefixP = yyyy + mm + dd;
    const dateForSheet = new Date(Number(yyyy), Number(mm) - 1, Number(dd)); // Date thật, không phụ thuộc locale

    const requestedIds = new Set(selectedIds.map(id => String(id).trim()));
    const idSet = new Set(); // chỉ gồm hồ sơ CHƯA chốt
    let skipped = 0;

    // 1. XỬ LÝ MẢNG 112
    const data112 = sh112.getRange(1, 1, sh112.getLastRow(), sh112.getLastColumn()).getValues();
    const mapSoLan = new Map();
    for (let k = 1; k < data112.length; k++) {
      let id = String(data112[k][0]).trim();
      if (!requestedIds.has(id)) continue;
      if (String(data112[k][20]).trim().toUpperCase() === "Y") { skipped++; continue; } // đã chốt trước đó
      data112[k][20] = "Y";
      idSet.add(id);
      mapSoLan.set(id, String(data112[k][18]).trim());
    }
    if (idSet.size === 0) return "⚠️ Không có hồ sơ nào cần chốt (" + skipped + " hồ sơ đã được chốt trước đó).";

    // 2. XỬ LÝ MẢNG NGUỒN
    const dataSrc = shSrc.getRange(1, 1, shSrc.getLastRow(), shSrc.getLastColumn()).getValues();
    for (let i = 1; i < dataSrc.length; i++) {
      let key = String(dataSrc[i][0]).trim();
      if (idSet.has(key) && dataSrc[i][17] !== "Y") {
        dataSrc[i][14] = "Đóng TT"; dataSrc[i][15] = prefixP + "_" + (mapSoLan.get(key) || "1"); 
        dataSrc[i][16] = dateForSheet; dataSrc[i][17] = "Y"; 
      }
    }

    // 3. XỬ LÝ MẢNG CHI TIẾT & LẤY MÃ PHIẾU CÂN
    const dataCT = shCT.getRange(1, 1, shCT.getLastRow(), shCT.getLastColumn()).getValues();
    const pcKeySet = new Set();
    for (let j = 1; j < dataCT.length; j++) {
      let parentId = String(dataCT[j][1]).trim();
      if (idSet.has(parentId)) {
        dataCT[j][18] = "Y"; dataCT[j][20] = dateForSheet; 
        if (dataCT[j][11]) pcKeySet.add(utils.standardize(dataCT[j][11]));
      }
    }

    // 4. XỬ LÝ MẢNG PHIẾU CÂN
    const dataPC = shPC.getRange(1, 1, shPC.getLastRow(), shPC.getLastColumn()).getValues();
    for (let m = 1; m < dataPC.length; m++) {
      if (pcKeySet.has(utils.standardize(dataPC[m][22]))) { 
        dataPC[m][24] = "OK"; dataPC[m][26] = "Đóng TT"; dataPC[m][27] = "Y"; 
      }
    }

    // GHI ĐỒNG LOẠT (Tối ưu tốc độ)
    sh112.getRange(1, 1, data112.length, data112[0].length).setValues(data112);
    shSrc.getRange(1, 1, dataSrc.length, dataSrc[0].length).setValues(dataSrc);
    shCT.getRange(1, 1, dataCT.length, dataCT[0].length).setValues(dataCT);
    shPC.getRange(1, 1, dataPC.length, dataPC[0].length).setValues(dataPC);

    SpreadsheetApp.flush(); 
    return `✅ Đã chốt thành công ${idSet.size} hồ sơ & khóa Phiếu Cân!` + (skipped ? ` (Bỏ qua ${skipped} hồ sơ đã chốt trước đó)` : "");
  } catch (e) { 
    return "❌ Lỗi: " + e.toString(); 
  } finally { if (lock) lock.releaseLock(); }
}

/**
 * NÚT 1: TÁCH PHIẾU CHI TIẾT (BẢN FIX LỖI CỘT L -> V VÀ TRẠNG THÁI Y)
 */
function runProcessDetail() {
  const lock = sysLock.acquire();
  try { return _runProcessDetail(); } finally { lock.releaseLock(); }
}

function _runProcessDetail() {
  const ss = SpreadsheetApp.getActive();
  const shSrc = ss.getSheetByName(CFG.DNTT_SRC);
  const shCT = ss.getSheetByName(CFG.DNTT_CT);
  const shPC = SpreadsheetApp.openById(CFG.PC_SS_ID).getSheetByName(CFG.PC_SHEET);
  
  const srcData = shSrc.getDataRange().getValues();
  const pcValues = shPC.getDataRange().getValues(); 
  const pcMap = new Map(pcValues.map((r, idx) => [utils.standardize(r[22]), { rowIdx: idx, data: r }]));

  let ctData = shCT.getDataRange().getValues();
  let currentCTRows = ctData.slice(1); 
  let ctMapIdx = new Map(currentCTRows.map((row, idx) => [String(row[0]).trim(), idx]));

  let countAdd = 0;
  let countUpdate = 0;

  const splitIdx = srcData[0].indexOf(CFG.SPLIT_COL);
  if (splitIdx < 0) return "❌ Không tìm thấy cột \"" + CFG.SPLIT_COL + "\" trong sheet " + CFG.DNTT_SRC;

  for (let i = 1; i < srcData.length; i++) {
    let r = srcData[i];
    let idKey = String(r[0]).trim();
    if (!idKey) continue;

    // LẤY DỮ LIỆU NGÀY ĐỀ NGHỊ TỪ CỘT L (INDEX 11) CỦA SHEET NGUỒN
    let ngayDeNghiNguon = r[11]; 

    // CHỈ XỬ LÝ NẾU CHƯA CHỐT (KHÁC Y VÀ KHÁC ĐÓNG TT)
    // Cũng bỏ qua hồ sơ đã lên 112 ("Đã ĐNTT") để không lùi trạng thái
    if (String(r[17]).toUpperCase() !== "Y" && String(r[14]) !== "Đóng TT" && String(r[14]) !== "Đã ĐNTT") {
      let listPC = String(r[splitIdx] || "").split(/[,;|]/).map(x => x.trim()).filter(Boolean);
      if (listPC.length === 0) continue; // không có phiếu cân => không đánh dấu trạng thái
      
      listPC.forEach((soP, idx) => {
        let newID = idKey + "-" + (idx + 1);
        let pcObj = pcMap.get(utils.standardize(soP));
        
        let newRow = new Array(22).fill("");
        newRow[0] = newID;
        newRow[1] = idKey;
        newRow[2] = r[1];   // Timestamp (Cột C)
        newRow[3] = r[3];   // Chủ rừng
        newRow[4] = "'" + r[4];
        newRow[5] = r[5];   // Người ĐN
        newRow[6] = r[7];   // Người nhận
        newRow[7] = "'" + r[8];
        newRow[8] = r[9];   // KL Tổng nguồn
        newRow[9] = r[10];  // Ngân hàng nguồn
        newRow[10] = idx + 1;
        newRow[11] = "'" + soP;

        if (pcObj) {
          let pcR = pcObj.data;
          newRow[12] = utils.parseNum(pcR[9]) / 1000;
          newRow[13] = utils.parseNum(pcR[19]);
          newRow[14] = utils.parseNum(pcR[17]);
          newRow[15] = newRow[13] + newRow[14];
          newRow[16] = utils.parseNum(pcR[25]);
          newRow[17] = "Đã lấy giá từ Phiếu Cân";
          
          pcValues[pcObj.rowIdx][26] = "Đã Lập ĐNTT"; 
          pcValues[pcObj.rowIdx][27] = "N";           
        }
        
        newRow[18] = "N";
        newRow[19] = "'" + r[13]; // Số HĐ
        newRow[20] = r[16];       // Ngày TT
        
        // --- FIX TẠI ĐÂY: GHI VÀO CỘT V (INDEX 21) ---
        newRow[21] = ngayDeNghiNguon; 

        if (ctMapIdx.has(newID)) {
          let oldIdx = ctMapIdx.get(newID);
          currentCTRows[oldIdx] = newRow;
          countUpdate++;
        } else {
          currentCTRows.push(newRow);
          ctMapIdx.set(newID, currentCTRows.length - 1);
          countAdd++;
        }
      });

      srcData[i][14] = "Đã Lập ĐNTT"; 
      srcData[i][17] = "N";           
    } 
    // NẾU LÀ TRẠNG THÁI Y: GIỮ NGUYÊN DỮ LIỆU CŨ, KHÔNG ĐÈ MẢNG TRỐNG
  }

  if (countAdd > 0 || countUpdate > 0) {
    // Không clear trước khi ghi: số dòng chỉ tăng, tránh mất dữ liệu nếu lỗi giữa chừng
    shCT.getRange(2, 1, currentCTRows.length, 22).setValues(currentCTRows);
    shSrc.getRange(1, 1, srcData.length, srcData[0].length).setValues(srcData);
    shPC.getRange(1, 1, pcValues.length, pcValues[0].length).setValues(pcValues);
    
    SpreadsheetApp.flush();
    return `✅ Thành công: Đã cập nhật Ngày Đề Nghị vào cột V.`;
  } else {
    return `⚠️ Không có hồ sơ mới cần xử lý.`;
  }
}

/**
 * NÚT 2: TỔNG HỢP 112 (BẢN HOÀN THIỆN DIỄN GIẢI CỘT P)
 */
function runCreate112() {
  const lock = sysLock.acquire();
  try { return _runCreate112(); } finally { lock.releaseLock(); }
}

function _runCreate112() {
  const ss = SpreadsheetApp.getActive();
  const shSrc = ss.getSheetByName(CFG.DNTT_SRC);
  const sh112 = ss.getSheetByName(CFG.DNTT_112);
  const shCT = ss.getSheetByName(CFG.DNTT_CT);
  
  // Mở file Phiếu Cân ngoại vi
  let ssPC, shPC;
  try {
    ssPC = SpreadsheetApp.openById(CFG.PC_SS_ID);
    shPC = ssPC.getSheetByName(CFG.PC_SHEET);
  } catch(e) { return "❌ Lỗi: Không thể kết nối File Phiếu Cân."; }

  const fmt = (num) => utils.parseNum(num).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});

  // 1. ĐỌC DỮ LIỆU CÓ GIỚI HẠN
  const readSheet = (sh) => {
    const lr = sh.getLastRow();
    return lr > 1 ? sh.getRange(1, 1, lr, sh.getLastColumn()).getValues() : [];
  };

  const data112 = readSheet(sh112);
  const srcData = readSheet(shSrc);
  const ctData = readSheet(shCT);
  const pcData = readSheet(shPC);

  if (!data112.length) return "⚠️ Sheet 112 trống.";

  // 2. TẠO MAPS TRA CỨU NHANH
  const pcMapIdx = new Map(pcData.map((r, idx) => [utils.standardize(r[22]), idx]));
  const srcMapIdx = new Map(srcData.map((r, idx) => [String(r[0]).trim(), idx]));
  
  // Nhóm Chi tiết theo Số Hợp Đồng để tính lũy kế
  const ctGroupedByHD = ctData.reduce((acc, row) => {
    let hdKey = utils.standardize(row[19]);
    if (hdKey) {
      if (!acc[hdKey]) acc[hdKey] = [];
      acc[hdKey].push(row);
    }
    return acc;
  }, {});

  // Nhóm Chi tiết theo ID hồ sơ để lấy danh sách phiếu hiện tại
  const ctGroupedById = ctData.reduce((acc, row) => {
    let id = String(row[1]).trim();
    if (!acc[id]) acc[id] = [];
    acc[id].push(row);
    return acc;
  }, {});

  let count = 0;

  // 3. VÒNG LẶP XỬ LÝ TRÊN MẢNG
  for (let i = 1; i < data112.length; i++) {
    let id = String(data112[i][0]).trim();
    if (!id || String(data112[i][20]).toUpperCase() === "Y") continue;

    let time112 = (data112[i][1] instanceof Date) ? data112[i][1].getTime() : 0;
    if (time112 === 0) continue; 

    let soHD = String(data112[i][8] || "").replace(/'/g, "");
    let hdKey = utils.standardize(soHD);
    let ngayHDRaw = data112[i][9];
    let ngayHDFmt = (ngayHDRaw instanceof Date) ? Utilities.formatDate(ngayHDRaw, CFG.TZ, "dd.MM.yyyy") : String(ngayHDRaw || "");

    // A. Lọc dữ liệu chi tiết của riêng hồ sơ này
    let currentDetails = ctGroupedById[id] || [];
    if (currentDetails.length === 0) continue;

    // B. Tính lũy kế theo Hợp đồng (Chặn theo thời gian)
    let relate = (ctGroupedByHD[hdKey] || []).filter(c => {
      let timeCT = (c[2] instanceof Date) ? c[2].getTime() : 0;
      return timeCT > 0 && timeCT <= time112;
    });

    let slHD_TuDong = 0, daTra_N = 0, kyNay_O = 0, listP = [];
    relate.forEach(c => {
      let kl = utils.parseNum(c[12]);
      slHD_TuDong += kl; 
      if (String(c[18]).trim().toUpperCase() === "Y") {
        daTra_N += kl; 
      } else { 
        kyNay_O += kl;
        listP.push(String(c[11]).replace(/'/g, "").split('/')[0]);
      }
    });

    // C. Chuẩn bị các biến cho chuỗi Diễn giải P
    let tongTien = currentDetails.reduce((sum, d) => sum + utils.parseNum(d[16]), 0);
    let valR = utils.parseNum(data112[i][17]); // SL HĐ (Cột R)
    let valL_Tan = utils.parseNum(data112[i][11]) / 1000; // Đề nghị đợt này (Cột L / 1000)
    let phieuStr = listP.length > 0 ? listP.join(", ") : "Trống";

    // --- HOÀN THIỆN CẤU TRÚC DIỄN GIẢI P THEO YÊU CẦU ---
    let ghiChuP = `SL HĐ: ${fmt(valR)} | Tổng KL: ${fmt(slHD_TuDong)} | Đã trả: ${fmt(daTra_N)} | Còn lại: ${fmt(kyNay_O)} | Đề nghị đợt này: ${fmt(valL_Tan)} | Phiếu: ${phieuStr}`;

    // D. Ghi vào mảng 112
    data112[i][6] = tongTien; // Cột G
    data112[i][7] = `Thanh toán tiền mua gỗ keo HĐ số ${soHD} ${ngayHDFmt ? 'ngày ' + ngayHDFmt : ''}`;
    data112[i][12] = slHD_TuDong; // Cột M
    data112[i][13] = daTra_N;      // Cột N
    data112[i][14] = kyNay_O;      // Cột O
    data112[i][15] = ghiChuP;      // Cột P (Index 15)
    data112[i][21] = (soHD !== "" && tongTien > 0) ? "Đủ ĐK TT" : "Không đủ ĐK TT"; // Cột V

    // E. Cập nhật trạng thái Nguồn & Phiếu Cân
    let srcIdx = srcMapIdx.get(id);
    if (srcIdx !== undefined) srcData[srcIdx][14] = "Đã ĐNTT";

    currentDetails.forEach(d => {
      let pcIdx = pcMapIdx.get(utils.standardize(d[11]));
      if (pcIdx !== undefined) {
        pcData[pcIdx][26] = "Đã ĐNTT"; 
        pcData[pcIdx][27] = "N";
      }
    });

    count++;
  }

  // 4. GHI DỮ LIỆU ĐỒNG LOẠT
  if (count > 0) {
    sh112.getRange(1, 1, data112.length, data112[0].length).setValues(data112);
    shSrc.getRange(1, 1, srcData.length, srcData[0].length).setValues(srcData);
    shPC.getRange(1, 1, pcData.length, pcData[0].length).setValues(pcData);
    
    SpreadsheetApp.flush();
    return `✅ Đã tổng hợp ${count} hồ sơ. Đã khóa Phiếu Cân và cập nhật diễn giải cột P.`;
  }
  return `⚠️ Không có hồ sơ mới.`;
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🚀 QUẢN LÝ HAK')
      .addItem('1. Tách Phiếu Chi Tiết', 'runProcessDetail')
      .addItem('2. Tổng Hợp 112', 'runCreate112')
      .addToUi();
  // Chốt thanh toán thực hiện trên giao diện web (cần chọn danh sách hồ sơ).
}
