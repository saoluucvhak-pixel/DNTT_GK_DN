/**
 * ============================================================
 * HỆ THỐNG QUẢN LÝ THANH TOÁN HAK - PHIÊN BẢN HOÀN THIỆN 2026.2
 * ------------------------------------------------------------
 * THAY ĐỔI SO VỚI BẢN 2026 (bổ sung theo yêu cầu mới):
 *
 * A) TỰ ĐỘNG ĐIỀN "NGÂN HÀNG CỦA NGƯỜI THỤ HƯỞNG" (cột E) TRONG
 *    DNTT_GK_DN_112 KHI ĐANG BỊ BỎ TRỐNG:
 *    - Thêm sheet tham chiếu "HD_STK" (nằm trong file Hợp Đồng NCC,
 *      CFG.HD_SS_ID) làm nguồn tra cứu Ngân hàng.
 *    - Khóa tra cứu = (Số hợp đồng - cột I của HD_STK) + (Số tài
 *      khoản nhận tiền - cột F của HD_STK) -> trả về Ngân hàng (cột G).
 *    - Áp dụng cho TOÀN BỘ các dòng của DNTT_GK_DN_112 đang trống cột E,
 *      kể cả những dòng đã "chốt" (Tình trạng thanh toán = Y) từ trước,
 *      vì vòng lặp tổng hợp chính chỉ xử lý dòng CHƯA chốt. Vì vậy đây
 *      là một lượt quét/vá dữ liệu ĐỘC LẬP, chạy mỗi lần bấm
 *      "2. Tổng Hợp 112", không phụ thuộc hồ sơ đó mới hay cũ.
 *    - Không bao giờ ghi đè lên ô đã có sẵn Ngân hàng - chỉ điền khi
 *      ô đang thực sự trống (rỗng hoặc chỉ có khoảng trắng).
 *    - Nếu không tìm thấy trong HD_STK (không khớp Số HĐ + Số TK) thì
 *      giữ nguyên trống, không suy đoán.
 *
 * B) RÀ SOÁT & CỦNG CỐ XỬ LÝ DÒNG RỖNG (empty rows):
 *    - Thêm hàm dùng chung utils.buildIndexMap() để build các Map tra
 *      cứu (số phiếu cân, số hợp đồng...) có loại bỏ dòng có khóa rỗng
 *      trước khi đưa vào Map - tránh trường hợp nhiều dòng trống "" đè
 *      lẫn nhau trong Map và trỏ sai index khi ghi ngược lại sheet.
 *      Áp dụng cho: pcMap (runProcessDetail), mapPC/mapHD
 *      (renderSheet2Detail), pcMapIdx/srcMapIdx (runCreate112).
 *    - ctGroupedById (runCreate112): thêm điều kiện bỏ qua dòng có
 *      ID_KEY rỗng trước khi gom nhóm (trước đây gom cả dòng rỗng vào
 *      nhóm "").
 *    - utils.standardize(): củng cố thêm để chuẩn hóa numeric-string
 *      dạng "123.0" -> "123" (phòng trường hợp dữ liệu số hợp đồng bị
 *      lưu lẫn kiểu chuỗi có đuôi thập phân từ nguồn ngoài).
 *    - Toàn bộ các vòng lặp ghi đè (runConfirmPayment, runProcessDetail,
 *      runCreate112) đã được kiểm tra lại: đều có điều kiện bỏ qua khi
 *      khóa định danh (ID_KEY / ID_CT / Số phiếu cân) rỗng, không có
 *      trường hợp nào ghi nhầm dữ liệu vào dòng trống.
 *
 * (Các ghi chú thay đổi của bản 2026 gốc vẫn giữ nguyên bên dưới các
 * hàm liên quan, không lặp lại ở đây.)
 * ============================================================
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
  HD_STK_SHEET: "HD_STK",           // MỚI: sheet tra cứu Ngân hàng theo Số HĐ + Số TK
  SPLIT_COL: "Số phiếu cân",
  LAN_TT: "Lanthanhtoan",
  LOG_SHEET: "NhatKyThaoTac"
};

const utils = {
  standardize: (v) => {
    if (!v && v !== 0) return "";
    let s = String(v).replace(/\s+/g, '').toUpperCase().trim();
    // Chuẩn hóa số dạng "123.0" -> "123" (phòng dữ liệu số hợp đồng bị
    // đọc lẫn kiểu chuỗi thập phân từ nguồn ngoài / API bên thứ 3)
    if (/^-?\d+\.0+$/.test(s)) s = s.replace(/\.0+$/, '');
    return s;
  },
  formatDate: (d) => (d instanceof Date && d.getFullYear() > 1900) ? Utilities.formatDate(d, "GMT+7", "dd/MM/yyyy") : (d ? String(d) : ""),
  parseNum: (v) => {
    if (typeof v === 'number') return isNaN(v) ? 0 : v;
    if (!v) return 0;
    let n = parseFloat(String(v).replace(/[^0-9.-]+/g, ""));
    return isNaN(n) ? 0 : n;
  },
  timeToSeconds: (d) => (d instanceof Date) ? (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) : null,
  isBlank: (v) => v === "" || v === null || v === undefined || (typeof v === 'string' && v.trim() === ""),
  /**
   * MỚI: Build Map tra cứu index từ 1 mảng 2 chiều, tự động BỎ QUA các
   * dòng có khóa (cột colIdx) rỗng - tránh nhiều dòng trống "" đè lẫn
   * nhau trong Map và trỏ sai index khi ghi ngược lại sheet.
   * @param {Array[]} rows mảng dữ liệu (đã hoặc chưa bỏ dòng header)
   * @param {number} colIdx cột dùng làm khóa
   * @param {boolean} standardizeKey true nếu cần chuẩn hóa qua utils.standardize
   */
  buildIndexMap: (rows, colIdx, standardizeKey) => {
    const map = new Map();
    for (let i = 0; i < rows.length; i++) {
      const raw = rows[i][colIdx];
      if (utils.isBlank(raw)) continue;
      const key = standardizeKey ? utils.standardize(raw) : String(raw).trim();
      if (!key) continue;
      map.set(key, standardizeKey ? rows[i] : i);
    }
    return map;
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

/**
 * Ghi nhật ký thao tác. Không bao giờ để lỗi ghi log làm gãy luồng chính.
 */
function logAction(actionType, refId, detail) {
  try {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(CFG.LOG_SHEET);
    if (!sh) {
      sh = ss.insertSheet(CFG.LOG_SHEET);
      sh.appendRow(["Thời gian", "Người thực hiện", "Hành động", "Mã hồ sơ", "Chi tiết"]);
      sh.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#d9d9d9");
      sh.setColumnWidths(1, 5, 180);
    }
    let user = "N/A";
    try { user = Session.getActiveUser().getEmail() || "N/A"; } catch (e) {}
    sh.appendRow([new Date(), user, actionType, refId, detail]);
  } catch (e) {
    // Không chặn luồng chính nếu ghi log lỗi
  }
}

function doGet(e) {
  if (e.parameter.action) {
    let action = e.parameter.action;
    try {
      // --- Các action ghi dữ liệu / trả về chuỗi thông báo (giữ nguyên hành vi cũ) ---
      if (action === "tach_phieu") {
        return ContentService.createTextOutput(JSON.stringify({ status: "success", message: runProcessDetail() })).setMimeType(ContentService.MimeType.JSON);
      }
      if (action === "lap_de_nghi") {
        return ContentService.createTextOutput(JSON.stringify({ status: "success", message: runCreate112() })).setMimeType(ContentService.MimeType.JSON);
      }
      // --- Các action tra cứu chỉ đọc, trả về "data" (dùng cho tích hợp ngoài, không bắt buộc) ---
      if (action === "tim_phieu_can") {
        return ContentService.createTextOutput(JSON.stringify({ status: "success", data: findAvailablePhieuCan(e.parameter.q || "") })).setMimeType(ContentService.MimeType.JSON);
      }
      if (action === "tra_cuu_hop_dong") {
        return ContentService.createTextOutput(JSON.stringify({ status: "success", data: getContractInfo(e.parameter.soHD || "") })).setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Hành động không hợp lệ." })).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return HtmlService.createTemplateFromFile('Index').evaluate()
      .setTitle("QUẢN LÝ THANH TOÁN HAK")
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// TRUY VẤN NGOẠI VI (mở an toàn có try/catch dùng chung)
// ============================================================
function openExternalSheet(ssId, sheetName, labelForError) {
  try {
    const ss = SpreadsheetApp.openById(ssId);
    const sh = ss.getSheetByName(sheetName);
    if (!sh) throw new Error(`Không tìm thấy sheet "${sheetName}" trong file ${labelForError}.`);
    return sh;
  } catch (e) {
    throw new Error(`Không thể kết nối File ${labelForError}: ` + e.toString());
  }
}

// --- DATA VIEW FOR WEB (giữ nguyên như bản gốc) ---
function get112ViewData(fDate, tDate) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.DNTT_112);
  if (!sh || sh.getLastRow() < 2) return [];
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 22).getValues();
  const filtered = data.filter(r => {
    if (utils.isBlank(r[0])) return false; // bỏ dòng rỗng hoàn toàn
    let d = r[16]; // Cột Q
    if (!(d instanceof Date)) return false;
    let iso = Utilities.formatDate(d, "GMT+7", "yyyy-MM-dd");
    return (!fDate || iso >= fDate) && (!tDate || iso <= tDate);
  });
  return filtered.map(r => ({
    idHeThong: String(r[0] || ""),
    ngayDN: utils.formatDate(r[16]),
    ngayISO: Utilities.formatDate(r[16], "GMT+7", "yyyy-MM-dd"),
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
    if (utils.isBlank(r[0])) return false; // bỏ dòng rỗng hoàn toàn
    let d = r[11]; // Cột L
    if (!(d instanceof Date)) return false;
    let iso = Utilities.formatDate(d, "GMT+7", "yyyy-MM-dd");
    return (!fDate || iso >= fDate) && (!tDate || iso <= tDate);
  });
  return filtered.map(r => ({
    idHeThong: String(r[0] || ""),
    ngayDN: utils.formatDate(r[11]),
    ngayISO: Utilities.formatDate(r[11], "GMT+7", "yyyy-MM-dd"),
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
// NÚT 6: TẠO BÁO CÁO & SHEET NGÂN HÀNG (giữ nguyên logic gốc)
// ============================================================
function createFinalReportFromFilteredData(filteredRows, dateRange) {
  try {
    const folder = DriveApp.getFolderById(CFG.FOLDER_REPORT_ID);
    const fileName = "BÁO CÁO ĐNTT (" + dateRange + ") - " + Utilities.formatDate(new Date(), "GMT+7", "HHmm");
    const newSS = SpreadsheetApp.create(fileName);
    const file = DriveApp.getFileById(newSS.getId());
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);

    const sheet1 = newSS.getSheets()[0];
    sheet1.setName("DeNghiThanhToanCK");
    renderSheet1Full(sheet1, filteredRows, dateRange);

    const sheet2 = newSS.insertSheet("BangKeChiTietCK");
    renderSheet2Detail(sheet2, filteredRows, dateRange);

    logAction("XUAT_BAO_CAO", dateRange, `Đã xuất báo cáo ${filteredRows.length} hồ sơ - ${newSS.getUrl()}`);

    return { success: true, url: newSS.getUrl() };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ============================================================
// SHEET 1: TỐI ƯU GIAO DIỆN IN ẤN (giữ nguyên logic gốc)
// ============================================================
function renderSheet1Full(sheet, filteredRows, dateRange) {
  sheet.clear();
  sheet.getRange("A1").setValue("HOÀNG ANH KHÔI ĐÀ NẴNG").setFontSize(10).setFontWeight("bold");
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
// (đã bọc try/catch khi mở file ngoại vi; map tra cứu đã lọc dòng rỗng)
// ============================================================
function renderSheet2Detail(sheet, filteredRows, dateRange) {
  const ss = SpreadsheetApp.getActive();
  const idSet = new Set(filteredRows.map(r => r.idHeThong));
  const shCT = ss.getSheetByName(CFG.DNTT_CT);
  if (!shCT) throw new Error(`Không tìm thấy sheet "${CFG.DNTT_CT}".`);
  const ctData = shCT.getDataRange().getValues();

  // Kết nối các file ngoại vi (an toàn hơn bản gốc: báo lỗi rõ ràng nếu mất quyền truy cập)
  const shPC = openExternalSheet(CFG.PC_SS_ID, CFG.PC_SHEET, "Phiếu Cân");
  const pcData = shPC.getDataRange().getValues();
  const shHD = openExternalSheet(CFG.HD_SS_ID, "HD_NCC", "Hợp Đồng NCC");
  const hdData = shHD.getDataRange().getValues();
  const ssNH = SpreadsheetApp.openById(CFG.UPDATE_NH_SS_ID);
  const shUpdateNH = ssNH.getSheetByName("Update_NganHang_DN");

  // MAP đã lọc dòng rỗng (utils.buildIndexMap) - tránh dòng trống "" đè
  // lẫn nhau và trỏ sai dữ liệu tham chiếu
  const mapPC = utils.buildIndexMap(pcData.slice(1), 22, true);
  const mapHD = utils.buildIndexMap(hdData.slice(1), 2, true);

  sheet.clear();
  sheet.getRange("A1:Z1").merge().setValue("BẢNG KÊ CHI TIẾT CHUYỂN KHOẢN").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("A2:Z2").merge().setValue("Thời gian: " + dateRange).setFontStyle("italic").setHorizontalAlignment("center");

  const headers = ["STT", "Lần TT", "Số phiếu cân", "Ngày CK", "Ngày nhập", "Đại lý", "Số hợp đồng", "Số xe", "Họ tên chủ rừng", "CCCD Chủ rừng", "Tên người thụ hưởng", "CCCD Người thụ hưởng", "Địa chỉ Chủ rừng", "Địa chỉ rừng", "Giờ vào", "Giờ ra", "KL hàng + xe", "Bì", "KL_KG", "KL_Tấn", "Đơn giá", "Thành tiền", "Nguồn gốc", "Ngân hàng", "Số tài khoản", "Thời gian nhập liệu"];
  sheet.getRange(4, 1, 1, headers.length).setValues([headers]).setBackground("#cfe2f3").setFontWeight("bold").setBorder(true, true, true, true, true, true).setHorizontalAlignment("center");

  let tempRows = [];
  ctData.slice(1).forEach(ctRow => {
    if (utils.isBlank(ctRow[0]) || utils.isBlank(ctRow[1])) return; // bỏ dòng rỗng
    let parentId = String(ctRow[1]).trim();
    if (idSet.has(parentId)) {
      let pc = mapPC.get(utils.standardize(ctRow[11])) || new Array(30).fill("");
      let hd = mapHD.get(utils.standardize(ctRow[19])) || new Array(30).fill("");
      let parentInfo = filteredRows.find(f => f.idHeThong === parentId);

      tempRows.push({
        data: [
          0, parentInfo?.soLan, "'" + ctRow[11], utils.formatDate(ctRow[20]), utils.formatDate(pc[1]),
          pc[13], "'" + ctRow[19], pc[5], ctRow[3], "'" + (hd[6]||""), hd[10], "'" + (hd[11]||""),
          hd[5], hd[18], pc[2], pc[4], utils.parseNum(pc[7]), utils.parseNum(pc[8]),
          utils.parseNum(pc[9]), utils.parseNum(pc[9])/1000, utils.parseNum(pc[23]),
          // NOTE: Đơn giá thanh toán = báo giá + khấu trừ - cần xác nhận dấu với kế toán (xem ghi chú đầu file)
          utils.parseNum(pc[25]), pc[14], hd[16], "'" + (hd[15]||""), ctRow[2]
        ],
        bank: { ngayCK: utils.formatDate(ctRow[20]), soPhieuCan: ctRow[11], noiDung: parentInfo?.noiDungCK, thanhTien: ctRow[16], soHD: ctRow[19], hd: hd }
      });
    }
  });

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

    // --- XỬ LÝ FILE UPDATE NGÂN HÀNG (XÓA VÀ GHI MỚI) ---
    if (shUpdateNH) {
      // Xóa toàn bộ dữ liệu cũ từ dòng 2
      if (shUpdateNH.getLastRow() > 1) {
        shUpdateNH.getRange(2, 1, shUpdateNH.getLastRow(), 33).clearContent();
      }

      let bankUpdateRows = tempRows.map(item => {
        let b = item.bank;
        let row = new Array(33).fill("");
        row[0] = ""; // Cột A bỏ trống
        row[1] = 0;       // Cột B: Giá trị mặc định là 0 (Số nguyên)
        row[2] = b.ngayCK; // Ngày hạch toán (Lấy từ ngày chốt)
        row[3] = b.ngayCK; // Ngày chứng từ
        row[4] = "'" + b.soPhieuCan;
        row[5] = "'8619299999"; row[6] = "BIDV"; row[7] = 43;
        row[8] = b.noiDung || ("Thanh toán phiếu cân " + b.soPhieuCan); // Cột I diễn giải
        row[9] = "'" + String(b.hd[6] || ""); row[10] = b.hd[4]; row[11] = b.hd[5];
        row[12] = "'" + String(b.hd[15] || ""); row[13] = b.hd[16]; row[14] = b.hd[10];
        row[15] = "'" + String(b.hd[11] || ""); row[19] = "VND"; row[21] = b.noiDung;
        row[22] = "33111"; row[23] = "1121"; row[24] = b.thanhTien; row[32] = "'" + b.soHD;
        return row;
      });
      shUpdateNH.getRange(2, 1, bankUpdateRows.length, 33).setValues(bankUpdateRows);
    }
  }
}

// ============================================================
// NÚT 8: CHỐT THANH TOÁN (KHÓA DỮ LIỆU ĐA TẦNG)
// (đã thêm validate định dạng ngày + ghi log)
// ============================================================
function runConfirmPayment(selectedIds, payDateStr) {
  const lock = sysLock.acquire();
  try {
    const ss = SpreadsheetApp.getActive();
    const shSrc = ss.getSheetByName(CFG.DNTT_SRC);
    const shCT = ss.getSheetByName(CFG.DNTT_CT);
    const sh112 = ss.getSheetByName(CFG.DNTT_112);
    const shPC = openExternalSheet(CFG.PC_SS_ID, CFG.PC_SHEET, "Phiếu Cân");

    if (!Array.isArray(selectedIds) || selectedIds.length === 0) {
      return "❌ Lỗi: Không có hồ sơ nào được chọn để chốt thanh toán.";
    }

    const p = String(payDateStr || "").trim().split('/');
    if (p.length !== 3 || isNaN(parseInt(p[0], 10)) || isNaN(parseInt(p[1], 10)) || isNaN(parseInt(p[2], 10))) {
      return "❌ Lỗi: Ngày thanh toán không hợp lệ. Định dạng đúng phải là dd/mm/yyyy.";
    }
    const prefixP = p[2] + p[1].padStart(2, '0') + p[0].padStart(2, '0'),
          dateForSheet = p[1] + "/" + p[0] + "/" + p[2];

    const idSet = new Set(selectedIds.map(id => String(id).trim()).filter(id => id !== ""));
    if (idSet.size === 0) {
      return "❌ Lỗi: Danh sách hồ sơ được chọn không hợp lệ (rỗng).";
    }

    // 1. XỬ LÝ MẢNG 112
    const data112 = sh112.getRange(1, 1, sh112.getLastRow(), sh112.getLastColumn()).getValues();
    const mapSoLan = new Map();
    for (let k = 1; k < data112.length; k++) {
      let id = String(data112[k][0]).trim();
      if (!id) continue; // bỏ dòng rỗng
      if (idSet.has(id)) {
        data112[k][20] = "Y";
        mapSoLan.set(id, String(data112[k][18]).trim());
      }
    }

    // 2. XỬ LÝ MẢNG NGUỒN
    const dataSrc = shSrc.getRange(1, 1, shSrc.getLastRow(), shSrc.getLastColumn()).getValues();
    for (let i = 1; i < dataSrc.length; i++) {
      let key = String(dataSrc[i][0]).trim();
      if (!key) continue; // bỏ dòng rỗng
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
      if (!parentId) continue; // bỏ dòng rỗng
      if (idSet.has(parentId)) {
        dataCT[j][18] = "Y"; dataCT[j][20] = dateForSheet;
        if (dataCT[j][11]) pcKeySet.add(utils.standardize(dataCT[j][11]));
      }
    }

    // 4. XỬ LÝ MẢNG PHIẾU CÂN
    const dataPC = shPC.getRange(1, 1, shPC.getLastRow(), shPC.getLastColumn()).getValues();
    for (let m = 1; m < dataPC.length; m++) {
      if (utils.isBlank(dataPC[m][22])) continue; // bỏ dòng rỗng
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
    logAction("CHOT_THANH_TOAN", selectedIds.join(","), `Chốt ${idSet.size} hồ sơ, ngày TT ${dateForSheet}`);
    return `✅ Đã chốt thành công ${idSet.size} hồ sơ & khóa Phiếu Cân!`;
  } catch (e) {
    return "❌ Lỗi: " + e.toString();
  } finally { if (lock) lock.releaseLock(); }
}

/**
 * NÚT 1: TÁCH PHIẾU CHI TIẾT
 * (đã thêm khóa; map tra cứu Phiếu Cân đã lọc dòng rỗng)
 */
function runProcessDetail() {
  const lock = sysLock.acquire();
  try {
    const ss = SpreadsheetApp.getActive();
    const shSrc = ss.getSheetByName(CFG.DNTT_SRC);
    const shCT = ss.getSheetByName(CFG.DNTT_CT);
    const shPC = openExternalSheet(CFG.PC_SS_ID, CFG.PC_SHEET, "Phiếu Cân");

    const srcData = shSrc.getDataRange().getValues();
    const pcValues = shPC.getDataRange().getValues();
    // MAP đã lọc dòng rỗng - tránh nhiều dòng Phiếu Cân trống số phiếu
    // đè lẫn nhau trong Map
    const pcMap = utils.buildIndexMap(pcValues, 22, true);
    // Lưu lại rowIdx thật (so với mảng pcValues) để ghi ngược lại đúng dòng
    const pcRowIdxMap = new Map();
    pcValues.forEach((r, idx) => {
      const key = utils.standardize(r[22]);
      if (key) pcRowIdxMap.set(key, idx);
    });

    let ctData = shCT.getDataRange().getValues();
    let currentCTRows = ctData.slice(1);
    let ctMapIdx = utils.buildIndexMap(currentCTRows, 0, false);

    const splitColIdx = srcData[0].indexOf(CFG.SPLIT_COL); // tính 1 lần ngoài vòng lặp
    if (splitColIdx === -1) {
      // BỔ SUNG: nếu vì lý do gì đó không tìm thấy cột "Số phiếu cân" (đổi
      // tên cột, sai chính tả...), dừng ngay và báo lỗi rõ ràng - thay vì
      // để mọi dòng lặng lẽ bị coi như "không có phiếu cân" và có nguy cơ
      // bị đánh dấu sai trạng thái hàng loạt.
      throw new Error(`Không tìm thấy cột "${CFG.SPLIT_COL}" trong sheet ${CFG.DNTT_SRC}. Vui lòng kiểm tra lại tiêu đề cột.`);
    }

    let countAdd = 0;
    let countUpdate = 0;
    let countSkippedEmptyPC = 0; // MỚI: đếm hồ sơ bị bỏ qua vì thiếu Số phiếu cân
    let skippedIds = [];

    for (let i = 1; i < srcData.length; i++) {
      let r = srcData[i];
      let idKey = String(r[0]).trim();
      if (!idKey) continue; // bỏ dòng rỗng

      // LẤY DỮ LIỆU NGÀY ĐỀ NGHỊ TỪ CỘT L (INDEX 11) CỦA SHEET NGUỒN
      let ngayDeNghiNguon = r[11];

      // CHỈ XỬ LÝ NẾU CHƯA CHỐT (KHÁC Y VÀ KHÁC ĐÓNG TT)
      if (String(r[17]).toUpperCase() !== "Y" && String(r[14]) !== "Đóng TT") {
        let listPC = String(r[splitColIdx] || "").split(/[,;|]/).map(x => x.trim()).filter(Boolean);

        // BỔ SUNG QUAN TRỌNG: nếu hồ sơ có ID hợp lệ nhưng cột "Số phiếu cân"
        // đang trống (hoặc chỉ chứa dấu phẩy/khoảng trắng), TRƯỚC ĐÂY code sẽ
        // vẫn đánh dấu srcData[i][14]="Đã Lập ĐNTT" dù KHÔNG tạo ra dòng chi
        // tiết nào trong DNTT_GK_DN_CT -> hồ sơ "biến mất" khỏi quy trình mà
        // không có cảnh báo. Nay: bỏ qua hồ sơ này, GIỮ NGUYÊN trạng thái để
        // hồ sơ còn hiển thị là "chưa xử lý" và được xử lý lại ở lần chạy sau
        // khi đã bổ sung Số phiếu cân.
        if (listPC.length === 0) {
          countSkippedEmptyPC++;
          if (skippedIds.length < 20) skippedIds.push(idKey);
          continue;
        }

        listPC.forEach((soP, idx) => {
          let newID = idKey + "-" + (idx + 1);
          let pcRow = pcMap.get(utils.standardize(soP));
          let pcRowIdx = pcRowIdxMap.get(utils.standardize(soP));

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
          newRow[9] = r[10];  // Số phiếu cân (danh sách gốc)
          newRow[10] = idx + 1;
          newRow[11] = "'" + soP;

          if (pcRow) {
            newRow[12] = utils.parseNum(pcRow[9]) / 1000;
            newRow[13] = utils.parseNum(pcRow[19]);
            newRow[14] = utils.parseNum(pcRow[17]);
            newRow[15] = newRow[13] + newRow[14];
            newRow[16] = utils.parseNum(pcRow[25]);
            newRow[17] = "Đã lấy giá từ Phiếu Cân";

            if (pcRowIdx !== undefined) {
              pcValues[pcRowIdx][26] = "Đã Lập ĐNTT";
              pcValues[pcRowIdx][27] = "N";
            }
          }

          newRow[18] = "N";
          newRow[19] = "'" + r[13]; // Số HĐ
          newRow[20] = r[16];       // Ngày TT

          // GHI VÀO CỘT V (INDEX 21)
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
      if (shCT.getLastRow() > 1) shCT.getRange(2, 1, shCT.getLastRow() - 1, 22).clearContent();
      shCT.getRange(2, 1, currentCTRows.length, 22).setValues(currentCTRows);
      shSrc.getRange(1, 1, srcData.length, srcData[0].length).setValues(srcData);
      shPC.getRange(1, 1, pcValues.length, pcValues[0].length).setValues(pcValues);

      SpreadsheetApp.flush();
      let logDetail = `Thêm mới ${countAdd}, cập nhật ${countUpdate} dòng chi tiết.`;
      if (countSkippedEmptyPC > 0) logDetail += ` Bỏ qua ${countSkippedEmptyPC} hồ sơ thiếu Số phiếu cân: ${skippedIds.join(", ")}${countSkippedEmptyPC > skippedIds.length ? "..." : ""}`;
      logAction("TACH_PHIEU", "-", logDetail);

      let msg = `✅ Thành công: Đã thêm ${countAdd}, cập nhật ${countUpdate} dòng. Đã cập nhật Ngày Đề Nghị vào cột V.`;
      if (countSkippedEmptyPC > 0) msg += ` ⚠️ Đã BỎ QUA ${countSkippedEmptyPC} hồ sơ vì thiếu "Số phiếu cân" (ID: ${skippedIds.join(", ")}${countSkippedEmptyPC > skippedIds.length ? "..." : ""}) - vui lòng bổ sung rồi chạy lại.`;
      return msg;
    } else if (countSkippedEmptyPC > 0) {
      return `⚠️ Không có dòng nào được tách. Có ${countSkippedEmptyPC} hồ sơ đang thiếu "Số phiếu cân" (ID: ${skippedIds.join(", ")}${countSkippedEmptyPC > skippedIds.length ? "..." : ""}) nên bị bỏ qua - vui lòng bổ sung Số phiếu cân rồi chạy lại.`;
    } else {
      return `⚠️ Không có hồ sơ mới cần xử lý.`;
    }
  } catch (e) {
    return "❌ Lỗi: " + e.toString();
  } finally { if (lock) lock.releaseLock(); }
}

// ============================================================
// MỚI: TRA CỨU NGÂN HÀNG THEO (SỐ HỢP ĐỒNG + SỐ TÀI KHOẢN) TỪ HD_STK
// Dùng để tự động điền cột E "Ngân hàng của người thụ hưởng" trong
// DNTT_GK_DN_112 khi ô đó đang bị bỏ trống.
// Khóa map = standardize(Số hợp đồng - cột I của HD_STK) + "|" +
//            standardize(Số TK nhận tiền - cột F của HD_STK)
// Giá trị   = Ngân hàng (cột G của HD_STK)
// Không chặn luồng chính của Tổng Hợp 112 nếu không đọc được HD_STK.
// ============================================================
function buildBankLookupFromHDSTK() {
  const map = new Map();
  try {
    const shSTK = openExternalSheet(CFG.HD_SS_ID, CFG.HD_STK_SHEET, "Hợp Đồng NCC (Sổ TK - HD_STK)");
    const data = shSTK.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const soHD = utils.standardize(row[8]);   // Cột I: Số hợp đồng
      const soTK = utils.standardize(row[5]);   // Cột F: Số TK nhận tiền
      const nganHang = String(row[6] || "").trim(); // Cột G: Ngân hàng
      if (!soHD || !soTK || !nganHang) continue; // bỏ dòng rỗng / thiếu dữ liệu
      map.set(soHD + "|" + soTK, nganHang);
    }
  } catch (e) {
    // Không có sheet HD_STK hoặc không truy cập được -> bỏ qua bước điền
    // Ngân hàng, các bước còn lại của Tổng Hợp 112 vẫn chạy bình thường.
  }
  return map;
}

/**
 * NÚT 2: TỔNG HỢP 112
 * (đã thêm khóa; map tra cứu đã lọc dòng rỗng; BỔ SUNG: tự động điền
 * Ngân hàng của người thụ hưởng (cột E) khi đang trống, tham chiếu
 * HD_STK theo cặp Số hợp đồng + Số tài khoản - áp dụng cho TOÀN BỘ
 * dòng dữ liệu, không chỉ hồ sơ mới, vì vòng lặp tổng hợp chính chỉ xử
 * lý các dòng chưa chốt (Tình trạng thanh toán khác "Y"))
 */
function runCreate112() {
  const lock = sysLock.acquire();
  try {
    const ss = SpreadsheetApp.getActive();
    const shSrc = ss.getSheetByName(CFG.DNTT_SRC);
    const sh112 = ss.getSheetByName(CFG.DNTT_112);
    const shCT = ss.getSheetByName(CFG.DNTT_CT);

    // Mở file Phiếu Cân ngoại vi
    let shPC;
    try {
      shPC = openExternalSheet(CFG.PC_SS_ID, CFG.PC_SHEET, "Phiếu Cân");
    } catch (e) { return "❌ Lỗi: Không thể kết nối File Phiếu Cân."; }

    const fmt = (num) => utils.parseNum(num).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

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

    // 2. TẠO MAPS TRA CỨU NHANH (đã lọc dòng rỗng)
    const pcMapIdx = utils.buildIndexMap(pcData, 22, true) && (() => {
      // buildIndexMap ở chế độ standardizeKey=true trả về row, ở đây cần idx
      // nên build riêng để giữ đúng hành vi gốc (idx trong pcData)
      const m = new Map();
      for (let i = 0; i < pcData.length; i++) {
        const key = utils.standardize(pcData[i][22]);
        if (key) m.set(key, i);
      }
      return m;
    })();
    const srcMapIdx = (() => {
      const m = new Map();
      for (let i = 0; i < srcData.length; i++) {
        const key = String(srcData[i][0] || "").trim();
        if (key) m.set(key, i);
      }
      return m;
    })();

    // Nhóm Chi tiết theo Số Hợp Đồng để tính lũy kế (bỏ dòng rỗng)
    const ctGroupedByHD = ctData.reduce((acc, row) => {
      let hdKey = utils.standardize(row[19]);
      if (hdKey) {
        if (!acc[hdKey]) acc[hdKey] = [];
        acc[hdKey].push(row);
      }
      return acc;
    }, {});

    // Nhóm Chi tiết theo ID hồ sơ để lấy danh sách phiếu hiện tại (bỏ dòng rỗng)
    const ctGroupedById = ctData.reduce((acc, row) => {
      let id = String(row[1] || "").trim();
      if (!id) return acc; // BỔ SUNG: bỏ qua dòng có ID_KEY rỗng
      if (!acc[id]) acc[id] = [];
      acc[id].push(row);
      return acc;
    }, {});

    // MỚI: bảng tra cứu Ngân hàng theo (Số hợp đồng + Số tài khoản) từ HD_STK
    const bankMap = buildBankLookupFromHDSTK();

    let count = 0;
    let filledBankCount = 0; // MỚI: đếm số dòng đã tự động điền Ngân hàng

    // 3. VÒNG LẶP XỬ LÝ TRÊN MẢNG
    for (let i = 1; i < data112.length; i++) {
      let id = String(data112[i][0]).trim();
      if (!id) continue; // bỏ dòng rỗng

      // MỚI: TỰ ĐỘNG ĐIỀN NGÂN HÀNG (CỘT E) NẾU ĐANG TRỐNG - áp dụng cho
      // MỌI dòng có ID (kể cả đã "Y"/đã chốt), không chỉ hồ sơ mới xử lý
      // bên dưới, vì các dòng đã chốt sẽ bị "continue" và không đi tới
      // đoạn xử lý chính.
      if (utils.isBlank(data112[i][4])) {
        const soHDRaw = String(data112[i][8] || "").replace(/'/g, "");
        const hdKeyLookup = utils.standardize(soHDRaw);
        const stkKeyLookup = utils.standardize(data112[i][5]);
        if (hdKeyLookup && stkKeyLookup) {
          const bankFound = bankMap.get(hdKeyLookup + "|" + stkKeyLookup);
          if (bankFound) {
            data112[i][4] = bankFound;
            filledBankCount++;
          }
        }
      }

      if (String(data112[i][20]).toUpperCase() === "Y") continue;

      let time112 = (data112[i][1] instanceof Date) ? data112[i][1].getTime() : 0;
      if (time112 === 0) continue;

      let soHD = String(data112[i][8] || "").replace(/'/g, "");
      let hdKey = utils.standardize(soHD);
      let ngayHDRaw = data112[i][9];
      let ngayHDFmt = (ngayHDRaw instanceof Date) ? Utilities.formatDate(ngayHDRaw, "GMT+7", "dd.MM.yyyy") : String(ngayHDRaw || "");

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
    // MỚI: ghi lại sh112 nếu có hồ sơ mới HOẶC có dòng vừa được tự động
    // điền Ngân hàng (kể cả khi không có hồ sơ mới nào)
    if (count > 0 || filledBankCount > 0) {
      sh112.getRange(1, 1, data112.length, data112[0].length).setValues(data112);

      if (count > 0) {
        shSrc.getRange(1, 1, srcData.length, srcData[0].length).setValues(srcData);
        shPC.getRange(1, 1, pcData.length, pcData[0].length).setValues(pcData);
      }

      SpreadsheetApp.flush();

      const parts = [];
      if (count > 0) parts.push(`Đã tổng hợp ${count} hồ sơ`);
      if (filledBankCount > 0) parts.push(`đã tự động điền Ngân hàng cho ${filledBankCount} dòng đang trống (tham chiếu HD_STK)`);
      const detailMsg = parts.join(", ") + ".";

      logAction("TONG_HOP_112", "-", detailMsg);

      let msg = "✅ " + detailMsg;
      if (count > 0) msg += " Đã khóa Phiếu Cân và cập nhật diễn giải cột P.";
      return msg;
    }
    return `⚠️ Không có hồ sơ mới và không có dòng nào cần điền Ngân hàng.`;
  } catch (e) {
    return "❌ Lỗi: " + e.toString();
  } finally { if (lock) lock.releaseLock(); }
}

// ============================================================
// MỚI: TÍNH LẦN THANH TOÁN + NGÀY DỰ KIẾN THEO KHUNG GIỜ
// Đọc động từ sheet "Lanthanhtoan" (trước đây khai báo CFG.LAN_TT
// nhưng không có hàm nào sử dụng).
// ============================================================
function getSessionInfo() {
  const now = new Date();
  const fallback = () => {
    const ngayDuKien = new Date(now);
    ngayDuKien.setDate(ngayDuKien.getDate() + 1);
    return { maPhien: "N/A", lan: 1, ngayDuKien };
  };
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(CFG.LAN_TT);
    if (!sh || sh.getLastRow() < 2) return fallback();
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
    const nowSec = now.getHours() * 3600 + now.getMinutes() * 60 + now.getSeconds();

    for (const row of data) {
      const [maPhien, tuGio, denGio, lan, ngayOffset] = row;
      if (!maPhien) continue;
      const tuSec = utils.timeToSeconds(tuGio);
      const denSec = utils.timeToSeconds(denGio);
      if (tuSec === null || denSec === null) continue;
      if (nowSec >= tuSec && nowSec <= denSec) {
        const ngayDuKien = new Date(now);
        ngayDuKien.setDate(ngayDuKien.getDate() + utils.parseNum(ngayOffset));
        return { maPhien: String(maPhien).trim(), lan: utils.parseNum(lan) || 1, ngayDuKien };
      }
    }
    return fallback();
  } catch (e) {
    return fallback();
  }
}

// ============================================================
// MỚI: TÌM SỐ PHIẾU CÂN CÒN KHẢ DỤNG (chưa lập ĐNTT / chưa đóng TT)
// Dùng cho form "Thêm mới đề nghị thanh toán" để tránh gõ tay sai
// và tránh dùng trùng phiếu cân.
// ============================================================
function findAvailablePhieuCan(query) {
  const shPC = openExternalSheet(CFG.PC_SS_ID, CFG.PC_SHEET, "Phiếu Cân");
  const data = shPC.getDataRange().getValues();
  const q = utils.standardize(query);
  const results = [];
  for (let i = 1; i < data.length && results.length < 25; i++) {
    const r = data[i];
    const soP = String(r[22] || "").trim();
    if (!soP) continue; // bỏ dòng rỗng
    const trangThai = String(r[26] || "").trim();
    const daKhoa = String(r[27] || "").trim().toUpperCase();
    if (trangThai || daKhoa === "Y" || daKhoa === "N") continue; // đã được dùng cho hồ sơ khác
    if (q && !utils.standardize(soP).includes(q)) continue;
    results.push({
      soPhieuCan: soP,
      ngayNhap: utils.formatDate(r[1]),
      soXe: String(r[5] || ""),
      klKg: utils.parseNum(r[9])
    });
  }
  return results;
}

// ============================================================
// MỚI: TRA CỨU THÔNG TIN HỢP ĐỒNG ĐỂ TỰ ĐIỀN FORM
// CẢNH BÁO: chỉ số cột (4,6,10,11,15,16) suy ra từ cách renderSheet2Detail
// đang dùng (hd[6], hd[10], hd[11], hd[15], hd[16]) - cần đối chiếu lại
// với cấu trúc thật của sheet "HD_NCC" trước khi dùng chính thức.
// ============================================================
function getContractInfo(soHopDong) {
  if (!soHopDong) return null;
  try {
    const shHD = openExternalSheet(CFG.HD_SS_ID, "HD_NCC", "Hợp Đồng NCC");
    const hdData = shHD.getDataRange().getValues();
    const key = utils.standardize(soHopDong);
    const row = hdData.slice(1).find(r => !utils.isBlank(r[2]) && utils.standardize(r[2]) === key);
    if (!row) return null;
    return {
      soHopDong: String(row[2] || ""),
      cccdChuRung: String(row[6] || ""),
      nguoiThuHuong: String(row[10] || ""),
      cccdThuHuong: String(row[11] || ""),
      soTaiKhoan: String(row[15] || ""),
      nganHang: String(row[16] || "")
    };
  } catch (e) {
    return null;
  }
}

// ============================================================
// MỚI: THÊM MỚI ĐỀ NGHỊ THANH TOÁN
// Tạo trực tiếp 1 hồ sơ đề nghị thanh toán mà KHÔNG cần qua Google Form,
// ghi đồng thời vào DNTT_GK_DN và DNTT_GK_DN_112 để tương thích 100% với
// "Tách Phiếu" (runProcessDetail) và "Tổng Hợp 112" (runCreate112) hiện có.
//
// payload = {
//   hoTenChuRung, cccdChuRung, nguoiDeNghi, chuRungUyQuyen ("Có"/"Không"),
//   nguoiNhanTien, soTKNhanTien, nganHang, soHopDong,
//   danhSachPhieuCan: ["15852/2025/NK", ...]
// }
// ============================================================
function createNewPaymentRequest(payload) {
  const lock = sysLock.acquire();
  try {
    if (!payload) throw new Error("Không nhận được dữ liệu đầu vào.");

    const required = ['hoTenChuRung', 'nguoiDeNghi', 'nguoiNhanTien', 'soTKNhanTien', 'nganHang', 'soHopDong'];
    for (const f of required) {
      if (!payload[f] || String(payload[f]).trim() === "") {
        throw new Error("Thiếu thông tin bắt buộc: " + f);
      }
    }
    const dsPC = (payload.danhSachPhieuCan || []).map(x => String(x).trim()).filter(Boolean);
    if (dsPC.length === 0) throw new Error("Vui lòng chọn ít nhất 1 số phiếu cân.");

    const ss = SpreadsheetApp.getActive();
    const shSrc = ss.getSheetByName(CFG.DNTT_SRC);
    const sh112 = ss.getSheetByName(CFG.DNTT_112);
    if (!shSrc || !sh112) throw new Error("Không tìm thấy sheet nguồn DNTT_GK_DN hoặc DNTT_GK_DN_112.");

    const shPC = openExternalSheet(CFG.PC_SS_ID, CFG.PC_SHEET, "Phiếu Cân");
    const pcData = shPC.getDataRange().getValues();
    const pcMap = utils.buildIndexMap(pcData, 22, true) && (() => {
      const m = new Map();
      for (let i = 0; i < pcData.length; i++) {
        const key = utils.standardize(pcData[i][22]);
        if (key) m.set(key, i);
      }
      return m;
    })();

    // 1. Kiểm tra & khóa tạm từng phiếu cân được chọn (chống dùng trùng)
    let tongKLKg = 0;
    const rowsToLock = [];
    for (const soP of dsPC) {
      const key = utils.standardize(soP);
      const idx = pcMap.get(key);
      if (idx === undefined) throw new Error(`Số phiếu cân "${soP}" không tồn tại trong hệ thống Phiếu Cân.`);
      const pcRow = pcData[idx];
      const trangThai = String(pcRow[26] || "").trim();
      const daKhoa = String(pcRow[27] || "").trim().toUpperCase();
      if (trangThai || daKhoa === "Y" || daKhoa === "N") {
        throw new Error(`Số phiếu cân "${soP}" đã được sử dụng cho một đề nghị thanh toán khác, không thể dùng lại.`);
      }
      tongKLKg += utils.parseNum(pcRow[9]);
      rowsToLock.push(idx);
    }

    // 2. Tính Lần thanh toán + Ngày dự kiến TT theo khung giờ hiện tại
    const session = getSessionInfo();
    const now = new Date();

    // 3. Sinh ID_KEY duy nhất
    const existingIdsRaw = shSrc.getLastRow() > 1
      ? shSrc.getRange(2, 1, shSrc.getLastRow() - 1, 1).getValues().flat()
      : [];
    const existingIds = new Set(existingIdsRaw.map(v => String(v).trim()).filter(Boolean));
    let newId;
    do { newId = Utilities.getUuid().split('-')[0]; } while (existingIds.has(newId));

    let userEmail = "";
    try { userEmail = Session.getActiveUser().getEmail() || ""; } catch (e) {}

    // 4. Ghi vào DNTT_GK_DN (nguồn) - đúng 18 cột như schema hiện có
    const newSrcRow = new Array(18).fill("");
    newSrcRow[0] = newId;
    newSrcRow[1] = now;
    newSrcRow[2] = userEmail;
    newSrcRow[3] = payload.hoTenChuRung;
    newSrcRow[4] = "'" + (payload.cccdChuRung || "");
    newSrcRow[5] = payload.nguoiDeNghi;
    newSrcRow[6] = payload.chuRungUyQuyen || "Không";
    newSrcRow[7] = payload.nguoiNhanTien;
    newSrcRow[8] = "'" + payload.soTKNhanTien;
    newSrcRow[9] = tongKLKg;
    newSrcRow[10] = dsPC.join(", ");
    newSrcRow[11] = now;
    newSrcRow[12] = newId;
    newSrcRow[13] = "'" + payload.soHopDong;
    newSrcRow[14] = "";
    newSrcRow[15] = "";
    newSrcRow[16] = "";
    newSrcRow[17] = "";
    shSrc.appendRow(newSrcRow);

    // 5. Ghi vào DNTT_GK_DN_112 (ngân hàng) - 23 cột (0..22, ID_DNTT_112 ở cuối)
    const id112 = Utilities.formatDate(now, "GMT+7", "yyyyMMddHHmmss") + ("0000" + Math.floor(Math.random() * 10000)).slice(-4);
    const new112Row = new Array(23).fill("");
    new112Row[0] = newId;
    new112Row[1] = now;
    new112Row[2] = payload.hoTenChuRung;
    new112Row[3] = payload.nguoiNhanTien;
    new112Row[4] = payload.nganHang; // Ngân hàng - đã có sẵn từ form, không cần tra cứu HD_STK
    new112Row[5] = "'" + payload.soTKNhanTien;
    new112Row[6] = ""; // Số tiền - sẽ được điền khi chạy "Tổng Hợp 112"
    new112Row[7] = ""; // Nội dung CK - sẽ được điền khi chạy "Tổng Hợp 112"
    new112Row[8] = "'" + payload.soHopDong;
    new112Row[9] = payload.ngayHopDong ? new Date(payload.ngayHopDong) : "";
    new112Row[10] = payload.chuRungUyQuyen || "Không";
    new112Row[11] = tongKLKg;
    new112Row[12] = "";
    new112Row[13] = "";
    new112Row[14] = "";
    new112Row[15] = "";
    new112Row[16] = now;
    new112Row[17] = "";
    new112Row[18] = session.lan;
    new112Row[19] = session.ngayDuKien;
    new112Row[20] = "";
    new112Row[21] = "";
    new112Row[22] = id112;
    sh112.appendRow(new112Row);

    // 6. Khóa tạm các phiếu cân đã dùng (ghi trực tiếp vào từng dòng - nhanh hơn ghi đè cả sheet)
    rowsToLock.forEach(idx => {
      shPC.getRange(idx + 1, 27, 1, 2).setValues([["Đang lập ĐNTT", "N"]]);
    });

    SpreadsheetApp.flush();
    logAction("TAO_MOI_DNTT", newId, `Tạo mới cho ${payload.hoTenChuRung}, HĐ ${payload.soHopDong}, ${dsPC.length} phiếu cân, KL ${tongKLKg}kg, Lần ${session.lan}`);

    return {
      success: true,
      idKey: newId,
      lan: session.lan,
      ngayDuKien: utils.formatDate(session.ngayDuKien),
      tongKLKg: tongKLKg,
      soLuongPhieu: dsPC.length
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    if (lock) lock.releaseLock();
  }
}

// ============================================================
// MỚI: ĐIỀN LẠI HÀNG LOẠT NGÂN HÀNG CÒN TRỐNG (chạy độc lập, thủ công)
// Dùng khi muốn chạy riêng việc vá dữ liệu Ngân hàng mà KHÔNG cần chạy
// toàn bộ "Tổng Hợp 112" (ví dụ: chỉ mới cập nhật thêm dữ liệu HD_STK).
// Có thể gán vào menu hoặc gọi tay từ Apps Script Editor.
// ============================================================
function runFillMissingBankOnly() {
  const lock = sysLock.acquire();
  try {
    const sh112 = SpreadsheetApp.getActive().getSheetByName(CFG.DNTT_112);
    if (!sh112 || sh112.getLastRow() < 2) return "⚠️ Sheet 112 trống.";

    const data112 = sh112.getRange(1, 1, sh112.getLastRow(), sh112.getLastColumn()).getValues();
    const bankMap = buildBankLookupFromHDSTK();

    let filledBankCount = 0;
    for (let i = 1; i < data112.length; i++) {
      const id = String(data112[i][0]).trim();
      if (!id) continue; // bỏ dòng rỗng
      if (!utils.isBlank(data112[i][4])) continue; // đã có Ngân hàng, bỏ qua

      const soHDRaw = String(data112[i][8] || "").replace(/'/g, "");
      const hdKeyLookup = utils.standardize(soHDRaw);
      const stkKeyLookup = utils.standardize(data112[i][5]);
      if (!hdKeyLookup || !stkKeyLookup) continue;

      const bankFound = bankMap.get(hdKeyLookup + "|" + stkKeyLookup);
      if (bankFound) {
        data112[i][4] = bankFound;
        filledBankCount++;
      }
    }

    if (filledBankCount > 0) {
      sh112.getRange(1, 1, data112.length, data112[0].length).setValues(data112);
      SpreadsheetApp.flush();
      logAction("VA_NGAN_HANG_112", "-", `Đã tự động điền Ngân hàng cho ${filledBankCount} dòng đang trống.`);
      return `✅ Đã tự động điền Ngân hàng cho ${filledBankCount} dòng đang trống (tham chiếu HD_STK).`;
    }
    return "⚠️ Không có dòng nào cần điền Ngân hàng (đã đầy đủ hoặc không khớp được với HD_STK).";
  } catch (e) {
    return "❌ Lỗi: " + e.toString();
  } finally { if (lock) lock.releaseLock(); }
}

// ============================================================
// MENU SHEET
// ============================================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 QUẢN LÝ HAK')
      .addItem('1. Tách Phiếu Chi Tiết', 'runProcessDetail')
      .addItem('2. Tổng Hợp 112', 'runCreate112')
      .addSeparator()
      .addItem('3. Chốt Thanh Toán', 'showPayDialog')
      .addSeparator()
      .addItem('4. Thêm Mới Đề Nghị Thanh Toán', 'showAddPaymentDialog')
      .addSeparator()
      .addItem('5. Vá Ngân Hàng Còn Trống (112)', 'runFillMissingBankOnly')
      .addToUi();
}

// Hộp thoại nhập ngày chốt thanh toán, áp dụng cho TẤT CẢ hồ sơ đã "Đã ĐNTT"
// nhưng chưa thanh toán (Trạng thái TT khác Y). Lưu ý: cách chọn hồ sơ chi tiết
// theo từng bộ lọc vẫn nên thực hiện từ Web App (nút "5. CHỐT & KHÓA"), hộp
// thoại này chỉ dùng cho thao tác nhanh/khẩn cấp ngay trong Google Sheet.
function showPayDialog() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt('Xác nhận chốt', 'Nhập ngày thanh toán (dd/mm/yyyy):', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() != ui.Button.OK) return;

  var payDateStr = result.getResponseText().trim();
  var ss = SpreadsheetApp.getActive();
  var shSrc = ss.getSheetByName(CFG.DNTT_SRC);
  var data = shSrc.getRange(2, 1, Math.max(shSrc.getLastRow() - 1, 0), 18).getValues();

  var selectedIds = data
    .filter(r => !utils.isBlank(r[0]) && String(r[14]) === "Đã ĐNTT" && String(r[17]).toUpperCase() !== "Y")
    .map(r => String(r[0]).trim());

  if (selectedIds.length === 0) {
    ui.alert("Không có hồ sơ nào ở trạng thái 'Đã ĐNTT' đang chờ thanh toán.");
    return;
  }

  var confirm = ui.alert('Xác nhận', `Sẽ chốt thanh toán cho ${selectedIds.length} hồ sơ. Tiếp tục?`, ui.ButtonSet.YES_NO);
  if (confirm != ui.Button.YES) return;

  var msg = runConfirmPayment(selectedIds, payDateStr);
  ui.alert(msg);
}

// Mở hộp thoại "Thêm Mới Đề Nghị Thanh Toán" ngay trong Google Sheet
// (dùng cùng logic với nút trên Web App).
function showAddPaymentDialog() {
  var html = HtmlService.createHtmlOutputFromFile('AddPaymentDialog')
      .setWidth(560).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Thêm Mới Đề Nghị Thanh Toán');
}
