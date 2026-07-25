/**
 * ============================================================
 * HỆ THỐNG QUẢN LÝ THANH TOÁN HAK - PHIÊN BẢN HOÀN THIỆN 2026.5
 * ------------------------------------------------------------
 * *** QUAN TRỌNG - CẦN LÀM TRƯỚC KHI DÙNG BẢN NÀY (chỉ 1 LẦN DUY NHẤT
 * cho cả đời dự án - KHÔNG cần lặp lại kể cả khi dán đè code mới sau
 * này) ***
 * Bản này thêm kiến trúc "FILE NHÁP" (xem mục J bên dưới). Trước khi
 * dùng các nút 1/2/3, bạn CHỈ cần:
 *   1. Mở Google Sheet, dùng menu "🚀 QUẢN LÝ HAK" > "⚙️ Khởi Tạo File
 *      Nháp (chạy 1 lần)" (mục này chỉ hiện khi CHƯA thiết lập) - HOẶC
 *      bấm nút tương ứng trong mục "Cài Đặt" của Web App.
 *   2. XONG. Từ giờ dùng menu/Web App như bình thường.
 * SỬA (mục S - 2026.5): ID File Nháp giờ được LƯU TỰ ĐỘNG vào Script
 * Properties (kho cấu hình riêng của dự án, KHÔNG nằm trong code) thay
 * vì hằng số CFG.DRAFT_SS_ID như bản cũ. Nhờ vậy, dù sau này bạn cập
 * nhật/dán đè TOÀN BỘ code mới, ID File Nháp vẫn được giữ nguyên - KHÔNG
 * còn tình trạng phải "thiết lập lại File Nháp" mỗi khi đổi code nữa.
 * KHÔNG chạy setupDraftSpreadsheet() lần thứ 2 khi ĐÃ có File Nháp (menu
 * sẽ không cho tạo lại đè lên - muốn tạo file MỚI HOÀN TOÀN KHÁC, phải
 * chủ động chạy resetDraftSsId_() trong Apps Script Editor trước).
 * ------------------------------------------------------------
 * THAY ĐỔI SO VỚI BẢN 2026.3:
 *
 * J) KIẾN TRÚC "FILE NHÁP" (DRAFT) RIÊNG BIỆT - THAY ĐỔI LỚN:
 *    - MỤC TIÊU: cho phép xem/sửa/xóa thoải mái ở giai đoạn "Tách
 *      Phiếu" + "Tổng Hợp 112" mà KHÔNG ảnh hưởng dữ liệu chính thức
 *      (DNTT_GK_DN_CT, DNTT_GK_DN_112) hay khóa Phiếu Cân thật. Dữ
 *      liệu chỉ được ghi CHÍNH THỨC (và không thể sửa qua lại nữa) khi
 *      bấm "Chốt Thanh Toán".
 *    - File Nháp là 1 Google Spreadsheet HOÀN TOÀN RIÊNG BIỆT (không
 *      phải tab trong file hiện tại), gồm 2 sheet:
 *        + DNTT_GK_DN_CT_DRAFT  (22 cột - giống schema DNTT_GK_DN_CT)
 *        + DNTT_GK_DN_112_DRAFT (23 cột - giống schema DNTT_GK_DN_112)
 *      Vì đây là 1 Google Sheet bình thường, bạn có thể mở và xem/sửa/
 *      xóa dòng trực tiếp bằng giao diện Sheets - KHÔNG cần thêm giao
 *      diện Web App nào cho việc này.
 *    - Nút 1 "Tách Phiếu": đọc DNTT_GK_DN (nguồn, THẬT) như cũ, nhưng
 *      giờ ghi dòng chi tiết vào DRAFT CT thay vì DNTT_GK_DN_CT thật,
 *      và KHÔNG khóa Phiếu Cân thật. Chỉ APPEND các hồ sơ CHƯA có trong
 *      Draft CT (tránh tạo trùng / đè lên dòng bạn đã sửa tay trong
 *      Nháp) - muốn tách lại 1 hồ sơ, phải xóa hồ sơ đó khỏi Nháp trước
 *      (menu "🗑️ Xóa 1 Hồ Sơ Khỏi Nháp").
 *    - Nút 2 "Tổng Hợp 112": đọc DRAFT CT, tính Số tiền / lũy kế theo
 *      hợp đồng / diễn giải P..., ghi vào DRAFT 112 (không đụng dữ liệu
 *      thật). Lũy kế theo hợp đồng được tính gộp cả lịch sử ĐÃ CHỐT
 *      (DNTT_GK_DN_CT thật) LẪN các hồ sơ khác đang nằm trong Nháp cho
 *      cùng hợp đồng, để "Còn lại" không bị tính sai khi có nhiều hồ sơ
 *      cùng chờ xử lý. Cột Ngân hàng chỉ tự điền khi đang trống (không
 *      ghi đè nếu bạn đã sửa tay trong Nháp); các cột khác (Số tiền,
 *      Nội dung CK, lũy kế...) LUÔN được tính lại mỗi lần bấm nút này
 *      (giống hành vi cũ của bản chính thức).
 *    - "KHÓA TẠM" Số phiếu cân trong lúc còn ở Nháp: KHÔNG dùng bảng
 *      khóa riêng. Một Số phiếu cân được coi là "đang bị giữ" nếu nó
 *      ĐANG XUẤT HIỆN trong DRAFT CT. Vì vậy: xóa dòng khỏi Draft CT
 *      (qua menu xóa, hoặc xóa tay trực tiếp trong Google Sheet) sẽ TỰ
 *      ĐỘNG giải phóng Số phiếu cân đó, không cần bước mở khóa riêng.
 *      findAvailablePhieuCan() và createNewPaymentRequest() đều kiểm
 *      tra thêm điều kiện này trước khi cho phép chọn 1 phiếu cân.
 *    - Nút 3 "Chốt Thanh Toán": đọc các hồ sơ đã chọn TỪ FILE NHÁP (chỉ
 *      chốt được hồ sơ đã có Số tiền > 0 trong Draft 112, tức đã chạy
 *      xong "Tổng Hợp 112" - hồ sơ chưa đủ điều kiện sẽ bị bỏ qua kèm
 *      cảnh báo, không chốt dữ liệu nửa vời). Với các hồ sơ hợp lệ:
 *        a. APPEND dòng chi tiết vào DNTT_GK_DN_CT thật (Ngày CK = ngày
 *           vừa nhập khi chốt).
 *        b. APPEND dòng tổng hợp vào DNTT_GK_DN_112 thật.
 *        c. Cập nhật DNTT_GK_DN (nguồn) thật: Đóng TT / Y / ngày TT.
 *        d. Khóa VĨNH VIỄN các Phiếu Cân liên quan (như bản cũ).
 *        e. XÓA các dòng vừa chốt khỏi File Nháp (Draft CT + Draft
 *           112) - dọn dẹp Nháp, các hồ sơ CHƯA chọn vẫn còn nguyên.
 *    - "Thêm Mới Đề Nghị Thanh Toán": vẫn ghi ngay vào DNTT_GK_DN
 *      (nguồn, THẬT - vì đây chỉ là "đơn xin", không phải số liệu tiền
 *      bạc), nhưng phần chi tiết (từng Số phiếu cân) và placeholder
 *      tổng hợp giờ được tạo THẲNG vào File Nháp (tương đương tự chạy
 *      "Tách Phiếu" cho riêng hồ sơ đó) - nhờ vậy hồ sơ mới tạo lập tức
 *      xuất hiện trong Nháp, đã giữ tạm Số phiếu cân, và có thể bấm
 *      "Tổng Hợp 112" để tính tiền như hồ sơ tách từ nguồn bình thường.
 *    - showPayDialog() (hộp thoại nhanh trong Sheet) giờ chọn danh sách
 *      hồ sơ để chốt DỰA TRÊN DRAFT 112 (Số tiền > 0) thay vì dựa trên
 *      cột trạng thái của nguồn như bản cũ (vì cột đó giờ chỉ mang tính
 *      thông tin, không còn phản ánh "đã sẵn sàng chốt" nữa).
 *    - THÊM MENU: "📂 Mở File Nháp", "🗑️ Xóa 1 Hồ Sơ Khỏi Nháp", và
 *      "⚙️ Khởi Tạo File Nháp" (chỉ hiện khi CFG.DRAFT_SS_ID còn trống).
 *    - GIẢ ĐỊNH CẦN LƯU Ý: với hồ sơ đến từ Google Form (không qua
 *      "Thêm Mới"), code KHÔNG có thông tin "Ngày HĐ" để tự điền vào
 *      Draft 112 khi tự tạo placeholder (trước đây, nếu có 1 trigger/
 *      quy trình khác điền sẵn "Ngày HĐ" thẳng vào DNTT_GK_DN_112 THẬT
 *      trước khi chạy Tổng Hợp 112, quy trình đó KHÔNG có trong file
 *      code này nên không thể tái tạo lại). Bạn có thể bổ sung tay cột
 *      "Ngày HĐ" ngay trong Draft 112 nếu cần, trước khi Chốt Thanh
 *      Toán - vì đây chỉ dùng để hiển thị diễn giải, không ảnh hưởng
 *      tính toán Số tiền.
 *
 * (Các ghi chú thay đổi C-I của bản 2026.3 vẫn còn hiệu lực và đã được
 * giữ nguyên trong các hàm liên quan, không lặp lại chi tiết ở đây.)
 * ============================================================
 */

// ============================================================
// MỚI (mục T - dễ bảo trì): CẤU HÌNH LINK/ID CÁC FILE LIÊN QUAN
// ------------------------------------------------------------
// Gom TOÀN BỘ ID Google Sheet/Drive Folder mà hệ thống cần kết nối tới
// vào 1 CHỖ DUY NHẤT - sau này cần đổi sang file khác (vd thay file
// Phiếu Cân mới, đổi thư mục lưu báo cáo...) chỉ cần sửa đúng 1 dòng ở
// đây, không phải tìm rải rác trong code. Lấy ID từ URL Google Sheet/
// Drive, ví dụ URL "https://docs.google.com/spreadsheets/d/ABC123/edit"
// thì ID là "ABC123".
// ============================================================
const LINKS = {
  PHIEU_CAN_SS_ID: "1vqMVxccBA7zlAMHrGsVBydGFwZJ6QuDZW10zJ74V29g",       // File "Phiếu Cân" (sheet PhieuCan_DN)
  HOP_DONG_SS_ID: "1cv11ORWuAF3Sit4f-kA0xrP6-ab4SF-7LEdkCvGi_gI",        // File "Hợp Đồng" (sheet HD_NCC, HD_STK)
  UPDATE_NGANHANG_SS_ID: "1id7m092q2OVY4dJgcavzPM5HNO14_QIUburVN7y7szg", // File cập nhật Ngân hàng (Update_NganHang_DN)
  BAOCAO_FOLDER_ID: "13_ed2jCP8CgL-UJk0eOn-vIH6LqQElxc"                  // Thư mục Drive lưu file Báo Cáo xuất ra
};

function _sheetUrl_(id) { return id ? `https://docs.google.com/spreadsheets/d/${id}/edit` : ""; }
function _driveFolderUrl_(id) { return id ? `https://drive.google.com/drive/folders/${id}` : ""; }

/** Danh sách link hiển thị trong mục Cài Đặt của Web App - mở nhanh các
 * file liên quan mà không cần lục tìm ID trong code. */
function getConfigLinksForSettings() {
  let draftUrl = "";
  try { draftUrl = SpreadsheetApp.openById(getDraftSsId_()).getUrl(); } catch (e) {}
  return [
    { label: "File Nháp (Draft)", url: draftUrl },
    { label: "File Phiếu Cân", url: _sheetUrl_(LINKS.PHIEU_CAN_SS_ID) },
    { label: "File Hợp Đồng (HD_NCC / HD_STK)", url: _sheetUrl_(LINKS.HOP_DONG_SS_ID) },
    { label: "File Cập Nhật Ngân Hàng", url: _sheetUrl_(LINKS.UPDATE_NGANHANG_SS_ID) },
    { label: "Thư mục Báo Cáo (Drive)", url: _driveFolderUrl_(LINKS.BAOCAO_FOLDER_ID) }
  ];
}

const CFG = {
  PC_SS_ID: LINKS.PHIEU_CAN_SS_ID,
  HD_SS_ID: LINKS.HOP_DONG_SS_ID,
  UPDATE_NH_SS_ID: LINKS.UPDATE_NGANHANG_SS_ID,
  FOLDER_REPORT_ID: LINKS.BAOCAO_FOLDER_ID,
  DNTT_SRC: "DNTT_GK_DN",
  DNTT_CT: "DNTT_GK_DN_CT",
  DNTT_112: "DNTT_GK_DN_112",
  PC_SHEET: "PhieuCan_DN",
  HD_STK_SHEET: "HD_STK",
  SPLIT_COL: "Số phiếu cân",
  LAN_TT: "Lanthanhtoan",
  LOG_SHEET: "NhatKyThaoTac",
  COMPANY_BANK_ACCOUNT: "8619299999",
  COMPANY_BANK_NAME: "BIDV",
  COMPANY_BANK_CODE: 43,

  // ===== MỚI (mục J): FILE NHÁP (DRAFT) =====
  // SỬA (mục S - 2026.5): ID File Nháp THẬT SỰ giờ được lưu tự động ở
  // Script Properties (PropertiesService), KHÔNG còn nằm ở dòng này -
  // nhờ vậy dán đè code mới sau này sẽ KHÔNG làm mất File Nháp đã tạo.
  // Để trống "" là bình thường - CHỈ giữ lại để tương thích ngược cho ai
  // đã lỡ dán tay ID vào đây ở bản cũ (sẽ tự đồng bộ sang Script
  // Properties trong lần chạy đầu). Không cần sửa dòng này nữa.
  DRAFT_SS_ID: "",
  DRAFT_CT_SHEET: "DNTT_GK_DN_CT_DRAFT",
  DRAFT_112_SHEET: "DNTT_GK_DN_112_DRAFT",

  // ===== MỚI (mục P - tối ưu tốc độ): CACHE PHIẾU CÂN CHƯA THANH TOÁN =====
  // Sheet này nằm CHUNG trong File Nháp (DRAFT_SS_ID ở trên), được tạo tự
  // động (không cần chạy setup riêng) - xem chi tiết ở khối mục P.
  DRAFT_PC_SHEET: "PhieuCan_DN_CHUA_TT_DRAFT",

  // ===== MỚI (mục Q/R - tối ưu tốc độ): GIỚI HẠN KHOẢNG NGÀY BÁO CÁO =====
  // Áp cho Báo Cáo / Công Nợ để tránh tải/hiển thị quá nhiều dòng cùng lúc.
  MAX_REPORT_RANGE_DAYS: 95 // ~3 tháng, chừa dư vài ngày
};

// ============================================================
// MỚI (mục U): QUY TRÌNH DUYỆT 3 BƯỚC TRONG DANH SÁCH ĐNTT
// ------------------------------------------------------------
// Draft 112 (DNTT_GK_DN_112_DRAFT) có thêm CỘT THỨ 24 (index 23, 0-based)
// "Trạng Thái ĐNTT" - CHỈ tồn tại trong File Nháp, KHÔNG có ở sheet
// DNTT_GK_DN_112 THẬT (vẫn giữ nguyên 23 cột như cũ).
// Giá trị có thể có: "" (Chờ ĐNTT) hoặc "Đang ĐNTT" (đã Xác Nhận, chờ
// Duyệt/Đóng Thanh Toán). 3 trạng thái hiển thị cho người dùng:
//   1. "Chưa ĐNTT"  - Số tiền (cột 6) <= 0, chưa chạy "Tổng Hợp 112".
//   2. "Chờ ĐNTT"   - Số tiền > 0 (đã tính xong) NHƯNG cột 23 rỗng.
//   3. "Đang ĐNTT"  - Số tiền > 0 VÀ cột 23 = "Đang ĐNTT" (đã bấm "Xác
//      Nhận") - ở bước này mới được phép in Báo Cáo ĐNTT & bấm "Duyệt"
//      để Đóng Thanh Toán (copy Nháp -> sheet chính).
// ============================================================
const COL_TRANG_THAI_DNTT = 23;

/** MỚI (mục W): CHỈ cho phép Sửa/Xóa hồ sơ khi đang ở "Chờ ĐNTT" - tức
 * ĐÃ tính tiền (Số tiền > 0) NHƯNG CHƯA Xác Nhận ("Đang ĐNTT"). Hồ sơ
 * "Chưa ĐNTT" (chưa tính tiền) và "Đang ĐNTT" (đã Xác Nhận) đều bị khóa
 * Sửa/Xóa - phải "Về Chờ ĐNTT" (runHuyXacNhanDNTT) trước nếu cần sửa. */
function _isRecordEditable_(row112) {
  return utils.parseNum(row112[6]) > 0 && String(row112[COL_TRANG_THAI_DNTT] || "").trim() !== "Đang ĐNTT";
}

const utils = {
  standardize: (v) => {
    if (!v && v !== 0) return "";
    let s = String(v).replace(/'/g, '').replace(/\s+/g, '').toUpperCase().trim();
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
   * Build Map tra cứu từ 1 mảng 2 chiều, tự động BỎ QUA các dòng có khóa
   * (cột colIdx) rỗng. LƯU Ý: hàm này KHÔNG tự loại dòng tiêu đề - người
   * gọi phải tự truyền mảng đã bỏ header (vd rows.slice(1)).
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
      if (action === "tach_phieu") {
        return ContentService.createTextOutput(JSON.stringify({ status: "success", message: runProcessDetail() })).setMimeType(ContentService.MimeType.JSON);
      }
      if (action === "lap_de_nghi") {
        return ContentService.createTextOutput(JSON.stringify({ status: "success", message: runCreate112() })).setMimeType(ContentService.MimeType.JSON);
      }
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

// ============================================================
// MỚI (mục J): TRUY CẬP FILE NHÁP
// ============================================================
// ============================================================
// SỬA (mục S - FIX LỖI MẤT FILE NHÁP KHI DÁN CODE MỚI): ID File Nháp
// giờ đọc/ghi vào SCRIPT PROPERTIES (kho cấu hình riêng của dự án Apps
// Script, KHÔNG nằm trong code) thay vì hằng số CFG.DRAFT_SS_ID trong
// code. Nhờ vậy, dù sau này dán ĐÈ code mới (cập nhật bản mới) bao nhiêu
// lần, ID File Nháp vẫn được giữ nguyên - KHÔNG cần "Khởi Tạo File Nháp"
// lại và KHÔNG cần dán tay ID vào CFG nữa.
// Vẫn ưu tiên đọc CFG.DRAFT_SS_ID nếu bạn đã lỡ dán tay ID vào đó trước
// đây (tương thích ngược) - script sẽ tự đồng bộ ID đó vào Script
// Properties trong lần chạy đầu tiên.
// ============================================================
function getDraftSsId_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('DRAFT_SS_ID');
  if (!id && CFG.DRAFT_SS_ID) {
    // Tương thích ngược: nếu trước đây đã dán tay vào CFG, đồng bộ 1 lần
    // sang Script Properties để từ nay không phụ thuộc code nữa.
    id = CFG.DRAFT_SS_ID;
    props.setProperty('DRAFT_SS_ID', id);
  }
  return id || "";
}

function setDraftSsId_(id) {
  PropertiesService.getScriptProperties().setProperty('DRAFT_SS_ID', id);
}

function getDraftSheets_() {
  const draftSsId = getDraftSsId_();
  if (!draftSsId) {
    throw new Error('Chưa thiết lập File Nháp. Vào menu "🚀 QUẢN LÝ HAK" > "⚙️ Khởi Tạo File Nháp (chạy 1 lần)" (hoặc mục Cài Đặt trên Web App) để tạo - hệ thống sẽ tự lưu lại, không cần thiết lập lại lần sau.');
  }
  let ss;
  try {
    ss = SpreadsheetApp.openById(draftSsId);
  } catch (e) {
    throw new Error("Không thể mở File Nháp (kiểm tra lại quyền truy cập, hoặc File Nháp đã bị xóa): " + e.toString());
  }
  const shCT = ss.getSheetByName(CFG.DRAFT_CT_SHEET);
  const sh112 = ss.getSheetByName(CFG.DRAFT_112_SHEET);
  if (!shCT || !sh112) {
    throw new Error(`File Nháp thiếu sheet "${CFG.DRAFT_CT_SHEET}" hoặc "${CFG.DRAFT_112_SHEET}". Kiểm tra lại file hoặc chạy lại setupDraftSpreadsheet().`);
  }
  _ensureDraft112Schema_(sh112); // mục U: tự thêm cột "Trạng Thái ĐNTT" nếu File Nháp được tạo từ bản cũ (23 cột)
  return { ss, shCT, sh112 };
}

/** MỚI (mục U): tự vá thêm cột 24 "Trạng Thái ĐNTT" cho File Nháp đã
 * tạo từ bản trước (chỉ có 23 cột) - để không ai phải chạy lại
 * setupDraftSpreadsheet() hay tạo File Nháp mới khi cập nhật code. */
function _ensureDraft112Schema_(sh112) {
  try {
    if (sh112.getLastColumn() < 24) {
      sh112.getRange(1, 24).setValue("Trạng Thái ĐNTT").setFontWeight("bold").setBackground("#fce5cd");
    }
  } catch (e) {
    // Không chặn luồng chính nếu không thêm được cột (vd thiếu quyền) -
    // các hàm đọc/ghi vẫn hoạt động, chỉ là cột mới có thể trống tiêu đề.
  }
}

/**
 * MỚI (mục J): dựng 1 dòng dữ liệu theo đúng schema DNTT_GK_DN_CT_DRAFT
 * (22 cột) - dùng chung cho "Tách Phiếu" và "Thêm Mới Đề Nghị Thanh Toán".
 */
function buildDraftCtRow_(idKey, stt, soP, srcInfo, pcMap) {
  const newID = idKey + "-" + stt;
  const pcRow = pcMap.get(utils.standardize(soP));

  const row = new Array(22).fill("");
  row[0] = newID;
  row[1] = idKey;
  row[2] = srcInfo.timestamp;
  row[3] = srcInfo.chuRung;
  row[4] = "'" + (srcInfo.cccd || "");
  row[5] = srcInfo.nguoiDN;
  row[6] = srcInfo.nguoiNhan;
  row[7] = "'" + (srcInfo.stkNguoiNhan || "");
  row[8] = srcInfo.klTongNguon;
  row[9] = srcInfo.dsPhieuCanGoc;
  row[10] = stt;
  row[11] = "'" + soP;

  if (pcRow) {
    row[12] = utils.parseNum(pcRow[9]) / 1000;
    row[13] = utils.parseNum(pcRow[19]);
    row[14] = utils.parseNum(pcRow[17]);
    row[15] = row[13] + row[14];
    row[16] = utils.parseNum(pcRow[25]);
    row[17] = "Đã lấy giá từ Phiếu Cân";
  }

  row[18] = "N";
  row[19] = "'" + (srcInfo.soHD || "");
  row[20] = srcInfo.ngayTT || "";
  row[21] = srcInfo.ngayDeNghi || "";

  return { row, found: !!pcRow };
}

// --- DATA VIEW FOR WEB (giữ nguyên như bản gốc) ---
function get112ViewData(fDate, tDate) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.DNTT_112);
  if (!sh || sh.getLastRow() < 2) return [];
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 22).getValues();
  const filtered = data.filter(r => {
    if (utils.isBlank(r[0])) return false;
    let d = r[16];
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
    soHD: String(r[8] || "").replace(/'/g, ""),
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
    if (utils.isBlank(r[0])) return false;
    let d = r[11];
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
// NÚT 6: TẠO BÁO CÁO & SHEET NGÂN HÀNG (giữ nguyên logic gốc - đọc từ
// dữ liệu CHÍNH THỨC, tức chỉ gồm các hồ sơ ĐÃ CHỐT qua File Nháp)
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
    const updateNhInfo = renderSheet2Detail(sheet2, filteredRows, dateRange);

    logAction("XUAT_BAO_CAO", dateRange, `Đã xuất báo cáo ${filteredRows.length} hồ sơ - ${newSS.getUrl()}`);

    return {
      success: true,
      url: newSS.getUrl(),
      // MỚI (mục V): "báo cáo kết xuất file import excel" - Sheet 2 ở
      // trên đã TỰ ĐỘNG xóa & ghi mới toàn bộ file Import Ngân Hàng
      // thật (sheet "Update_NganHang_DN", STK/Ngân hàng công ty lấy từ
      // CFG.COMPANY_BANK_*) - trả kèm link để hiển thị ngay cho người dùng.
      updateNhUrl: updateNhInfo ? updateNhInfo.url : "",
      updateNhCount: updateNhInfo ? updateNhInfo.count : 0
    };
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

    sheet.setColumnWidth(1, 30);
    sheet.setColumnWidth(11, 150);
    sheet.getRange(5, 10, body.length, 2).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

    sheet.getRange(5, 8, body.length, 1).setNumberFormat("#,##0.00");
    sheet.getRange(5, 9, body.length, 1).setNumberFormat("#,##0");

    const lastR = 5 + body.length;
    sheet.getRange(lastR, 1, 1, 7).merge().setValue("TỔNG CỘNG").setFontWeight("bold").setHorizontalAlignment("right");
    sheet.getRange(lastR, 8, 1, 2).setFormulaR1C1(`=SUM(R5C:R[-1]C)`).setFontWeight("bold");
    sheet.getRange(lastR, 1, 1, 11).setBackground("#FFF2CC").setBorder(true, true, true, true, true, true);

    sheet.autoResizeRows(5, body.length);
  }
}

// ============================================================
// SHEET 2: CHI TIẾT VÀ ĐỊNH DẠNG TIME/NUMBER
// ============================================================
function renderSheet2Detail(sheet, filteredRows, dateRange) {
  const ss = SpreadsheetApp.getActive();
  const idSet = new Set(filteredRows.map(r => r.idHeThong));
  const shCT = ss.getSheetByName(CFG.DNTT_CT);
  if (!shCT) throw new Error(`Không tìm thấy sheet "${CFG.DNTT_CT}".`);
  const ctData = shCT.getDataRange().getValues();

  const shPC = openExternalSheet(CFG.PC_SS_ID, CFG.PC_SHEET, "Phiếu Cân");
  const pcData = shPC.getDataRange().getValues();
  const shHD = openExternalSheet(CFG.HD_SS_ID, "HD_NCC", "Hợp Đồng NCC");
  const hdData = shHD.getDataRange().getValues();
  const ssNH = SpreadsheetApp.openById(CFG.UPDATE_NH_SS_ID);
  const shUpdateNH = ssNH.getSheetByName("Update_NganHang_DN");
  // MỚI (mục V): trả về link + số dòng đã ghi vào "Update_NganHang_DN"
  // (file import ngân hàng THẬT, đã có sẵn từ trước) để hiển thị cho
  // người dùng ngay sau khi xuất báo cáo - xem cuối hàm.
  const updateNhInfo = { url: "", count: 0 };

  const mapPC = utils.buildIndexMap(pcData.slice(1), 22, true);
  const mapHD = utils.buildIndexMap(hdData.slice(1), 2, true);

  sheet.clear();
  sheet.getRange("A1:Z1").merge().setValue("BẢNG KÊ CHI TIẾT CHUYỂN KHOẢN").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("A2:Z2").merge().setValue("Thời gian: " + dateRange).setFontStyle("italic").setHorizontalAlignment("center");

  const headers = ["STT", "Lần TT", "Số phiếu cân", "Ngày CK", "Ngày nhập", "Đại lý", "Số hợp đồng", "Số xe", "Họ tên chủ rừng", "CCCD Chủ rừng", "Tên người thụ hưởng", "CCCD Người thụ hưởng", "Địa chỉ Chủ rừng", "Địa chỉ rừng", "Giờ vào", "Giờ ra", "KL hàng + xe", "Bì", "KL_KG", "KL_Tấn", "Đơn giá", "Thành tiền", "Nguồn gốc", "Ngân hàng", "Số tài khoản", "Thời gian nhập liệu"];
  sheet.getRange(4, 1, 1, headers.length).setValues([headers]).setBackground("#cfe2f3").setFontWeight("bold").setBorder(true, true, true, true, true, true).setHorizontalAlignment("center");

  let tempRows = [];
  ctData.slice(1).forEach(ctRow => {
    if (utils.isBlank(ctRow[0]) || utils.isBlank(ctRow[1])) return;
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

    sheet.getRange(5, 15, finalRows.length, 2).setNumberFormat("HH:mm");
    sheet.getRange(5, 17, finalRows.length, 3).setNumberFormat("#,##0");
    sheet.getRange(5, 20, finalRows.length, 1).setNumberFormat("#,##0.000");
    sheet.getRange(5, 21, finalRows.length, 2).setNumberFormat("#,##0");

    const totalRow = lastR + 1;
    sheet.getRange(totalRow, 1, 1, 16).merge().setValue("TỔNG CỘNG").setFontWeight("bold").setHorizontalAlignment("right");
    sheet.getRange(totalRow, 17, 1, 3).setFormulaR1C1(`=SUM(R5C:R[-1]C)`);
    sheet.getRange(totalRow, 20).setFormulaR1C1(`=SUM(R5C:R[-1]C)`);
    sheet.getRange(totalRow, 22).setFormulaR1C1(`=SUM(R5C:R[-1]C)`);
    sheet.getRange(totalRow, 1, 1, headers.length).setBackground("#fff2cc").setBorder(true, true, true, true, true, true);

    if (shUpdateNH) {
      if (shUpdateNH.getLastRow() > 1) {
        shUpdateNH.getRange(2, 1, shUpdateNH.getLastRow() - 1, 33).clearContent();
      }

      let bankUpdateRows = tempRows.map(item => {
        let b = item.bank;
        let row = new Array(33).fill("");
        row[0] = "";
        row[1] = 0;
        row[2] = b.ngayCK;
        row[3] = b.ngayCK;
        row[4] = "'" + b.soPhieuCan;
        row[5] = "'" + CFG.COMPANY_BANK_ACCOUNT; row[6] = CFG.COMPANY_BANK_NAME; row[7] = CFG.COMPANY_BANK_CODE;
        row[8] = b.noiDung || ("Thanh toán phiếu cân " + b.soPhieuCan);
        row[9] = "'" + String(b.hd[6] || ""); row[10] = b.hd[4]; row[11] = b.hd[5];
        row[12] = "'" + String(b.hd[15] || ""); row[13] = b.hd[16]; row[14] = b.hd[10];
        row[15] = "'" + String(b.hd[11] || ""); row[19] = "VND"; row[21] = b.noiDung;
        row[22] = "33111"; row[23] = "1121"; row[24] = b.thanhTien; row[32] = "'" + b.soHD;
        return row;
      });
      shUpdateNH.getRange(2, 1, bankUpdateRows.length, 33).setValues(bankUpdateRows);
      updateNhInfo.url = ssNH.getUrl() + "#gid=" + shUpdateNH.getSheetId();
      updateNhInfo.count = bankUpdateRows.length;
    }
  }
  return updateNhInfo;
}

// ============================================================
// TRA CỨU NGÂN HÀNG THEO (SỐ HỢP ĐỒNG + SỐ TÀI KHOẢN) TỪ HD_STK
// ============================================================
function buildBankLookupFromHDSTK() {
  const map = new Map();
  try {
    _hdStkData_().forEach(row => {
      const soHD = utils.standardize(row[8]);
      const soTK = utils.standardize(row[5]);
      const nganHang = String(row[6] || "").trim();
      if (!soHD || !soTK || !nganHang) return;
      map.set(soHD + "|" + soTK, nganHang);
    });
  } catch (e) {
    // Không có sheet HD_STK hoặc không truy cập được -> bỏ qua, không chặn luồng chính
  }
  return map;
}

// ============================================================
// NÚT 8: CHỐT THANH TOÁN
// SỬA (mục J): giờ đọc từ FILE NHÁP, đẩy dữ liệu về sheet chính thức,
// rồi XÓA phần vừa chốt khỏi Nháp. Chỉ chốt được hồ sơ đã đủ điều kiện
// (đã có Số tiền > 0 trong Draft 112, tức đã chạy xong "Tổng Hợp 112").
// ============================================================
function runConfirmPayment(selectedIds, payDateStr) {
  let lock;
  try {
    lock = sysLock.acquire();
    const ss = SpreadsheetApp.getActive();
    const shSrc = ss.getSheetByName(CFG.DNTT_SRC);
    const shCTReal = ss.getSheetByName(CFG.DNTT_CT);
    const sh112Real = ss.getSheetByName(CFG.DNTT_112);
    const shPC = openExternalSheet(CFG.PC_SS_ID, CFG.PC_SHEET, "Phiếu Cân");
    const { shCT: shDraftCT, sh112: shDraft112 } = getDraftSheets_();

    if (!Array.isArray(selectedIds) || selectedIds.length === 0) {
      return "❌ Lỗi: Không có hồ sơ nào được chọn để chốt thanh toán.";
    }

    const p = String(payDateStr || "").trim().split('/');
    if (p.length !== 3 || isNaN(parseInt(p[0], 10)) || isNaN(parseInt(p[1], 10)) || isNaN(parseInt(p[2], 10))) {
      return "❌ Lỗi: Ngày thanh toán không hợp lệ. Định dạng đúng phải là dd/mm/yyyy.";
    }
    // p[0]=dd, p[1]=mm, p[2]=yyyy - SỬA (mục C, xem 2026.3): ghép đúng
    // thứ tự dd/mm/yyyy, khớp với toàn bộ định dạng ngày còn lại của hệ
    // thống (utils.formatDate).
    const prefixP = p[2] + p[1].padStart(2, '0') + p[0].padStart(2, '0');
    const dateForSheet = p[0] + "/" + p[1] + "/" + p[2];

    const idSet = new Set(selectedIds.map(id => String(id).trim()).filter(id => id !== ""));
    if (idSet.size === 0) return "❌ Lỗi: Danh sách hồ sơ được chọn không hợp lệ (rỗng).";

    // 1. ĐỌC TOÀN BỘ FILE NHÁP
    const draftCTLastRow = shDraftCT.getLastRow();
    const draftCTAll = draftCTLastRow > 1 ? shDraftCT.getRange(2, 1, draftCTLastRow - 1, 22).getValues() : [];
    const draft112LastRow = shDraft112.getLastRow();
    // SỬA (mục U): Draft 112 giờ có 24 cột (thêm cột "Trạng Thái ĐNTT" ở
    // cuối) - sheet DNTT_GK_DN_112 THẬT vẫn giữ nguyên 23 cột như cũ.
    const draft112All = draft112LastRow > 1 ? shDraft112.getRange(2, 1, draft112LastRow - 1, 24).getValues() : [];

    // 2. CHỈ CHỐT HỒ SƠ ĐÃ CÓ Draft 112 VỚI SỐ TIỀN > 0 (đã Tổng Hợp 112
    // đầy đủ) VÀ ĐÃ QUA BƯỚC "Xác Nhận" (Trạng Thái ĐNTT = "Đang ĐNTT") -
    // hồ sơ chưa đủ điều kiện bị BỎ QUA kèm cảnh báo, không chốt dữ liệu
    // nửa vời.
    const draft112ById = new Map();
    draft112All.forEach(r => { const id = String(r[0] || "").trim(); if (id) draft112ById.set(id, r); });

    const validIds = [], skippedNoData = [], skippedNotCalculated = [], skippedNotConfirmed = [];
    idSet.forEach(id => {
      const r112 = draft112ById.get(id);
      if (!r112) { skippedNoData.push(id); return; }
      if (utils.parseNum(r112[6]) <= 0) { skippedNotCalculated.push(id); return; }
      if (String(r112[COL_TRANG_THAI_DNTT] || "").trim() !== "Đang ĐNTT") { skippedNotConfirmed.push(id); return; }
      validIds.push(id);
    });

    if (validIds.length === 0) {
      let msg = "❌ Không có hồ sơ nào đủ điều kiện để chốt.";
      if (skippedNoData.length) msg += ` Không tìm thấy trong File Nháp: ${skippedNoData.join(", ")}.`;
      if (skippedNotCalculated.length) msg += ` Chưa chạy "Tổng Hợp 112": ${skippedNotCalculated.join(", ")}.`;
      if (skippedNotConfirmed.length) msg += ` Chưa "Xác Nhận" ĐNTT: ${skippedNotConfirmed.join(", ")}.`;
      return msg;
    }
    const validIdSet = new Set(validIds);

    // 3. TÁCH DÒNG NHÁP: dòng thuộc hồ sơ được chốt (đẩy về THẬT) và
    // dòng còn lại (giữ nguyên trong Nháp)
    const ctToCommit = [], ctToKeepInDraft = [];
    draftCTAll.forEach(row => {
      const id = String(row[1] || "").trim();
      if (id && validIdSet.has(id)) {
        row[18] = "Y";
        row[20] = dateForSheet;
        ctToCommit.push(row);
      } else {
        ctToKeepInDraft.push(row);
      }
    });

    const the112ToCommit = [], the112ToKeepInDraft = [];
    const mapSoLan = new Map();
    draft112All.forEach(row => {
      const id = String(row[0] || "").trim();
      if (id && validIdSet.has(id)) {
        row[20] = "Y";
        mapSoLan.set(id, String(row[18]).trim());
        the112ToCommit.push(row);
      } else {
        the112ToKeepInDraft.push(row);
      }
    });

    // 4. GHI VÀO DỮ LIỆU THẬT (APPEND)
    if (ctToCommit.length > 0) {
      const startRow = shCTReal.getLastRow() + 1;
      shCTReal.getRange(startRow, 1, ctToCommit.length, 22).setValues(ctToCommit);
    }
    if (the112ToCommit.length > 0) {
      // Sheet 112 THẬT chỉ có 23 cột - cắt bỏ cột "Trạng Thái ĐNTT" (chỉ
      // có ý nghĩa trong lúc còn ở Nháp) trước khi ghi vào THẬT.
      const startRow112 = sh112Real.getLastRow() + 1;
      sh112Real.getRange(startRow112, 1, the112ToCommit.length, 23).setValues(the112ToCommit.map(r => r.slice(0, 23)));
    }

    // 5. CẬP NHẬT NGUỒN (DNTT_GK_DN)
    const dataSrc = shSrc.getRange(1, 1, shSrc.getLastRow(), shSrc.getLastColumn()).getValues();
    for (let i = 1; i < dataSrc.length; i++) {
      const key = String(dataSrc[i][0]).trim();
      if (!key || !validIdSet.has(key)) continue;
      if (dataSrc[i][17] !== "Y") {
        dataSrc[i][14] = "Đóng TT";
        dataSrc[i][15] = prefixP + "_" + (mapSoLan.get(key) || "1");
        dataSrc[i][16] = dateForSheet;
        dataSrc[i][17] = "Y";
      }
    }
    shSrc.getRange(1, 1, dataSrc.length, dataSrc[0].length).setValues(dataSrc);

    // 6. KHÓA VĨNH VIỄN PHIẾU CÂN LIÊN QUAN
    const pcKeySet = new Set();
    ctToCommit.forEach(row => { if (row[11]) pcKeySet.add(utils.standardize(row[11])); });
    const dataPC = shPC.getRange(1, 1, shPC.getLastRow(), shPC.getLastColumn()).getValues();
    for (let m = 1; m < dataPC.length; m++) {
      if (utils.isBlank(dataPC[m][22])) continue;
      if (pcKeySet.has(utils.standardize(dataPC[m][22]))) {
        dataPC[m][24] = "OK"; dataPC[m][26] = "Đóng TT"; dataPC[m][27] = "Y";
      }
    }
    shPC.getRange(1, 1, dataPC.length, dataPC[0].length).setValues(dataPC);
    _invalidatePcCache_();
    // MỚI (mục P): xóa NGAY các phiếu cân vừa Chốt Thanh Toán khỏi cache
    // "chưa thanh toán" trong File Nháp - không cần đợi tới lần làm mới
    // định kỳ, đúng yêu cầu "đóng thanh toán xong thì xóa khỏi Nháp".
    _removeFromPcUnpaidCache_(pcKeySet);

    // 7. DỌN DẸP FILE NHÁP - chỉ giữ lại các hồ sơ CHƯA chốt
    if (draftCTAll.length > 0) {
      shDraftCT.getRange(2, 1, draftCTAll.length, 22).clearContent();
      if (ctToKeepInDraft.length > 0) shDraftCT.getRange(2, 1, ctToKeepInDraft.length, 22).setValues(ctToKeepInDraft);
    }
    if (draft112All.length > 0) {
      shDraft112.getRange(2, 1, draft112All.length, 24).clearContent();
      if (the112ToKeepInDraft.length > 0) shDraft112.getRange(2, 1, the112ToKeepInDraft.length, 24).setValues(the112ToKeepInDraft);
    }

    SpreadsheetApp.flush();

    let msg = `✅ Đã chốt thành công ${validIds.length} hồ sơ, đẩy dữ liệu về sheet chính & khóa Phiếu Cân, đồng thời dọn khỏi File Nháp!`;
    if (skippedNoData.length) msg += ` ⚠️ Không tìm thấy trong Nháp (bỏ qua): ${skippedNoData.join(", ")}.`;
    if (skippedNotCalculated.length) msg += ` ⚠️ Chưa chạy Tổng Hợp 112 (bỏ qua): ${skippedNotCalculated.join(", ")}.`;
    if (skippedNotConfirmed.length) msg += ` ⚠️ Chưa "Xác Nhận" ĐNTT (bỏ qua): ${skippedNotConfirmed.join(", ")}.`;

    logAction("CHOT_THANH_TOAN", validIds.join(","), `Chốt ${validIds.length} hồ sơ từ File Nháp, ngày TT ${dateForSheet}.`);
    return msg;
  } catch (e) {
    return "❌ Lỗi: " + e.toString();
  } finally { if (lock) lock.releaseLock(); }
}

/**
 * NÚT 1: TÁCH PHIẾU CHI TIẾT
 * SỬA (mục J): giờ ghi vào FILE NHÁP thay vì DNTT_GK_DN_CT thật, và
 * KHÔNG khóa Phiếu Cân thật nữa. Chỉ xử lý các hồ sơ CHƯA có trong Draft
 * CT (tránh tạo trùng / đè lên dòng đã sửa tay trong Nháp).
 */
function runProcessDetail() {
  let lock;
  try {
    lock = sysLock.acquire();
    const ss = SpreadsheetApp.getActive();
    const shSrc = ss.getSheetByName(CFG.DNTT_SRC);
    const { shCT: shDraftCT } = getDraftSheets_();

    const srcData = shSrc.getDataRange().getValues();
    // SỬA (mục P - tối ưu tốc độ): pcMap chỉ cần tra cứu phiếu cân CHƯA
    // thanh toán (đúng đối tượng được phép tách) -> dùng cache nhỏ, nhanh.
    const pcMap = utils.buildIndexMap(_pcUnpaidData_(), 22, true);

    // Hồ sơ đã có trong Draft CT -> đã tách rồi, bỏ qua (tránh tạo
    // trùng / đè lên dòng có thể đã bị sửa tay). Cũng lấy luôn danh sách
    // Số phiếu cân đang "bị giữ" bởi BẤT KỲ hồ sơ nào trong Nháp.
    const draftLastRow = shDraftCT.getLastRow();
    const draftIdSet = new Set();
    const heldPhieuCan = new Set();
    if (draftLastRow > 1) {
      shDraftCT.getRange(2, 1, draftLastRow - 1, 12).getValues().forEach(r => {
        if (!utils.isBlank(r[1])) draftIdSet.add(String(r[1]).trim());
        if (!utils.isBlank(r[11])) heldPhieuCan.add(utils.standardize(r[11]));
      });
    }

    const splitColIdx = srcData[0].indexOf(CFG.SPLIT_COL);
    if (splitColIdx === -1) {
      throw new Error(`Không tìm thấy cột "${CFG.SPLIT_COL}" trong sheet ${CFG.DNTT_SRC}. Vui lòng kiểm tra lại tiêu đề cột.`);
    }

    let newDraftRows = [];
    let countNewIds = 0;
    let countSkippedEmptyPC = 0, skippedIds = [];
    let countSkippedHeld = 0, heldWarnings = [];

    for (let i = 1; i < srcData.length; i++) {
      let r = srcData[i];
      let idKey = String(r[0]).trim();
      if (!idKey) continue;
      if (String(r[17]).toUpperCase() === "Y" || String(r[14]) === "Đóng TT") continue; // đã chốt thật rồi
      if (draftIdSet.has(idKey)) continue; // đã có trong Nháp, bỏ qua

      let listPC = String(r[splitColIdx] || "").split(/[,;|]/).map(x => x.trim()).filter(Boolean);
      if (listPC.length === 0) {
        countSkippedEmptyPC++;
        if (skippedIds.length < 20) skippedIds.push(idKey);
        continue;
      }

      const srcInfo = {
        timestamp: r[1], chuRung: r[3], cccd: r[4], nguoiDN: r[5], nguoiNhan: r[7],
        stkNguoiNhan: r[8], klTongNguon: r[9], dsPhieuCanGoc: r[10], soHD: r[13],
        ngayTT: r[16], ngayDeNghi: r[11]
      };

      let anyAdded = false;
      listPC.forEach((soP, idx) => {
        const key = utils.standardize(soP);
        if (heldPhieuCan.has(key)) {
          countSkippedHeld++;
          if (heldWarnings.length < 20) heldWarnings.push(`${soP} (hồ sơ ${idKey})`);
          return;
        }
        const { row } = buildDraftCtRow_(idKey, idx + 1, soP, srcInfo, pcMap);
        newDraftRows.push(row);
        heldPhieuCan.add(key); // giữ ngay trong phạm vi lượt chạy này, tránh tự trùng giữa các hồ sơ đang xử lý cùng lượt
        anyAdded = true;
      });

      if (anyAdded) {
        countNewIds++;
        // Thông tin (KHÔNG dùng để xét điều kiện xử lý - chỉ để dễ nhận
        // biết trạng thái khi xem sheet nguồn).
        srcData[i][14] = "Đang xử lý (Nháp)";
      }
    }

    if (newDraftRows.length > 0) {
      const startRow = shDraftCT.getLastRow() + 1;
      shDraftCT.getRange(startRow, 1, newDraftRows.length, 22).setValues(newDraftRows);
      shSrc.getRange(1, 1, srcData.length, srcData[0].length).setValues(srcData);
      SpreadsheetApp.flush();

      let logDetail = `Đã đưa ${countNewIds} hồ sơ (${newDraftRows.length} dòng chi tiết) vào File Nháp.`;
      if (countSkippedEmptyPC > 0) logDetail += ` Bỏ qua ${countSkippedEmptyPC} hồ sơ thiếu Số phiếu cân.`;
      if (countSkippedHeld > 0) logDetail += ` Bỏ qua ${countSkippedHeld} phiếu cân đang bị giữ bởi hồ sơ Nháp khác.`;
      logAction("TACH_PHIEU_NHAP", "-", logDetail);

      let msg = `✅ Đã đưa ${countNewIds} hồ sơ (${newDraftRows.length} dòng) vào File Nháp để kiểm tra/chỉnh sửa trước khi Tổng Hợp 112.`;
      if (countSkippedEmptyPC > 0) msg += ` ⚠️ Bỏ qua ${countSkippedEmptyPC} hồ sơ thiếu "Số phiếu cân" (ID: ${skippedIds.join(", ")}${countSkippedEmptyPC > skippedIds.length ? "..." : ""}).`;
      if (countSkippedHeld > 0) msg += ` ⚠️ Bỏ qua ${countSkippedHeld} phiếu cân đang bị giữ bởi hồ sơ Nháp khác: ${heldWarnings.join(", ")}${countSkippedHeld > heldWarnings.length ? "..." : ""}.`;
      return msg;
    } else if (countSkippedEmptyPC > 0 || countSkippedHeld > 0) {
      return `⚠️ Không có hồ sơ mới nào được đưa vào Nháp. (Thiếu Số phiếu cân: ${countSkippedEmptyPC}, đang bị giữ bởi Nháp khác: ${countSkippedHeld})`;
    } else {
      return `⚠️ Không có hồ sơ mới cần xử lý (mọi hồ sơ hợp lệ đã có trong File Nháp hoặc đã chốt).`;
    }
  } catch (e) {
    return "❌ Lỗi: " + e.toString();
  } finally { if (lock) lock.releaseLock(); }
}

/**
 * NÚT 2: TỔNG HỢP 112
 * SỬA (mục J): giờ đọc/ghi hoàn toàn trong FILE NHÁP (Draft CT -> Draft
 * 112). Lũy kế theo hợp đồng gộp cả CT THẬT (lịch sử đã chốt) lẫn CT
 * NHÁP hiện tại (các hồ sơ khác đang chờ xử lý cho cùng hợp đồng).
 * Không đụng đến dữ liệu chính thức hay khóa Phiếu Cân.
 */
function runCreate112() {
  let lock;
  try {
    lock = sysLock.acquire();
    const ss = SpreadsheetApp.getActive();
    const shCTReal = ss.getSheetByName(CFG.DNTT_CT); // chỉ ĐỌC - lấy lịch sử "Đã trả"
    const { shCT: shDraftCT, sh112: shDraft112 } = getDraftSheets_();

    const fmt = (num) => utils.parseNum(num).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

    const readSheet = (sh) => {
      const lr = sh.getLastRow();
      return lr > 1 ? sh.getRange(1, 1, lr, sh.getLastColumn()).getValues() : [];
    };

    const draftCTData = readSheet(shDraftCT);
    const draftCTRows = draftCTData.length ? draftCTData.slice(1) : [];
    if (draftCTRows.length === 0) {
      return "⚠️ File Nháp chưa có dòng chi tiết nào (chưa chạy 'Tách Phiếu' hoặc chưa 'Thêm Mới Đề Nghị Thanh Toán').";
    }

    const realCTData = shCTReal ? readSheet(shCTReal) : [];
    const realCTRows = realCTData.length ? realCTData.slice(1) : [];

    const draft112Data = readSheet(shDraft112);
    let draft112Rows = draft112Data.length ? draft112Data.slice(1) : [];

    const draft112MapIdx = new Map();
    draft112Rows.forEach((r, idx) => {
      const id = String(r[0] || "").trim();
      if (id) draft112MapIdx.set(id, idx);
    });

    // Nhóm CT NHÁP theo ID_KEY
    const ctGroupedById = draftCTRows.reduce((acc, row) => {
      const id = String(row[1] || "").trim();
      if (!id) return acc;
      if (!acc[id]) acc[id] = [];
      acc[id].push(row);
      return acc;
    }, {});

    // Nhóm CT theo Số HĐ để tính lũy kế - GỘP CẢ dữ liệu THẬT (lịch sử
    // đã chốt) VÀ dữ liệu NHÁP hiện tại (mục J), tránh tính sai "Còn lại"
    // khi có nhiều hồ sơ đang cùng chờ xử lý cho 1 hợp đồng.
    const ctGroupedByHD = {};
    const addToGroup = (row) => {
      const hdKey = utils.standardize(row[19]);
      if (!hdKey) return;
      if (!ctGroupedByHD[hdKey]) ctGroupedByHD[hdKey] = [];
      ctGroupedByHD[hdKey].push(row);
    };
    realCTRows.forEach(addToGroup);
    draftCTRows.forEach(addToGroup);

    const bankMap = buildBankLookupFromHDSTK();

    // Tự tạo Draft 112 placeholder cho các ID_KEY có trong Draft CT
    // nhưng chưa có dòng 112 tương ứng (vd vừa "Tách Phiếu" xong)
    Object.keys(ctGroupedById).forEach(id => {
      if (draft112MapIdx.has(id)) return;
      const first = ctGroupedById[id][0];
      const session = getSessionInfo();
      const newRow = new Array(24).fill(""); // mục U: +1 cột "Trạng Thái ĐNTT" (mặc định rỗng = Chờ ĐNTT)
      newRow[0] = id;
      newRow[1] = first[2];
      newRow[2] = first[3];
      newRow[3] = first[6];
      newRow[4] = "";
      newRow[5] = String(first[7] || "").replace(/'/g, "");
      newRow[8] = String(first[19] || "").replace(/'/g, "");
      newRow[16] = first[2];
      newRow[18] = session.lan;
      newRow[19] = session.ngayDuKien;
      newRow[22] = id + "_112";
      draft112Rows.push(newRow);
      draft112MapIdx.set(id, draft112Rows.length - 1);
    });

    let count = 0, filledBankCount = 0;

    Object.keys(ctGroupedById).forEach(id => {
      const idx = draft112MapIdx.get(id);
      const row112 = draft112Rows[idx];

      // Tự động điền Ngân hàng NẾU ĐANG TRỐNG - không ghi đè giá trị đã
      // có sẵn (kể cả do người dùng tự sửa tay trong Nháp).
      if (utils.isBlank(row112[4])) {
        const hdKeyLookup = utils.standardize(String(row112[8] || ""));
        const stkKeyLookup = utils.standardize(row112[5]);
        if (hdKeyLookup && stkKeyLookup) {
          const bankFound = bankMap.get(hdKeyLookup + "|" + stkKeyLookup);
          if (bankFound) { row112[4] = bankFound; filledBankCount++; }
        }
      }

      const currentDetails = ctGroupedById[id];
      const soHD = String(row112[8] || "").replace(/'/g, "");
      const hdKey = utils.standardize(soHD);
      const time112 = (row112[1] instanceof Date) ? row112[1].getTime()
        : ((currentDetails[0][2] instanceof Date) ? currentDetails[0][2].getTime() : Date.now());

      const relate = (ctGroupedByHD[hdKey] || []).filter(c => {
        const timeCT = (c[2] instanceof Date) ? c[2].getTime() : 0;
        return timeCT > 0 && timeCT <= time112;
      });

      let slHD_TuDong = 0, daTra_N = 0, kyNay_O = 0, listP = [];
      relate.forEach(c => {
        const kl = utils.parseNum(c[12]);
        slHD_TuDong += kl;
        if (String(c[18]).trim().toUpperCase() === "Y") {
          daTra_N += kl;
        } else {
          kyNay_O += kl;
          listP.push(String(c[11]).replace(/'/g, "").split('/')[0]);
        }
      });

      const tongTien = currentDetails.reduce((sum, d) => sum + utils.parseNum(d[16]), 0);
      const valL_Tan = currentDetails.reduce((sum, d) => sum + utils.parseNum(d[12]), 0);
      const phieuStr = listP.length > 0 ? listP.join(", ") : "Trống";

      row112[6] = tongTien;
      row112[7] = `Thanh toán tiền mua gỗ keo HĐ số ${soHD}`;
      row112[11] = valL_Tan * 1000; // KL Tổng (kg)
      row112[12] = slHD_TuDong;
      row112[13] = daTra_N;
      row112[14] = kyNay_O;
      row112[15] = `Tổng KL: ${fmt(slHD_TuDong)} | Đã trả: ${fmt(daTra_N)} | Còn lại: ${fmt(kyNay_O)} | Đề nghị đợt này: ${fmt(valL_Tan)} | Phiếu: ${phieuStr}`;
      row112[21] = (soHD !== "" && tongTien > 0) ? "Đủ ĐK TT" : "Không đủ ĐK TT";

      count++;
    });

    if (draft112Rows.length > 0) {
      if (shDraft112.getLastRow() > 1) shDraft112.getRange(2, 1, shDraft112.getLastRow() - 1, 24).clearContent();
      shDraft112.getRange(2, 1, draft112Rows.length, 24).setValues(draft112Rows);
      SpreadsheetApp.flush();
    }

    const parts = [];
    if (count > 0) parts.push(`Đã tính toán/cập nhật ${count} hồ sơ trong File Nháp`);
    if (filledBankCount > 0) parts.push(`tự động điền Ngân hàng cho ${filledBankCount} hồ sơ`);
    const detailMsg = (parts.join(", ") || "Không có gì thay đổi") + ".";
    logAction("TONG_HOP_112_NHAP", "-", detailMsg);
    return "✅ " + detailMsg + " Vào File Nháp để kiểm tra/chỉnh sửa trước khi Chốt Thanh Toán.";
  } catch (e) {
    return "❌ Lỗi: " + e.toString();
  } finally { if (lock) lock.releaseLock(); }
}

// ============================================================
// TÍNH LẦN THANH TOÁN + NGÀY DỰ KIẾN THEO KHUNG GIỜ
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
// TÌM SỐ PHIẾU CÂN CÒN KHẢ DỤNG
// SỬA (mục J): loại thêm các Số phiếu cân đang bị "giữ tạm" bởi 1 hồ sơ
// NHÁP khác (đang có mặt trong Draft CT).
// ============================================================
function findAvailablePhieuCan(query) {
  // SỬA (mục P - tối ưu tốc độ): đọc từ cache "chưa thanh toán" (nhỏ,
  // nhanh) thay vì quét toàn bộ PhieuCan_DN thật.
  const data = _pcUnpaidData_();
  const q = utils.standardize(query);

  const heldByDraft = new Set();
  try {
    const { shCT } = getDraftSheets_();
    const lastRow = shCT.getLastRow();
    if (lastRow > 1) {
      shCT.getRange(2, 12, lastRow - 1, 1).getValues().forEach(r => {
        const key = utils.standardize(r[0]);
        if (key) heldByDraft.add(key);
      });
    }
  } catch (e) {
    // Chưa thiết lập File Nháp -> bỏ qua bước lọc này
  }

  const results = [];
  for (let i = 0; i < data.length && results.length < 25; i++) {
    const r = data[i];
    const soP = String(r[22] || "").trim();
    if (!soP) continue;
    const trangThai = String(r[26] || "").trim();
    const daKhoa = String(r[27] || "").trim().toUpperCase();
    if (trangThai || daKhoa === "Y" || daKhoa === "N") continue; // đã dùng cho hồ sơ CHÍNH THỨC khác
    if (heldByDraft.has(utils.standardize(soP))) continue; // đang bị giữ tạm bởi hồ sơ NHÁP khác
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
// TRA CỨU THÔNG TIN HỢP ĐỒNG ĐỂ TỰ ĐIỀN FORM
// ============================================================
function getContractInfo(soHopDong) {
  if (!soHopDong) return null;
  try {
    const key = utils.standardize(soHopDong);
    const row = _hdNccData_().find(r => !utils.isBlank(r[2]) && utils.standardize(r[2]) === key);
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
// MỚI (mục M): CÁC HÀM TRA CỨU THEO CHUỖI PHỤ THUỘC CHO "TẠO ĐỀ NGHỊ
// THANH TOÁN" (Họ tên Chủ rừng -> CCCD -> Người đề nghị/Ủy quyền ->
// Số hợp đồng -> Số phiếu cân -> Người nhận tiền -> Số tài khoản).
// CHỈ SỐ CỘT dưới đây lấy CHÍNH XÁC theo dòng tiêu đề thật của 3 sheet
// (do người dùng cung cấp) - không suy đoán.
// ============================================================
const HDNCC_COL = {
  TIMESTAMP:0, EMAIL:1, SO_HD:2, NGAY_KY:3, HO_TEN:4, DIA_CHI:5, CCCD:6, NGAY_CAP:7,
  NOI_CAP:8, SDT:9, NGUOI_UQ:10, CCCD_UQ:11, NOI_CAP_UQ:12, DIA_CHI_UQ:13, SDT_UQ:14,
  STK:15, NGAN_HANG:16, EMAIL_UQ:17, DIA_CHI_RUNG:18, DT_RUNG:19, LOCATION:20,
  HSO_NGUON_GOC:21, SO_GIAY_TO:22, DT_GPS:23, UY_QUYEN_TT:24, SL_DU_KIEN:25,
  DON_GIA:26, NHOM_KH:27, CHI_NHANH_NH:28, ID_HD:29, TINH_TRANG:30
};
const HDSTK_COL = { ID_HD:0, ID_STK:1, HO_TEN:2, CCCD:3, NGUOI_UQ:4, STK:5, NGAN_HANG:6, UY_QUYEN_TT:7, SO_HD:8, TIMESTAMP:9 };
const PC_COL = { SO_PHIEU:0, NGAY_CAN_1:1, BIEN_SO_1:5, KL_KG:9, NGUON_GOC:10, KHACH_HANG:11, DAI_LY:13, SO_CT:22, DON_GIA_TC:23, TRANG_THAI:24, THANH_TIEN:25, ID_DNTT:26, CHON_TT:27 };

// ============================================================
// MỚI (mục N - tối ưu tốc độ): LỚP CACHE cho 3 sheet tham chiếu
// (HD_NCC, HD_STK, PhieuCan_DN) - vốn được đọc lại RẤT NHIỀU LẦN trong
// cùng 1 phiên làm việc (mỗi bước chọn ở "Tạo ĐNTT" đều cần tra cứu)
// nhưng dữ liệu thay đổi không thường xuyên.
//
// AN TOÀN VỚI KIỂU DATE: CacheService chỉ lưu được chuỗi (string), nếu
// JSON.stringify trực tiếp thì các ô kiểu Date sẽ bị biến thành chuỗi
// ISO và MẤT kiểu Date khi đọc lại (làm hỏng mọi chỗ dùng `instanceof
// Date` trong code, vd lọc theo thời gian). Vì vậy trước khi lưu cache,
// mỗi ô Date được bọc thành {__d: <timestamp>}, và khi đọc lại từ cache
// sẽ được dựng lại đúng thành đối tượng Date như ban đầu.
//
// AN TOÀN VỚI SHEET LỚN: CacheService giới hạn ~100KB/khóa. Nếu dữ liệu
// vượt ngưỡng, hàm chỉ đơn giản BỎ QUA việc lưu cache (không throw lỗi,
// không chặn luồng chính) - hệ thống vẫn chạy đúng, chỉ là không được
// tăng tốc cho riêng sheet đó.
//
// THỜI GIAN SỐNG (TTL) ngắn (90 giây) để hạn chế độ trễ dữ liệu nếu ai
// đó sửa tay trực tiếp trong HD_NCC/HD_STK/PhieuCan_DN ở nơi khác.
// ============================================================
const REF_CACHE_TTL_SECONDS = 90;
const REF_CACHE_MAX_BYTES = 95000; // chừa margin dưới giới hạn ~100KB của CacheService

function _serializeRowsForCache_(rows) {
  return rows.map(row => row.map(cell => (cell instanceof Date) ? { __d: cell.getTime() } : cell));
}
function _deserializeRowsFromCache_(rows) {
  return rows.map(row => row.map(cell =>
    (cell && typeof cell === "object" && cell.__d !== undefined) ? new Date(cell.__d) : cell
  ));
}

function _getCachedRefData_(cacheKey, fetchFn) {
  const cache = CacheService.getScriptCache();
  try {
    const raw = cache.get(cacheKey);
    if (raw) return _deserializeRowsFromCache_(JSON.parse(raw));
  } catch (e) {
    // Lỗi đọc cache (hỏng dữ liệu, hết hạn giữa chừng...) -> đọc trực
    // tiếp từ sheet, không chặn luồng chính.
  }

  const data = fetchFn();

  try {
    const serialized = JSON.stringify(_serializeRowsForCache_(data));
    if (serialized.length < REF_CACHE_MAX_BYTES) {
      cache.put(cacheKey, serialized, REF_CACHE_TTL_SECONDS);
    }
    // Nếu vượt ngưỡng: bỏ qua cache cho lượt này, vẫn trả dữ liệu vừa đọc.
  } catch (e) {
    // Lỗi serialize (hiếm) -> bỏ qua cache, không chặn luồng chính.
  }

  return data;
}

/** Xóa cache PhieuCan_DN ngay sau khi hệ thống ghi đè trạng thái/khóa
 * phiếu cân (vd khi Chốt Thanh Toán) - tránh hiển thị "còn khả dụng"
 * cho phiếu cân vừa bị khóa trong lúc cache còn hạn. */
function _invalidatePcCache_() {
  try { CacheService.getScriptCache().remove("pc_data_v1"); } catch (e) {}
}

function _hdNccData_() {
  return _getCachedRefData_("hdncc_data_v1", () => {
    const sh = openExternalSheet(CFG.HD_SS_ID, "HD_NCC", "Hợp Đồng NCC");
    const lr = sh.getLastRow();
    return lr > 1 ? sh.getRange(2, 1, lr - 1, sh.getLastColumn()).getValues() : [];
  });
}
function _hdStkData_() {
  return _getCachedRefData_("hdstk_data_v1", () => {
    const sh = openExternalSheet(CFG.HD_SS_ID, CFG.HD_STK_SHEET, "Hợp Đồng NCC (HD_STK)");
    const lr = sh.getLastRow();
    return lr > 1 ? sh.getRange(2, 1, lr - 1, sh.getLastColumn()).getValues() : [];
  });
}

// ============================================================
// MỚI (mục Q - tối ưu tốc độ): CHỈ TẢI HỢP ĐỒNG "ĐANG THỰC HIỆN"
// ------------------------------------------------------------
// Các hợp đồng "Đã Thanh lý" / "Đã Hủy" không cần xuất hiện khi TẠO MỚI
// đề nghị thanh toán (gợi ý Chủ rừng, Số hợp đồng, Người nhận tiền...) -
// lọc bớt các dòng này giúp các gợi ý tải nhanh hơn và danh sách gọn
// hơn. CHỈ áp dụng cho các hàm LIỆT KÊ/GỢI Ý bên dưới. Các hàm TRA CỨU
// 1 Số HĐ CỤ THỂ đã biết trước (getContractInfo, getHopDongSummary) và
// các hàm CÔNG NỢ/BÁO CÁO (getDebtByContract...) vẫn dùng _hdNccData_()
// ĐẦY ĐỦ như cũ - để không bỏ sót lịch sử/số liệu của hợp đồng đã xong.
// ============================================================
const HD_TRANG_THAI_LOAI_TRU = new Set(
  ["Đã Thanh lý", "Đã thanh lý", "Đã Hủy", "Đã hủy", "Đã Huỷ", "Đã huỷ"].map(s => utils.standardize(s))
);

function _hdNccActiveData_() {
  return _hdNccData_().filter(r => {
    const tt = utils.standardize(String(r[HDNCC_COL.TINH_TRANG] || ""));
    return !tt || !HD_TRANG_THAI_LOAI_TRU.has(tt);
  });
}
function _pcData_() {
  return _getCachedRefData_("pc_data_v1", () => {
    const sh = openExternalSheet(CFG.PC_SS_ID, CFG.PC_SHEET, "Phiếu Cân");
    const lr = sh.getLastRow();
    return lr > 1 ? sh.getRange(2, 1, lr - 1, sh.getLastColumn()).getValues() : [];
  });
}

// ============================================================
// MỚI (mục P - tối ưu tốc độ): CACHE "PHIẾU CÂN CHƯA THANH TOÁN"
// ------------------------------------------------------------
// LÝ DO: PhieuCan_DN (nguồn) càng ngày càng nhiều dòng (tích lũy nhiều
// năm), nhưng các thao tác "CHỌN phiếu cân" (Tạo Đề Nghị Thanh Toán,
// Tách Phiếu, Thêm phiếu vào 1 hồ sơ Nháp...) CHỈ cần các dòng CHƯA
// thanh toán (ID_DNTT trống, Chọn TT khác Y/N). Quét lại TOÀN BỘ
// PhieuCan_DN mỗi lần (kể cả các dòng đã trả từ nhiều năm trước) là
// nguyên nhân chính khiến Web App tải chậm.
// GIẢI PHÁP: duy trì 1 BẢN SAO thu gọn - CHỈ gồm các dòng CHƯA thanh
// toán - ngay trong File Nháp (sheet CFG.DRAFT_PC_SHEET). Các hàm chọn
// phiếu cân (findAvailablePhieuCan, getAvailablePhieuCanForChuRung,
// pcMap trong runProcessDetail/createNewPaymentRequest/
// addPhieuCanToDraft) đọc từ bản sao NHỎ này thay vì PhieuCan_DN thật.
// QUAN TRỌNG: Công nợ/Báo cáo/Phân tích (getDebtByCustomer,
// getPaymentAnalysis, searchChuRungNames...) CẦN dữ liệu ĐẦY ĐỦ (kể cả
// đã trả) để tính đúng lũy kế - các hàm đó VẪN dùng _pcData_() như cũ,
// KHÔNG đụng tới bản sao này.
// LÀM MỚI (refresh) bản sao:
//  a. Tự động làm mới 1 lần nếu bản sao đang TRỐNG (vd lần đầu dùng).
//  b. Bấm menu "🔄 Làm Mới Cache Phiếu Cân (chưa TT)" trong Google Sheet,
//     hoặc nút tương ứng trong mục Cài Đặt của Web App.
//  c. (Khuyến khích) chạy 1 lần hàm setupPcCacheAutoRefreshTrigger() để
//     hệ thống tự làm mới mỗi 10 phút (Trigger theo thời gian) - dữ liệu
//     "chưa thanh toán" khi đó luôn mới trong vòng tối đa 10 phút.
// Ngoài ra, ngay sau khi "Chốt Thanh Toán" (runConfirmPayment), các
// phiếu cân vừa dùng được XÓA NGAY khỏi bản sao (không cần đợi tới lần
// làm mới định kỳ) - đúng yêu cầu "đóng thanh toán xong thì xóa khỏi
// Nháp".
// ============================================================
const PC_COL_SO_CT_IDX = 22;
const PC_COL_ID_DNTT_IDX = 26;
const PC_COL_CHON_TT_IDX = 27;

function getPcCacheSheet_() {
  const { ss } = getDraftSheets_(); // vẫn yêu cầu File Nháp (CT/112) đã thiết lập
  let sh = ss.getSheetByName(CFG.DRAFT_PC_SHEET);
  if (!sh) sh = ss.insertSheet(CFG.DRAFT_PC_SHEET);
  return sh;
}

/** Quét TOÀN BỘ PhieuCan_DN thật (thao tác nặng - chỉ nên chạy định kỳ,
 * không nên gọi trong luồng xử lý nhanh) và ghi đè bản sao "chưa TT". */
function refreshPhieuCanUnpaidCache_() {
  const shSrcPC = openExternalSheet(CFG.PC_SS_ID, CFG.PC_SHEET, "Phiếu Cân");
  const lastRow = shSrcPC.getLastRow(), lastCol = shSrcPC.getLastColumn();
  const header = lastRow >= 1 ? shSrcPC.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  const all = lastRow > 1 ? shSrcPC.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];

  const unpaid = all.filter(r => {
    const soP = String(r[PC_COL_SO_CT_IDX] || "").trim();
    if (!soP) return false;
    const idDNTT = String(r[PC_COL_ID_DNTT_IDX] || "").trim();
    const chonTT = String(r[PC_COL_CHON_TT_IDX] || "").trim().toUpperCase();
    return !idDNTT && chonTT !== "Y" && chonTT !== "N";
  });

  const sh = getPcCacheSheet_();
  sh.clear();
  if (header.length) sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight("bold").setBackground("#d9d2e9");
  if (unpaid.length) sh.getRange(2, 1, unpaid.length, lastCol).setValues(unpaid);
  sh.setFrozenRows(1);
  CacheService.getScriptCache().remove("pc_unpaid_data_v1");
  logAction("REFRESH_CACHE_PC_CHUA_TT", "-", `Đã làm mới cache Phiếu Cân chưa TT: ${unpaid.length}/${all.length} dòng.`);
  return unpaid.length;
}

/** Dữ liệu Phiếu Cân CHƯA thanh toán - dùng cho mọi chức năng "chọn
 * phiếu cân" để tạo/tách hồ sơ mới (KHÔNG dùng cho công nợ/báo cáo). */
function _pcUnpaidData_() {
  return _getCachedRefData_("pc_unpaid_data_v1", () => {
    const sh = getPcCacheSheet_();
    let lastRow = sh.getLastRow();
    if (lastRow < 2) {
      // Bản sao đang trống (chưa từng làm mới) -> tự động làm mới 1 lần
      // để không trả về rỗng oan cho người dùng đầu tiên trong ngày.
      refreshPhieuCanUnpaidCache_();
      lastRow = sh.getLastRow();
    }
    return lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues() : [];
  });
}

/** Xóa NGAY các phiếu cân vừa được Chốt Thanh Toán khỏi bản sao "chưa
 * TT" (soCTKeySet: Set các "Số phiếu cân" đã utils.standardize()). */
function _removeFromPcUnpaidCache_(soCTKeySet) {
  if (!soCTKeySet || soCTKeySet.size === 0) return;
  try {
    const sh = getPcCacheSheet_();
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;
    const lastCol = sh.getLastColumn();
    const all = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const kept = all.filter(r => !soCTKeySet.has(utils.standardize(r[PC_COL_SO_CT_IDX])));
    if (kept.length === all.length) return; // không có gì cần xóa
    sh.getRange(2, 1, all.length, lastCol).clearContent();
    if (kept.length > 0) sh.getRange(2, 1, kept.length, lastCol).setValues(kept);
    CacheService.getScriptCache().remove("pc_unpaid_data_v1");
  } catch (e) {
    // Không chặn luồng chính nếu dọn cache lỗi - lần làm mới định kỳ kế
    // tiếp (hoặc thủ công) sẽ tự sửa lại đúng.
  }
}

/** Wrapper cho Web App (mục Cài Đặt) - làm mới cache thủ công. */
function webRefreshPhieuCanCache() {
  try {
    const n = refreshPhieuCanUnpaidCache_();
    return { success: true, message: `✅ Đã làm mới cache: hiện có ${n} phiếu cân CHƯA thanh toán sẵn sàng để chọn.` };
  } catch (e) {
    return { success: false, message: "❌ Lỗi: " + e.toString() };
  }
}

/** Chạy 1 LẦN (menu, hoặc Apps Script Editor) để bật tự động làm mới
 * cache Phiếu Cân chưa TT mỗi 10 phút bằng Trigger theo thời gian. */
function setupPcCacheAutoRefreshTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'refreshPhieuCanUnpaidCache_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('refreshPhieuCanUnpaidCache_').timeBased().everyMinutes(10).create();
  const msg = "✅ Đã bật tự động làm mới Cache Phiếu Cân (chưa TT) mỗi 10 phút.";
  logAction("SETUP_TRIGGER_CACHE_PC", "-", msg);
  return msg;
}

// ============================================================
// MỚI (mục O): BÁO CÁO CÔNG NỢ
// ------------------------------------------------------------
// NGUYÊN TẮC TÍNH:
//  - "Giá trị NHẬP" (phát sinh nợ) = Thành tiền trong PhieuCan_DN, tính
//    cho MỌI phiếu cân đã cân xong (Thành tiền > 0), KHÔNG phụ thuộc đã
//    lập ĐNTT hay chưa - vì công nợ phát sinh ngay khi nhận gỗ.
//  - "Đã THANH TOÁN" = tổng Thành tiền (cột 16) của các dòng
//    DNTT_GK_DN_CT THẬT. Trong kiến trúc File Nháp hiện tại, dòng CT
//    THẬT chỉ được tạo ra ĐỒNG THỜI với lúc Chốt Thanh Toán (luôn ở
//    trạng thái đã chốt) - nên tổng CT thật CHÍNH LÀ số đã trả thực tế,
//    dùng đúng "Ngày CK" (ngày chốt) làm mốc thời gian.
//  - "Đang chờ thanh toán" = giá trị đang nằm trong Draft CT (File
//    Nháp) - đã cân, đã gán hồ sơ/hợp đồng, nhưng CHƯA chốt.
//  - "Công nợ" = Giá trị nhập LŨY KẾ (đến hết kỳ) - Đã thanh toán LŨY
//    KẾ (đến hết kỳ). "Phát sinh trong kỳ" chỉ để tham khảo biến động,
//    không dùng để suy ra số dư (vì có thể có nợ tồn từ trước kỳ).
// ============================================================

function _inDateRange_(d, fDate, tDate) {
  if (!(d instanceof Date)) return false;
  const iso = Utilities.formatDate(d, "GMT+7", "yyyy-MM-dd");
  return (!fDate || iso >= fDate) && (!tDate || iso <= tDate);
}
function _onOrBefore_(d, tDate) {
  if (!(d instanceof Date)) return false;
  if (!tDate) return true;
  const iso = Utilities.formatDate(d, "GMT+7", "yyyy-MM-dd");
  return iso <= tDate;
}

/**
 * MỚI (mục R - tối ưu tốc độ): giới hạn khoảng "Từ ngày - Đến ngày" tối
 * đa CFG.MAX_REPORT_RANGE_DAYS (~3 tháng) để Báo Cáo/Công Nợ tải nhanh
 * hơn (đỡ phải trả về + hiển thị quá nhiều dòng cùng lúc). Đây là lớp
 * phòng vệ phía server - giao diện Web App đã tự giới hạn/co dãn khoảng
 * ngày ở phía trình duyệt trước khi gọi lên, hàm này chỉ đảm bảo hành vi
 * đúng dù có gọi thẳng API (vd qua doGet) hay không.
 * LƯU Ý: KHÔNG áp cho các chỉ số LŨY KẾ (klLuyKe/giaTriLuyKe/duCuoiKy...)
 * - các chỉ số đó tính theo "Đến ngày" (tDate) và cần TOÀN BỘ lịch sử
 * trước đó để đúng số liệu công nợ, không phụ thuộc "Từ ngày" (fDate).
 */
function _clampReportRange_(fDate, tDate) {
  if (!fDate || !tDate) return fDate;
  try {
    const f = new Date(fDate + "T00:00:00"), t = new Date(tDate + "T00:00:00");
    const diffDays = Math.round((t - f) / 86400000);
    if (diffDays <= CFG.MAX_REPORT_RANGE_DAYS) return fDate;
    const newF = new Date(t.getTime() - CFG.MAX_REPORT_RANGE_DAYS * 86400000);
    return Utilities.formatDate(newF, "GMT+7", "yyyy-MM-dd");
  } catch (e) {
    return fDate; // ngày không hợp lệ -> để nguyên, các hàm gọi tự xử lý
  }
}

/** #1 + #3: Công nợ theo Khách hàng mua gỗ keo, trong khoảng ngày. */
function getDebtByCustomer(fDate, tDate) {
  fDate = _clampReportRange_(fDate, tDate); // mục R - xem ghi chú tại _clampReportRange_
  const byCustomer = new Map();

  _pcData_().forEach(r => {
    const ten = String(r[PC_COL.KHACH_HANG] || "").trim();
    if (!ten) return;
    const thanhTien = utils.parseNum(r[PC_COL.THANH_TIEN]);
    if (thanhTien <= 0) return;
    const ngay = r[PC_COL.NGAY_CAN_1];
    const kl = utils.parseNum(r[PC_COL.KL_KG]) / 1000;
    const key = utils.standardize(ten);
    if (!byCustomer.has(key)) byCustomer.set(key, { ten, klTrongKy:0, giaTriTrongKy:0, klLuyKe:0, giaTriLuyKe:0, daTTTrongKy:0, daTTLuyKe:0 });
    const o = byCustomer.get(key);
    if (_onOrBefore_(ngay, tDate)) { o.klLuyKe += kl; o.giaTriLuyKe += thanhTien; }
    if (_inDateRange_(ngay, fDate, tDate)) { o.klTrongKy += kl; o.giaTriTrongKy += thanhTien; }
  });

  try {
    const shCTReal = SpreadsheetApp.getActive().getSheetByName(CFG.DNTT_CT);
    if (shCTReal && shCTReal.getLastRow() > 1) {
      shCTReal.getRange(2, 1, shCTReal.getLastRow() - 1, 22).getValues().forEach(r => {
        const ten = String(r[3] || "").trim();
        if (!ten) return;
        const key = utils.standardize(ten);
        const thanhTien = utils.parseNum(r[16]);
        const ngayCK = r[20];
        if (!byCustomer.has(key)) byCustomer.set(key, { ten, klTrongKy:0, giaTriTrongKy:0, klLuyKe:0, giaTriLuyKe:0, daTTTrongKy:0, daTTLuyKe:0 });
        const o = byCustomer.get(key);
        if (_onOrBefore_(ngayCK, tDate)) o.daTTLuyKe += thanhTien;
        if (_inDateRange_(ngayCK, fDate, tDate)) o.daTTTrongKy += thanhTien;
      });
    }
  } catch (e) {}

  return Array.from(byCustomer.values())
    .map(o => ({
      khachHang: o.ten,
      klNhapTrongKy: o.klTrongKy, giaTriNhapTrongKy: o.giaTriTrongKy, daTTTrongKy: o.daTTTrongKy,
      klNhapLuyKe: o.klLuyKe, giaTriNhapLuyKe: o.giaTriLuyKe, daTTLuyKe: o.daTTLuyKe,
      congNo: o.giaTriLuyKe - o.daTTLuyKe
    }))
    .filter(o => o.giaTriNhapTrongKy > 0 || o.daTTTrongKy > 0 || Math.abs(o.congNo) > 0.01)
    .sort((a, b) => b.congNo - a.congNo);
}

/** #2: Công nợ theo Hợp đồng mua gỗ keo. */
function getDebtByContract(fDate, tDate) {
  fDate = _clampReportRange_(fDate, tDate); // mục R - xem ghi chú tại _clampReportRange_
  const byHD = new Map();
  _hdNccData_().forEach(r => {
    const soHD = String(r[HDNCC_COL.SO_HD] || "").trim();
    if (!soHD) return;
    const key = utils.standardize(soHD);
    if (byHD.has(key)) return;
    byHD.set(key, {
      soHD, chuRung: String(r[HDNCC_COL.HO_TEN] || ""), ngayKy: utils.formatDate(r[HDNCC_COL.NGAY_KY]),
      slDuKien: utils.parseNum(r[HDNCC_COL.SL_DU_KIEN]),
      klThucHienLuyKe: 0, giaTriThucHienLuyKe: 0, giaTriTrongKy: 0, giaTriDangCho: 0
    });
  });

  try {
    const shCTReal = SpreadsheetApp.getActive().getSheetByName(CFG.DNTT_CT);
    if (shCTReal && shCTReal.getLastRow() > 1) {
      shCTReal.getRange(2, 1, shCTReal.getLastRow() - 1, 22).getValues().forEach(r => {
        const key = utils.standardize(r[19]);
        if (!key || !byHD.has(key)) return;
        const o = byHD.get(key);
        o.klThucHienLuyKe += utils.parseNum(r[12]);
        o.giaTriThucHienLuyKe += utils.parseNum(r[16]);
        if (_inDateRange_(r[20], fDate, tDate)) o.giaTriTrongKy += utils.parseNum(r[16]);
      });
    }
  } catch (e) {}

  try {
    const { shCT: shDraftCT } = getDraftSheets_();
    const lr = shDraftCT.getLastRow();
    if (lr > 1) {
      shDraftCT.getRange(2, 1, lr - 1, 22).getValues().forEach(r => {
        const key = utils.standardize(r[19]);
        if (!key || !byHD.has(key)) return;
        byHD.get(key).giaTriDangCho += utils.parseNum(r[16]);
      });
    }
  } catch (e) {}

  return Array.from(byHD.values())
    .map(o => ({ ...o, conLaiSanLuong: o.slDuKien - o.klThucHienLuyKe, congNo: o.giaTriDangCho }))
    .filter(o => o.klThucHienLuyKe > 0 || o.giaTriDangCho > 0)
    .sort((a, b) => b.congNo - a.congNo);
}

/**
 * #4: Sổ chi tiết công nợ.
 * type = 'customer': sổ cái Nợ/Có/Số dư lũy kế theo thời gian (chuẩn
 *   kế toán) - Nợ = giá trị nhập (PhieuCan_DN), Có = đã thanh toán (CT
 *   thật).
 * type = 'contract': danh sách tiến độ từng phiếu cân đã gán vào hợp
 *   đồng (CT thật = "Đã thanh toán", Draft CT = "Đang chờ (Nháp)") - vì
 *   CT thật luôn phát sinh Nợ/Có cùng lúc (không có độ trễ), nên trình
 *   bày dạng theo dõi tiến độ dễ hiểu hơn dạng sổ cái thuần túy.
 */
function getDebtLedgerDetail(type, key, tDate) {
  const keyStd = utils.standardize(key);
  if (!keyStd) return { success: false, message: "Thiếu tên khách hàng / số hợp đồng." };

  if (type === 'contract') {
    const rows = [];
    try {
      const shCTReal = SpreadsheetApp.getActive().getSheetByName(CFG.DNTT_CT);
      if (shCTReal && shCTReal.getLastRow() > 1) {
        shCTReal.getRange(2, 1, shCTReal.getLastRow() - 1, 22).getValues().forEach(r => {
          if (utils.standardize(r[19]) !== keyStd) return;
          rows.push({
            ngay: utils.formatDate(r[20]), soPhieuCan: String(r[11] || "").replace(/'/g, ""),
            klTan: utils.parseNum(r[12]), thanhTien: utils.parseNum(r[16]), trangThai: "Đã thanh toán"
          });
        });
      }
    } catch (e) {}
    try {
      const { shCT: shDraftCT } = getDraftSheets_();
      const lr = shDraftCT.getLastRow();
      if (lr > 1) {
        shDraftCT.getRange(2, 1, lr - 1, 22).getValues().forEach(r => {
          if (utils.standardize(r[19]) !== keyStd) return;
          rows.push({
            ngay: utils.formatDate(r[2]), soPhieuCan: String(r[11] || "").replace(/'/g, ""),
            klTan: utils.parseNum(r[12]), thanhTien: utils.parseNum(r[16]), trangThai: "Đang chờ (Nháp)"
          });
        });
      }
    } catch (e) {}
    const tongDaTT = rows.filter(r => r.trangThai === "Đã thanh toán").reduce((s, r) => s + r.thanhTien, 0);
    const tongDangCho = rows.filter(r => r.trangThai === "Đang chờ (Nháp)").reduce((s, r) => s + r.thanhTien, 0);
    return { success: true, mode: "contract", key, rows, tongDaTT, tongDangCho, congNo: tongDangCho };
  }

  if (type === 'customer') {
    const entries = [];
    _pcData_().forEach(r => {
      if (utils.standardize(r[PC_COL.KHACH_HANG]) !== keyStd) return;
      const tien = utils.parseNum(r[PC_COL.THANH_TIEN]);
      if (tien <= 0) return;
      const ngay = r[PC_COL.NGAY_CAN_1];
      if (!_onOrBefore_(ngay, tDate)) return;
      entries.push({ ngay, dienGiai: `Nhập gỗ - Phiếu cân ${String(r[PC_COL.SO_CT] || "").trim()}`, no: tien, co: 0 });
    });
    try {
      const shCTReal = SpreadsheetApp.getActive().getSheetByName(CFG.DNTT_CT);
      if (shCTReal && shCTReal.getLastRow() > 1) {
        shCTReal.getRange(2, 1, shCTReal.getLastRow() - 1, 22).getValues().forEach(r => {
          if (utils.standardize(r[3]) !== keyStd) return;
          const ngay = r[20];
          if (!_onOrBefore_(ngay, tDate)) return;
          entries.push({
            ngay, dienGiai: `Thanh toán - HĐ ${String(r[19] || "").replace(/'/g, "")} - Phiếu ${String(r[11] || "").replace(/'/g, "")}`,
            no: 0, co: utils.parseNum(r[16])
          });
        });
      }
    } catch (e) {}

    entries.sort((a, b) => {
      const ta = a.ngay instanceof Date ? a.ngay.getTime() : 0;
      const tb = b.ngay instanceof Date ? b.ngay.getTime() : 0;
      return ta - tb;
    });

    let du = 0;
    const rows = entries.map(e => {
      du += e.no - e.co;
      return { ngay: utils.formatDate(e.ngay), dienGiai: e.dienGiai, no: e.no, co: e.co, du };
    });

    return {
      success: true, mode: "customer", key, rows,
      tongNo: rows.reduce((s, r) => s + r.no, 0),
      tongCo: rows.reduce((s, r) => s + r.co, 0),
      duCuoiKy: du
    };
  }

  return { success: false, message: "Loại sổ chi tiết không hợp lệ (phải là 'customer' hoặc 'contract')." };
}

/** #5: Bảng phân tích tình hình thanh toán trong khoảng ngày. */
function getPaymentAnalysis(fDate, tDate) {
  fDate = _clampReportRange_(fDate, tDate); // mục R - xem ghi chú tại _clampReportRange_
  let tongKL = 0, tongGiaTri = 0;
  const byNguonGoc = new Map(), byDaiLy = new Map();

  _pcData_().forEach(r => {
    const tien = utils.parseNum(r[PC_COL.THANH_TIEN]);
    if (tien <= 0) return;
    if (!_inDateRange_(r[PC_COL.NGAY_CAN_1], fDate, tDate)) return;
    const kl = utils.parseNum(r[PC_COL.KL_KG]) / 1000;
    tongKL += kl; tongGiaTri += tien;

    const ng = String(r[PC_COL.NGUON_GOC] || "").trim() || "(Chưa rõ)";
    if (!byNguonGoc.has(ng)) byNguonGoc.set(ng, { nhan: ng, klTan: 0, giaTri: 0 });
    const og = byNguonGoc.get(ng); og.klTan += kl; og.giaTri += tien;

    const dl = String(r[PC_COL.DAI_LY] || "").trim() || "(Chưa rõ)";
    if (!byDaiLy.has(dl)) byDaiLy.set(dl, { nhan: dl, klTan: 0, giaTri: 0 });
    const od = byDaiLy.get(dl); od.klTan += kl; od.giaTri += tien;
  });

  let tongDaTT = 0;
  const hopDongDuocTT = new Set();
  try {
    const shCTReal = SpreadsheetApp.getActive().getSheetByName(CFG.DNTT_CT);
    if (shCTReal && shCTReal.getLastRow() > 1) {
      shCTReal.getRange(2, 1, shCTReal.getLastRow() - 1, 22).getValues().forEach(r => {
        if (!_inDateRange_(r[20], fDate, tDate)) return;
        tongDaTT += utils.parseNum(r[16]);
        const hd = utils.standardize(r[19]);
        if (hd) hopDongDuocTT.add(hd);
      });
    }
  } catch (e) {}

  return {
    tongKLNhap: tongKL,
    tongGiaTriNhap: tongGiaTri,
    tongGiaTriDaTT: tongDaTT,
    soLuongHopDongDuocTT: hopDongDuocTT.size,
    theoNguonGoc: Array.from(byNguonGoc.values()).sort((a, b) => b.giaTri - a.giaTri),
    theoDaiLy: Array.from(byDaiLy.values()).sort((a, b) => b.giaTri - a.giaTri)
  };
}


/**
 * #1: Gợi ý Họ tên Chủ rừng - GIAO của PhieuCan_DN (cột Khách hàng) và
 * HD_NCC (cột Họ tên Chủ rừng). Loại các tên chung chung kiểu "KH"/"KL"
 * ra khỏi danh sách gợi ý (đó là khách lẻ, không phải chủ rừng cụ thể).
 */
function searchChuRungNames(query) {
  const q = utils.standardize(query);
  const namesInHD = new Set();
  _hdNccActiveData_().forEach(r => {
    const n = String(r[HDNCC_COL.HO_TEN] || "").trim();
    if (n) namesInHD.add(utils.standardize(n));
  });
  const genericNames = new Set(["KH", "KL", "KHACHLE"]);
  const namesInPC = new Map();
  _pcData_().forEach(r => {
    const n = String(r[PC_COL.KHACH_HANG] || "").trim();
    if (!n) return;
    const std = utils.standardize(n);
    if (genericNames.has(std)) return;
    if (!namesInPC.has(std)) namesInPC.set(std, n);
  });
  const results = [];
  namesInPC.forEach((original, std) => {
    if (!namesInHD.has(std)) return;
    if (q && !std.includes(q)) return;
    results.push(original);
  });
  return results.sort((a, b) => a.localeCompare(b)).slice(0, 30);
}

/** #2: Danh sách CCCD của 1 Họ tên Chủ rừng, tra trong HD_NCC. */
function getCccdListByChuRung(hoTen) {
  const name = utils.standardize(hoTen);
  if (!name) return [];
  const seen = new Set(), out = [];
  _hdNccActiveData_().forEach(r => {
    if (utils.standardize(r[HDNCC_COL.HO_TEN]) !== name) return;
    const cccd = String(r[HDNCC_COL.CCCD] || "").trim();
    if (!cccd || seen.has(cccd)) return;
    seen.add(cccd);
    out.push(cccd);
  });
  return out;
}

/**
 * #3 + #4: Theo Tên + CCCD đã chọn, gợi ý "Người đề nghị" (= Họ và tên
 * người được ủy quyền trong HD_NCC) và "Chủ rừng ủy quyền" (= Ủy quyền
 * thanh toán). Nếu có nhiều hợp đồng cùng Tên+CCCD, lấy hợp đồng có
 * Ngày ký MỚI NHẤT làm gợi ý mặc định - người dùng vẫn có thể sửa tay.
 */
function getNguoiDeNghiInfo(hoTen, cccd) {
  const name = utils.standardize(hoTen), cc = String(cccd || "").trim();
  if (!name || !cc) return null;
  const rows = _hdNccActiveData_().filter(r =>
    utils.standardize(r[HDNCC_COL.HO_TEN]) === name && String(r[HDNCC_COL.CCCD] || "").trim() === cc);
  if (!rows.length) return null;
  rows.sort((a, b) => {
    const ta = a[HDNCC_COL.NGAY_KY] instanceof Date ? a[HDNCC_COL.NGAY_KY].getTime() : 0;
    const tb = b[HDNCC_COL.NGAY_KY] instanceof Date ? b[HDNCC_COL.NGAY_KY].getTime() : 0;
    return tb - ta;
  });
  const r = rows[0];
  return {
    nguoiDeNghi: String(r[HDNCC_COL.NGUOI_UQ] || ""),
    chuRungUyQuyen: String(r[HDNCC_COL.UY_QUYEN_TT] || "").trim() || "Không"
  };
}

/**
 * #5: Số phiếu cân khả dụng cho 1 Chủ rừng. Điều kiện:
 *  - Khách hàng (PhieuCan_DN) khớp đúng Họ tên Chủ rừng ĐÃ CHỌN, HOẶC là
 *    tên chung "KH"/"KL"/"Khách lẻ"/"Khách Lẻ".
 *  - Thành tiền > 0.
 *  - CHƯA có trong DNTT_GK_DN_CT THẬT (đã chốt) và CHƯA có trong Draft
 *    CT (File Nháp) - đối chiếu theo "So_CT".
 *  - Chưa bị khóa ngay trong PhieuCan_DN (như findAvailablePhieuCan cũ).
 */
function getAvailablePhieuCanForChuRung(hoTen, query) {
  const nameStd = utils.standardize(hoTen);
  const genericNames = new Set(["KH", "KL", "KHACHLE"]);
  const q = utils.standardize(query);

  const usedReal = new Set();
  try {
    const shCTReal = SpreadsheetApp.getActive().getSheetByName(CFG.DNTT_CT);
    if (shCTReal && shCTReal.getLastRow() > 1) {
      shCTReal.getRange(2, 12, shCTReal.getLastRow() - 1, 1).getValues().forEach(r => {
        const k = utils.standardize(r[0]); if (k) usedReal.add(k);
      });
    }
  } catch (e) {}

  const heldByDraft = new Set();
  try {
    const { shCT } = getDraftSheets_();
    const lr = shCT.getLastRow();
    if (lr > 1) {
      shCT.getRange(2, 12, lr - 1, 1).getValues().forEach(r => {
        const k = utils.standardize(r[0]); if (k) heldByDraft.add(k);
      });
    }
  } catch (e) {}

  const results = [];
  // SỬA (mục P - tối ưu tốc độ): đọc từ cache "chưa thanh toán" (nhỏ,
  // nhanh) thay vì quét toàn bộ PhieuCan_DN thật.
  const data = _pcUnpaidData_();
  for (let i = 0; i < data.length && results.length < 60; i++) {
    const r = data[i];
    const soCT = String(r[PC_COL.SO_CT] || "").trim();
    if (!soCT) continue;

    const trangThai = String(r[PC_COL.ID_DNTT] || "").trim();
    const daKhoa = String(r[PC_COL.CHON_TT] || "").trim().toUpperCase();
    if (trangThai || daKhoa === "Y" || daKhoa === "N") continue;

    const key = utils.standardize(soCT);
    if (usedReal.has(key) || heldByDraft.has(key)) continue;

    const khach = utils.standardize(r[PC_COL.KHACH_HANG]);
    const khopTen = !!nameStd && khach === nameStd;
    const khopKhachLe = genericNames.has(khach);
    if (!khopTen && !khopKhachLe) continue;

    if (utils.parseNum(r[PC_COL.THANH_TIEN]) <= 0) continue;
    if (q && !utils.standardize(soCT).includes(q)) continue;

    results.push({
      soPhieuCan: soCT,
      ngayNhap: utils.formatDate(r[PC_COL.NGAY_CAN_1]),
      soXe: String(r[PC_COL.BIEN_SO_1] || ""),
      klKg: utils.parseNum(r[PC_COL.KL_KG]),
      donGia: utils.parseNum(r[PC_COL.DON_GIA_TC]),
      thanhTien: utils.parseNum(r[PC_COL.THANH_TIEN]),
      khachHang: String(r[PC_COL.KHACH_HANG] || "")
    });
  }
  return results;
}

/** #6: Người nhận tiền khả dụng theo Tên + CCCD, tra trong HD_STK. */
function getNguoiNhanTienOptions(hoTen, cccd) {
  const name = utils.standardize(hoTen), cc = String(cccd || "").trim();
  if (!name || !cc) return [];
  const seen = new Set(), out = [];
  _hdStkData_().forEach(r => {
    if (utils.standardize(r[HDSTK_COL.HO_TEN]) !== name || String(r[HDSTK_COL.CCCD] || "").trim() !== cc) return;
    const nguoi = String(r[HDSTK_COL.NGUOI_UQ] || "").trim();
    if (!nguoi || seen.has(nguoi)) return;
    seen.add(nguoi);
    out.push(nguoi);
  });
  return out;
}

/** #7: Số tài khoản khả dụng theo Tên + CCCD + Người nhận, tra trong HD_STK. */
function getSoTaiKhoanOptions(hoTen, cccd, nguoiNhan) {
  const name = utils.standardize(hoTen), cc = String(cccd || "").trim(), nguoi = String(nguoiNhan || "").trim();
  if (!name || !cc) return [];
  const seen = new Set(), out = [];
  _hdStkData_().forEach(r => {
    if (utils.standardize(r[HDSTK_COL.HO_TEN]) !== name || String(r[HDSTK_COL.CCCD] || "").trim() !== cc) return;
    if (nguoi && String(r[HDSTK_COL.NGUOI_UQ] || "").trim() !== nguoi) return;
    const stk = String(r[HDSTK_COL.STK] || "").trim();
    if (!stk) return;
    const key = stk + "|" + String(r[HDSTK_COL.NGAN_HANG] || "").trim();
    if (seen.has(key)) return;
    seen.add(key);
    out.push({ stk, nganHang: String(r[HDSTK_COL.NGAN_HANG] || "").trim(), soHD: String(r[HDSTK_COL.SO_HD] || "").trim() });
  });
  return out;
}

/** #12: Số hợp đồng khả dụng theo Tên + CCCD, tra trong HD_NCC. */
function getSoHopDongOptions(hoTen, cccd) {
  const name = utils.standardize(hoTen), cc = String(cccd || "").trim();
  if (!name || !cc) return [];
  const seen = new Set(), out = [];
  _hdNccActiveData_().forEach(r => {
    if (utils.standardize(r[HDNCC_COL.HO_TEN]) !== name || String(r[HDNCC_COL.CCCD] || "").trim() !== cc) return;
    const soHD = String(r[HDNCC_COL.SO_HD] || "").trim();
    if (!soHD || seen.has(soHD)) return;
    seen.add(soHD);
    out.push({
      soHD,
      ngayKy: utils.formatDate(r[HDNCC_COL.NGAY_KY]),
      slDuKien: utils.parseNum(r[HDNCC_COL.SL_DU_KIEN])
    });
  });
  return out;
}

/**
 * MỚI (mục N - tối ưu tốc độ): gộp 3 lượt tra cứu (Người đề nghị/Ủy
 * quyền, Danh sách Số hợp đồng, Danh sách Người nhận tiền) thành 1 lượt
 * gọi google.script.run DUY NHẤT, đọc HD_NCC và HD_STK MỖI SHEET CHỈ 1
 * LẦN thay vì 3 lần riêng rẽ như trước (mỗi hàm getNguoiDeNghiInfo/
 * getSoHopDongOptions/getNguoiNhanTienOptions gọi riêng đều tự mở lại
 * HD_NCC/HD_STK từ đầu). Gọi ngay khi người dùng vừa chọn CCCD, để khi
 * sang Bước 2 dữ liệu đã có sẵn, không phải chờ tải lại.
 */
function getChuRungContext(hoTen, cccd) {
  const name = utils.standardize(hoTen), cc = String(cccd || "").trim();
  const result = { nguoiDeNghi: "", chuRungUyQuyen: "Không", soHopDongOptions: [], nguoiNhanOptions: [] };
  if (!name || !cc) return result;

  const hdRows = _hdNccActiveData_().filter(r =>
    utils.standardize(r[HDNCC_COL.HO_TEN]) === name && String(r[HDNCC_COL.CCCD] || "").trim() === cc);

  if (hdRows.length) {
    const sorted = hdRows.slice().sort((a, b) => {
      const ta = a[HDNCC_COL.NGAY_KY] instanceof Date ? a[HDNCC_COL.NGAY_KY].getTime() : 0;
      const tb = b[HDNCC_COL.NGAY_KY] instanceof Date ? b[HDNCC_COL.NGAY_KY].getTime() : 0;
      return tb - ta;
    });
    const newest = sorted[0];
    result.nguoiDeNghi = String(newest[HDNCC_COL.NGUOI_UQ] || "");
    result.chuRungUyQuyen = String(newest[HDNCC_COL.UY_QUYEN_TT] || "").trim() || "Không";

    const seenHD = new Set();
    hdRows.forEach(r => {
      const soHD = String(r[HDNCC_COL.SO_HD] || "").trim();
      if (!soHD || seenHD.has(soHD)) return;
      seenHD.add(soHD);
      result.soHopDongOptions.push({
        soHD,
        ngayKy: utils.formatDate(r[HDNCC_COL.NGAY_KY]),
        slDuKien: utils.parseNum(r[HDNCC_COL.SL_DU_KIEN])
      });
    });
  }

  const seenNguoi = new Set();
  _hdStkData_().forEach(r => {
    if (utils.standardize(r[HDSTK_COL.HO_TEN]) !== name || String(r[HDSTK_COL.CCCD] || "").trim() !== cc) return;
    const nguoi = String(r[HDSTK_COL.NGUOI_UQ] || "").trim();
    if (!nguoi || seenNguoi.has(nguoi)) return;
    seenNguoi.add(nguoi);
    result.nguoiNhanOptions.push(nguoi);
  });

  return result;
}

/**
 * #13: Tóm lược hợp đồng - Ngày ký / SL dự kiến / SL đã thực hiện (tổng
 * hợp từ DNTT_GK_DN_CT THẬT) / Còn lại hoặc Đã vượt / các lần đã thanh
 * toán (ngày, STK, ngân hàng, số tiền) từ DNTT_GK_DN_112 THẬT đã chốt.
 * stkDangDeNghi (tùy chọn): nếu truyền vào, tính thêm số lần đã chuyển
 * riêng cho đúng tài khoản đang đề nghị lần này.
 */
function getHopDongSummary(soHD, stkDangDeNghi) {
  const key = utils.standardize(soHD);
  if (!key) return null;

  const hdRow = _hdNccData_().find(r => utils.standardize(r[HDNCC_COL.SO_HD]) === key);

  let slThucHien = 0;
  try {
    const shCTReal = SpreadsheetApp.getActive().getSheetByName(CFG.DNTT_CT);
    if (shCTReal && shCTReal.getLastRow() > 1) {
      shCTReal.getRange(2, 1, shCTReal.getLastRow() - 1, 22).getValues().forEach(r => {
        if (utils.standardize(r[19]) === key) slThucHien += utils.parseNum(r[12]);
      });
    }
  } catch (e) {}

  const danhSachLanTT = [];
  try {
    const sh112Real = SpreadsheetApp.getActive().getSheetByName(CFG.DNTT_112);
    if (sh112Real && sh112Real.getLastRow() > 1) {
      sh112Real.getRange(2, 1, sh112Real.getLastRow() - 1, 23).getValues().forEach(r => {
        if (utils.standardize(String(r[8] || "").replace(/'/g, "")) !== key) return;
        if (String(r[20]).toUpperCase() !== "Y") return;
        danhSachLanTT.push({
          ngay: utils.formatDate(r[16]),
          stk: String(r[5] || "").replace(/'/g, ""),
          nganHang: String(r[4] || ""),
          soTien: utils.parseNum(r[6])
        });
      });
    }
  } catch (e) {}

  const slDuKien = hdRow ? utils.parseNum(hdRow[HDNCC_COL.SL_DU_KIEN]) : 0;
  const stkKey = utils.standardize(stkDangDeNghi);
  const soLanChoTKDangDeNghi = stkKey ? danhSachLanTT.filter(x => utils.standardize(x.stk) === stkKey).length : 0;

  return {
    soHD: soHD,
    ngayKy: hdRow ? utils.formatDate(hdRow[HDNCC_COL.NGAY_KY]) : "",
    slDuKien: slDuKien,
    slThucHien: slThucHien,
    conLaiHayVuot: slDuKien - slThucHien, // >=0: còn lại; <0: đã vượt |giá trị|
    soLanDaThanhToan: danhSachLanTT.length,
    danhSachLanThanhToan: danhSachLanTT,
    soLanChoTKDangDeNghi: soLanChoTKDangDeNghi
  };
}

// ============================================================
// THÊM MỚI ĐỀ NGHỊ THANH TOÁN
// SỬA (mục J): nguồn (DNTT_GK_DN) vẫn ghi thẳng vào dữ liệu THẬT (chỉ là
// "đơn xin"), nhưng phần chi tiết (Draft CT) và placeholder (Draft 112)
// giờ tạo THẲNG vào File Nháp - tương đương tự chạy "Tách Phiếu" cho
// riêng hồ sơ này, đồng thời giữ tạm Số phiếu cân ngay lập tức.
// SỬA (mục M): thêm "Ngày đề nghị" là trường BẮT BUỘC do người dùng
// nhập (khác với Timestamp = giờ hệ thống tự động).
// ============================================================
function createNewPaymentRequest(payload) {
  let lock;
  try {
    lock = sysLock.acquire();
    if (!payload) throw new Error("Không nhận được dữ liệu đầu vào.");

    const required = ['hoTenChuRung', 'cccdChuRung', 'nguoiDeNghi', 'nguoiNhanTien', 'soTKNhanTien', 'nganHang', 'soHopDong', 'ngayDeNghi'];
    for (const f of required) {
      if (!payload[f] || String(payload[f]).trim() === "") throw new Error("Thiếu thông tin bắt buộc: " + f);
    }
    const ngayDeNghiDate = new Date(payload.ngayDeNghi);
    if (isNaN(ngayDeNghiDate.getTime())) throw new Error("Ngày đề nghị không hợp lệ.");
    const dsPC = (payload.danhSachPhieuCan || []).map(x => String(x).trim()).filter(Boolean);
    if (dsPC.length === 0) throw new Error("Vui lòng chọn ít nhất 1 số phiếu cân.");

    const ss = SpreadsheetApp.getActive();
    const shSrc = ss.getSheetByName(CFG.DNTT_SRC);
    if (!shSrc) throw new Error("Không tìm thấy sheet nguồn DNTT_GK_DN.");
    const { shCT: shDraftCT, sh112: shDraft112 } = getDraftSheets_();

    // SỬA (mục P - tối ưu tốc độ): chỉ cần tra cứu trong các phiếu cân
    // CHƯA thanh toán (đúng đối tượng được phép chọn) -> dùng cache nhỏ.
    const pcMap = utils.buildIndexMap(_pcUnpaidData_(), 22, true);

    const heldByDraft = new Set();
    const draftLastRow = shDraftCT.getLastRow();
    if (draftLastRow > 1) {
      shDraftCT.getRange(2, 12, draftLastRow - 1, 1).getValues().forEach(r => {
        const k = utils.standardize(r[0]);
        if (k) heldByDraft.add(k);
      });
    }

    // 1. Kiểm tra từng phiếu cân được chọn
    let tongKLKg = 0;
    for (const soP of dsPC) {
      const key = utils.standardize(soP);
      const pcRow = pcMap.get(key);
      if (!pcRow) throw new Error(`Số phiếu cân "${soP}" không tồn tại trong hệ thống Phiếu Cân.`);
      const trangThai = String(pcRow[26] || "").trim();
      const daKhoa = String(pcRow[27] || "").trim().toUpperCase();
      if (trangThai || daKhoa === "Y" || daKhoa === "N") {
        throw new Error(`Số phiếu cân "${soP}" đã được sử dụng cho một đề nghị thanh toán CHÍNH THỨC khác.`);
      }
      if (heldByDraft.has(key)) {
        throw new Error(`Số phiếu cân "${soP}" đang được giữ tạm bởi 1 hồ sơ NHÁP khác chưa chốt.`);
      }
      tongKLKg += utils.parseNum(pcRow[9]);
    }

    // 2. Tính Lần thanh toán + Ngày dự kiến TT
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

    // 4. Ghi vào DNTT_GK_DN (nguồn - THẬT)
    const newSrcRow = new Array(18).fill("");
    newSrcRow[0] = newId; newSrcRow[1] = now; newSrcRow[2] = userEmail;
    newSrcRow[3] = payload.hoTenChuRung; newSrcRow[4] = "'" + (payload.cccdChuRung || "");
    newSrcRow[5] = payload.nguoiDeNghi; newSrcRow[6] = payload.chuRungUyQuyen || "Không";
    newSrcRow[7] = payload.nguoiNhanTien; newSrcRow[8] = "'" + payload.soTKNhanTien;
    newSrcRow[9] = tongKLKg; newSrcRow[10] = dsPC.join(", "); newSrcRow[11] = ngayDeNghiDate;
    newSrcRow[12] = newId; newSrcRow[13] = "'" + payload.soHopDong;
    newSrcRow[14] = "Đang xử lý (Nháp)"; newSrcRow[15] = ""; newSrcRow[16] = ""; newSrcRow[17] = "";
    shSrc.appendRow(newSrcRow);

    // 5. Tạo trực tiếp các dòng Draft CT (tương đương "Tách Phiếu" cho
    // riêng hồ sơ này) - đẩy ngay vào File Nháp, giữ tạm Phiếu Cân luôn.
    const srcInfo = {
      timestamp: now, chuRung: payload.hoTenChuRung, cccd: payload.cccdChuRung,
      nguoiDN: payload.nguoiDeNghi, nguoiNhan: payload.nguoiNhanTien,
      stkNguoiNhan: payload.soTKNhanTien, klTongNguon: tongKLKg, dsPhieuCanGoc: dsPC.join(", "),
      soHD: payload.soHopDong, ngayTT: "", ngayDeNghi: ngayDeNghiDate
    };
    const draftRows = dsPC.map((soP, idx) => buildDraftCtRow_(newId, idx + 1, soP, srcInfo, pcMap).row);
    shDraftCT.getRange(shDraftCT.getLastRow() + 1, 1, draftRows.length, 22).setValues(draftRows);

    // 6. Tạo placeholder Draft 112 (Số tiền/Nội dung CK sẽ được tính khi
    // chạy "Tổng Hợp 112")
    const id112 = Utilities.formatDate(now, "GMT+7", "yyyyMMddHHmmss") + ("0000" + Math.floor(Math.random() * 10000)).slice(-4);
    const new112Row = new Array(24).fill(""); // mục U: +1 cột "Trạng Thái ĐNTT" (mặc định rỗng = Chờ ĐNTT)
    new112Row[0] = newId; new112Row[1] = now; new112Row[2] = payload.hoTenChuRung;
    new112Row[3] = payload.nguoiNhanTien; new112Row[4] = payload.nganHang;
    new112Row[5] = "'" + payload.soTKNhanTien; new112Row[6] = ""; new112Row[7] = "";
    new112Row[8] = "'" + payload.soHopDong;
    new112Row[9] = payload.ngayHopDong ? new Date(payload.ngayHopDong) : "";
    new112Row[10] = payload.chuRungUyQuyen || "Không"; new112Row[11] = tongKLKg;
    new112Row[16] = ngayDeNghiDate; new112Row[18] = session.lan; new112Row[19] = session.ngayDuKien;
    new112Row[22] = id112;
    shDraft112.getRange(shDraft112.getLastRow() + 1, 1, 1, 24).setValues([new112Row]);

    SpreadsheetApp.flush();
    logAction("TAO_MOI_DNTT_NHAP", newId, `Tạo mới (vào File Nháp) cho ${payload.hoTenChuRung}, HĐ ${payload.soHopDong}, ${dsPC.length} phiếu cân, KL ${tongKLKg}kg, Lần ${session.lan}`);

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
// MỚI (mục J): XÓA 1 HỒ SƠ KHỎI FILE NHÁP
// Xóa toàn bộ Draft CT + Draft 112 của 1 ID_KEY. Vì "giữ tạm" Số phiếu
// cân được suy ra TRỰC TIẾP từ nội dung Draft CT, xóa xong thì các Số
// phiếu cân của hồ sơ đó TỰ ĐỘNG được giải phóng - không cần bước "mở
// khóa" riêng. KHÔNG đụng đến dữ liệu CHÍNH THỨC.
// ============================================================
function runDeleteDraftRecord(idKey) {
  let lock;
  try {
    lock = sysLock.acquire();
    const id = String(idKey || "").trim();
    if (!id) return "❌ Lỗi: Vui lòng nhập ID hồ sơ cần xóa khỏi Nháp.";

    const { shCT: shDraftCT, sh112: shDraft112 } = getDraftSheets_();

    // MỚI (mục W): CHỈ được xóa khi hồ sơ đang ở "Chờ ĐNTT".
    const c112LastRowCheck = shDraft112.getLastRow();
    if (c112LastRowCheck > 1) {
      const all112Check = shDraft112.getRange(2, 1, c112LastRowCheck - 1, 24).getValues();
      const row112Check = all112Check.find(r => String(r[0] || "").trim() === id);
      if (row112Check && !_isRecordEditable_(row112Check)) {
        return `❌ Không thể xóa: hồ sơ "${id}" chưa/đã qua trạng thái "Chờ ĐNTT". Hãy bấm "Về Chờ ĐNTT" trước nếu đang ở "Đang ĐNTT".`;
      }
    }

    const ctLastRow = shDraftCT.getLastRow();
    let ctRemoved = 0;
    if (ctLastRow > 1) {
      const ctAll = shDraftCT.getRange(2, 1, ctLastRow - 1, 22).getValues();
      const kept = ctAll.filter(r => String(r[1] || "").trim() !== id);
      ctRemoved = ctAll.length - kept.length;
      shDraftCT.getRange(2, 1, ctAll.length, 22).clearContent();
      if (kept.length > 0) shDraftCT.getRange(2, 1, kept.length, 22).setValues(kept);
    }

    const c112LastRow = shDraft112.getLastRow();
    let c112Removed = 0;
    if (c112LastRow > 1) {
      const all112 = shDraft112.getRange(2, 1, c112LastRow - 1, 24).getValues();
      const kept112 = all112.filter(r => String(r[0] || "").trim() !== id);
      c112Removed = all112.length - kept112.length;
      shDraft112.getRange(2, 1, all112.length, 24).clearContent();
      if (kept112.length > 0) shDraft112.getRange(2, 1, kept112.length, 24).setValues(kept112);
    }

    // Reset trạng thái thông tin (Nháp) trên sheet nguồn THẬT (nếu có)
    // để hồ sơ có thể được "Tách Phiếu" lại từ đầu.
    try {
      const shSrc = SpreadsheetApp.getActive().getSheetByName(CFG.DNTT_SRC);
      if (shSrc) {
        const lastRow = shSrc.getLastRow();
        if (lastRow > 1) {
          const idsCol = shSrc.getRange(2, 1, lastRow - 1, 1).getValues();
          for (let i = 0; i < idsCol.length; i++) {
            if (String(idsCol[i][0]).trim() === id) {
              const cell = shSrc.getRange(i + 2, 15);
              if (cell.getValue() === "Đang xử lý (Nháp)") cell.setValue("");
              break;
            }
          }
        }
      }
    } catch (e) { /* không chặn luồng chính */ }

    SpreadsheetApp.flush();

    if (ctRemoved === 0 && c112Removed === 0) {
      return `⚠️ Không tìm thấy hồ sơ "${id}" trong File Nháp.`;
    }
    logAction("XOA_NHAP", id, `Đã xóa ${ctRemoved} dòng CT nháp và ${c112Removed} dòng 112 nháp.`);
    return `✅ Đã xóa hồ sơ "${id}" khỏi File Nháp (${ctRemoved} dòng chi tiết, ${c112Removed} dòng tổng hợp). Các Số phiếu cân liên quan đã được giải phóng.`;
  } catch (e) {
    return "❌ Lỗi: " + e.toString();
  } finally { if (lock) lock.releaseLock(); }
}

// ============================================================
// MỚI (mục K): CÁC HÀM PHỤC VỤ WEB APP (menu từng bước)
// ============================================================

/**
 * Số liệu tổng quan cho Trang chủ.
 */
function getDashboardStats() {
  const result = {
    draftSetup: !!getDraftSsId_(),
    draftCount: 0, draftReady: 0, draftPending: 0, draftTotalTien: 0,
    nguonChoXuLy: 0
  };
  try {
    const list = getDraftListSummary();
    result.draftCount = list.length;
    result.draftReady = list.filter(x => x.sanSangChot).length;
    result.draftPending = result.draftCount - result.draftReady;
    result.draftTotalTien = list.reduce((s, x) => s + (x.sanSangChot ? x.soTien : 0), 0);
  } catch (e) { /* File Nháp chưa thiết lập -> giữ giá trị mặc định */ }

  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(CFG.DNTT_SRC);
    if (sh && sh.getLastRow() > 1) {
      const data = sh.getRange(2, 1, sh.getLastRow() - 1, 18).getValues();
      result.nguonChoXuLy = data.filter(r => !utils.isBlank(r[0]) && String(r[17]).toUpperCase() !== "Y" && String(r[14]) !== "Đóng TT").length;
    }
  } catch (e) {}

  return result;
}

/**
 * Danh sách hồ sơ đang có trong File Nháp (gộp CT + 112), dùng cho màn
 * "Danh Sách Nháp" của Web App.
 */
function getDraftListSummary() {
  const { shCT, sh112 } = getDraftSheets_();
  const ctLastRow = shCT.getLastRow();
  const ctAll = ctLastRow > 1 ? shCT.getRange(2, 1, ctLastRow - 1, 22).getValues() : [];
  const the112LastRow = sh112.getLastRow();
  const the112All = the112LastRow > 1 ? sh112.getRange(2, 1, the112LastRow - 1, 24).getValues() : [];

  const ctGrouped = {};
  ctAll.forEach(r => {
    const id = String(r[1] || "").trim();
    if (!id) return;
    if (!ctGrouped[id]) ctGrouped[id] = [];
    ctGrouped[id].push(r);
  });

  return the112All.filter(r => !utils.isBlank(r[0])).map(r => {
    const id = String(r[0]).trim();
    const details = ctGrouped[id] || [];
    const soTien = utils.parseNum(r[6]);
    const sanSangChot = soTien > 0;
    const daXacNhan = String(r[COL_TRANG_THAI_DNTT] || "").trim() === "Đang ĐNTT";
    // mục U: 3 trạng thái - "cho_tinh" (chưa tính tiền) / "cho_dntt" (đã
    // tính tiền, chưa Xác Nhận) / "dang_dntt" (đã Xác Nhận, chờ Duyệt).
    const trangThaiKey = !sanSangChot ? "cho_tinh" : (daXacNhan ? "dang_dntt" : "cho_dntt");
    const trangThaiLabel = !sanSangChot ? "Chưa ĐNTT" : (daXacNhan ? "Đang ĐNTT" : "Chờ ĐNTT");
    return {
      idKey: id,
      timestamp: utils.formatDate(r[1]),
      chuRung: String(r[2] || ""),
      nguoiNhan: String(r[3] || ""),
      nganHang: String(r[4] || ""),
      stk: String(r[5] || "").replace(/'/g, ""),
      soTien: soTien,
      noiDungCK: String(r[7] || ""),
      soHD: String(r[8] || "").replace(/'/g, ""),
      klTongKg: utils.parseNum(r[11]),
      soLuongPhieu: details.length,
      danhSachPhieuCan: details.map(d => String(d[11] || "").replace(/'/g, "")),
      ghiChu: String(r[15] || ""),
      sanSangChot: sanSangChot,
      daXacNhan: daXacNhan,
      trangThaiKey: trangThaiKey,
      trangThaiLabel: trangThaiLabel
    };
  }).sort((a, b) => (a.chuRung || "").localeCompare(b.chuRung || ""));
}

/**
 * Chi tiết 1 hồ sơ trong Nháp (dùng cho modal Xem/Sửa).
 */
function getDraftRecordDetail(idKey) {
  const id = String(idKey || "").trim();
  if (!id) return null;
  const { shCT, sh112 } = getDraftSheets_();

  const ctLastRow = shCT.getLastRow();
  const ctAll = ctLastRow > 1 ? shCT.getRange(2, 1, ctLastRow - 1, 22).getValues() : [];
  const chiTiet = ctAll.filter(r => String(r[1] || "").trim() === id).map(r => ({
    idCT: String(r[0] || ""),
    soPhieuCan: String(r[11] || "").replace(/'/g, ""),
    klTan: utils.parseNum(r[12]),
    thanhTien: utils.parseNum(r[16])
  }));

  const the112LastRow = sh112.getLastRow();
  const all112 = the112LastRow > 1 ? sh112.getRange(2, 1, the112LastRow - 1, 24).getValues() : [];
  const row = all112.find(r => String(r[0] || "").trim() === id);
  if (!row) return null;

  return {
    idKey: id,
    chuRung: String(row[2] || ""),
    nguoiNhan: String(row[3] || ""),
    nganHang: String(row[4] || ""),
    stk: String(row[5] || "").replace(/'/g, ""),
    soTien: utils.parseNum(row[6]),
    noiDungCK: String(row[7] || ""),
    soHD: String(row[8] || "").replace(/'/g, ""),
    klTongKg: utils.parseNum(row[11]),
    ghiChu: String(row[15] || ""),
    daXacNhan: String(row[COL_TRANG_THAI_DNTT] || "").trim() === "Đang ĐNTT",
    chiTiet: chiTiet
  };
}

/**
 * Sửa thông tin 1 hồ sơ Nháp (các trường cho phép sửa tay: Chủ rừng,
 * Người nhận, Ngân hàng, STK, Số HĐ, Nội dung CK). Số tiền/lũy kế vẫn
 * do "Tổng Hợp 112" tính - không sửa tay ở đây để tránh sai lệch.
 * SỬA (mục L): Draft CT lưu bản sao RIÊNG của Chủ rừng/Người nhận/STK/
 * Số HĐ cho từng dòng chi tiết (không tham chiếu chéo tới Draft 112).
 * "Đề Nghị Thanh Toán" (runCreate112) dùng CHÍNH Số HĐ trong Draft CT để
 * gộp nhóm tính lũy kế - nếu chỉ sửa ở Draft 112 mà không đồng bộ xuống
 * Draft CT, lũy kế "Còn lại" sẽ bị tính SAI theo Số HĐ cũ. Vì vậy hàm
 * này giờ đồng bộ các trường liên quan xuống MỌI dòng Draft CT của
 * cùng ID_KEY.
 */
function updateDraft112Info(idKey, updates) {
  let lock;
  try {
    lock = sysLock.acquire();
    const id = String(idKey || "").trim();
    if (!id) throw new Error("Thiếu ID hồ sơ.");
    updates = updates || {};

    const { shCT, sh112 } = getDraftSheets_();
    const lastRow = sh112.getLastRow();
    if (lastRow < 2) throw new Error("File Nháp trống.");
    const data = sh112.getRange(2, 1, lastRow - 1, 24).getValues();
    const idx = data.findIndex(r => String(r[0] || "").trim() === id);
    if (idx === -1) throw new Error(`Không tìm thấy hồ sơ "${id}" trong Nháp.`);

    const row = data[idx];
    // MỚI (mục W): CHỈ được sửa khi hồ sơ đang ở "Chờ ĐNTT".
    if (!_isRecordEditable_(row)) {
      throw new Error(`Không thể sửa: hồ sơ "${id}" chưa/đã qua trạng thái "Chờ ĐNTT". Hãy bấm "Về Chờ ĐNTT" trước nếu đang ở "Đang ĐNTT".`);
    }
    if (updates.chuRung !== undefined) row[2] = updates.chuRung;
    if (updates.nguoiNhan !== undefined) row[3] = updates.nguoiNhan;
    if (updates.nganHang !== undefined) row[4] = updates.nganHang;
    if (updates.stk !== undefined) row[5] = "'" + String(updates.stk).replace(/'/g, "");
    if (updates.noiDungCK !== undefined) row[7] = updates.noiDungCK;
    if (updates.soHD !== undefined) row[8] = "'" + String(updates.soHD).replace(/'/g, "");

    sh112.getRange(idx + 2, 1, 1, 24).setValues([row]);

    // MỚI (mục L): đồng bộ xuống Draft CT
    const ctLastRow = shCT.getLastRow();
    if (ctLastRow > 1) {
      const ctAll = shCT.getRange(2, 1, ctLastRow - 1, 22).getValues();
      let changed = false;
      ctAll.forEach(r => {
        if (String(r[1] || "").trim() !== id) return;
        if (updates.chuRung !== undefined) { r[3] = updates.chuRung; changed = true; }
        if (updates.nguoiNhan !== undefined) { r[6] = updates.nguoiNhan; changed = true; }
        if (updates.stk !== undefined) { r[7] = "'" + String(updates.stk).replace(/'/g, ""); changed = true; }
        if (updates.soHD !== undefined) { r[19] = "'" + String(updates.soHD).replace(/'/g, ""); changed = true; }
      });
      if (changed) shCT.getRange(2, 1, ctAll.length, 22).setValues(ctAll);
    }

    SpreadsheetApp.flush();
    logAction("SUA_NHAP", id, "Cập nhật thông tin hồ sơ nháp: " + JSON.stringify(updates));
    return { success: true, message: `✅ Đã cập nhật hồ sơ "${id}" (đã đồng bộ xuống dòng chi tiết).` };
  } catch (e) {
    return { success: false, message: "❌ Lỗi: " + e.toString() };
  } finally { if (lock) lock.releaseLock(); }
}

/**
 * Thêm 1 Số phiếu cân vào 1 hồ sơ đang có trong Nháp (mở rộng Draft CT).
 * Sau khi thêm, cần chạy lại "Tổng Hợp 112" để tính lại Số tiền.
 */
function addPhieuCanToDraft(idKey, soPhieuCan) {
  let lock;
  try {
    lock = sysLock.acquire();
    const id = String(idKey || "").trim();
    const soP = String(soPhieuCan || "").trim();
    if (!id || !soP) throw new Error("Thiếu thông tin.");

    const { shCT, sh112 } = getDraftSheets_();

    // MỚI (mục W): CHỈ được thêm phiếu cân khi hồ sơ đang ở "Chờ ĐNTT".
    const c112LastRowChk = sh112.getLastRow();
    if (c112LastRowChk > 1) {
      const all112Chk = sh112.getRange(2, 1, c112LastRowChk - 1, 24).getValues();
      const row112Chk = all112Chk.find(r => String(r[0] || "").trim() === id);
      if (row112Chk && !_isRecordEditable_(row112Chk)) {
        throw new Error(`Không thể sửa: hồ sơ "${id}" chưa/đã qua trạng thái "Chờ ĐNTT". Hãy bấm "Về Chờ ĐNTT" trước nếu đang ở "Đang ĐNTT".`);
      }
    }

    const lastRow = shCT.getLastRow();
    const ctAll = lastRow > 1 ? shCT.getRange(2, 1, lastRow - 1, 22).getValues() : [];

    const key = utils.standardize(soP);
    if (ctAll.some(r => utils.standardize(r[11]) === key)) {
      throw new Error(`Số phiếu cân "${soP}" đang được dùng bởi 1 hồ sơ trong Nháp rồi.`);
    }

    // SỬA (mục P - tối ưu tốc độ): dùng cache "chưa thanh toán" (nhỏ, nhanh).
    const pcMap = utils.buildIndexMap(_pcUnpaidData_(), 22, true);
    const pcRow = pcMap.get(key);
    if (!pcRow) throw new Error(`Số phiếu cân "${soP}" không tồn tại trong hệ thống Phiếu Cân (hoặc đã được thanh toán).`);
    const trangThai = String(pcRow[26] || "").trim();
    const daKhoa = String(pcRow[27] || "").trim().toUpperCase();
    if (trangThai || daKhoa === "Y" || daKhoa === "N") throw new Error(`Số phiếu cân "${soP}" đã được sử dụng CHÍNH THỨC.`);

    const own = ctAll.filter(r => String(r[1] || "").trim() === id);
    if (own.length === 0) throw new Error(`Không tìm thấy hồ sơ "${id}" trong Nháp.`);
    const first = own[0];
    const srcInfo = {
      timestamp: first[2], chuRung: first[3], cccd: String(first[4] || "").replace(/'/g, ""),
      nguoiDN: first[5], nguoiNhan: first[6], stkNguoiNhan: String(first[7] || "").replace(/'/g, ""),
      klTongNguon: first[8], dsPhieuCanGoc: first[9], soHD: String(first[19] || "").replace(/'/g, ""),
      ngayTT: first[20], ngayDeNghi: first[21]
    };
    const stt = own.length + 1;
    const { row } = buildDraftCtRow_(id, stt, soP, srcInfo, pcMap);
    shCT.getRange(shCT.getLastRow() + 1, 1, 1, 22).setValues([row]);
    SpreadsheetApp.flush();
    logAction("THEM_PHIEU_NHAP", id, `Thêm phiếu cân ${soP}`);
    return { success: true, message: `✅ Đã thêm phiếu cân "${soP}" vào hồ sơ "${id}". Hãy chạy lại "Đề Nghị Thanh Toán" để tính lại số tiền.` };
  } catch (e) {
    return { success: false, message: "❌ Lỗi: " + e.toString() };
  } finally { if (lock) lock.releaseLock(); }
}

/**
 * Xóa 1 dòng phiếu cân (1 dòng chi tiết) khỏi 1 hồ sơ trong Nháp - vẫn
 * giữ lại các phiếu cân khác của hồ sơ đó. idCT = "ID_KEY-STT".
 * SỬA (mục L): nếu đây là phiếu cân CUỐI CÙNG của hồ sơ, tự động xóa
 * luôn dòng Draft 112 tương ứng - tránh để sót hồ sơ "mồ côi" (còn dòng
 * 112 nhưng 0 phiếu cân, không thể chốt cũng không hiển thị đúng).
 */
function removePhieuCanFromDraft(idCT) {
  let lock;
  try {
    lock = sysLock.acquire();
    const idct = String(idCT || "").trim();
    if (!idct) throw new Error("Thiếu ID dòng chi tiết.");
    const { shCT, sh112 } = getDraftSheets_();
    const lastRow = shCT.getLastRow();
    if (lastRow < 2) throw new Error("File Nháp trống.");
    const ctAll = shCT.getRange(2, 1, lastRow - 1, 22).getValues();
    const removedRow = ctAll.find(r => String(r[0] || "").trim() === idct);
    if (!removedRow) throw new Error(`Không tìm thấy dòng "${idct}".`);
    const ownerId = String(removedRow[1] || "").trim();

    // MỚI (mục W): CHỈ được xóa phiếu cân khi hồ sơ CHỦ đang ở "Chờ ĐNTT".
    const c112LastRowChk = sh112.getLastRow();
    if (c112LastRowChk > 1 && ownerId) {
      const all112Chk = sh112.getRange(2, 1, c112LastRowChk - 1, 24).getValues();
      const row112Chk = all112Chk.find(r => String(r[0] || "").trim() === ownerId);
      if (row112Chk && !_isRecordEditable_(row112Chk)) {
        throw new Error(`Không thể sửa: hồ sơ "${ownerId}" chưa/đã qua trạng thái "Chờ ĐNTT". Hãy bấm "Về Chờ ĐNTT" trước nếu đang ở "Đang ĐNTT".`);
      }
    }

    const kept = ctAll.filter(r => String(r[0] || "").trim() !== idct);
    shCT.getRange(2, 1, ctAll.length, 22).clearContent();
    if (kept.length > 0) shCT.getRange(2, 1, kept.length, 22).setValues(kept);

    // Còn phiếu cân nào khác của cùng hồ sơ này không?
    const stillHasOther = kept.some(r => String(r[1] || "").trim() === ownerId);
    let alsoRemoved112 = false;
    if (!stillHasOther && ownerId) {
      const c112LastRow = sh112.getLastRow();
      if (c112LastRow > 1) {
        const all112 = sh112.getRange(2, 1, c112LastRow - 1, 24).getValues();
        const kept112 = all112.filter(r => String(r[0] || "").trim() !== ownerId);
        if (kept112.length !== all112.length) {
          sh112.getRange(2, 1, all112.length, 24).clearContent();
          if (kept112.length > 0) sh112.getRange(2, 1, kept112.length, 24).setValues(kept112);
          alsoRemoved112 = true;
        }
      }
    }

    SpreadsheetApp.flush();
    logAction("XOA_PHIEU_NHAP", "-", `Xóa dòng chi tiết ${idct} khỏi Nháp` + (alsoRemoved112 ? ` (hồ sơ ${ownerId} hết phiếu cân, đã xóa luôn dòng tổng hợp)` : ""));
    const msg = alsoRemoved112
      ? `✅ Đã xóa phiếu cân cuối cùng - hồ sơ "${ownerId}" đã hết dữ liệu và được xóa luôn khỏi Nháp.`
      : `✅ Đã xóa dòng "${idct}" khỏi Nháp. Hãy chạy lại "Đề Nghị Thanh Toán" để tính lại số tiền.`;
    return { success: true, message: msg, recordRemoved: alsoRemoved112 };
  } catch (e) {
    return { success: false, message: "❌ Lỗi: " + e.toString() };
  } finally { if (lock) lock.releaseLock(); }
}

/**
 * Wrapper cho Web App: chốt thanh toán cho danh sách ID được chọn từ
 * màn "Danh Sách Nháp", trả về object thay vì chuỗi để JS dễ xử lý.
 */
/**
 * MỚI (mục U): "Xác Nhận" - chuyển các hồ sơ đã CÓ Số tiền (Chờ ĐNTT)
 * sang trạng thái "Đang ĐNTT" (đã Xác Nhận, chờ Duyệt). Chỉ sau bước
 * này mới được phép in Báo Cáo ĐNTT (mẫu excel) và bấm "Duyệt" (Đóng
 * Thanh Toán). Hồ sơ chưa tính tiền (Số tiền <= 0) bị bỏ qua kèm cảnh báo.
 */
function runXacNhanDNTT(selectedIds) {
  let lock;
  try {
    lock = sysLock.acquire();
    if (!Array.isArray(selectedIds) || selectedIds.length === 0) {
      return { success: false, message: "❌ Lỗi: Không có hồ sơ nào được chọn để Xác Nhận." };
    }
    const idSet = new Set(selectedIds.map(id => String(id).trim()).filter(Boolean));
    if (idSet.size === 0) return { success: false, message: "❌ Lỗi: Danh sách hồ sơ được chọn không hợp lệ." };

    const { sh112 } = getDraftSheets_();
    const lastRow = sh112.getLastRow();
    if (lastRow < 2) return { success: false, message: "⚠️ File Nháp trống." };
    const data = sh112.getRange(2, 1, lastRow - 1, 24).getValues();

    let count = 0;
    const skippedNotCalculated = [], skippedNoData = [];
    const foundIds = new Set();
    data.forEach(row => {
      const id = String(row[0] || "").trim();
      if (!id || !idSet.has(id)) return;
      foundIds.add(id);
      if (utils.parseNum(row[6]) <= 0) { skippedNotCalculated.push(id); return; }
      row[COL_TRANG_THAI_DNTT] = "Đang ĐNTT";
      count++;
    });
    idSet.forEach(id => { if (!foundIds.has(id)) skippedNoData.push(id); });

    if (count > 0) {
      sh112.getRange(2, 1, data.length, 24).setValues(data);
      SpreadsheetApp.flush();
    }

    let msg;
    if (count === 0) {
      msg = "❌ Không có hồ sơ nào đủ điều kiện Xác Nhận (phải đã có Số tiền > 0, tức đã chạy 'Tổng Hợp 112').";
    } else {
      msg = `✅ Đã Xác Nhận ${count} hồ sơ - chuyển sang trạng thái "Đang ĐNTT", có thể in Báo Cáo ĐNTT và chờ Duyệt để Đóng Thanh Toán.`;
    }
    if (skippedNotCalculated.length) msg += ` ⚠️ Chưa tính tiền (bỏ qua): ${skippedNotCalculated.join(", ")}.`;
    if (skippedNoData.length) msg += ` ⚠️ Không tìm thấy trong Nháp: ${skippedNoData.join(", ")}.`;
    logAction("XAC_NHAN_DNTT", Array.from(idSet).join(","), `Xác Nhận ${count} hồ sơ.`);
    return { success: count > 0, message: msg };
  } catch (e) {
    return { success: false, message: "❌ Lỗi: " + e.toString() };
  } finally { if (lock) lock.releaseLock(); }
}

/**
 * MỚI (mục W): "Về Chờ ĐNTT" - chuyển 1 hồ sơ đang "Đang ĐNTT" (đã Xác
 * Nhận) NGƯỢC LẠI về "Chờ ĐNTT" - vì Sửa/Xóa CHỈ được phép thực hiện ở
 * trạng thái "Chờ ĐNTT" (mục W: khóa Sửa/Xóa ở "Chưa ĐNTT" và "Đang
 * ĐNTT" để tránh sửa nhầm hồ sơ đã tính tiền/đã xác nhận).
 */
function runHuyXacNhanDNTT(idKey) {
  let lock;
  try {
    lock = sysLock.acquire();
    const id = String(idKey || "").trim();
    if (!id) return { success: false, message: "❌ Lỗi: Thiếu ID hồ sơ." };

    const { sh112 } = getDraftSheets_();
    const lastRow = sh112.getLastRow();
    if (lastRow < 2) return { success: false, message: "⚠️ File Nháp trống." };
    const data = sh112.getRange(2, 1, lastRow - 1, 24).getValues();
    const idx = data.findIndex(r => String(r[0] || "").trim() === id);
    if (idx === -1) return { success: false, message: `❌ Không tìm thấy hồ sơ "${id}" trong Nháp.` };

    data[idx][COL_TRANG_THAI_DNTT] = "";
    sh112.getRange(idx + 2, 1, 1, 24).setValues([data[idx]]);
    SpreadsheetApp.flush();

    logAction("HUY_XAC_NHAN_DNTT", id, `Chuyển hồ sơ "${id}" về "Chờ ĐNTT" để cho phép Sửa/Xóa.`);
    return { success: true, message: `✅ Đã chuyển hồ sơ "${id}" về trạng thái "Chờ ĐNTT" - giờ có thể Sửa/Xóa.` };
  } catch (e) {
    return { success: false, message: "❌ Lỗi: " + e.toString() };
  } finally { if (lock) lock.releaseLock(); }
}

function webConfirmPayment(selectedIds, payDateStr) {
  const msg = runConfirmPayment(selectedIds, payDateStr);
  return { success: msg.indexOf("✅") === 0, message: msg };
}

/**
 * Wrapper cho Web App: chạy "Đề Nghị Thanh Toán" (Tổng Hợp 112).
 */
function webRunCreate112() {
  const msg = runCreate112();
  return { success: msg.indexOf("✅") === 0, message: msg };
}

/**
 * Báo cáo: danh sách hồ sơ CHÍNH THỨC (đã chốt) trong khoảng ngày, dùng
 * cho màn "Báo Cáo" của Web App. fDate/tDate dạng "yyyy-MM-dd" hoặc rỗng.
 */
function getReportList(fDate, tDate) {
  fDate = _clampReportRange_(fDate, tDate); // mục R - xem ghi chú tại _clampReportRange_
  return get112ViewData(fDate, tDate);
}

/**
 * Xuất báo cáo (tạo file Google Sheet mới) cho các dòng đã chọn trên
 * màn Báo Cáo của Web App.
 */
function webExportReport(rows, dateRange) {
  return createFinalReportFromFilteredData(rows, dateRange || "");
}

/**
 * MỚI (mục U): "in Báo cáo ĐNTT theo mẫu excel" - xuất Bảng Đề Xuất
 * Thanh Toán (cùng mẫu Sheet1 với Báo Cáo chính thức) NHƯNG lấy dữ liệu
 * từ File Nháp, CHỈ cho các hồ sơ đã "Đang ĐNTT" (đã Xác Nhận). Dùng để
 * in/trình ký TRƯỚC KHI Duyệt/Đóng Thanh Toán - vì vậy KHÔNG kèm sheet
 * "Bảng Kê Chi Tiết CK" (sheet đó cần dữ liệu đã chốt ở sheet chính thức
 * thật, chưa có ở giai đoạn này).
 */
function exportBaoCaoDNTTFromDraft(selectedIds) {
  try {
    if (!Array.isArray(selectedIds) || selectedIds.length === 0) {
      return { success: false, message: "❌ Lỗi: Không có hồ sơ nào được chọn để xuất Báo Cáo ĐNTT." };
    }
    const idSet = new Set(selectedIds.map(id => String(id).trim()).filter(Boolean));
    const { sh112 } = getDraftSheets_();
    const lastRow = sh112.getLastRow();
    if (lastRow < 2) return { success: false, message: "⚠️ File Nháp trống." };
    const data = sh112.getRange(2, 1, lastRow - 1, 24).getValues();

    const rows = [];
    const skippedNotConfirmed = [];
    data.forEach(r => {
      const id = String(r[0] || "").trim();
      if (!id || !idSet.has(id)) return;
      if (String(r[COL_TRANG_THAI_DNTT] || "").trim() !== "Đang ĐNTT") { skippedNotConfirmed.push(id); return; }
      rows.push({
        idHeThong: id,
        ngayDN: utils.formatDate(r[16]),
        soLan: String(r[18] || ""),
        chuRung: String(r[2] || ""),
        nguoiNhan: String(r[3] || ""),
        stk: String(r[5] || "").replace(/'/g, ""),
        nganHang: String(r[4] || ""),
        klTan: utils.parseNum(r[11]) / 1000,
        soTien: utils.parseNum(r[6]),
        noiDungCK: String(r[7] || ""),
        ghiChu: String(r[15] || "")
      });
    });

    if (rows.length === 0) {
      let msg = "❌ Không có hồ sơ nào ở trạng thái 'Đang ĐNTT' để xuất báo cáo (phải bấm 'Xác Nhận' trước).";
      if (skippedNotConfirmed.length) msg += ` Hồ sơ chưa Xác Nhận: ${skippedNotConfirmed.join(", ")}.`;
      return { success: false, message: msg };
    }

    const dateRange = "CHỜ DUYỆT - " + Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy");
    const folder = DriveApp.getFolderById(CFG.FOLDER_REPORT_ID);
    const fileName = "BÁO CÁO ĐNTT (chờ duyệt) - " + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd_HHmm");
    const newSS = SpreadsheetApp.create(fileName);
    const file = DriveApp.getFileById(newSS.getId());
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);

    const sheet1 = newSS.getSheets()[0];
    sheet1.setName("DeNghiThanhToanCK");
    renderSheet1Full(sheet1, rows, dateRange);

    let msg = `✅ Đã xuất Báo Cáo ĐNTT (${rows.length} hồ sơ - CHƯA chốt thanh toán chính thức, chỉ để in/trình duyệt).`;
    if (skippedNotConfirmed.length) msg += ` ⚠️ Bỏ qua (chưa Xác Nhận): ${skippedNotConfirmed.join(", ")}.`;
    logAction("XUAT_BAO_CAO_DNTT_NHAP", Array.from(idSet).join(","), `Xuất Báo Cáo ĐNTT nháp ${rows.length} hồ sơ - ${newSS.getUrl()}`);
    return { success: true, url: newSS.getUrl(), message: msg };
  } catch (e) {
    return { success: false, message: "❌ Lỗi: " + e.toString() };
  }
}

// ============================================================
// MỚI (mục J): KHỞI TẠO FILE NHÁP - CHỈ CHẠY 1 LẦN DUY NHẤT
// ============================================================
function setupDraftSpreadsheet() {
  const existingId = getDraftSsId_();
  if (existingId) {
    const msg = `File Nháp đã được thiết lập từ trước (ID: ${existingId}). Nếu bạn CHẮC CHẮN muốn tạo file Nháp MỚI (bỏ toàn bộ hồ sơ nháp hiện có), hãy chạy hàm resetDraftSsId_() 1 lần trong Apps Script Editor rồi chạy lại hàm này.`;
    Logger.log(msg);
    return { alreadySetup: true, message: msg, id: existingId };
  }

  const ss = SpreadsheetApp.create("HAK - FILE NHAP (DRAFT) - " + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd_HHmm"));

  const shCT = ss.getSheets()[0];
  shCT.setName(CFG.DRAFT_CT_SHEET);
  shCT.getRange(1, 1, 1, 22).setValues([[
    "ID_CT", "ID_KEY", "Timestamp", "Chủ rừng", "CCCD", "Người ĐN", "Người nhận", "STK người nhận",
    "KL Tổng nguồn(kg)", "DS Phiếu cân gốc", "STT", "Số phiếu cân", "KL Tấn", "Cột N(PC)", "Cột O(PC)",
    "Cột P(N+O)", "Thành tiền", "Trạng thái giá", "Đã chốt(Y/N)", "Số HĐ", "Ngày CK/TT", "Ngày Đề Nghị"
  ]]).setFontWeight("bold").setBackground("#fce5cd");
  shCT.setFrozenRows(1);

  const sh112 = ss.insertSheet(CFG.DRAFT_112_SHEET);
  sh112.getRange(1, 1, 1, 24).setValues([[
    "ID_KEY", "Timestamp", "Chủ rừng", "Người nhận", "Ngân hàng", "STK", "Số tiền", "Nội dung CK",
    "Số HĐ", "Ngày HĐ", "Ủy quyền", "KL Tổng(kg)", "SL HĐ lũy kế", "Đã trả", "Còn lại", "Ghi chú",
    "Ngày ĐN", "SL HĐ(cột R)", "Lần TT", "Ngày dự kiến TT", "Đã chốt(Y/N)", "Đủ ĐK TT", "ID_112", "Trạng Thái ĐNTT"
  ]]).setFontWeight("bold").setBackground("#fce5cd");
  sh112.setFrozenRows(1);

  const id = ss.getId(), url = ss.getUrl();
  // SỬA (mục S): tự lưu ID vào Script Properties - KHÔNG cần dán tay vào
  // code nữa, và sẽ KHÔNG mất khi cập nhật/dán đè code mới sau này.
  setDraftSsId_(id);
  const msg = `✅ Đã tạo File Nháp thành công và LƯU LẠI tự động!\n\nID: ${id}\nURL: ${url}\n\nTừ nay không cần thiết lập lại nữa, kể cả khi cập nhật code mới.`;
  Logger.log(msg);
  return { alreadySetup: false, message: msg, id, url };
}

/** Chỉ dùng khi CHẮC CHẮN muốn bỏ File Nháp hiện tại và tạo file MỚI
 * HOÀN TOÀN KHÁC (mất hết hồ sơ đang chờ xử lý trong Nháp cũ). Chạy 1
 * lần trong Apps Script Editor rồi mới chạy lại setupDraftSpreadsheet(). */
function resetDraftSsId_() {
  PropertiesService.getScriptProperties().deleteProperty('DRAFT_SS_ID');
  return "Đã gỡ liên kết File Nháp cũ. Chạy setupDraftSpreadsheet() (hoặc menu Khởi Tạo File Nháp) để tạo file mới.";
}

function showSetupDraftDialog() {
  const ui = SpreadsheetApp.getUi();
  const result = setupDraftSpreadsheet();
  if (result.alreadySetup) {
    ui.alert(result.message);
    return;
  }
  const html = HtmlService.createHtmlOutput(
    `<div style="font-family:Arial;padding:10px;line-height:1.6">
      <p>✅ Đã tạo <b>File Nháp</b> thành công và <b>lưu lại tự động</b>!</p>
      <p><a href="${result.url}" target="_blank">Mở File Nháp</a></p>
      <p>ID: <code>${result.id}</code></p>
      <p style="color:#2C4636">Từ nay có thể dùng ngay các nút Tách Phiếu / Tổng Hợp 112 / Chốt Thanh Toán - không cần thiết lập lại, kể cả khi cập nhật code mới sau này.</p>
    </div>`
  ).setWidth(440).setHeight(240);
  ui.showModalDialog(html, "Khởi Tạo File Nháp");
}

function showOpenDraftDialog() {
  const ui = SpreadsheetApp.getUi();
  try {
    const { ss } = getDraftSheets_();
    const html = HtmlService.createHtmlOutput(
      `<div style="font-family:Arial;padding:10px"><a href="${ss.getUrl()}" target="_blank">Mở File Nháp trong tab mới</a></div>`
    ).setWidth(360).setHeight(80);
    ui.showModalDialog(html, "File Nháp");
  } catch (e) {
    ui.alert(e.toString());
  }
}

function showDeleteDraftDialog() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Xóa Hồ Sơ Nháp', 'Nhập ID hồ sơ (ID_KEY) cần xóa khỏi File Nháp:', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() != ui.Button.OK) return;
  const msg = runDeleteDraftRecord(result.getResponseText().trim());
  ui.alert(msg);
}

// MỚI (mục P): làm mới thủ công cache Phiếu Cân chưa thanh toán.
function showRefreshPcCacheDialog() {
  const ui = SpreadsheetApp.getUi();
  try {
    const n = refreshPhieuCanUnpaidCache_();
    ui.alert(`✅ Đã làm mới cache: hiện có ${n} phiếu cân CHƯA thanh toán.`);
  } catch (e) {
    ui.alert("❌ Lỗi: " + e.toString());
  }
}

// MỚI (mục P): bật Trigger tự động làm mới cache mỗi 10 phút (chỉ cần chạy 1 lần).
function showSetupPcCacheTriggerDialog() {
  const ui = SpreadsheetApp.getUi();
  try {
    ui.alert(setupPcCacheAutoRefreshTrigger());
  } catch (e) {
    ui.alert("❌ Lỗi: " + e.toString());
  }
}

// ============================================================
// ĐIỀN LẠI HÀNG LOẠT NGÂN HÀNG CÒN TRỐNG TRONG DỮ LIỆU CHÍNH THỨC
// (không đổi so với 2026.3 - vẫn dùng để vá dữ liệu LỊCH SỬ đã chốt)
// ============================================================
function runFillMissingBankOnly() {
  let lock;
  try {
    lock = sysLock.acquire();
    const sh112 = SpreadsheetApp.getActive().getSheetByName(CFG.DNTT_112);
    if (!sh112 || sh112.getLastRow() < 2) return "⚠️ Sheet 112 trống.";

    const data112 = sh112.getRange(1, 1, sh112.getLastRow(), sh112.getLastColumn()).getValues();
    const bankMap = buildBankLookupFromHDSTK();

    let filledBankCount = 0;
    for (let i = 1; i < data112.length; i++) {
      const id = String(data112[i][0]).trim();
      if (!id) continue;
      if (!utils.isBlank(data112[i][4])) continue;

      const soHDRaw = String(data112[i][8] || "");
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
// SỬA (mục J): thêm menu cho File Nháp; menu "Khởi Tạo File Nháp" chỉ
// hiện khi CFG.DRAFT_SS_ID còn trống.
// ============================================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('🚀 QUẢN LÝ HAK')
      .addItem('1. Tách Phiếu (vào File Nháp)', 'runProcessDetail')
      .addItem('2. Tổng Hợp 112 (trong File Nháp)', 'runCreate112')
      .addSeparator()
      .addItem('3. Chốt Thanh Toán (đẩy Nháp -> Chính thức)', 'showPayDialog')
      .addSeparator()
      .addItem('4. Thêm Mới Đề Nghị Thanh Toán', 'showAddPaymentDialog')
      .addSeparator()
      .addItem('📂 Mở File Nháp', 'showOpenDraftDialog')
      .addItem('🗑️ Xóa 1 Hồ Sơ Khỏi Nháp', 'showDeleteDraftDialog')
      .addSeparator()
      .addItem('5. Vá Ngân Hàng Còn Trống (dữ liệu Chính thức)', 'runFillMissingBankOnly')
      .addSeparator()
      .addItem('🔄 Làm Mới Cache Phiếu Cân (chưa TT)', 'showRefreshPcCacheDialog')
      .addItem('⏱️ Bật Tự Động Làm Mới Cache (mỗi 10 phút)', 'showSetupPcCacheTriggerDialog');

  if (!getDraftSsId_()) {
    menu.addSeparator().addItem('⚙️ Khởi Tạo File Nháp (chạy 1 lần)', 'showSetupDraftDialog');
  }
  menu.addToUi();
}

// SỬA (mục J): chọn danh sách hồ sơ để chốt dựa trên FILE NHÁP (Draft
// 112 đã có Số tiền > 0), không còn dựa vào cột trạng thái của nguồn.
function showPayDialog() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt('Xác nhận chốt', 'Nhập ngày thanh toán (dd/mm/yyyy):', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() != ui.Button.OK) return;

  var payDateStr = result.getResponseText().trim();

  var selectedIds;
  try {
    var draft = getDraftSheets_();
    var lastRow = draft.sh112.getLastRow();
    if (lastRow < 2) { ui.alert("File Nháp chưa có hồ sơ nào sẵn sàng để chốt."); return; }
    var data = draft.sh112.getRange(2, 1, lastRow - 1, 7).getValues();
    selectedIds = data
      .filter(r => !utils.isBlank(r[0]) && utils.parseNum(r[6]) > 0)
      .map(r => String(r[0]).trim());
  } catch (e) {
    ui.alert("Lỗi khi đọc File Nháp: " + e.toString());
    return;
  }

  if (selectedIds.length === 0) {
    ui.alert("Không có hồ sơ nào trong File Nháp đã sẵn sàng (đã chạy Tổng Hợp 112) để chốt thanh toán.");
    return;
  }

  var confirm = ui.alert('Xác nhận', `Sẽ chốt thanh toán cho ${selectedIds.length} hồ sơ đang có trong File Nháp. Tiếp tục?`, ui.ButtonSet.YES_NO);
  if (confirm != ui.Button.YES) return;

  var msg = runConfirmPayment(selectedIds, payDateStr);
  ui.alert(msg);
}

// Mở hộp thoại "Thêm Mới Đề Nghị Thanh Toán" ngay trong Google Sheet.
function showAddPaymentDialog() {
  var html = HtmlService.createHtmlOutputFromFile('AddPaymentDialog')
      .setWidth(560).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Thêm Mới Đề Nghị Thanh Toán');
}
