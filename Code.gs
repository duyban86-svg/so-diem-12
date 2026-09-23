/**
 * =========================================================================
 * GOOGLE APPS SCRIPT - HỆ THỐNG NHẬP ĐIỂM & ĐÁNH GIÁ (NĂM HỌC 2026-2027)
 * TÍNH NĂNG NÂNG CẤP:
 * 1. XÁC THỰC TÀI KHOẢN GOOGLE & TỰ ĐỘNG LIÊN KẾT HỌC SINH (TAB 'TaiKhoan').
 * 2. ĐIỂM CỘNG TUẦN: Mỗi học sinh trong 1 tuần chỉ ghi 1 dòng (tự động GHI ĐÈ KẾT QUẢ CŨ).
 * 3. ĐIỂM KIỂM TRA: Gọn gàng gồm Tên học sinh, Điểm, Ghi chú.
 * 4. ĐIỂM NHÓM: Hỗ trợ tùy chỉnh điểm riêng biệt cho từng thành viên trong nhóm.
 * 5. LƯU EMAIL XÁC THỰC: Cột I lưu Email Google gửi điểm để giáo viên dễ dàng hậu kiểm.
 * =========================================================================
 */

var SHEET_NAME_SCORES = "DiemSo";
var SHEET_NAME_ACCOUNTS = "TaiKhoan";

/**
 * HÀM TỰ ĐỘNG CĂN CHỈNH & ĐỊNH DẠNG SHEET THEO ĐÚNG MẪU BẢNG ĐIỂM
 */
function formatSheetLikeImage() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_SCORES) || ss.getActiveSheet();
  
  var headers = [
    "Thời gian",
    "Lớp",
    "STT",
    "Họ và tên",
    "Loại điểm",
    "Chi tiết / Hạng mục",
    "Điểm số",
    "Nội dung / Ghi chú",
    "Email Google"
  ];

  var numCols = headers.length;
  sheet.getRange(1, 1, 1, numCols).setValues([headers]);

  // Tiêu đề: Nền xanh Navy Steel, chữ trắng đậm
  var headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange.setFontWeight("bold").setFontSize(10.5).setFontFamily("Arial");
  headerRange.setBackground("#3b5998").setFontColor("#ffffff");
  headerRange.setHorizontalAlignment("center").setVerticalAlignment("middle");
  headerRange.setBorder(true, true, true, true, true, true, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(1, 36);
  sheet.setFrozenRows(1);

  // Định dạng dữ liệu các dòng
  var lastRow = Math.max(sheet.getLastRow(), 2);
  if (lastRow >= 2) {
    var dataRange = sheet.getRange(2, 1, lastRow - 1, numCols);
    dataRange.setBackground("#ffffff").setFontColor("#000000").setFontFamily("Arial").setFontSize(10).setFontWeight("normal").setVerticalAlignment("middle");
    dataRange.setBorder(true, true, true, true, true, true, "#e2e8f0", SpreadsheetApp.BorderStyle.SOLID);

    for (var r = 2; r <= lastRow; r++) {
      sheet.setRowHeight(r, 28);
    }

    sheet.getRange(2, 1, lastRow - 1, 1).setHorizontalAlignment("center"); // Cột A: Thời gian
    sheet.getRange(2, 2, lastRow - 1, 1).setHorizontalAlignment("center"); // Cột B: Lớp
    sheet.getRange(2, 3, lastRow - 1, 1).setHorizontalAlignment("center"); // Cột C: STT
    sheet.getRange(2, 4, lastRow - 1, 1).setHorizontalAlignment("left");   // Cột D: Họ tên
    sheet.getRange(2, 5, lastRow - 1, 1).setHorizontalAlignment("center"); // Cột E: Loại điểm
    sheet.getRange(2, 6, lastRow - 1, 1).setHorizontalAlignment("left");   // Cột F: Chi tiết
    sheet.getRange(2, 7, lastRow - 1, 1).setHorizontalAlignment("center").setFontWeight("bold"); // Cột G: Điểm số
    sheet.getRange(2, 8, lastRow - 1, 1).setHorizontalAlignment("left");   // Cột H: Ghi chú
    sheet.getRange(2, 9, lastRow - 1, 1).setHorizontalAlignment("left").setFontColor("#64748b"); // Cột I: Email Google
  }

  for (var i = 1; i <= numCols; i++) {
    sheet.autoResizeColumn(i);
    var currentWidth = sheet.getColumnWidth(i);
    sheet.setColumnWidth(i, Math.max(currentWidth + 20, 80));
  }
}

/**
 * HÀM KHỞI TẠO HOẶC LẤY TAB TÀI KHOẢN
 */
function getOrCreateAccountsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_ACCOUNTS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME_ACCOUNTS);
    var headers = ["Thời gian liên kết", "Lớp", "STT", "Họ và tên", "Email Google", "Ghi chú"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    var hRange = sheet.getRange(1, 1, 1, headers.length);
    hRange.setFontWeight("bold").setFontSize(10.5).setFontFamily("Arial");
    hRange.setBackground("#1e293b").setFontColor("#38bdf8");
    hRange.setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.setRowHeight(1, 36);
    sheet.setFrozenRows(1);
    
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 90);
    sheet.setColumnWidth(3, 70);
    sheet.setColumnWidth(4, 200);
    sheet.setColumnWidth(5, 250);
    sheet.setColumnWidth(6, 200);
  }
  return sheet;
}

/**
 * XỬ LÝ YÊU CẦU GET (TRA CỨU TÀI KHOẢN LIÊN KẾT)
 */
function doGet(e) {
  try {
    var action = e && e.parameter ? e.parameter.action : "";
    var accSheet = getOrCreateAccountsSheet();
    var lastRow = accSheet.getLastRow();
    var accounts = [];

    if (lastRow >= 2) {
      var data = accSheet.getRange(2, 1, lastRow - 1, 6).getValues();
      for (var i = 0; i < data.length; i++) {
        var row = data[i];
        if (row[4]) { // Có Email
          accounts.push({
            timestamp: row[0],
            lop: String(row[1]).trim(),
            stt: String(row[2]).trim(),
            name: String(row[3]).trim(),
            email: String(row[4]).trim().toLowerCase(),
            note: String(row[5] || "").trim()
          });
        }
      }
    }

    if (action === "check_email") {
      var targetEmail = (e.parameter.email || "").trim().toLowerCase();
      var found = accounts.find(function(a) { return a.email === targetEmail; });
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        bound: !!found,
        student: found || null
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Mặc định trả về toàn bộ danh sách liên kết
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      accounts: accounts
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * XỬ LÝ DỮ LIỆU GỬI TỪ WEB APP (POST)
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var payload = JSON.parse(e.postData.contents);
    var timestamp = new Date().toLocaleString("vi-VN");

    // TRƯỜNG HỢP 0: LIÊN KẾT TÀI KHOẢN GOOGLE LẦN ĐẦU (SELF-BINDING)
    if (payload.action === "bind_account") {
      var accSheet = getOrCreateAccountsSheet();
      var email = String(payload.email || "").trim().toLowerCase();
      var lop = String(payload.lop || "").trim();
      var stt = String(payload.stt || "").trim();
      var name = String(payload.name || "").trim();
      var note = String(payload.note || "Tự liên kết qua Google Sign-In").trim();

      if (!email || !lop || !stt || !name) {
        return ContentService.createTextOutput(JSON.stringify({
          status: "error",
          message: "Thiếu thông tin bắt buộc (Email, Lớp, STT, Tên)!"
        })).setMimeType(ContentService.MimeType.JSON);
      }

      var lastRowAcc = accSheet.getLastRow();
      var isUpdated = false;

      if (lastRowAcc >= 2) {
        var accData = accSheet.getRange(2, 1, lastRowAcc - 1, 6).getValues();
        for (var idx = 0; idx < accData.length; idx++) {
          var rowLop = String(accData[idx][1]).trim();
          var rowStt = String(accData[idx][2]).trim();
          var rowEmail = String(accData[idx][4]).trim().toLowerCase();

          // 1. Nếu học sinh này (Lớp + STT) đã bị email khác liên kết trước đó
          if (rowLop === lop && rowStt === stt && rowEmail !== email && rowEmail !== "") {
            return ContentService.createTextOutput(JSON.stringify({
              status: "error",
              message: "Học sinh [" + name + " - " + lop + "] đã được liên kết với email: " + rowEmail + ". Nếu có sự nhầm lẫn, vui lòng báo Giáo viên."
            })).setMimeType(ContentService.MimeType.JSON);
          }

          // 2. Nếu email này đã từng liên kết trước đó -> Cập nhật sang học sinh mới
          if (rowEmail === email) {
            var updateRow = idx + 2;
            accSheet.getRange(updateRow, 1).setValue(timestamp);
            accSheet.getRange(updateRow, 2).setValue(lop);
            accSheet.getRange(updateRow, 3).setValue(stt);
            accSheet.getRange(updateRow, 4).setValue(name);
            accSheet.getRange(updateRow, 6).setValue("Cập nhật lại: " + note);
            isUpdated = true;
            break;
          }
        }
      }

      // Nếu email hoàn toàn mới -> Thêm dòng mới vào sheet TaiKhoan
      if (!isUpdated) {
        accSheet.appendRow([timestamp, lop, stt, name, email, note]);
        var newAccRow = accSheet.getLastRow();
        accSheet.getRange(newAccRow, 1, 1, 6).setBackground("#ffffff").setFontFamily("Arial").setFontSize(10).setVerticalAlignment("middle");
        accSheet.setRowHeight(newAccRow, 28);
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "Liên kết tài khoản thành công cho học sinh: " + name + " (" + lop + ")",
        student: { lop: lop, stt: stt, name: name, email: email }
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // LẤY HOẶC TẠO SHEET ĐIỂM SỐ
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME_SCORES);
    if (!sheet) {
      sheet = ss.getActiveSheet();
      if (sheet.getName() === SHEET_NAME_ACCOUNTS) {
        sheet = ss.insertSheet(SHEET_NAME_SCORES);
      }
    }

    var submitterEmail = String(payload.email || "").trim().toLowerCase();

    // TRƯỜNG HỢP 1: ĐIỂM NHÓM (CÓ THỂ CÓ ĐIỂM RIÊNG TỪNG THÀNH VIÊN)
    if (payload.students && Array.isArray(payload.students) && payload.students.length > 0) {
      var rows = [];
      for (var i = 0; i < payload.students.length; i++) {
        var hs = payload.students[i];
        rows.push([
          timestamp,
          payload.lop || "",
          hs.stt || "",
          hs.name || "",
          payload.scoreType || "Điểm nhóm",
          payload.category || "",
          hs.score !== undefined ? hs.score : (payload.score || ""),
          hs.note || payload.note || "",
          submitterEmail
        ]);
      }
      
      if (rows.length > 0) {
        var startRow = sheet.getLastRow() + 1;
        sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
        formatDataRows(sheet, startRow, rows.length);
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "Đã ghi nhận điểm cho " + rows.length + " thành viên trong nhóm!",
        count: rows.length
      })).setMimeType(ContentService.MimeType.JSON);

    } 
    // TRƯỜNG HỢP 2: ĐIỂM CỘNG THEO TUẦN (KIỂM TRA ĐỂ GHI ĐÈ KẾT QUẢ CŨ TRONG CÙNG TUẦN)
    else if (payload.scoreType === "Điểm cộng theo tuần") {
      var lastRow = sheet.getLastRow();
      var isOverwritten = false;
      var overwrittenRow = -1;

      if (lastRow >= 2) {
        // Lấy toàn bộ dữ liệu cột B (Lớp), C (STT), E (Loại điểm), F (Tuần)
        var data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
        
        for (var idx = 0; idx < data.length; idx++) {
          var rowClass = String(data[idx][1]).trim();
          var rowStt = String(data[idx][2]).trim();
          var rowType = String(data[idx][4]).trim();
          var rowWeek = String(data[idx][5]).trim();

          // Nếu cùng Lớp, cùng STT, cùng là Điểm cộng tuần, và cùng Tuần -> GHI ĐÈ
          if (rowClass === String(payload.lop).trim() && 
              rowStt === String(payload.stt).trim() && 
              rowType === "Điểm cộng theo tuần" && 
              rowWeek === String(payload.category).trim()) {
            
            overwrittenRow = idx + 2; // Số hàng thực tế trên Sheet
            sheet.getRange(overwrittenRow, 1).setValue(timestamp); // Cập nhật thời gian mới
            sheet.getRange(overwrittenRow, 7).setValue(payload.score); // Cập nhật điểm mới
            sheet.getRange(overwrittenRow, 8).setValue(payload.note || ""); // Cập nhật ghi chú mới
            sheet.getRange(overwrittenRow, 9).setValue(submitterEmail); // Cập nhật Email người gửi
            
            formatDataRows(sheet, overwrittenRow, 1);
            isOverwritten = true;
            break;
          }
        }
      }

      // Nếu chưa có kết quả trong tuần này -> Thêm dòng mới
      if (!isOverwritten) {
        sheet.appendRow([
          timestamp,
          payload.lop || "",
          payload.stt || "",
          payload.name || "",
          payload.scoreType || "",
          payload.category || "",
          payload.score !== undefined ? payload.score : "",
          payload.note || "",
          submitterEmail
        ]);
        var newRowNum = sheet.getLastRow();
        formatDataRows(sheet, newRowNum, 1);
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        isOverwritten: isOverwritten,
        message: isOverwritten ? "Đã cập nhật (ghi đè) lại điểm cộng tuần thành công!" : "Đã ghi nhận điểm cộng tuần thành công!"
      })).setMimeType(ContentService.MimeType.JSON);

    } 
    // TRƯỜNG HỢP 3: ĐIỂM KIỂM TRA & CÁC LOẠI ĐIỂM ĐƠN KHÁC
    else {
      sheet.appendRow([
        timestamp,
        payload.lop || "",
        payload.stt || "",
        payload.name || "",
        payload.scoreType || "Điểm kiểm tra",
        payload.category || "Kiểm tra",
        payload.score !== undefined ? payload.score : "",
        payload.note || "",
        submitterEmail
      ]);

      var newRowNum2 = sheet.getLastRow();
      formatDataRows(sheet, newRowNum2, 1);

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "Ghi nhận điểm kiểm tra thành công cho: " + (payload.name || "Học sinh")
      })).setMimeType(ContentService.MimeType.JSON);
    }

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}

/**
 * ĐỊNH DẠNG CÁC DÒNG MỚI ĐỒNG BỘ (9 CỘT)
 */
function formatDataRows(sheet, startRow, numRows) {
  var range = sheet.getRange(startRow, 1, numRows, 9);
  range.setBackground("#ffffff");
  range.setFontColor("#000000");
  range.setFontFamily("Arial");
  range.setFontSize(10);
  range.setVerticalAlignment("middle");
  range.setBorder(true, true, true, true, true, true, "#e2e8f0", SpreadsheetApp.BorderStyle.SOLID);

  for (var r = startRow; r < startRow + numRows; r++) {
    sheet.setRowHeight(r, 28);
  }

  sheet.getRange(startRow, 1, numRows, 1).setHorizontalAlignment("center");
  sheet.getRange(startRow, 2, numRows, 1).setHorizontalAlignment("center");
  sheet.getRange(startRow, 3, numRows, 1).setHorizontalAlignment("center");
  sheet.getRange(startRow, 4, numRows, 1).setHorizontalAlignment("left");
  sheet.getRange(startRow, 5, numRows, 1).setHorizontalAlignment("center");
  sheet.getRange(startRow, 6, numRows, 1).setHorizontalAlignment("left");
  sheet.getRange(startRow, 7, numRows, 1).setHorizontalAlignment("center").setFontWeight("bold");
  sheet.getRange(startRow, 8, numRows, 1).setHorizontalAlignment("left");
  sheet.getRange(startRow, 9, numRows, 1).setHorizontalAlignment("left").setFontColor("#64748b");
}
