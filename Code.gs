/**
 * =========================================================================
 * GOOGLE APPS SCRIPT - HỆ THỐNG NHẬP ĐIỂM KHỐI 12 (NĂM HỌC 2026-2027)
 * TÍNH NĂNG MỚI:
 * 1. ĐIỂM CỘNG TUẦN: Mỗi học sinh trong 1 tuần chỉ ghi 1 dòng. Nộp lại sẽ tự động GHI ĐÈ KẾT QUẢ CŨ.
 * 2. ĐIỂM KIỂM TRA: Gọn gàng gồm Tên học sinh, Điểm, Ghi chú.
 * 3. ĐIỂM NHÓM: Hỗ trợ tùy chỉnh điểm riêng biệt cho từng thành viên trong nhóm.
 * =========================================================================
 */

/**
 * HÀM TỰ ĐỘNG CĂN CHỈNH & ĐỊNH DẠNG SHEET THEO ĐÚNG MẪU BẢNG ĐIỂM
 */
function formatSheetLikeImage() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  var headers = [
    "Thời gian",
    "Lớp",
    "STT",
    "Họ và tên",
    "Loại điểm",
    "Chi tiết / Hạng mục",
    "Điểm số",
    "Nội dung / Ghi chú"
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
  }

  for (var i = 1; i <= numCols; i++) {
    sheet.autoResizeColumn(i);
    var currentWidth = sheet.getColumnWidth(i);
    sheet.setColumnWidth(i, Math.max(currentWidth + 20, 80));
  }
}

/**
 * XỬ LÝ DỮ LIỆU GỬI TỪ WEB APP
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var payload = JSON.parse(e.postData.contents);
    var timestamp = new Date().toLocaleString("vi-VN");

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
          hs.note || payload.note || ""
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
        var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
        
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
          payload.note || ""
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
        payload.note || ""
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
 * ĐỊNH DẠNG CÁC DÒNG MỚI ĐỒNG BỘ
 */
function formatDataRows(sheet, startRow, numRows) {
  var range = sheet.getRange(startRow, 1, numRows, 8);
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
}
