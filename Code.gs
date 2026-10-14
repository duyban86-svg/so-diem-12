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
 * HÀM CHUYỂN ĐỔI CHUỖI ĐIỂM SANG SỐ THỰC CHUẨN HOẶC ĐỊNH DẠNG ĐÚNG CHO SHEET
 * Chấp nhận cả "8.5", "8,5", "+1.5", "+1,5", v.v.
 */
function normalizeScoreValue(rawVal) {
  if (rawVal === undefined || rawVal === null || rawVal === "") return "";
  var str = String(rawVal).trim();
  var isPlus = str.startsWith("+");
  var cleanStr = str.replace(/^\+/, "").replace(/,/g, ".");
  var num = parseFloat(cleanStr);
  if (!isNaN(num)) {
    // Nếu là điểm cộng có dấu + thì trả về chuỗi có dấu + kèm dấu phẩy (VD: "+1,5") hoặc số tuỳ ngữ cảnh
    if (isPlus) {
      return "+" + String(num).replace(/\./g, ",");
    }
    return num; // Trả về dạng Number để Google Sheet nhận diện chuẩn là Số
  }
  return rawVal;
}

/**
 * XỬ LÝ YÊU CẦU GET (TRA CỨU TÀI KHOẢN LIÊN KẾT & XEM ĐIỂM CÁ NHÂN)
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

    // 1. Kiểm tra trạng thái liên kết của email
    if (action === "check_email") {
      var targetEmail = (e.parameter.email || "").trim().toLowerCase();
      var found = accounts.find(function(a) { return a.email === targetEmail; });
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        bound: !!found,
        student: found || null
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 2. Tra cứu điểm cá nhân của học sinh (BẢO MẬT: ĐỌC TỪ TRANG TÍNH TỔNG HỢP CỦA TỪNG LỚP)
    if (action === "get_my_scores") {
      var queryEmail = (e.parameter.email || "").trim().toLowerCase();
      var queryLop = (e.parameter.lop || "").trim();
      var queryStt = (e.parameter.stt || "").trim();

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var myScores = [];
      var studentFoundName = "";
      var studentFoundNote = "";

      // 1. Tìm trang tính riêng của lớp (VD: "12A13", "12A08", "10A01", v.v.)
      var classSheet = ss.getSheetByName(queryLop);
      if (!classSheet) {
        // Thử tìm các biến thể nếu có (VD: "Lớp 12A13")
        classSheet = ss.getSheetByName("Lớp " + queryLop);
      }

      if (classSheet && classSheet.getLastRow() >= 2 && classSheet.getLastColumn() >= 4) {
        var numRows = classSheet.getLastRow();
        var numCols = classSheet.getLastColumn();
        
        // Đọc toàn bộ tiêu đề (Dòng 1) và dữ liệu học sinh
        var allValues = classSheet.getRange(1, 1, numRows, numCols).getValues();
        var headers = allValues[0]; // Dòng 1: [STT, Lớp, Họ và tên, Điểm cộng theo tuần, Điểm nhóm (1), Email, ...]

        // Tìm dòng của học sinh theo STT (Cột A - index 0)
        var targetRowIndex = -1;
        for (var r = 1; r < allValues.length; r++) {
          var rowStt = String(allValues[r][0]).trim();
          var rowEmail = (numCols >= 6 && allValues[r][5]) ? String(allValues[r][5]).trim().toLowerCase() : "";

          if (queryStt && rowStt === queryStt) {
            targetRowIndex = r;
            break;
          } else if (queryEmail && rowEmail === queryEmail) {
            targetRowIndex = r;
            break;
          }
        }

        if (targetRowIndex !== -1) {
          var targetRow = allValues[targetRowIndex];
          studentFoundName = String(targetRow[2] || "").trim(); // Cột C: Họ và tên

          // Duyệt tự động từ Cột D (index 3) đến hết cột cuối cùng bên phải
          for (var c = 3; c < numCols; c++) {
            var colTitle = String(headers[c] || "").trim();
            var colVal = targetRow[c];

            // Bỏ qua các cột chứa email hoặc không có tiêu đề
            var lowerTitle = colTitle.toLowerCase();
            if (!colTitle || lowerTitle.includes("email") || lowerTitle.includes("tài khoản") || lowerTitle.includes("tai khoan")) {
              continue;
            }

            // Nếu là cột Ghi chú / Nhận xét của Thầy
            if (lowerTitle.includes("ghi chú") || lowerTitle.includes("nhận xét") || lowerTitle.includes("lời dặn") || lowerTitle.includes("dặn dò")) {
              if (colVal !== undefined && colVal !== null && String(colVal).trim() !== "") {
                studentFoundNote += (studentFoundNote ? " | " : "") + String(colVal).trim();
              }
              continue;
            }

            // Nếu là cột điểm (có giá trị)
            if (colVal !== undefined && colVal !== null && String(colVal).trim() !== "") {
              var displayScore = colVal;
              // Nếu là số thì định dạng dấu phẩy tiếng Việt cho đẹp
              if (typeof colVal === "number") {
                displayScore = String(colVal).replace(/\./g, ",");
              }

              myScores.push({
                time: "Đã tổng hợp",
                lop: queryLop,
                stt: String(targetRow[0]).trim(),
                name: studentFoundName,
                scoreType: colTitle,
                category: colTitle,
                score: displayScore,
                note: ""
              });
            }
          }

          // Gắn nhận xét vào mục đầu tiên nếu có
          if (studentFoundNote && myScores.length > 0) {
            myScores[0].note = studentFoundNote;
          }
        }
      } 
      // 2. Dự phòng: Nếu chưa có trang lớp tổng hợp, đọc từ trang DiemSo như cũ
      else {
        var scoreSheet = ss.getSheetByName(SHEET_NAME_SCORES);
        if (scoreSheet && scoreSheet.getLastRow() >= 2) {
          var numScoreRows = scoreSheet.getLastRow() - 1;
          var scoreData = scoreSheet.getRange(2, 1, numScoreRows, 9).getValues();

          for (var sIdx = 0; sIdx < scoreData.length; sIdx++) {
            var sRow = scoreData[sIdx];
            var rTime = sRow[0];
            var rLop = String(sRow[1]).trim();
            var rStt = String(sRow[2]).trim();
            var rName = String(sRow[3]).trim();
            var rType = String(sRow[4]).trim();
            var rCategory = String(sRow[5]).trim();
            var rScore = sRow[6];
            var rNote = String(sRow[7] || "").trim();
            var rEmail = String(sRow[8] || "").trim().toLowerCase();

            var isMatch = false;
            if (queryLop && queryStt && rLop === queryLop && rStt === queryStt) {
              isMatch = true;
            } else if (queryEmail && rEmail === queryEmail) {
              isMatch = true;
            }

            if (isMatch) {
              myScores.push({
                time: rTime instanceof Date ? Utilities.formatDate(rTime, "GMT+7", "dd/MM/yyyy HH:mm") : String(rTime),
                lop: rLop,
                stt: rStt,
                name: rName,
                scoreType: rType,
                category: rCategory,
                score: rScore,
                note: rNote
              });
            }
          }
        }
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        count: myScores.length,
        studentName: studentFoundName,
        teacherNote: studentFoundNote,
        scores: myScores
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 3. Toàn bộ danh sách liên kết tài khoản
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
        var itemScore = hs.score !== undefined ? hs.score : (payload.score || "");
        rows.push([
          timestamp,
          payload.lop || "",
          hs.stt || "",
          hs.name || "",
          payload.scoreType || "Điểm nhóm",
          payload.category || "",
          normalizeScoreValue(itemScore),
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
      var normalizedScore = normalizeScoreValue(payload.score);

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
            sheet.getRange(overwrittenRow, 7).setValue(normalizedScore); // Cập nhật điểm mới
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
          normalizedScore,
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
        normalizeScoreValue(payload.score),
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
