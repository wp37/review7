/**
 * Google Apps Script - Lưu kết quả bài kiểm tra vào Google Sheets
 * 
 * HƯỚNG DẪN:
 * 1. Dán code này vào thay thế code cũ.
 * 2. Lưu lại (Ctrl+S).
 * 3. Bấm "Deploy" > "Manage deployments" > Bấm icon bút chì > "New version" > "Deploy".
 * 4. Kiểm tra lại URL, nếu URL đổi thì copy URL mới.
 */

const SHEET_NAME = "Kết Quả Quiz";

// --- Xử lý dữ liệu gửi đến (GET Request) ---
function doGet(e) {
  try {
    // Nếu có dữ liệu gửi kèm (người dùng nộp bài)
    if (e && e.parameter && e.parameter.studentName) {
      
      // TÍNH ĐIỂM SERVER-SIDE (AN TOÀN)
      // Mỗi câu đúng 0.5 điểm. Tổng 20 câu = 10 điểm.
      const correctCount = parseInt(e.parameter.correctCount) || 0;
      const calculatedScore = (correctCount * 0.5).toFixed(1);

      const data = {
        timestamp: new Date().toISOString(), // Dùng giờ gửi request
        studentName: e.parameter.studentName,
        studentClass: e.parameter.studentClass || '',
        parentPhone: e.parameter.parentPhone || '',
        score: calculatedScore, // Dùng điểm do server tính
        correctCount: correctCount,
        totalQuestions: e.parameter.totalQuestions || 20,
        timeUsed: e.parameter.timeUsed,
        details: e.parameter.details || ''
      };
      
      saveToSheet(data);
      
      // Trả về thông báo thành công
      return HtmlService.createHtmlOutput('<html><body><script>console.log("Saved");</script>OK</body></html>');
    }
    
    // Nếu không có dữ liệu (test link)
    return ContentService.createTextOutput(JSON.stringify({status: "ok", message: "Quiz API is ready"}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  // Fallback nếu có request POST
  return doGet(e); 
}

// --- Hàm lưu vào Sheet ---
function saveToSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    setupSheet();
    sheet = ss.getSheetByName(SHEET_NAME);
  }
  
  const timestamp = new Date(); // Lấy giờ server (Việt Nam)
  const vietnamTime = Utilities.formatDate(timestamp, "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm:ss");
  
  const newRow = [
    vietnamTime,
    data.studentName,
    data.studentClass,
    data.parentPhone, // Số điện thoại
    "'" + data.score, // Thêm dấu ' để Sheet hiểu là chuỗi/số hiển thị đúng
    data.correctCount,
    data.totalQuestions,
    data.timeUsed,
    data.details
  ];
  
  sheet.appendRow(newRow);
  
  // Tô màu điểm
  const lastRow = sheet.getLastRow();
  const scoreCell = sheet.getRange(lastRow, 5); // Cột E là cột điểm (số 5)
  const score = parseFloat(data.score);
  
  if (score >= 9) scoreCell.setBackground("#00b894").setFontColor("white"); // Xuất sắc
  else if (score >= 7) scoreCell.setBackground("#0984e3").setFontColor("white"); // Giỏi
  else if (score >= 5) scoreCell.setBackground("#ffeaa7").setFontColor("black"); // Khá
  else scoreCell.setBackground("#ff7675").setFontColor("white"); // Yếu/TB
}

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  
  const headers = ["Thời gian nộp", "Họ và tên", "Lớp", "SĐT Cha/Mẹ", "Điểm", "Số câu đúng", "Tổng số câu", "Thời gian làm bài", "Chi tiết"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
       .setFontWeight("bold").setBackground("#667eea").setFontColor("white").setHorizontalAlignment("center");
  sheet.setFrozenRows(1);
  
  // Set độ rộng cột
  sheet.setColumnWidth(1, 160); // Thời gian
  sheet.setColumnWidth(2, 180); // Tên
  sheet.setColumnWidth(3, 60);  // Lớp
  sheet.setColumnWidth(4, 110); // SĐT
  sheet.setColumnWidth(5, 60);  // Điểm
}
