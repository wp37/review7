/**
 * Google Apps Script - Lưu kết quả bài kiểm tra vào Google Sheets
 * 
 * HƯỚNG DẪN SỬ DỤNG:
 * 1. Mở Google Sheets mới
 * 2. Vào Extensions > Apps Script
 * 3. Dán toàn bộ code này vào
 * 4. Lưu và chạy hàm setupSheet() một lần
 * 5. Vào Deploy > New deployment
 * 6. Chọn "Web app"
 * 7. Execute as: Me
 * 8. Who has access: Anyone
 * 9. Copy URL và dán vào file HTML (thay YOUR_APPS_SCRIPT_URL_HERE)
 */

// Tên sheet lưu kết quả
const SHEET_NAME = "Kết Quả Quiz";

/**
 * Chạy một lần để thiết lập sheet
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  
  // Thiết lập header
  const headers = [
    "Thời gian nộp",
    "Họ và tên",
    "Lớp",
    "SĐT Cha/Mẹ",
    "Điểm",
    "Số câu đúng",
    "Tổng số câu",
    "Thời gian làm bài",
    "Chi tiết"
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Định dạng header
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#667eea");
  headerRange.setFontColor("#ffffff");
  headerRange.setHorizontalAlignment("center");
  
  // Cố định hàng header
  sheet.setFrozenRows(1);
  
  // Điều chỉnh độ rộng cột
  sheet.setColumnWidth(1, 180); // Thời gian
  sheet.setColumnWidth(2, 200); // Họ tên
  sheet.setColumnWidth(3, 80);  // Lớp
  sheet.setColumnWidth(4, 120); // SĐT Cha/Mẹ
  sheet.setColumnWidth(5, 70);  // Điểm
  sheet.setColumnWidth(6, 100); // Số câu đúng
  sheet.setColumnWidth(7, 100); // Tổng câu
  sheet.setColumnWidth(8, 120); // Thời gian làm
  sheet.setColumnWidth(9, 300); // Chi tiết
  
  Logger.log("✅ Đã thiết lập sheet thành công!");
}

/**
 * Xử lý POST request từ web app
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    saveToSheet(data);
    
    return ContentService
      .createTextOutput(JSON.stringify({status: "success"}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    Logger.log("Error: " + error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({status: "error", message: error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Xử lý GET request (để test)
 */
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({status: "ok", message: "Quiz API is running"}))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Lưu dữ liệu vào sheet
 */
function saveToSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    setupSheet();
    sheet = ss.getSheetByName(SHEET_NAME);
  }
  
  // Chuyển đổi timestamp sang múi giờ Việt Nam
  const timestamp = new Date(data.timestamp);
  const vietnamTime = Utilities.formatDate(timestamp, "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm:ss");
  
  // Tạo hàng dữ liệu mới
  const newRow = [
    vietnamTime,
    data.studentName,
    data.studentClass,
    data.parentPhone || "",
    data.score,
    data.correctCount,
    data.totalQuestions,
    data.timeUsed,
    data.details || ""
  ];
  
  // Thêm vào cuối sheet
  sheet.appendRow(newRow);
  
  // Định dạng điểm (tô màu theo mức điểm)
  const lastRow = sheet.getLastRow();
  const scoreCell = sheet.getRange(lastRow, 5);
  const score = parseFloat(data.score);
  
  if (score >= 9) {
    scoreCell.setBackground("#00b894"); // Xanh lá - Xuất sắc
    scoreCell.setFontColor("#ffffff");
  } else if (score >= 7) {
    scoreCell.setBackground("#0984e3"); // Xanh dương - Tốt
    scoreCell.setFontColor("#ffffff");
  } else if (score >= 5) {
    scoreCell.setBackground("#ffeaa7"); // Vàng - Khá
    scoreCell.setFontColor("#2d3436");
  } else {
    scoreCell.setBackground("#ff7675"); // Đỏ - Cần cố gắng
    scoreCell.setFontColor("#ffffff");
  }
  
  Logger.log("✅ Đã lưu kết quả: " + data.studentName + " - " + data.score + " điểm");
}

/**
 * Lấy thống kê tổng quan
 */
function getStatistics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() < 2) {
    return { message: "Chưa có dữ liệu" };
  }
  
  const data = sheet.getDataRange().getValues();
  const scores = data.slice(1).map(row => parseFloat(row[3]));
  
  const stats = {
    totalStudents: scores.length,
    averageScore: (scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(2),
    highestScore: Math.max(...scores),
    lowestScore: Math.min(...scores),
    excellentCount: scores.filter(s => s >= 9).length,
    goodCount: scores.filter(s => s >= 7 && s < 9).length,
    averageCount: scores.filter(s => s >= 5 && s < 7).length,
    belowAverageCount: scores.filter(s => s < 5).length
  };
  
  Logger.log(JSON.stringify(stats, null, 2));
  return stats;
}

/**
 * Tạo báo cáo theo lớp
 */
function getClassReport(className) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() < 2) {
    return { message: "Chưa có dữ liệu" };
  }
  
  const data = sheet.getDataRange().getValues();
  const classData = data.slice(1).filter(row => row[2] === className);
  
  if (classData.length === 0) {
    return { message: "Không tìm thấy dữ liệu lớp " + className };
  }
  
  const scores = classData.map(row => parseFloat(row[3]));
  
  return {
    className: className,
    totalStudents: scores.length,
    averageScore: (scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(2),
    highestScore: Math.max(...scores),
    lowestScore: Math.min(...scores)
  };
}

/**
 * Xóa tất cả dữ liệu (cẩn thận!)
 */
function clearAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (sheet && sheet.getLastRow() > 1) {
    sheet.deleteRows(2, sheet.getLastRow() - 1);
    Logger.log("✅ Đã xóa tất cả dữ liệu!");
  }
}
