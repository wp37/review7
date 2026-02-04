# Hướng dẫn kết nối Google Sheets

## Bước 1: Tạo Google Sheet mới

1. Truy cập [Google Sheets](https://sheets.google.com)
2. Tạo một bảng tính mới
3. Đặt tên cho sheet là "Kết quả kiểm tra Unit 7"
4. Tạo các cột header ở hàng 1:
   - A1: `Thời gian`
   - B1: `Họ và tên`
   - C1: `Lớp`
   - D1: `Số câu đúng`
   - E1: `Số câu sai`
   - F1: `Tổng câu`
   - G1: `Điểm`
   - H1: `Chi tiết đáp án`

## Bước 2: Tạo Google Apps Script

1. Trong Google Sheet, vào **Extensions** > **Apps Script**
2. Xóa code mặc định và dán đoạn code sau:

```javascript
function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);
    
    sheet.appendRow([
      data.timestamp,
      data.studentName,
      data.studentClass,
      data.correct,
      data.incorrect,
      data.total,
      data.grade,
      data.answers
    ]);
    
    return ContentService
      .createTextOutput(JSON.stringify({status: 'success'}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch(error) {
    return ContentService
      .createTextOutput(JSON.stringify({status: 'error', message: error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput("Quiz API is running!")
    .setMimeType(ContentService.MimeType.TEXT);
}
```

3. Lưu file (Ctrl + S)
4. Đặt tên project là "Quiz API"

## Bước 3: Deploy Web App

1. Click **Deploy** > **New deployment**
2. Chọn loại: **Web app**
3. Cấu hình:
   - **Description**: Quiz Results API
   - **Execute as**: Me (email của bạn)
   - **Who has access**: Anyone
4. Click **Deploy**
5. **Authorize access** khi được yêu cầu
6. Copy **Web app URL** (ví dụ: `https://script.google.com/macros/s/xxxxx/exec`)

## Bước 4: Cập nhật file app.js

Mở file `app.js` và thay thế dòng:

```javascript
const GOOGLE_SCRIPT_URL = 'YOUR_GOOGLE_APPS_SCRIPT_URL_HERE';
```

Bằng URL bạn vừa copy, ví dụ:

```javascript
const GOOGLE_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxxxxxxxxxxxxxxxxxxxxxxxxxx/exec';
```

## Bước 5: Kiểm tra

1. Mở file `index.html` trong trình duyệt
2. Làm bài kiểm tra thử
3. Kiểm tra Google Sheet xem kết quả đã được ghi chưa

---

## Lưu ý quan trọng

- ⚠️ **Bảo mật**: URL của Google Apps Script là công khai. Chỉ chia sẻ cho học sinh khi cần.
- 📊 **Dữ liệu**: Tất cả kết quả sẽ được lưu tự động vào Google Sheet.
- 🔄 **Cập nhật**: Nếu cần thay đổi script, bạn phải deploy lại và cập nhật URL mới.

## Hỗ trợ

Nếu gặp vấn đề, vui lòng liên hệ:
- 💬 Zalo: [0914666040](https://zalo.me/0914666040)
- 📘 Facebook: [facebook.com/vongoctungthcs](https://facebook.com/vongoctungthcs)
