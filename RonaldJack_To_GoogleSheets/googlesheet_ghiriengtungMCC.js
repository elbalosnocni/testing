// Google Sheets ID của bạn: 1LshZZc_Pqf2hTq6CCC-6y69RjOPZoylYbGroaAqUrNM

function doPost(e) {
  var ss = SpreadsheetApp.openById("1LshZZc_Pqf2hTq6CCC-6y69RjOPZoylYbGroaAqUrNM");
  
  try {
    var data = JSON.parse(e.postData.contents);
    var now = new Date();
    
    // Duyệt qua từng dòng dữ liệu gửi về từ file .exe
    for (var i = 0; i < data.length; i++) {
      var record = data[i];
      var machineName = record.MachineName ? record.MachineName.trim() : "Khac";
      
      // Tìm xem đã có tab (sheet) nào trùng tên với Tên Máy Chấm Công chưa
      var sheet = ss.getSheetByName(machineName);
      
      // Nếu chưa có tab mang tên máy này, tự động tạo mới tab đó
      if (!sheet) {
        sheet = ss.insertSheet(machineName);
        // Chèn luôn dòng tiêu đề cho tab mới tạo
        sheet.appendRow(["Thời Gian Đồng Bộ", "Tên Máy Chấm Công", "Mã Nhân Viên", "Thời Gian Bấm Giờ", "Kiểu Bấm"]);
      }
      
      // Ghi dữ liệu chấm công vào đúng tab của máy đó
      sheet.appendRow([
        now, 
        machineName, 
        record.UserID, 
        record.DateTimeCheck, 
        record.CheckType
      ]);
    }
    
    return ContentService.createTextOutput(JSON.stringify({"status": "Thành công", "count": data.length}))
                         .setMimeType(ContentService.MimeType.JSON);
                         
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({"status": "Lỗi", "message": error.toString()}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}
