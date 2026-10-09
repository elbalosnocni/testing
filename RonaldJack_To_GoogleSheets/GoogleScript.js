// Google Sheets ID cua ban: 1LshZZc_Pqf2hTq6CCC-6y69RjOPZoylYbGroaAqUrNM

function doPost(e) {
  // Mo dung file Google Sheet cua ban bang ID thay vi MacDinh
  var ss = SpreadsheetApp.openById("1LshZZc_Pqf2hTq6CCC-6y69RjOPZoylYbGroaAqUrNM");
  var sheet = ss.getSheets()[0]; // Lay sheet dau tien
  
  try {
    var data = JSON.parse(e.postData.contents);
    
    // Tao dong tieu de neu Sheet chua co gi
    if (sheet.getLastRow() == 0) {
      sheet.appendRow(["Thời Gian Đồng Bộ", "Tên Máy Chấm Công", "Mã Nhân Viên", "Thời Gian Bấm Giờ", "Kiểu Bấm"]);
    }
    
    var now = new Date();
    // Ghi thong tin tung records vao Sheet
    for (var i = 0; i < data.length; i++) {
      sheet.appendRow([
        now, 
        data[i].MachineName, 
        data[i].UserID, 
        data[i].DateTimeCheck, 
        data[i].CheckType
      ]);
    }
    
    return ContentService.createTextOutput(JSON.stringify({"status": "Thành công", "count": data.length}))
                         .setMimeType(ContentService.MimeType.JSON);
                         
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({"status": "Lỗi", "message": error.toString()}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}
