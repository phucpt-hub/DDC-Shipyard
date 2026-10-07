/**
 * CỬA RA DỮ LIỆU cho báo cáo chuỗi thép (dán vào project Apps Script của "Dashboard Tồn Kho Thép - VT02")
 *
 * Chỉ ĐỌC dữ liệu, không ghi/sửa gì trong bảng tính. Dashboard hiện tại vẫn chạy như cũ.
 *
 * Cách gắn (người có quyền sửa script làm, khoảng 5 phút):
 *  1. Mở project Apps Script của dashboard → thêm file mới "XuatDuLieu" → dán toàn bộ nội dung file này.
 *  2. Nếu script KHÔNG gắn sẵn với bảng tính tồn kho: dán mã bảng tính vào BANG_TINH_ID
 *     (đoạn giữa /d/ và /edit trong link Google Sheet). Gắn sẵn thì để trống.
 *  3. Project Settings → Script properties → thêm thuộc tính KHOA_XUAT = một chuỗi bí mật bất kỳ
 *     (ví dụ 24 ký tự ngẫu nhiên). Gửi chuỗi này cho người quản lý báo cáo để dán vào GitHub Secrets.
 *  4. Trong hàm doGet(e) đang có, thêm MỘT dòng ở đầu hàm:
 *         var _x = xuatDuLieu_(e); if (_x) return _x;
 *  5. Deploy → Manage deployments → bút chì (Edit) → Version: New version → Deploy.
 *     (Sửa bản deploy cũ để GIỮ NGUYÊN link /exec; lần đầu Google sẽ hỏi cấp quyền → Allow.)
 *
 * Kiểm tra: mở  <link /exec>?xuat=kiemtra&khoa=<KHOA_XUAT>  → thấy tên bảng tính và danh sách sheet là được.
 */
var BANG_TINH_ID = '';   // để trống nếu script gắn với bảng tính tồn kho

function bangTinhId_() {
  return BANG_TINH_ID || SpreadsheetApp.getActive().getId();
}

function xuatDuLieu_(e) {
  var p = (e && e.parameter) || {};
  if (!p.xuat) return null;                                   // không phải yêu cầu xuất → dashboard chạy bình thường
  var khoa = PropertiesService.getScriptProperties().getProperty('KHOA_XUAT');
  if (!khoa || p.khoa !== khoa) return ContentService.createTextOutput('LOI: sai khoa');
  var id = bangTinhId_();
  if (p.xuat === 'kiemtra') {
    var ss = SpreadsheetApp.openById(id);
    return ContentService.createTextOutput(JSON.stringify({
      ten: ss.getName(), sheet: ss.getSheets().map(function (s) { return s.getName() + ' (' + s.getLastRow() + ' dòng)'; }),
      luc: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
  if (p.xuat === 'xlsx') {                                    // nguyên bảng tính dạng Excel, mã hóa base64
    var url = 'https://docs.google.com/spreadsheets/d/' + id + '/export?format=xlsx';
    var blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob();
    return ContentService.createTextOutput(Utilities.base64Encode(blob.getBytes()));
  }
  return ContentService.createTextOutput('LOI: xuat=kiemtra | xlsx');
}

// Gợi ý cho Apps Script biết cần quyền đọc Drive để xuất file (không chạy, chỉ để khai báo quyền):
// DriveApp.getFiles();
