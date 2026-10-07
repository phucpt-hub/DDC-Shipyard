# Báo cáo chuỗi thép – Sơ chế Vũng Tàu

Kho → Nesting → Cutting → Bàn giao. Trang báo cáo: **https://phucpt-hub.github.io/DDC-Shipyard/**

## Cập nhật báo cáo (không cần cài gì)

1. Mở thư mục cần cập nhật trong `du_lieu/` (bảng dưới).
2. Bấm **Add file → Upload files**, kéo file Excel vào, bấm **Commit changes**.
3. Chờ khoảng 3 phút. Tab **Actions** hiện dấu tích xanh là xong, mở lại trang báo cáo (Ctrl+F5).

| Thư mục | File | Ghi chú |
|---|---|---|
| `du_lieu/1_CUTTING` | `DDC_SHIPYARD - KHGC - SC.xlsx` | Bắt buộc. Chỉ giữ 1 file, tải bản mới cùng tên để ghi đè. |
| `du_lieu/2_TON_KHO` | `TỒN KHO dd.mm.xlsx` | Bắt buộc. Ngày trong tên file là ngày báo cáo. Nếu có nhiều file, lấy file tải lên sau cùng. |
| `du_lieu/3_NESTING` | `VT-xx-TT-nnn.xlsx` | Nên có. Mỗi dự án một file bản mới nhất; dự án tự nhận theo mã LSX. |
| `du_lieu/4_CAU_HINH` | `MAP_VA_QUY_TAC.xlsx` | Chỉ sửa khi có dự án mới (map công trình Cutting → dự án kho). |

Có thể tải nhiều file trong một lần commit. Muốn chạy lại mà không đổi file: tab **Actions → Cập nhật báo cáo chuỗi thép → Run workflow**.

## Công cụ làm sạch dữ liệu

Mỗi lần chạy, `cong_cu/clean.py` tự chuẩn hóa dữ liệu trước khi tính báo cáo (ngày ghi kèm chữ, ngày đảo ngày/tháng, bàn giao thiếu ngày cắt, thiếu dấu "x", phân loại thép trống, ngày xuất kho bị đảo, số nesting lặp, mác viết khác kiểu…). File gốc không bị sửa.
Kết quả ghi vào **`nhat_ky_lam_sach.xlsx`** (tải từ trang báo cáo, mục *Dữ liệu & làm sạch*): sheet TONG_HOP theo quy tắc, sheet CHI_TIET từng ô (file, sheet, dòng Excel, cột, giá trị cũ → mới) để các phòng sửa file gốc. Các dòng loại *Cần xác nhận* được giữ nguyên như file.

## Trang Tóm tắt điều hành (trang mở đầu)

- **Tình hình cần chú ý**: 3–5 vấn đề lớn nhất, tự rút ra từ số liệu (thiếu thép, dự án đỏ, bàn giao chậm, LSX tồn lâu, nesting chưa nạp Cutting, phôi chờ cắt).
- **Chỉ số chính**: so với báo cáo đã lưu gần nhất của ngày trước (thư mục `lich_su/`). Xanh = tốt lên, đỏ = xấu đi.
- **Sức khỏe dự án**: Đỏ = thiếu thép ≥ 5 t, hoặc LSX quá 60 ngày còn ≥ 5 t, hoặc ≥ 50 t đã nesting chưa nạp Cutting. Cam = có vấn đề nhỏ hơn (thiếu thép, LSX quá 30 ngày, nesting chưa nạp, hàng cắt chờ bàn giao nhiều). Xanh = còn lại.
- **Việc cần xử lý / cần quyết định**: ghi trong sheet `VIEC_CAN_XU_LY` của `du_lieu/4_CAU_HINH/MAP_VA_QUY_TAC.xlsx` (Việc · Dự án · Phụ trách · Hạn · Trạng thái · Ghi chú). Trạng thái "Xong" sẽ tự ẩn. Muốn sửa nhanh trên Google Sheets: tạo Google Sheet cùng các cột, chia sẻ "ai có link đều xem được", dán link vào secret `NGUON_VIEC`. Bên dưới có thêm các việc **đề xuất** tự rút ra từ số liệu.

## Lấy dữ liệu tự động từ link

Secrets trong Settings → Secrets and variables → Actions: `NGUON_CUTTING` (Google Sheet Cutting), `NGUON_KHO` (+ `NGUON_KHO_KHOA` nếu là Apps Script), `NGUON_NESTING`, `NGUON_VIEC`. Báo cáo tự chạy 07:30 và 17:30 giờ Việt Nam; nút Run workflow để chạy ngay.

## Báo cáo cũ

Mỗi lần chạy, báo cáo được lưu thêm vào `lich_su/<ngày>/`. Xem danh sách tại `/lich_su/` trên trang báo cáo.

## Khi chạy lỗi

Tab **Actions** hiện dấu X đỏ → bấm vào lần chạy → mở bước **Chạy phân tích và dựng báo cáo** để xem thông báo (thường là thiếu file, sai tên sheet, hoặc sheet bị đổi cấu trúc). Trang báo cáo cũ vẫn giữ nguyên khi lỗi.

## Thành phần

- `cong_cu/chay.py`: đọc Excel, chạy phân tích, dựng `index.html` và `lich_su/`.
- `cong_cu/data.py`: tổng hợp số liệu Cutting, tồn kho, nesting, cân đối thép.
- `cong_cu/template.html`: giao diện báo cáo.
- `.github/workflows/cap-nhat-bao-cao.yml`: tự chạy khi có file mới trong `du_lieu/`.
