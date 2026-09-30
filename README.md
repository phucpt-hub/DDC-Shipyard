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

## Báo cáo cũ

Mỗi lần chạy, báo cáo được lưu thêm vào `lich_su/<ngày>/`. Xem danh sách tại `/lich_su/` trên trang báo cáo.

## Khi chạy lỗi

Tab **Actions** hiện dấu X đỏ → bấm vào lần chạy → mở bước **Chạy phân tích và dựng báo cáo** để xem thông báo (thường là thiếu file, sai tên sheet, hoặc sheet bị đổi cấu trúc). Trang báo cáo cũ vẫn giữ nguyên khi lỗi.

## Thành phần

- `cong_cu/chay.py`: đọc Excel, chạy phân tích, dựng `index.html` và `lich_su/`.
- `cong_cu/data.py`: tổng hợp số liệu Cutting, tồn kho, nesting, cân đối thép.
- `cong_cu/template.html`: giao diện báo cáo.
- `.github/workflows/cap-nhat-bao-cao.yml`: tự chạy khi có file mới trong `du_lieu/`.
