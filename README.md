# Dashboard BH Một đổi một · Viettel Store × BH PVI

## Cập nhật hằng ngày
1. Xuất báo cáo từ hệ thống, lưu file `bhkh-YYYY-MM-DD.xlsx` vào thư mục này.
2. Đóng file Excel (khuyến nghị), bấm đúp `update.bat`.
3. Script đọc **tất cả** file `bhkh-*.xlsx`, gộp và loại trùng theo Số đơn BH (file mới nhất được ưu tiên), tạo lại `data.js`, rồi commit + push lên GitHub.

Không cần cài Python hay Excel — chỉ dùng PowerShell có sẵn trên Windows.

## Quy tắc số liệu
- Đơn có hậu tố `/E01 SĐBS` (mọi `/Exx`) = đơn hủy, gắn với đơn gốc cùng số.
- Đơn thực tế = Đơn ghi nhận − Đơn hủy. Phí thực tế = phí ghi nhận − phí đơn hủy.
- Nút "Ghi nhận hủy theo": *Ngày SĐBS* (hủy tính vào ngày phát sinh SĐBS) hoặc *Ngày phát hành gốc* (tính ngược về ngày đơn gốc).
- Tên khách hàng được ẩn thành chữ cái đầu trước khi ghi vào `data.js`. File `.xlsx` gốc không được đẩy lên GitHub (`.gitignore`).

## Kết nối GitHub lần đầu
Repo: https://github.com/PVI-AviationDivision/dashboard_VST_BHMDM_PVI

1. Cài Git for Windows nếu chưa có: https://git-scm.com/download/win
2. Bấm đúp `setup-github.bat` (chỉ chạy 1 lần). Lần đầu push, Git sẽ mở cửa sổ đăng nhập GitHub — dùng tài khoản có quyền ghi vào tổ chức PVI-AviationDivision.
3. Trên GitHub: **Settings → Pages → Build and deployment → Deploy from a branch → `main` / `(root)` → Save**.
4. Link dashboard: https://pvi-aviationdivision.github.io/dashboard_VST_BHMDM_PVI/

Từ đó trở đi mỗi ngày chỉ cần bấm `update.bat`.

> Lưu ý: repo đang để **Public** nên ai có link đều xem được `data.js` (số đơn, số tiền BH, tên KH đã ẩn chữ cái đầu). File Excel gốc không bao giờ được đẩy lên.
