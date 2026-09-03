# Thiết lập lần đầu

Làm **một lần** trên máy Windows có Excel. Sau bước này, mọi bản phát hành tiếp
theo sẽ tự động — xem "Quy trình phát hành" trong `README.md`.

## 1. Nhúng module tự cập nhật vào LINK.xlam

1. Tải `release/LINK.xlam` từ repo về máy, mở bằng Excel (Enable Content nếu hỏi).
2. `Alt+F11` mở VBE.
3. Cửa sổ **Project Explorer** → chọn project của `LINK.xlam`.
4. `File → Import File…` → chọn `src/modAutoUpdate.bas`. Module `modAutoUpdate`
   xuất hiện trong nhánh *Modules*.
5. Double‑click **ThisWorkbook** → dán nội dung `src/ThisWorkbook.snippet.txt`.
   Nếu `ThisWorkbook` đã có sẵn `Sub Workbook_Open`, chỉ thêm dòng
   `modAutoUpdate.ScheduleUpdateCheck` vào trong sub đó.
6. Kiểm tra `ADDIN_VERSION` trong `modAutoUpdate` = `1.0.0` (khớp
   `release/version.txt` hiện tại).
7. `Debug → Compile VBAProject` — phải **không lỗi**.
8. `Ctrl+S` lưu `LINK.xlam`. Đóng Excel.

## 2. Test nhanh cơ chế (khuyến nghị)

Trên một máy test:

1. Cài `LINK.xlam` vừa sửa làm add-in (`File → Options → Add-ins → Manage: Excel
   Add-ins → Browse`).
2. Trên repo, tạm sửa `release/version.txt` thành `9.9.9`, commit + push (hoặc
   dùng một nhánh phụ và trỏ `GH_BASE` sang nhánh đó).
3. Mở Excel → sau ~2 giây hiện hộp thoại "Đã tải bản cập nhật 9.9.9…".
4. Kiểm tra có file `LINK.update.xlam` cạnh `LINK.xlam`.
5. Đóng **toàn bộ** Excel → chờ vài giây → Excel tự mở lại.
6. `Alt+F11`, cửa sổ Immediate: `? modAutoUpdate.ADDIN_VERSION` — vẫn `1.0.0` vì
   file trên repo vẫn là bản 1.0.0 (chỉ `version.txt` bị đổi số). Log
   `%TEMP%\link_addin_update.log` phải có `OK=1`.
7. Trả `release/version.txt` về `1.0.0`, commit + push.

> Để test đầy đủ (đổi cả version lẫn nội dung), bump `ADDIN_VERSION` lên `1.0.1`,
> lưu `LINK.xlam`, đẩy cả file + `version.txt=1.0.1` lên nhánh phụ rồi chạy lại.

## 3. Phát hành bản gốc

1. Commit `LINK.xlam` (đã có `modAutoUpdate`) vào `release/LINK.xlam`.
2. Giữ `release/version.txt` = `1.0.0`.
3. `git push`.

## 4. Rollout bản gốc tới toàn bộ người dùng (một lần)

Auto-update chỉ chạy được sau khi máy người dùng đã có `modAutoUpdate`. Vì vậy bản
gốc phải phát tay một lần:

1. Gửi `release/LINK.xlam` (bản có `modAutoUpdate`) cho tất cả người dùng, hoặc
   đặt ở nơi hiện đang phân phối add-in.
2. Người dùng thay add-in cũ bằng file này (gỡ add-in cũ → Browse tới file mới),
   hoặc ghi đè file `.xlam` ở đúng vị trí đang dùng rồi mở lại Excel.
3. Từ lần này trở đi, mọi bản mới đẩy lên `release/` sẽ tự về máy họ.

## Xử lý sự cố

| Hiện tượng | Kiểm tra |
|---|---|
| Không thấy hộp thoại cập nhật | Máy có chặn `raw.githubusercontent.com` không? Thử mở URL trong `GH_BASE` bằng trình duyệt. |
| Có `LINK.update.xlam` nhưng không cập nhật | Chưa đóng hết Excel. Đóng toàn bộ cửa sổ Excel. |
| `link_addin_update.log` = `OK=0` | File add-in ở thư mục không có quyền ghi, hoặc đường dẫn mạng. Đặt add-in ở `%APPDATA%\Microsoft\AddIns\`. |
| Excel không tự mở lại | `start "" excel.exe` không tìm thấy Excel trong PATH/App Paths. Người dùng tự mở Excel; add-in vẫn đã được cập nhật. |
