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
gốc phải phát tay một lần — dùng **[install.bat](./install.bat)**, không cần
người dùng làm thao tác VBE/Options nào cả.

> **Click đúp vào `.xlam` KHÔNG cài đặt add-in** — chỉ mở file cho phiên đang
> chạy, không đăng ký, không tự load lại lần sau. `install.bat` tránh hẳn vấn đề
> này bằng cách đặt file vào thư mục `XLSTART` — nơi Excel tự mở mọi file trong
> đó ở mỗi lần khởi động, không cần đăng ký qua `File → Options → Add-ins`,
> không cần quyền admin.

Gửi người dùng:

1. Gửi link tải `install.bat` (raw URL:
   `https://raw.githubusercontent.com/namtao/add-in/main/install.bat` — hoặc gửi
   trực tiếp file qua Zalo/email).
2. Người dùng double-click `install.bat`. Script tự:
   - tải `release/LINK.xlam` mới nhất từ GitHub,
   - kiểm tra kích thước file hợp lệ,
   - nếu Excel đang mở, nhắc đóng hết rồi mới ghi file,
   - copy vào `%APPDATA%\Microsoft\Excel\XLSTART\LINK.xlam`.
3. Mở Excel để xác nhận ribbon LINK hiện ra.

Từ đây, `ThisWorkbook.FullName` mà `modAutoUpdate` dùng làm đích ghi đè luôn trỏ
vào `XLSTART\LINK.xlam` — ổn định, có quyền ghi, không phụ thuộc người dùng có
dọn Downloads hay không. Mọi bản mới đẩy lên `release/` sau đó tự về máy họ,
không cần gửi lại file hay chạy lại `install.bat`.

> Nếu người dùng đã từng cài `LINK` theo cách cũ (qua `File → Options →
> Add-ins → Browse`), `install.bat` sẽ cảnh báo và nhắc gỡ bản cũ để tránh
> ribbon hiện 2 lần.

## 5. Người khác (không phải bạn) muốn phát hành bản mới

Cơ chế auto-update nghĩa là **không cần "gửi cho từng người" nữa** — chỉ cần
cập nhật `release/` trên GitHub, mọi máy đã bootstrap (mục 4) tự nhận ở lần mở
Excel kế tiếp. Người phát hành mới cần:

1. **Quyền ghi vào repo:** bạn (chủ repo) vào
   `github.com/namtao/add-in → Settings → Collaborators` → thêm họ làm
   collaborator (Write), hoặc để họ fork repo rồi gửi Pull Request cho bạn duyệt.
2. **Sửa add-in:** họ tải `release/LINK.xlam` hiện tại (đã có sẵn
   `modAutoUpdate` — không cần làm lại mục 1) về máy, mở trong Excel, sửa nội
   dung/ribbon/VBA cần thiết, tăng `ADDIN_VERSION` trong `modAutoUpdate`
   (ví dụ `1.0.0` → `1.0.1`), `Debug → Compile`, lưu.
3. **Đẩy lên GitHub** — hai cách, không bắt buộc biết Git:
   - **Không cần cài Git:** vào `github.com/namtao/add-in/tree/main/release`
     trên trình duyệt → `Add file → Upload files` → kéo thả `LINK.xlam` mới đè
     lên file cũ → commit. Sau đó mở `version.txt`, bấm biểu tượng bút chì, sửa
     thành `1.0.1`, commit **riêng, sau khi file `.xlam` đã lên**.
   - **Có Git:** `git add release/LINK.xlam && git commit` trước, rồi mới sửa
     `version.txt` và commit lần hai, `git push`.
4. Xong — không cần thao tác gì thêm trên từng máy người dùng. Nếu có quyền
   truy cập 1–2 máy pilot, nên kiểm tra vòng cập nhật chạy đúng trước khi yên
   tâm (mục 2 ở trên) trước khi coi bản phát hành là ổn định.

Thứ tự file `.xlam` trước, `version.txt` sau **luôn phải giữ** dù ai phát hành —
nếu đảo ngược, máy người dùng có thể thấy version mới trước khi file mới có
mặt và tải về file `.xlam` cũ.

## Xử lý sự cố

| Hiện tượng | Kiểm tra |
|---|---|
| Không thấy hộp thoại cập nhật | Máy có chặn `raw.githubusercontent.com` không? Thử mở URL trong `GH_BASE` bằng trình duyệt. |
| Có `LINK.update.xlam` nhưng không cập nhật | Chưa đóng hết Excel. Đóng toàn bộ cửa sổ Excel. |
| `link_addin_update.log` = `OK=0` | File add-in ở thư mục không có quyền ghi, hoặc đường dẫn mạng. Đặt add-in ở `%APPDATA%\Microsoft\AddIns\`. |
| Excel không tự mở lại | `start "" excel.exe` không tìm thấy Excel trong PATH/App Paths. Người dùng tự mở Excel; add-in vẫn đã được cập nhật. |
