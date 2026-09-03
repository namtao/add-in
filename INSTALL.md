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
6. `Debug → Compile VBAProject` — phải **không lỗi**.
7. `Ctrl+S` lưu `LINK.xlam`. Đóng Excel.

> Không có hằng số version nào cần chỉnh ở bước này — `modAutoUpdate` so sánh
> bằng checksum nội dung file, không phải số version thủ công. Xem
> `README.md` mục "Cách hoạt động".

## 2. Test nhanh cơ chế (khuyến nghị)

Trên một máy test:

1. Cài `LINK.xlam` vừa sửa làm add-in (`File → Options → Add-ins → Manage: Excel
   Add-ins → Browse`).
2. Trên repo (hoặc nhánh phụ), tạm sửa `release/version.txt` thành một chuỗi bất
   kỳ khác với checksum thật (ví dụ `test123`), commit + push — mục đích chỉ để
   giả lập "server báo có bản khác", không cần đúng checksum thật cho bước test
   cơ chế tải + swap.
3. Mở Excel → sau ~2 giây hiện hộp thoại "Đã có bản cập nhật mới…".
4. Kiểm tra có file `LINK.update.xlam` cạnh `LINK.xlam`.
5. Đóng **toàn bộ** Excel → chờ vài giây → Excel tự mở lại.
6. Log `%TEMP%\link_addin_update.log` phải có `OK=1`.
7. Trả `release/version.txt` về đúng checksum thật của `release/LINK.xlam`
   (`publish.bat` sẽ tự làm việc này ở lần phát hành thật — xem mục 4 và
   `README.md`), commit + push.

> Để test đầy đủ (nội dung file thật sự đổi), sửa gì đó trong `LINK.xlam`
> (ví dụ đổi 1 label ribbon), lưu, rồi chạy `publish.bat` lên nhánh test —
> checksum sẽ tự khác, không cần sửa `version.txt` tay.

## 3. Phát hành bản gốc

1. Commit `LINK.xlam` (đã có `modAutoUpdate`) vào `release/LINK.xlam`.
2. Tính checksum thật và ghi vào `release/version.txt` — cách nhanh nhất là
   chạy `publish.bat` một lần (xem mục 4 của `README.md`), nó tự tính và
   commit cả hai file đúng cặp với nhau.
3. `git push` (nếu không dùng `publish.bat`).

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

1. **Tài khoản GitHub:** nếu họ chưa có, tạo giúp tại `github.com/signup`
   (email + mật khẩu, ~1 phút, miễn phí, không cần thẻ).
2. **Quyền ghi vào repo:** bạn (chủ repo) vào
   `github.com/namtao/add-in → Settings → Collaborators → Add people`, nhập
   username/email GitHub của họ, họ bấm chấp nhận lời mời qua email.
3. **Sửa add-in:** họ mở `LINK.xlam` trong Excel (đã có sẵn `modAutoUpdate` —
   không cần làm lại mục 1), sửa nội dung/ribbon/VBA cần thiết, lưu. Không có
   số version nào phải nhớ bump — sửa nội dung là đủ.
4. **Đẩy lên GitHub bằng [publish.bat](./publish.bat)** — không cần biết lệnh
   git, không cần gõ gì:
   - Chỉ cần cài **Git for Windows** một lần
     ([git-scm.com/download/win](https://git-scm.com/download/win), bấm Next
     liên tục là xong).
   - Chạy `publish.bat` lần đầu để nó tự tạo thư mục
     `%USERPROFILE%\LINK-addin-publish\` — từ đó về sau, họ luôn mở và **lưu
     đè** `LINK.xlam` đúng vào `...\LINK-addin-publish\release\LINK.xlam`.
   - Mỗi lần có bản mới: double-click `publish.bat`. Script tự tính checksum,
     commit, push — **không hỏi gì, không cần gõ gì**. Lần push đầu tiên sẽ
     bật cửa sổ đăng nhập GitHub trong trình duyệt, đăng nhập xong là dùng
     được mãi.
   - (Không cài Git cũng được: vào
     `github.com/namtao/add-in/tree/main/release` trên trình duyệt →
     `Add file → Upload files` → kéo thả `LINK.xlam` mới đè lên file cũ →
     commit; nhưng khi đó phải tự tính lại checksum SHA256 của file và dán
     vào `version.txt` — `publish.bat` làm việc này tự động nên là cách được
     khuyến nghị.)
5. Xong — không cần thao tác gì thêm trên từng máy người dùng. Nếu có quyền
   truy cập 1–2 máy pilot, nên kiểm tra vòng cập nhật chạy đúng trước khi yên
   tâm (mục 2 ở trên) trước khi coi bản phát hành là ổn định.

## Xử lý sự cố

| Hiện tượng | Kiểm tra |
|---|---|
| Không thấy hộp thoại cập nhật | Máy có chặn `raw.githubusercontent.com` không? Thử mở URL trong `GH_BASE` bằng trình duyệt. |
| Có `LINK.update.xlam` nhưng không cập nhật | Chưa đóng hết Excel. Đóng toàn bộ cửa sổ Excel. |
| `link_addin_update.log` = `OK=0` | File add-in ở thư mục không có quyền ghi, hoặc đường dẫn mạng. Đặt add-in ở `%APPDATA%\Microsoft\Excel\XLSTART\` (xem mục 4). |
| Excel không tự mở lại | `start "" excel.exe` không tìm thấy Excel trong PATH/App Paths. Người dùng tự mở Excel; add-in vẫn đã được cập nhật. |
| `publish.bat` báo "khong co gi thay doi" dù đã sửa file | So sánh bằng checksum nội dung — nếu lưu Excel mà nội dung nhị phân giống hệt (hiếm) sẽ không đổi checksum. Sửa thêm gì đó rồi lưu lại. |
