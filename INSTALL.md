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

## 2. Test cơ chế (khuyến nghị, trên máy test)

Add-in chỉ chấp nhận file tải về khi checksum của nó **đúng bằng** giá trị trong
`version.txt`. Nên không thể test bằng cách đặt `version.txt` thành chuỗi giả —
phải phát hành một bản thật sự khác. Cách test đúng:

1. Cài add-in lên máy test bằng `install.bat` (xem mục 4).
2. Trên máy phát hành, mở `LINK.xlam`, đổi một thứ dễ nhận ra (ví dụ sửa nhãn
   một nút trong sheet `Setting`), lưu, đóng Excel.
3. Kéo thả file đó vào `publish.bat`.
4. Máy test: mở Excel, đợi ~5 giây → kiểm tra có file `LINK.update.xlam` trong
   `%APPDATA%\Microsoft\Excel\XLSTART\`.
5. Đóng **toàn bộ** Excel, chờ ~5 giây → `LINK.update.xlam` biến mất và
   `%TEMP%\link_addin_update.log` có dòng `OK=1`.
6. Mở lại Excel → thấy nhãn mới. Xong.

> Cập nhật chạy **im lặng**: không có hộp thoại nào, Excel không tự mở lại. Nếu
> muốn thấy thông báo trong lúc test, tạm đặt `NOTIFY_USER = True` trong
> `modAutoUpdate.bas`.

Kiểm tra nhanh repo có nhất quán không (chạy ở thư mục clone):

```
powershell -NoProfile -Command "(Get-FileHash -Algorithm SHA256 release\LINK.xlam).Hash; Get-Content release\version.txt"
```

Hai dòng in ra phải giống hệt nhau. Nếu lệch, không máy nào cập nhật được (đúng
theo thiết kế — thà không cập nhật còn hơn cập nhật sai).

## 3. Phát hành bản gốc

Kéo thả `LINK.xlam` (bản đã có `modAutoUpdate` ở mục 1) vào `publish.bat`.
Script tự tính checksum, ghi `version.txt`, commit và push cả hai file.

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

> **SmartScreen:** file `.bat` tải từ trình duyệt sẽ bị Windows chặn một lần với
> thông báo *"Windows protected your PC"* → bấm **More info → Run anyway**. Muốn
> tránh hẳn thì gửi file qua Zalo/ổ mạng nội bộ thay vì link tải, hoặc kèm ảnh
> chụp hai cú bấm đó vào hướng dẫn gửi người dùng.

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

1. **Tài khoản GitHub:** nếu họ chưa có, tạo tại `github.com/signup`
   (email + mật khẩu, ~1 phút, miễn phí, không cần thẻ).
2. **Quyền ghi vào repo:** bạn (chủ repo) vào
   `github.com/namtao/add-in → Settings → Collaborators → Add people`, nhập
   username/email GitHub của họ, họ bấm chấp nhận lời mời qua email.
3. **Cài Git for Windows một lần**
   ([git-scm.com/download/win](https://git-scm.com/download/win), bấm Next
   liên tục là xong).
4. **Chạy `publish.bat` lần đầu** — nó tự tải repo về
   `%USERPROFILE%\LINK-addin-publish\` rồi dừng lại.
5. **Mỗi lần có bản mới:** mở `LINK.xlam` (đã có sẵn `modAutoUpdate` — không cần
   làm lại mục 1), sửa, lưu, rồi **kéo thả file đó vào icon `publish.bat`**.
   Không hỏi gì, không cần gõ gì, không có số version nào phải nhớ. Lần push đầu
   tiên sẽ bật cửa sổ đăng nhập GitHub trong trình duyệt, đăng nhập xong là dùng
   được mãi.
6. Nếu trong lúc họ sửa mà người khác vừa phát hành, `publish.bat` in cảnh báo và
   đợi 10 giây để họ kịp bấm Ctrl+C (xem `README.md` mục "Nhiều người cùng phát
   hành").

> **Không nên** upload `LINK.xlam` thẳng qua giao diện web GitHub: khi đó
> `version.txt` không được tính lại, checksum sẽ lệch và **không máy nào cập
> nhật được** cho tới khi có người chạy `publish.bat`. Luôn dùng `publish.bat`.

## Xử lý sự cố

| Hiện tượng | Kiểm tra |
|---|---|
| Không máy nào cập nhật, cũng không lỗi gì | Checksum trong `version.txt` lệch với `release/LINK.xlam`. Chạy lệnh kiểm tra ở mục 2; sửa bằng cách chạy `publish.bat` lại. |
| Không có `LINK.update.xlam` sau khi mở Excel | Máy có chặn `raw.githubusercontent.com` không? Thử mở URL trong `GH_BASE` bằng trình duyệt. |
| Có `LINK.update.xlam` nhưng file không được thay | Chưa đóng hết Excel. Đóng toàn bộ cửa sổ Excel rồi chờ vài giây. |
| `link_addin_update.log` = `OK=0` | File add-in ở thư mục không có quyền ghi, hoặc đường dẫn mạng. Đặt add-in ở `%APPDATA%\Microsoft\Excel\XLSTART\` (xem mục 4). |
| Excel mở chậm hẳn đi ~1–3 giây | .NET COM không dùng được nên add-in phải lùi về PowerShell để tính checksum. Kiểm tra máy có .NET Framework không. |
| `publish.bat` báo "khong co gi thay doi" dù đã sửa file | So sánh bằng checksum nội dung — nếu lưu Excel mà nội dung nhị phân giống hệt (hiếm) sẽ không đổi checksum. Sửa thêm gì đó rồi lưu lại. |
| `publish.bat` báo lỗi checksum | PowerShell không chạy được trên máy đó. Script cố tình từ chối publish thay vì ghi `version.txt` rác. |
