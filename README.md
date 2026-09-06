# LINK add-in

Add-in ribbon Excel dùng chung. Cài một lần, sau đó **tự cập nhật** — người dùng
không phải làm gì khi có bản mới.

Tài liệu này chỉ có 3 việc: chuẩn bị lần đầu, cài cho người dùng, phát hành bản mới.

---

## 0. Chuẩn bị lần đầu — chỉ làm MỘT lần cho cả dự án

Cần một máy Windows có Excel.

**Kiểm tra đã làm chưa:** mở `release/LINK.xlam`, bấm `Alt+F11`, nhìn cột trái —
có module tên `modAutoUpdate` không? Có rồi thì bỏ qua toàn bộ mục này.

1. Tải `release/LINK.xlam` về máy, mở bằng Excel (Enable Content nếu hỏi).
2. `Alt+F11` mở VBE.
3. `File → Import File…` → chọn `src/modAutoUpdate.bas`.
4. Double-click **ThisWorkbook** ở cột trái → dán nội dung `src/ThisWorkbook.snippet.txt`.
   (Nếu `ThisWorkbook` đã có sẵn `Sub Workbook_Open`, chỉ thêm dòng
   `modAutoUpdate.ScheduleUpdateCheck` vào trong sub đó.)
5. `Debug → Compile VBAProject` — phải **không báo lỗi**.
6. `Ctrl+S` lưu. Đóng Excel.
7. Phát hành file này theo **mục 2**.

Xong. Từ đây không bao giờ phải làm lại.

---

## 1. Cài cho người dùng — mỗi máy một lần

Gửi họ file `install.bat`, qua Zalo/email hoặc link:

```
https://raw.githubusercontent.com/namtao/add-in/main/install.bat
```

Người dùng **double-click `install.bat`**. Script tự tải add-in về và đặt vào
đúng chỗ Excel tự nạp mỗi lần khởi động.

Mở Excel, thấy tab **LINK** trên ribbon là xong. Từ đó họ **không phải làm gì
nữa** — mọi bản mới tự về.

> Nếu Windows hiện *"Windows protected your PC"* → bấm **More info → Run anyway**.
> Muốn tránh hẳn thì gửi thẳng file qua Zalo thay vì gửi link tải.

> Click đúp vào file `.xlam` **không cài được add-in** — chỉ mở tạm cho phiên đang
> chạy, đóng Excel là mất. Phải dùng `install.bat`.

---

## 2. Phát hành bản mới

### Chuẩn bị — chủ repo làm MỘT lần

Tạo token phát hành:

1. `github.com` → ảnh đại diện góc phải → **Settings** → cuối cột trái → **Developer settings**
2. **Personal access tokens → Fine-grained tokens → Generate new token**
3. **Token name** `link-addin-publish` · **Expiration** 1 năm · **Resource owner** `namtao`
4. **Repository access**: chọn **Only select repositories** → `namtao/add-in`
5. **Permissions → Repository permissions → Contents: Read and write**
   *(mọi mục khác để No access; Metadata tự bật Read-only là bình thường)*
6. **Generate token** → copy ngay, token chỉ hiện **một lần**

Gửi token này **riêng tư** (Zalo/tin nhắn) cho người được phép phát hành.
Đừng đưa token vào file nào commit lên repo — GitHub tự thu hồi token nếu phát
hiện nó trong mã nguồn công khai. Đặt lịch tạo token mới mỗi năm.

### Mỗi lần phát hành

**Bước 1.** Mở `LINK.xlam` trong Excel, sửa nội dung/ribbon/VBA, lưu, đóng Excel.
Không có số version nào phải tăng — sửa nội dung là đủ.

**Bước 2.** **Kéo thả file `LINK.xlam` vào icon `publish.bat`.** Xong.

Lần chạy đầu tiên script hỏi token — dán vào một lần, nó lưu mã hoá trên máy đó
và không hỏi lại nữa. **Không cần tài khoản GitHub, không cần cài Git.**

> **Cách thay thế** — nếu người phát hành có tài khoản GitHub và đã được add làm
> Collaborator: vào https://github.com/namtao/add-in/tree/main/release →
> `Add file → Upload files` → kéo thả `LINK.xlam` mới đè lên file cũ →
> `Commit changes`.

> `.xlam` là file nhị phân, không merge được — chỉ nên **một người sửa tại một
> thời điểm**.

### Kiểm tra sau khi phát hành

Trên một máy đã cài add-in:

1. Mở Excel, đợi ~5 giây → phải có file `LINK.update.xlam` trong
   `%APPDATA%\Microsoft\Excel\XLSTART\`
2. Đóng **toàn bộ** Excel, chờ vài giây → file đó biến mất, và
   `%TEMP%\link_addin_update.log` có dòng `OK=1`
3. Mở lại Excel → thấy thay đổi vừa phát hành

---

## 3. Người dùng nhận bản mới thế nào

**Không thao tác gì.**

Mở Excel → add-in tự kiểm tra ngầm → có bản mới thì tải về im lặng → **thay file
khi họ đóng hết Excel** → lần mở Excel sau đã là bản mới.

Không hộp thoại, không gián đoạn công việc. Không có mạng thì bỏ qua, lần sau
thử lại.

---

## 4. Quay về bản cũ

Lưu đè `LINK.xlam` bằng bản tốt trước đó rồi phát hành lại theo mục 2. Máy đang
dùng bản lỗi sẽ tự "cập nhật" ngược về bản cũ, qua đúng cơ chế đó.

Nên giữ vài bản `.xlam` cũ (đặt tên kèm ngày) ở một thư mục riêng để rollback nhanh.

---

## Xử lý sự cố

| Hiện tượng | Cách xử lý |
|---|---|
| Không máy nào cập nhật, cũng không báo lỗi gì | `version.txt` lệch với `LINK.xlam`. Phát hành lại theo mục 2 là tự khớp. |
| Mở Excel mãi không thấy `LINK.update.xlam` | Máy có chặn `raw.githubusercontent.com`? Thử mở link đó bằng trình duyệt. |
| Có `LINK.update.xlam` nhưng file không được thay | Chưa đóng hết Excel. Đóng **toàn bộ** cửa sổ Excel rồi chờ vài giây. |
| `link_addin_update.log` ghi `OK=0` | Thư mục add-in không có quyền ghi. Chạy lại `install.bat`. |
| Excel mở chậm hẳn ~1–3 giây | Máy thiếu .NET Framework nên phải dùng cách tính checksum chậm hơn. |
| Ribbon LINK hiện 2 lần | Máy còn bản cài kiểu cũ. Gỡ ở `File → Options → Add-ins → Manage: Excel Add-ins → Go`, bỏ tick, xoá file cũ. |
| `publish.bat` báo "noi dung giong het ban dang phat hanh" | File không đổi so với bản đang chạy. Sửa gì đó rồi lưu lại. |
| `publish.bat` báo token không hợp lệ / hết hạn | Token đã hết hạn hoặc bị thu hồi. Chủ repo tạo token mới (mục 2) và gửi lại. |
| `publish.bat` báo "khong du quyen" | Token thiếu quyền `Contents: Read and write`. Tạo lại token đúng quyền. |
