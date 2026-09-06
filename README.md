# LINK add-in

Add-in ribbon Excel dùng chung của công ty. Cài một lần, sau đó **tự cập nhật** —
người dùng không phải làm gì khi có bản mới.

## Bạn cần đọc phần nào

| Bạn là | Lần đầu | Các lần sau |
|---|---|---|
| **Người dùng add-in** | [A1 — Cài đặt](#a1-lần-đầu--cài-add-in) | [A2 — Không phải làm gì](#a2-các-lần-sau--không-phải-làm-gì) |
| **Người phát hành bản mới** | [B1 — Chuẩn bị](#b1-lần-đầu--chuẩn-bị-một-lần) | [B2 — Kéo thả để phát hành](#b2-mỗi-lần-phát-hành) |
| **Chủ repo, dựng hệ thống** | [C — Thiết lập ban đầu](#c-thiết-lập-ban-đầu--chủ-repo-làm-một-lần-duy-nhất) | — |

> Đây là ba **vai trò**, không phải ba người khác nhau. Một người có thể kiêm cả
> ba. Tách ra để khi giao việc phát hành cho đồng nghiệp, họ chỉ cần đọc mục B.

---

# A. Người dùng add-in

## A1. Lần đầu — cài add-in

Bạn sẽ nhận được file **`install.bat`** (qua Zalo/email), hoặc tải tại:

```
https://raw.githubusercontent.com/namtao/add-in/main/install.bat
```

1. **Double-click `install.bat`**
2. Nếu Windows hiện *"Windows protected your PC"* → bấm **More info** → **Run anyway**
3. Nếu đang mở Excel, script sẽ nhắc đóng hết Excel lại
4. Mở Excel → thấy tab **LINK** trên ribbon là xong

Cài đúng **một lần**. Không cần tài khoản gì, không cần quyền admin.

> ⚠️ **Click đúp vào file `.xlam` KHÔNG cài được add-in** — nó chỉ mở tạm cho
> phiên Excel đang chạy, đóng Excel là mất. Phải chạy `install.bat`.

## A2. Các lần sau — không phải làm gì

**Bạn không phải làm gì cả.** Không phải tải lại, không phải chạy lại `install.bat`.

Cơ chế chạy ngầm: mỗi lần mở Excel, add-in tự kiểm tra có bản mới không. Có thì
tải về im lặng, rồi **thay file lúc bạn đóng hết Excel**. Lần mở Excel sau đã là
bản mới.

Không hộp thoại, không gián đoạn công việc. Không có mạng thì bỏ qua, lần sau thử lại.

---

# B. Người phát hành bản mới

## B1. Lần đầu — chuẩn bị (một lần)

Bạn cần 2 thứ, xin từ người quản trị repo:

1. **File `publish.bat`** — tải tại
   `https://raw.githubusercontent.com/namtao/add-in/main/publish.bat`
   (lưu vào chỗ dễ tìm, ví dụ Desktop)
2. **Token phát hành** — người quản trị gửi riêng cho bạn qua Zalo. Là một chuỗi
   dài bắt đầu bằng `github_pat_`

Chưa cần làm gì thêm. Lần đầu chạy `publish.bat` nó sẽ hỏi token — dán vào một
lần, script lưu mã hoá trên máy bạn và **không hỏi lại nữa**.

> **Không cần tài khoản GitHub. Không cần cài Git.** Chỉ cần đúng 1 file
> (`publish.bat`) và 1 chuỗi token.

## B2. Mỗi lần phát hành

**Bước 1.** Mở `LINK.xlam` bằng Excel, sửa nội dung / ribbon / VBA, **lưu**, đóng Excel.

Không có số version nào phải tăng — sửa nội dung là đủ.

**Bước 2.** **Kéo file `LINK.xlam` vừa lưu, thả vào icon `publish.bat`.**

Xong. Script báo *"Da phat hanh thanh cong"* là đã lên.

> Lần chạy đầu tiên script hỏi token — dán vào rồi Enter.

> `.xlam` là file nhị phân, không gộp thay đổi của 2 người được — **chỉ nên một
> người sửa tại một thời điểm**. Ai phát hành sau sẽ ghi đè lên bản của người
> trước (bản cũ vẫn còn trong lịch sử, lấy lại được).

**Cách thay thế** — nếu bạn có tài khoản GitHub và đã được add làm Collaborator:
vào https://github.com/namtao/add-in/tree/main/release → `Add file → Upload files`
→ kéo thả `LINK.xlam` mới đè lên file cũ → `Commit changes`.

## B3. Kiểm tra bản vừa phát hành

Trên một máy đã cài add-in:

1. Mở Excel, đợi ~5 giây → phải xuất hiện file `LINK.update.xlam` trong
   `%APPDATA%\Microsoft\Excel\XLSTART\`
2. Đóng **toàn bộ** cửa sổ Excel, chờ vài giây → file đó biến mất, và
   `%TEMP%\link_addin_update.log` có dòng `OK=1`
3. Mở lại Excel → thấy thay đổi bạn vừa phát hành

## B4. Quay về bản cũ

Mở bản `.xlam` tốt trước đó, kéo thả vào `publish.bat` như bình thường. Máy đang
dùng bản lỗi sẽ tự "cập nhật" ngược về bản cũ, qua đúng cơ chế đó.

> Nên giữ vài bản `.xlam` cũ (đặt tên kèm ngày) trong một thư mục riêng để
> rollback nhanh khi cần.

---

# C. Thiết lập ban đầu — chủ repo làm một lần duy nhất

## C1. Nhúng module tự cập nhật vào LINK.xlam

**Vì sao cần bước này:** code tự cập nhật (`modAutoUpdate`) nằm *bên trong* chính
file `LINK.xlam`. Một file không thể tự cập nhật trước khi nó chứa đoạn code biết
cách tự cập nhật — nên phải nhét code vào bằng tay đúng một lần đầu tiên. Sau khi
làm xong và phát hành, mọi bản sau đều tự có sẵn, **không bao giờ phải làm lại**.

Bước này bắt buộc làm thủ công vì VBA nằm trong `vbaProject.bin` — file nhị phân
biên dịch, chỉ ghi được bằng chính Excel trên Windows.

Cần một máy Windows có Excel.

**Kiểm tra đã làm chưa:** mở `release/LINK.xlam`, bấm `Alt+F11`, nhìn cột trái —
có module tên `modAutoUpdate` không? Có rồi thì bỏ qua mục này.

1. Tải `release/LINK.xlam` về máy, mở bằng Excel (Enable Content nếu hỏi)
2. `Alt+F11` mở VBE
3. `File → Import File…` → chọn `src/modAutoUpdate.bas`
4. Double-click **ThisWorkbook** ở cột trái → dán nội dung `src/ThisWorkbook.snippet.txt`
   *(nếu `ThisWorkbook` đã có sẵn `Sub Workbook_Open`, chỉ thêm dòng
   `modAutoUpdate.ScheduleUpdateCheck` vào trong sub đó)*
5. `Debug → Compile VBAProject` — phải **không báo lỗi**
6. `Ctrl+S` lưu, đóng Excel
7. Phát hành file này theo **B2**

## C2. Tạo token phát hành

1. `github.com` → ảnh đại diện góc phải → **Settings**
2. Cuối cột trái → **Developer settings**
3. **Personal access tokens → Fine-grained tokens → Generate new token**
4. **Token name** `link-addin-publish` · **Expiration** 1 năm · **Resource owner** `namtao`
5. **Repository access** → **Only select repositories** → `namtao/add-in`
6. **Permissions → Repository permissions → Contents: Read and write**
   *(mọi mục khác để No access; Metadata tự bật Read-only là bắt buộc, bình thường)*
7. **Generate token** → **copy ngay**, token chỉ hiện đúng một lần

Gửi token này **riêng tư** cho từng người được phép phát hành.

> ⚠️ Đừng đưa token vào file nào commit lên repo — GitHub tự thu hồi token nếu
> phát hiện nó trong mã nguồn công khai. Đặt lịch tạo token mới trước khi hết hạn.

## C3. Rollout cho toàn bộ người dùng

Gửi `install.bat` cho mọi người (Zalo/email hoặc link raw ở mục A1). Mỗi người
chạy một lần. Từ đó về sau họ tự nhận mọi bản mới.

---

# Xử lý sự cố

| Hiện tượng | Cách xử lý |
|---|---|
| Cài xong nhưng không thấy tab LINK | Đóng hết Excel rồi mở lại. Vẫn không có → chạy lại `install.bat`. |
| Ribbon LINK hiện 2 lần | Máy còn bản cài kiểu cũ. Gỡ ở `File → Options → Add-ins → Manage: Excel Add-ins → Go`, bỏ tick, xoá file cũ. |
| Máy không nhận bản mới, cũng không báo lỗi | Máy có chặn `raw.githubusercontent.com`? Thử mở link đó bằng trình duyệt. |
| Có `LINK.update.xlam` nhưng file không được thay | Chưa đóng hết Excel. Đóng **toàn bộ** cửa sổ Excel rồi chờ vài giây. |
| `link_addin_update.log` ghi `OK=0` | Thư mục add-in không có quyền ghi. Chạy lại `install.bat`. |
| Excel mở chậm hẳn ~1–3 giây | Máy thiếu .NET Framework nên add-in phải dùng cách tính checksum chậm hơn. |
| Không máy nào cập nhật được | `version.txt` lệch với `LINK.xlam`. Phát hành lại theo B2 là tự khớp. |
| `publish.bat` báo *"noi dung giong het ban dang phat hanh"* | File không đổi so với bản đang chạy. Sửa gì đó rồi lưu lại. |
| `publish.bat` báo token không hợp lệ / hết hạn | Token hết hạn hoặc bị thu hồi. Xin token mới từ chủ repo (C2). |
| `publish.bat` báo *"khong du quyen"* | Token thiếu quyền `Contents: Read and write`. Tạo lại token theo C2. |
| `publish.bat` báo có người push cùng lúc | Chạy lại `publish.bat` một lần nữa. |
