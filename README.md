# LINK add-in

Add-in ribbon Excel dùng chung của công ty. Cài một lần, sau đó **tự cập nhật** —
khi có bản mới, add-in hỏi một câu lúc mở Excel, bấm OK là xong.

## Bạn cần đọc phần nào

| Bạn là | Lần đầu | Các lần sau |
|---|---|---|
| **Người dùng add-in** | [A1 — Cài đặt](#a1-lần-đầu--cài-add-in) | [A2 — Bấm OK khi được hỏi](#a2-các-lần-sau--chỉ-bấm-ok-khi-được-hỏi) |
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

## A2. Các lần sau — chỉ bấm OK khi được hỏi

Không phải tải lại, không phải chạy lại `install.bat`.

Mỗi lần mở Excel, add-in âm thầm kiểm tra có bản mới không. **Chỉ khi thực sự có
bản mới đã tải xong và kiểm tra đạt**, nó mới hiện một hộp thoại:

> Đã có bản cập nhật mới cho add-in LINK.
>
> Bấm **OK**: Excel sẽ lưu các file đang mở, đóng lại để cập nhật, rồi tự mở lại.
> Bấm **Cancel**: bỏ qua lần này, lần mở Excel sau sẽ hỏi lại.

**Bấm OK** — Excel lưu mọi file bạn từng lưu trước đó, tự đóng, thay add-in, rồi
tự mở lại. Mất khoảng 5–10 giây và bạn không phải thao tác gì thêm.

**Bấm Cancel** nếu đang dở việc — add-in cũ chạy tiếp bình thường, lần mở Excel
sau sẽ hỏi lại.

> File **chưa từng được lưu lần nào** (Book1 mới gõ, chưa đặt tên) thì Excel vẫn
> hỏi bạn chọn chỗ lưu như mọi khi — add-in không tự quyết định thay bạn và cũng
> không vứt dữ liệu đi. Nếu bạn bấm Cancel ở hộp thoại lưu đó thì Excel không
> đóng, lần cập nhật này bỏ dở và lần mở Excel sau sẽ hỏi lại.

Không có mạng thì bỏ qua trong im lặng, lần sau thử lại. Không bao giờ có hộp
thoại báo lỗi.

## A3. Gỡ add-in

Tải và chạy **`uninstall.bat`**:

```
https://raw.githubusercontent.com/namtao/add-in/main/uninstall.bat
```

1. **Double-click `uninstall.bat`**
2. Đóng hết Excel khi script nhắc
3. Mở Excel → tab **LINK** đã biến mất

Script xoá file add-in và dọn các file tạm. Không đụng registry, không cần quyền
admin. Gỡ rồi cài lại lúc nào cũng được, không mất gì.

> Nếu máy từng cài theo kiểu cũ (`File → Options → Add-ins → Browse`),
> `uninstall.bat` sẽ **báo cho bạn biết** chứ không tự xoá — vì bản đó có đăng ký
> trong registry, xoá thẳng file sẽ làm Excel báo lỗi thiếu file mỗi lần mở.
> Script chỉ dẫn cách gỡ đúng.

## A4. Cài lại add-in

Chạy lại **`install.bat`** (mục A1). Nó luôn tải bản mới nhất về, nên cài lại
cũng là cách **sửa lỗi nhanh nhất** trong 3 tình huống:

| Tình huống | Làm gì |
|---|---|
| Đã gỡ, giờ muốn dùng lại | Chạy `install.bat` |
| Máy mãi không nhận bản mới | Chạy `install.bat` — ép lấy bản mới nhất ngay |
| Add-in lỗi, nghi file hỏng | Chạy `install.bat` — ghi đè bằng bản sạch từ repo |

Không cần gỡ trước. `install.bat` tự ghi đè lên bản cũ.

> Không mất gì khi cài lại: mọi cấu hình của add-in nằm trong chính file `.xlam`
> tải từ repo, không có dữ liệu riêng nào lưu trên máy.

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

**Bước 1.** Mở file của bạn bằng Excel, sửa nội dung / ribbon / VBA, **lưu**, đóng Excel.

Dùng file `.xlsm` hay `.xlam` đều được — cứ làm việc trên file bạn quen dùng.
Không có số version nào phải tăng, sửa nội dung là đủ.

**Bước 2.** **Kéo file vừa lưu, thả vào icon `publish.bat`.**

Tên file đặt gì cũng được. Script luôn phát hành vào đúng `release/LINK.xlam`.

**Bước 3.** Script hỏi đúng một câu:

```
Phat hanh? (y/n):
```

Gõ `y` rồi Enter để phát hành. Bấm phím khác hoặc Enter trống là huỷ, không có gì
được đẩy lên.

Xong. Script báo *"Da phat hanh thanh cong"* là đã lên.

> 🧪 **Muốn chạy thử mà không phát hành?** Cứ kéo thả như bình thường rồi trả lời
> `n`. File vẫn được kiểm tra và vá đầy đủ nhưng không có gì rời khỏi máy bạn.

> ⏱️ Máy người dùng có thể mất **tới 5 phút** mới thấy bản mới (GitHub cache file
> khoảng 5 phút). Trong lúc đó họ vẫn dùng bản cũ bình thường — không lỗi gì.

> Lần chạy đầu tiên script hỏi token — dán vào rồi Enter.

> `.xlam` là file nhị phân, không gộp thay đổi của 2 người được — **chỉ nên một
> người sửa tại một thời điểm**. Ai phát hành sau sẽ ghi đè lên bản của người
> trước (bản cũ vẫn còn trong lịch sử, lấy lại được).

> 🛡️ **Script tự lo hết phần kỹ thuật, trong im lặng.** Nếu file thiếu module tự
> cập nhật, thiếu đoạn gọi trong `ThisWorkbook`, quên bật `IsAddin`, hoặc còn là
> `.xlsm` chưa chuyển sang add-in, nó tự xử lý rồi mới phát hành — không hỏi, không
> báo cáo. Bạn không phải nhớ thao tác Excel nào, kể cả bước `Save As` sang `.xlam`.
>
> Script làm việc trên một **bản sao** trong thư mục tạm, file gốc của bạn không bị
> đụng tới. Sau khi vá nó kiểm tra lại lần nữa; chưa đạt thì từ chối phát hành chứ
> không đẩy bừa lên. Và [cổng chặn trên GitHub](#c1-nhúng-module-tự-cập-nhật-vào-linkxlam)
> vẫn soi lại lần nữa sau đó.
>
> Thói quen an toàn vẫn nên giữ: **sửa trên bản mới nhất lấy từ repo**, đừng dùng
> file cũ nằm sẵn trong Downloads.

**Cách thay thế** — nếu bạn có tài khoản GitHub và đã được add làm Collaborator:
vào https://github.com/namtao/add-in/tree/main/release → `Add file → Upload files`
→ kéo thả `LINK.xlam` mới đè lên file cũ → `Commit changes`.

## B3. Kiểm tra bản vừa phát hành

Trên một máy đã cài add-in:

1. Mở Excel, đợi ~5 giây → hiện hộp thoại *"Đã có bản cập nhật mới cho add-in LINK"*
2. Bấm **OK** → Excel tự đóng, vài giây sau tự mở lại
3. Kiểm tra `%TEMP%\link_addin_update.log` có dòng `OK=1`, và ribbon LINK đã có
   thay đổi bạn vừa phát hành

> Muốn kiểm tra lại mà không phải khởi động lại Excel: `Alt+F8`, gõ
> `LINK_CheckUpdateNow` rồi bấm Run. Nó chạy đúng luồng kiểm tra đó ngay lập tức.

## B4. Quay về bản cũ

Mở bản `.xlam` tốt trước đó, kéo thả vào `publish.bat` như bình thường. Máy đang
dùng bản lỗi sẽ tự "cập nhật" ngược về bản cũ, qua đúng cơ chế đó.

> Nên giữ vài bản `.xlam` cũ (đặt tên kèm ngày) trong một thư mục riêng để
> rollback nhanh khi cần.

---

# C. Thiết lập ban đầu — chủ repo làm một lần duy nhất

## C1. Nhúng module tự cập nhật vào LINK.xlam

**Vì sao cần:** code tự cập nhật (`modAutoUpdate`) nằm *bên trong* chính file
`LINK.xlam`, và nó chỉ chạy khi `ThisWorkbook` gọi nó lúc Excel mở file. Thiếu
một trong hai thứ đó thì add-in mất khả năng tự cập nhật — và hỏng âm thầm, nhìn
bề ngoài vẫn chạy bình thường.

**Hiện `publish.bat` tự lo việc này**, nên bạn không phải làm tay nữa. Mỗi lần
phát hành, nó kiểm tra file và tự thêm những gì còn thiếu. Mục này giữ lại để
bạn biết chuyện gì đang diễn ra bên dưới, và để làm thủ công khi cần.

**Kiểm tra một file bất kỳ đã đủ chưa**, trên máy Linux/macOS hoặc trong CI:

```bash
python3 scripts/validate-addin.py duong-dan-toi-file.xlam
```

Script đọc thẳng mã nguồn VBA bên trong file (giải nén `vbaProject.bin` theo đúng
định dạng MS-OVBA) nên kết luận là chắc chắn, không phải đoán. Cũng chính script
này chạy trên GitHub Actions mỗi lần `release/LINK.xlam` thay đổi.

**Làm thủ công**, nếu muốn tự tay nhúng thay vì để `publish.bat` làm:

1. Tải `release/LINK.xlam` về máy, mở bằng Excel (Enable Content nếu hỏi)
2. `Alt+F11` mở VBE
3. `File → Import File…` → chọn `src/modAutoUpdate.bas`
4. Double-click **ThisWorkbook** ở cột trái → dán nội dung `src/ThisWorkbook.snippet.txt`
   *(nếu `ThisWorkbook` đã có sẵn `Sub Workbook_Open`, chỉ thêm dòng
   `modAutoUpdate.ScheduleUpdateCheck` vào trong sub đó)*
5. `Debug → Compile VBAProject` — phải **không báo lỗi**
6. `Ctrl+S` lưu, đóng Excel
7. Phát hành file này theo **B2**

> ⚠️ **Sửa trên file gốc, đừng chỉ sửa bản `.xlam` xuất ra.** Nếu bạn làm việc
> trên một file `.xlsm` rồi mới xuất sang `.xlam`, hãy dán đoạn `Workbook_Open`
> vào **chính file `.xlsm` gốc** đó. Chỉ sửa bản xuất ra thì lần sau xuất lại,
> thiếu sót quay về y nguyên.

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
| Máy không hỏi cập nhật, cũng không báo lỗi | `Alt+F8` → chạy `LINK_CheckUpdateNow` để kiểm tra ngay. Vẫn im thì chạy lại `install.bat` để ép lấy bản mới nhất. Vẫn không được → máy có chặn `raw.githubusercontent.com`? Thử mở link đó bằng trình duyệt. |
| Bấm OK nhưng Excel không đóng | Còn hộp thoại Excel hỏi chỗ lưu cho file chưa từng lưu. Trả lời hộp thoại đó là Excel đóng và cập nhật chạy tiếp. |
| Excel đã đóng nhưng file không được thay | Còn cửa sổ Excel khác đang mở. Đóng **toàn bộ** rồi chờ vài giây. Quá ~10 phút thì script bỏ cuộc, ghi `TIMEOUT` vào log và add-in sẽ hỏi lại ở lần mở Excel sau. |
| `link_addin_update.log` ghi `OK=0` | Thư mục add-in không có quyền ghi. Chạy lại `install.bat`. |
| Excel mở chậm hẳn ~1–3 giây | Máy thiếu .NET Framework nên add-in phải dùng cách tính checksum chậm hơn. |
| Không máy nào cập nhật được | `version.txt` lệch với `LINK.xlam`. Phát hành lại theo B2 là tự khớp. |
| `publish.bat` báo *"noi dung giong het ban dang phat hanh"* | File không đổi so với bản đang chạy. Sửa gì đó rồi lưu lại. |
| `publish.bat` báo token không hợp lệ / hết hạn | Token hết hạn hoặc bị thu hồi. Xin token mới từ chủ repo (C2). |
| `publish.bat` báo đuôi file không nhận được | Chỉ nhận `.xlam` và `.xlsm`. File `.xlsx` không chứa macro nên không phải add-in LINK. |
| `publish.bat` báo *"Excel dang chan script doc phan code VBA"* | Đóng **hết** Excel rồi chạy lại — Excel chỉ đọc thiết lập bảo mật lúc khởi động. Vẫn lỗi thì máy có thể bị chính sách công ty khoá, script in sẵn các bước kiểm tra. |
| GitHub Action báo đỏ, `version.txt` không đổi | File vừa đẩy lên không qua được kiểm tra. Máy người dùng **vẫn an toàn** — không máy nào đổi sang bản hỏng. Xem log Action để biết thiếu gì, sửa file gốc rồi phát hành lại. |
| `publish.bat` báo *"khong du quyen"* | Token thiếu quyền `Contents: Read and write`. Tạo lại token theo C2. |
| `publish.bat` báo có người push cùng lúc | Chạy lại `publish.bat` một lần nữa. |
