# LINK add-in — tự động cập nhật

Add-in ribbon Excel `LINK.xlam` tự kiểm tra phiên bản mới trên repo này mỗi lần
mở Excel và tự cập nhật, không cần gửi file thủ công.

## Cách hoạt động

1. `ThisWorkbook.Workbook_Open` gọi `modAutoUpdate.ScheduleUpdateCheck`.
2. Sau ~2 giây (qua `Application.OnTime`, không chặn Excel khởi động),
   `CheckForUpdate` tính checksum SHA256 của chính file add-in đang chạy, tải
   `release/version.txt` từ GitHub raw (checksum của bản đang phát hành) và so
   hai giá trị này.
3. Nếu khác nhau → tải `release/LINK.xlam` về `LINK.update.xlam` cạnh file
   add-in, rồi kiểm tra 3 lớp: kích thước tối thiểu, chữ ký ZIP (`PK`), và
   **checksum của file vừa tải phải đúng bằng checksum server báo**. Sai bất kỳ
   lớp nào → xoá file tải, bỏ qua lần này.
4. Ghi một script `.bat` ra `%TEMP%`, chạy ẩn. Script:
   - đợi **toàn bộ** `EXCEL.EXE` thoát;
   - `copy /y` bản mới đè lên file add-in thật (thử lại tối đa 5 lần);
   - ghi log `%TEMP%\link_addin_update.log`;
   - tự xoá.
5. **Im lặng hoàn toàn** — không hộp thoại, không tự mở lại Excel. Người dùng tắt
   Excel lúc nào thì file được thay lúc đó; lần mở Excel kế tiếp đã là bản mới.

Mọi lỗi mạng / HTTPS / file đều bị nuốt lặng — add-in luôn chạy tiếp ở phiên bản
hiện tại, Excel không bị chậm hay hiện lỗi khi offline.

**Không có số version thủ công:** `release/version.txt` là checksum SHA256 của
`release/LINK.xlam`, `publish.bat` tự tính — người phát hành chỉ cần sửa nội
dung add-in và lưu, không cần nhớ bump version, không thể quên.

**Vì sao phải kiểm checksum file vừa tải (bước 3):** nếu `version.txt` trên
server lệch với `LINK.xlam` thật (push thiếu, sửa tay nhầm), mà cứ swap bừa thì
sau khi swap xong checksum local vẫn khác remote → lần mở Excel sau lại tải, lại
swap… lặp mãi mãi trên **mọi máy**. Chỉ swap khi file tải về đúng bằng checksum
server báo thì lỗi đó không thể xảy ra. `publish.bat` cũng chặn ở đầu nguồn bằng
cách từ chối ghi `version.txt` nếu checksum không phải đúng 64 ký tự hex.

## Bố cục repo

```
install.bat         # gửi cho người dùng cuối — double-click là cài xong
publish.bat         # người phát hành: kéo thả LINK.xlam vào là xong
release/
  LINK.xlam         # bản đang phát hành (install.bat tải file này)
  version.txt       # checksum SHA256 của LINK.xlam ở trên — publish.bat tự tính
src/
  modAutoUpdate.bas             # module cần import vào LINK.xlam (làm 1 lần, xem INSTALL.md)
  ThisWorkbook.snippet.txt      # đoạn Workbook_Open cần dán
  link_addin_update.reference.bat  # bản tham khảo của .bat sinh lúc chạy
.github/workflows/
  update-version-checksum.yml   # Action tự tính version.txt khi LINK.xlam đổi
INSTALL.md       # thiết lập lần đầu (import module) + cách rollout bằng install.bat
```

## Quy trình phát hành bản mới

> Ai phát hành cũng được (không chỉ chủ repo) miễn có quyền ghi (collaborator)
> — xem "Người khác muốn phát hành bản mới" trong `INSTALL.md`. Người dùng cuối
> **không cần làm gì** khi có bản mới, chỉ cần đã cài add-in đúng cách một lần
> (xem `INSTALL.md` mục 4).

Bước 1 chung cho cả hai cách: mở `LINK.xlam` trong Excel, sửa nội dung/ribbon/VBA
cần thiết, lưu, đóng Excel. Không có hằng số version nào phải sửa. Sửa ở thư mục
nào cũng được.

### Cách A — trên trình duyệt (không cần cài gì)

2. Vào `github.com/namtao/add-in/tree/main/release` → `Add file → Upload files`
   → kéo thả `LINK.xlam` mới đè lên file cũ → `Commit changes`.
3. Xong. Action **Cap nhat version.txt** tự chạy, tính SHA256 và commit
   `version.txt` đúng sau ~20–30 giây.

Trong ~30 giây đó `version.txt` còn là checksum cũ, nên máy người dùng tải file
mới về sẽ thấy không khớp và **bỏ qua** — an toàn, và tự nhận ở lần mở Excel sau.

### Cách B — kéo thả trên máy (cần Git for Windows)

2. **Kéo thả file `LINK.xlam` vừa sửa vào icon [publish.bat](./publish.bat)** —
   không hỏi gì, không cần gõ gì. Script tự: lấy bản mới nhất từ GitHub → đặt
   file của bạn lên trên → tính checksum SHA256 → ghi vào `version.txt` → commit
   cả hai file cùng lúc → push.
   (Hoặc lưu đè vào `%USERPROFILE%\LINK-addin-publish\release\LINK.xlam` rồi
   double-click `publish.bat` — cùng kết quả.)
3. `publish.bat` commit `LINK.xlam` và `version.txt` cùng một lần push nên không
   có cả khoảng hở 30 giây như Cách A. Action chạy sau đó thấy checksum đã khớp
   và thoát, không commit gì thêm.

### Kiểm tra sau khi phát hành (cả hai cách)

4. Pilot 1–2 máy: mở Excel, đợi ~5 giây, kiểm tra có file `LINK.update.xlam` cạnh
   `%APPDATA%\Microsoft\Excel\XLSTART\LINK.xlam`. Đóng hết Excel → mở lại → file
   staging đã biến mất và `%TEMP%\link_addin_update.log` có dòng `OK=1`.
5. Sạch → các máy còn lại tự nhận ở lần mở Excel kế tiếp.

## Nhiều người cùng phát hành

`publish.bat` luôn đồng bộ với GitHub trước khi commit, nên không bao giờ kẹt ở
lỗi git. Nếu phát hiện người khác vừa phát hành trong lúc bạn đang sửa, script
**in cảnh báo và đợi 10 giây** để bạn kịp bấm Ctrl+C — vì bản của bạn dựa trên
bản cũ hơn nên sẽ ghi đè thay đổi của họ. Không dừng thì script tiếp tục theo
nguyên tắc "người phát hành sau thắng"; bản bị ghi đè vẫn còn nguyên trong lịch
sử git, lấy lại được bất cứ lúc nào.

`.xlam` là file nhị phân nên git không merge được — chỉ nên một người sửa tại
một thời điểm.

## Rollback

Lưu đè `LINK.xlam` bằng bản tốt trước đó rồi chạy `publish.bat` lại — checksum
mới sẽ tự khớp với bản cũ đó, máy đang ở bản lỗi "cập nhật" ngược về bản cũ qua
đúng cơ chế — không có nhánh code riêng. Nên giữ 1–2 bản `.xlam` cũ (đặt tên
kèm ngày) ở một thư mục riêng để có gì rollback ngay.

## Cấu hình trong `modAutoUpdate.bas`

| Hằng số | Ý nghĩa |
|---|---|
| `GH_BASE` | URL thư mục `release/` trên GitHub raw. Đổi khi đổi repo/nhánh. |
| `MIN_VALID_BYTES` | Ngưỡng kích thước tối thiểu coi file tải là hợp lệ (mặc định 50 KB). |
| `AUTO_REOPEN_EXCEL` | `True` = script tự mở lại Excel sau khi cập nhật. Mặc định `False`. |
| `NOTIFY_USER` | `True` = hiện hộp thoại báo có bản mới. Mặc định `False` (im lặng). |

## Giới hạn đã biết

- Checksum được tính **ngay trong tiến trình Excel** bằng .NET
  (`System.Security.Cryptography.SHA256Managed` qua COM, có sẵn trên mọi máy có
  .NET Framework — gần như mọi Windows 7+). Chỉ tốn vài mili giây. Nếu .NET
  không dùng được, add-in lùi về gọi PowerShell `Get-FileHash` — cách này chặn
  luồng chính Excel ~1–3 giây mỗi lần mở, nên chỉ là đường lui.
- Hai tiến trình Excel cùng lúc có thể cùng tải + cùng bung script; `copy /y`
  idempotent nên không hỏng file, chỉ dư một dòng `OK=0` trong log.
- Nếu file add-in nằm ở thư mục người dùng không có quyền ghi, script log `OK=0`
  và add-in giữ nguyên bản cũ; lần mở Excel sau sẽ thử lại.
- Cơ chế swap dùng `cmd.exe` (không phụ thuộc Windows Script Host).
- Vì cập nhật im lặng, người dùng không biết mình đang chạy bản nào. Muốn hiện
  thông báo thì đặt `NOTIFY_USER = True`.
