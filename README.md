# LINK add-in — tự động cập nhật

Add-in ribbon Excel `LINK.xlam` tự kiểm tra phiên bản mới trên repo này mỗi lần
mở Excel và tự cập nhật, không cần gửi file thủ công.

## Cách hoạt động

1. `ThisWorkbook.Workbook_Open` gọi `modAutoUpdate.ScheduleUpdateCheck`.
2. Sau ~2 giây (qua `Application.OnTime`, không chặn Excel khởi động),
   `CheckForUpdate` tính checksum SHA256 của chính file add-in đang chạy, tải
   `release/version.txt` từ GitHub raw (checksum của bản đang phát hành) và so
   hai giá trị này.
3. Nếu khác nhau → tải `release/LINK.xlam` về `LINK.update.xlam` cạnh file add-in,
   kiểm tra kích thước tối thiểu + chữ ký ZIP (`PK`).
4. Ghi một script `.bat` ra `%TEMP%`, chạy ẩn. Script:
   - đợi **toàn bộ** `EXCEL.EXE` thoát;
   - `copy /y` bản mới đè lên file add-in thật (thử lại tối đa 5 lần);
   - ghi log `%TEMP%\link_addin_update.log`;
   - mở lại Excel (tắt bằng `AUTO_REOPEN_EXCEL = False` trong module);
   - tự xoá.
5. Add-in báo cho người dùng bằng một hộp thoại tiếng Việt rồi để họ đóng Excel
   khi thuận tiện.

Mọi lỗi mạng / HTTPS / file đều bị nuốt lặng — add-in luôn chạy tiếp ở phiên bản
hiện tại, Excel không bị chậm hay hiện lỗi khi offline.

**Không có số version thủ công:** `release/version.txt` là checksum SHA256 của
`release/LINK.xlam`, `publish.bat` tự tính — người phát hành chỉ cần sửa nội
dung add-in và lưu, không cần nhớ bump version, không thể quên.

## Bố cục repo

```
install.bat        # gửi cho người dùng cuối — double-click là cài xong
publish.bat         # gửi cho người phát hành bản mới — double-click là xong, không gõ gì
release/
  LINK.xlam         # bản đang phát hành (install.bat tải file này)
  version.txt       # checksum SHA256 của LINK.xlam ở trên — publish.bat tự tính
src/
  modAutoUpdate.bas             # module cần import vào LINK.xlam (làm 1 lần, xem INSTALL.md)
  ThisWorkbook.snippet.txt      # đoạn Workbook_Open cần dán
  link_addin_update.reference.bat  # bản tham khảo của .bat sinh lúc chạy
INSTALL.md       # thiết lập lần đầu (import module) + cách rollout bằng install.bat
```

## Quy trình phát hành bản mới

> Ai phát hành cũng được (không chỉ chủ repo) miễn có quyền ghi (collaborator)
> — xem "Người khác muốn phát hành bản mới" trong `INSTALL.md`. Người dùng cuối
> **không cần làm gì** khi có bản mới, chỉ cần đã cài add-in đúng cách một lần
> (xem `INSTALL.md` mục 4).

1. Mở `LINK.xlam` trong Excel, sửa nội dung/ribbon/VBA cần thiết, lưu, đóng Excel.
   Không có hằng số version nào phải sửa.
2. Double-click **[publish.bat](./publish.bat)** — không hỏi gì, không cần gõ
   gì. Script tự: `git pull` → tính checksum SHA256 của `LINK.xlam` → ghi vào
   `version.txt` → commit cả hai file cùng lúc → `git push`.
3. Pilot 1–2 máy: mở Excel → thấy hộp thoại cập nhật → đóng hết Excel → Excel tự
   mở lại ở bản mới. Kiểm tra `%TEMP%\link_addin_update.log` = `OK=1`.
4. Sạch → các máy còn lại tự nhận ở lần mở Excel kế tiếp.

`publish.bat` commit `LINK.xlam` và `version.txt` cùng một lần push nên không
có khoảng hở "version mới báo trước khi file mới sẵn sàng".

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
| `AUTO_REOPEN_EXCEL` | `True` = script tự mở lại Excel sau khi cập nhật. |

## Giới hạn đã biết

- Checksum tính bằng PowerShell (`Get-FileHash`, có sẵn từ Windows 7 SP1) qua
  `WScript.Shell.Run` đồng bộ — nếu máy không có PowerShell (rất hiếm), việc
  tính checksum thất bại và add-in bỏ qua lần kiểm tra đó, thử lại lần mở sau.
- Hai tiến trình Excel cùng lúc có thể cùng tải + cùng bung script; `copy /y`
  idempotent nên không hỏng file, chỉ dư một dòng `OK=0` trong log.
- Nếu file add-in nằm ở thư mục người dùng không có quyền ghi, script log `OK=0`
  và add-in giữ nguyên bản cũ; lần mở Excel sau sẽ thử lại.
- Cơ chế swap dùng `cmd.exe` (không phụ thuộc Windows Script Host).
