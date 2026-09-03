# LINK add-in — tự động cập nhật

Add-in ribbon Excel `LINK.xlam` tự kiểm tra phiên bản mới trên repo này mỗi lần
mở Excel và tự cập nhật, không cần gửi file thủ công.

## Cách hoạt động

1. `ThisWorkbook.Workbook_Open` gọi `modAutoUpdate.ScheduleUpdateCheck`.
2. Sau ~2 giây (qua `Application.OnTime`, không chặn Excel khởi động),
   `CheckForUpdate` tải `release/version.txt` từ GitHub raw và so với hằng số
   `ADDIN_VERSION` nhúng trong add-in.
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

## Bố cục repo

```
install.bat        # gửi cho người dùng cuối — double-click là cài xong
release/
  LINK.xlam         # bản đang phát hành (install.bat tải file này)
  version.txt       # 1 dòng, ví dụ 1.0.3 — phải khớp ADDIN_VERSION trong LINK.xlam
src/
  modAutoUpdate.bas             # module cần import vào LINK.xlam (làm 1 lần, xem INSTALL.md)
  ThisWorkbook.snippet.txt      # đoạn Workbook_Open cần dán
  link_addin_update.reference.bat  # bản tham khảo của .bat sinh lúc chạy
INSTALL.md       # thiết lập lần đầu (import module) + cách rollout bằng install.bat
```

## Quy trình phát hành bản mới

> Làm trên máy Windows có Excel. Luôn test trên nhánh phụ trước khi đụng `main`.
> Ai phát hành cũng được (không chỉ chủ repo) miễn có quyền ghi (collaborator)
> hoặc gửi Pull Request — xem "Người khác muốn phát hành bản mới" trong
> `INSTALL.md`. Người dùng cuối **không cần làm gì** khi có bản mới, chỉ cần đã
> cài add-in đúng cách một lần (xem `INSTALL.md` mục 4).

1. Mở `release/LINK.xlam`, VBE (`Alt+F11`).
2. Sửa `ADDIN_VERSION` trong `modAutoUpdate` sang version mới (ví dụ `1.0.3`).
3. `Debug → Compile VBAProject` (không lỗi) → `Ctrl+S` → đóng Excel.
4. (Khuyến nghị) test vòng cập nhật trên nhánh phụ: đẩy `LINK.xlam` + `version.txt`
   mới lên nhánh đó, tạm sửa `GH_BASE` trỏ nhánh đó, chạy thử.
5. Commit **theo đúng thứ tự**:
   - commit `release/LINK.xlam` trước;
   - rồi commit `release/version.txt = 1.0.3` (để client không thấy version mới
     trước khi file mới có mặt).
6. `git push` lên `main`. Tuỳ chọn `git tag v1.0.3`.
7. Pilot 1–2 máy: mở Excel → thấy hộp thoại cập nhật → đóng hết Excel → Excel tự
   mở lại ở version mới. Kiểm tra `%TEMP%\link_addin_update.log` = `OK=1`.
8. Sạch → các máy còn lại tự nhận ở lần mở Excel kế tiếp.

## Rollback

Publish lại bản tốt trước đó: đưa `release/LINK.xlam` cũ về + đặt
`release/version.txt` về version cũ, push. Máy đang ở bản lỗi sẽ "cập nhật" ngược
về bản cũ qua đúng cơ chế đó — không có nhánh code riêng.

## Cấu hình trong `modAutoUpdate.bas`

| Hằng số | Ý nghĩa |
|---|---|
| `ADDIN_VERSION` | Version của bản build hiện tại. **Bump mỗi lần phát hành.** |
| `GH_BASE` | URL thư mục `release/` trên GitHub raw. Đổi khi đổi repo/nhánh. |
| `MIN_VALID_BYTES` | Ngưỡng kích thước tối thiểu coi file tải là hợp lệ (mặc định 50 KB). |
| `AUTO_REOPEN_EXCEL` | `True` = script tự mở lại Excel sau khi cập nhật. |

## Giới hạn đã biết

- Hai tiến trình Excel cùng lúc có thể cùng tải + cùng bung script; `copy /y`
  idempotent nên không hỏng file, chỉ dư một dòng `OK=0` trong log.
- Nếu file add-in nằm ở thư mục người dùng không có quyền ghi, script log `OK=0`
  và add-in giữ nguyên bản cũ; lần mở Excel sau sẽ thử lại.
- Cơ chế dùng `cmd.exe` (không phụ thuộc Windows Script Host).
