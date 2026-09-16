@echo off
rem ============================================================
rem  BAN THAM KHAO - noi dung .bat nay do modAutoUpdate.LaunchSwap
rem  tu sinh ra %TEMP%\link_addin_update_<timestamp>.bat luc chay.
rem  KHONG dung file nay truc tiep; sua logic o modAutoUpdate.bas.
rem
rem  Tham so:  %1 = duong dan file staging (LINK.update.xlam)
rem            %2 = duong dan file add-in that (ThisWorkbook.FullName)
rem
rem  Script chay ngoai Excel nen no van song sau khi Excel thoat:
rem    1. doi toan bo EXCEL.EXE thoat (toi da 300 vong, ~2 giay/vong)
rem    2. copy file staging de len add-in that
rem    3. mo lai Excel de nguoi dung dung tiep voi ban moi
rem
rem  Nhanh TIMEOUT: nguoi dung bam Cancel o hop thoai luu file nen Excel
rem  khong dong. Script bo cuoc, xoa staging va tu xoa minh thay vi treo
rem  mai trong nen. Lan mo Excel sau add-in se hoi lai.
rem ============================================================
setlocal
set "SRC=%~1"
set "DST=%~2"
set "RUNNING=1"
for /l %%i in (1,1,300) do (
  tasklist /fi "imagename eq excel.exe" 2>nul | find /i "excel.exe" >nul
  if errorlevel 1 ( set "RUNNING=0" & goto closed )
  ping -n 3 127.0.0.1 >nul
)
:closed
if "%RUNNING%"=="1" (
  >>"%TEMP%\link_addin_update.log" echo %date% %time% TIMEOUT "%DST%"
  del /q "%SRC%" >nul 2>&1
  goto cleanup
)
set "OK=0"
for /l %%i in (1,1,5) do (
  copy /y "%SRC%" "%DST%" >nul 2>&1
  if not errorlevel 1 ( set "OK=1" & goto done )
  ping -n 3 127.0.0.1 >nul
)
:done
>>"%TEMP%\link_addin_update.log" echo %date% %time% OK=%OK% "%DST%"
if "%OK%"=="1" (
  del /q "%SRC%" >nul 2>&1
  start "" excel.exe
)
:cleanup
(goto) 2>nul & del "%~f0"
