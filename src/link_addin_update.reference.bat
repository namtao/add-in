@echo off
rem ============================================================
rem  BAN THAM KHAO - noi dung .bat nay do modAutoUpdate.LaunchSwap
rem  tu sinh ra %TEMP%\link_addin_update_<timestamp>.bat luc chay.
rem  KHONG dung file nay truc tiep; sua logic o modAutoUpdate.bas.
rem
rem  Tham so:  %1 = duong dan file staging (LINK.update.xlam)
rem            %2 = duong dan file add-in that (ThisWorkbook.FullName)
rem ============================================================
setlocal
set "SRC=%~1"
set "DST=%~2"

:waitloop
tasklist /fi "imagename eq excel.exe" 2>nul | find /i "excel.exe" >nul
if not errorlevel 1 (
  ping -n 3 127.0.0.1 >nul
  goto waitloop
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
(goto) 2>nul & del "%~f0"
