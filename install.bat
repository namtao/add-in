@echo off
rem ============================================================
rem  Cai dat add-in LINK cho Excel - chi can double-click file nay.
rem  Tai ban moi nhat tu GitHub, dat vao thu muc XLSTART cua Excel
rem  (Excel tu mo add-in trong thu muc nay o moi lan khoi dong -
rem  khong can qua File > Options > Add-ins, khong can quyen admin).
rem
rem  Chay lai file nay bat cu luc nao de cai lai / lay ban moi nhat
rem  thu cong (binh thuong add-in se tu cap nhat, khong can chay lai).
rem ============================================================
setlocal EnableDelayedExpansion

set "ADDIN_URL=https://raw.githubusercontent.com/namtao/add-in/main/release/LINK.xlam"
set "DST_DIR=%APPDATA%\Microsoft\Excel\XLSTART"
set "DST=%DST_DIR%\LINK.xlam"
set "TMP=%DST%.tmp"

echo ============================================
echo   Cai dat add-in LINK cho Excel
echo ============================================
echo.

if not exist "%DST_DIR%" mkdir "%DST_DIR%" >nul 2>&1

if exist "%APPDATA%\Microsoft\AddIns\LINK.xlam" (
    echo Luu y: phat hien LINK.xlam cu tai:
    echo   %APPDATA%\Microsoft\AddIns\LINK.xlam
    echo Neu ribbon LINK hien 2 lan sau khi cai, hay go bo ban cu do qua
    echo File ^> Options ^> Add-ins ^> Manage: Excel Add-ins ^> Go, bo tick, roi xoa file.
    echo.
)

echo Dang tai ban moi nhat...
del /q "%TMP%" >nul 2>&1
where curl >nul 2>&1
if not errorlevel 1 (
    curl -L -f -s -o "%TMP%" "%ADDIN_URL%?nc=%RANDOM%%RANDOM%"
) else (
    powershell -NoProfile -Command "try { Invoke-WebRequest -Uri '%ADDIN_URL%?nc=%RANDOM%%RANDOM%' -OutFile '%TMP%' -UseBasicParsing } catch { exit 1 }"
)

if not exist "%TMP%" (
    echo.
    echo LOI: khong tai duoc file. Kiem tra ket noi mang / firewall roi thu lai.
    pause
    exit /b 1
)

for %%A in ("%TMP%") do set "SZ=%%~zA"
if !SZ! LSS 51200 (
    echo.
    echo LOI: file tai ve qua nho ^(!SZ! bytes^), co the bi loi mang.
    del /q "%TMP%" >nul 2>&1
    pause
    exit /b 1
)

tasklist /fi "imagename eq excel.exe" 2>nul | find /i "excel.exe" >nul
if not errorlevel 1 (
    echo.
    echo Excel dang mo. Dong TOAN BO cua so Excel roi bam phim bat ky de tiep tuc...
    pause >nul
    :waitclose
    tasklist /fi "imagename eq excel.exe" 2>nul | find /i "excel.exe" >nul
    if not errorlevel 1 (
        ping -n 2 127.0.0.1 >nul
        goto waitclose
    )
)

move /y "%TMP%" "%DST%" >nul

echo.
echo Da cai dat xong.
echo File:  %DST%
echo Mo Excel de kiem tra ribbon LINK hien ra la thanh cong.
echo.
pause
