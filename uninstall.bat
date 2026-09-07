@echo off
rem ============================================================
rem  Go add-in LINK khoi Excel - chi can double-click file nay.
rem
rem  Xoa file add-in trong thu muc XLSTART va don rac tam. Khong
rem  dung den registry, khong can quyen admin.
rem
rem  Muon cai lai luc nao cung duoc bang install.bat.
rem ============================================================
setlocal EnableDelayedExpansion

set "XLSTART=%APPDATA%\Microsoft\Excel\XLSTART"
set "DST=%XLSTART%\LINK.xlam"
set "STAGING=%XLSTART%\LINK.update.xlam"
set "LEGACY=%APPDATA%\Microsoft\AddIns\LINK.xlam"
set "TOKENFILE=%USERPROFILE%\.link-addin-token"

echo ============================================
echo   Go add-in LINK khoi Excel
echo ============================================
echo.

rem --- Co gi de go khong? ---
set "FOUND=0"
if exist "%DST%" set "FOUND=1"
if exist "%LEGACY%" set "FOUND=1"
if "!FOUND!"=="0" (
    echo Khong tim thay add-in LINK tren may nay - khong co gi de go.
    echo.
    pause
    exit /b 0
)

rem --- Excel phai dong han moi xoa duoc file ---
tasklist /fi "imagename eq excel.exe" 2>nul | find /i "excel.exe" >nul
if not errorlevel 1 (
    echo Excel dang mo. Dong TOAN BO cua so Excel roi bam phim bat ky de tiep tuc...
    pause >nul
    :waitclose
    tasklist /fi "imagename eq excel.exe" 2>nul | find /i "excel.exe" >nul
    if not errorlevel 1 (
        ping -n 2 127.0.0.1 >nul
        goto waitclose
    )
)

rem --- Xoa file add-in chinh ---
if exist "%DST%" (
    del /q "%DST%" >nul 2>&1
    if exist "%DST%" (
        echo.
        echo LOI: khong xoa duoc file:
        echo   %DST%
        echo Kiem tra Excel da dong han chua roi chay lai file nay.
        echo.
        pause
        exit /b 1
    )
    echo Da xoa: %DST%
)

rem --- Don rac tam do co che tu cap nhat sinh ra ---
del /q "%STAGING%" >nul 2>&1
del /q "%TEMP%\link_addin_update.log" >nul 2>&1
del /q "%TEMP%\link_addin_update_*.bat" >nul 2>&1
del /q "%TEMP%\link_hash_*.txt" >nul 2>&1
echo Da don cac file tam.

rem --- Ban cai theo kieu cu (Options ^> Add-ins ^> Browse) ---
if exist "%LEGACY%" (
    echo.
    echo ******************** LUU Y ********************
    echo Con mot ban cai theo kieu cu tai:
    echo   %LEGACY%
    echo.
    echo File nay duoc dang ky trong registry cua Excel. Neu xoa thang
    echo file ma khong bo dang ky, moi lan mo Excel se bao loi thieu file.
    echo.
    echo Hay go dung cach: mo Excel ^> File ^> Options ^> Add-ins ^>
    echo Manage: Excel Add-ins ^> Go... ^> BO TICK dong LINK ^> OK.
    echo Sau do dong Excel va xoa file o duong dan tren.
    echo ***********************************************
)

rem --- Token phat hanh: chi bao, KHONG tu xoa ---
if exist "%TOKENFILE%" (
    echo.
    echo Ghi chu: may nay con file token phat hanh tai
    echo   %TOKENFILE%
    echo Day la thong tin dang nhap de phat hanh ban moi, khong lien quan den
    echo add-in vua go. Neu ban khong con lam nguoi phat hanh nua thi xoa no di.
)

echo.
echo ============================================
echo   Da go xong. Mo Excel se khong con tab LINK.
echo   Muon cai lai: chay install.bat.
echo ============================================
echo.
pause
