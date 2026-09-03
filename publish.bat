@echo off
rem ============================================================
rem  Phat hanh ban LINK.xlam moi len GitHub - double-click la xong,
rem  KHONG can go gi ca. Khong can biet lenh git.
rem
rem  Cach dung:
rem   - Sua add-in trong Excel, luon LUU DE vao dung file:
rem       %USERPROFILE%\LINK-addin-publish\release\LINK.xlam
rem     (lan dau chua co thu muc nay thi chay publish.bat 1 lan de no
rem     tu tao, roi copy/luu file vao dung cho do).
rem   - Roi double-click publish.bat. Xong - khong hoi gi them.
rem
rem  Script tu tinh checksum SHA256 cua LINK.xlam lam "version" - khong
rem  can go so version tay, khong the quen bump gi ca. Sua noi dung file
rem  la du de nguoi dung khac tu nhan duoc ban moi.
rem
rem  Yeu cau: da cai "Git for Windows" (https://git-scm.com/download/win)
rem  va da duoc them vao repo GitHub voi quyen ghi (Settings > Collaborators).
rem ============================================================
setlocal EnableDelayedExpansion

set "REPO_URL=https://github.com/namtao/add-in.git"
set "WORKDIR=%USERPROFILE%\LINK-addin-publish"

echo ============================================
echo   Phat hanh ban moi cho add-in LINK
echo ============================================
echo.

where git >nul 2>&1
if errorlevel 1 (
    echo Chua cai Git. Tai va cai "Git for Windows" tai:
    echo   https://git-scm.com/download/win
    echo Cai xong ^(giu nguyen mac dinh, bam Next lien tuc^) roi chay lai file nay.
    pause
    exit /b 1
)

if not exist "%WORKDIR%\.git" (
    echo Lan dau chay - dang tai kho luu tru ve:
    echo   %WORKDIR%
    git clone "%REPO_URL%" "%WORKDIR%"
    if errorlevel 1 (
        echo.
        echo LOI: khong tai duoc. Kiem tra ket noi mang, hoac ban chua duoc cap
        echo quyen ghi vao repo ^(nho nguoi quan tri them ban vao Collaborators
        echo tren GitHub, muc Settings ^> Collaborators^).
        pause
        exit /b 1
    )
    echo.
    echo Da tai xong. Mo Excel, sua add-in xong thi LUU DE vao dung file:
    echo   %WORKDIR%\release\LINK.xlam
    echo Roi chay lai publish.bat lan nua.
    pause
    exit /b 0
)

cd /d "%WORKDIR%"

git config user.name >nul 2>&1
if errorlevel 1 git config user.name "LINK Publisher"
git config user.email >nul 2>&1
if errorlevel 1 git config user.email "link-addin-publisher@localhost"

git pull --ff-only
if errorlevel 1 (
    echo.
    echo LOI: khong dong bo duoc voi GitHub. Kiem tra ket noi mang.
    pause
    exit /b 1
)

if not "%~1"=="" (
    echo Dang copy file ban keo tha vao release\LINK.xlam ...
    copy /y "%~1" "release\LINK.xlam" >nul
)

if not exist "release\LINK.xlam" (
    echo.
    echo Chua co file de phat hanh. Sua add-in trong Excel roi LUU DE vao:
    echo   %WORKDIR%\release\LINK.xlam
    echo Roi chay lai publish.bat. ^(Hoac keo tha file LINK.xlam vao thang
    echo icon publish.bat nay.^)
    pause
    exit /b 1
)

echo Dang tinh checksum...
powershell -NoProfile -NonInteractive -Command "(Get-FileHash -Algorithm SHA256 -LiteralPath 'release\LINK.xlam').Hash" > "release\version.txt.tmp"
if errorlevel 1 (
    echo.
    echo LOI: khong tinh duoc checksum ^(can Windows co PowerShell^).
    del /q "release\version.txt.tmp" >nul 2>&1
    pause
    exit /b 1
)
rem Bo dong trong / khoang trang thua PowerShell co the them vao cuoi file
for /f "usebackq delims=" %%H in ("release\version.txt.tmp") do set "HASH=%%H"
del /q "release\version.txt.tmp" >nul 2>&1
> "release\version.txt" echo !HASH!

git add release\LINK.xlam release\version.txt
git commit -m "release: cap nhat LINK.xlam" >nul
if errorlevel 1 (
    echo.
    echo Khong co gi thay doi de phat hanh ^(file giong het ban truoc^).
    pause
    exit /b 0
)

git push
if errorlevel 1 (
    echo.
    echo LOI: khong push len GitHub duoc.
    echo Neu day la lan dau, mot cua so dang nhap GitHub co the da hien ra -
    echo dang nhap xong roi chay lai publish.bat.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   Da phat hanh ban moi thanh cong.
echo   May nguoi dung se tu nhan ban nay o lan mo Excel ke tiep.
echo ============================================
pause
