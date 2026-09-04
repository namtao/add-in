@echo off
rem ============================================================
rem  Phat hanh ban LINK.xlam moi len GitHub - KHONG can go gi ca,
rem  khong can biet lenh git.
rem
rem  CACH NHANH NHAT: keo tha file LINK.xlam da sua vao thang icon
rem  publish.bat nay. Sua file o dau cung duoc.
rem
rem  Cach 2: sua truc tiep va LUU DE vao
rem      %USERPROFILE%\LINK-addin-publish\release\LINK.xlam
rem    roi double-click publish.bat.
rem
rem  Script tu tinh checksum SHA256 cua LINK.xlam lam "version" - khong
rem  can go so version tay, khong the quen bump. Sua noi dung file la du
rem  de moi may nguoi dung tu nhan duoc ban moi.
rem
rem  Yeu cau: da cai "Git for Windows" (https://git-scm.com/download/win)
rem  va da duoc them vao repo GitHub voi quyen ghi (Settings > Collaborators).
rem ============================================================
setlocal EnableDelayedExpansion

set "REPO_URL=https://github.com/namtao/add-in.git"
set "WORKDIR=%USERPROFILE%\LINK-addin-publish"
set "KEEP=%TEMP%\link_publish_keep.xlam"
set "HASHTMP=%TEMP%\link_publish_hash.txt"

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

git -C "%WORKDIR%" rev-parse --is-inside-work-tree >nul 2>&1
if errorlevel 1 (
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
    echo Da tai xong. Gio hay keo tha file LINK.xlam da sua vao icon
    echo publish.bat nay, hoac luu de vao:
    echo   %WORKDIR%\release\LINK.xlam
    echo roi chay lai publish.bat.
    pause
    exit /b 0
)

cd /d "%WORKDIR%"

git config user.name >nul 2>&1
if errorlevel 1 git config user.name "LINK Publisher"
git config user.email >nul 2>&1
if errorlevel 1 git config user.email "link-addin-publisher@localhost"

rem --- Xac dinh file nguon: keo tha, hoac ban da luu de trong release\ ---
set "SRC=%~1"
if "%SRC%"=="" set "SRC=%WORKDIR%\release\LINK.xlam"

if not exist "%SRC%" (
    echo Chua co file de phat hanh.
    echo Keo tha file LINK.xlam da sua vao icon publish.bat, hoac luu de vao:
    echo   %WORKDIR%\release\LINK.xlam
    echo roi chay lai.
    pause
    exit /b 1
)

rem Giu ban cua nguoi phat hanh ra ngoai TRUOC khi dong bo voi GitHub, de buoc
rem "reset --hard" duoi day khong bao gio lam mat cong cua ho.
copy /y "%SRC%" "%KEEP%" >nul
if errorlevel 1 (
    echo LOI: khong doc duoc file nguon: %SRC%
    pause
    exit /b 1
)

rem --- Canh bao neu nguoi khac vua phat hanh trong luc ban dang sua ---
echo Dang dong bo voi GitHub...
git fetch origin >nul 2>&1
if errorlevel 1 (
    echo.
    echo LOI: khong ket noi duoc GitHub. Kiem tra mang roi thu lai.
    pause
    exit /b 1
)

set "BEHIND=0"
for /f %%N in ('git rev-list --count HEAD..origin/main 2^>nul') do set "BEHIND=%%N"
if not "!BEHIND!"=="0" (
    echo.
    echo ********************* CANH BAO *********************
    echo  Trong luc ban dang sua, da co !BEHIND! ban phat hanh moi tu nguoi khac.
    echo  Ban cua ban dua tren ban CU hon, nen thay doi cua ho se bi ghi de.
    echo.
    echo  Neu khong chac, bam Ctrl+C ngay bay gio de dung lai, lay ban moi
    echo  nhat ve sua lai. ^(Ban cu van con trong lich su git, khong mat han.^)
    echo.
    echo  Tu dong tiep tuc sau 10 giay...
    echo ****************************************************
    timeout /t 10
)

call :sync_and_commit
if "%ERRORLEVEL%"=="3" (
    echo.
    echo Khong co gi thay doi de phat hanh ^(file giong het ban dang phat hanh^).
    goto :cleanup_exit
)
if not "%ERRORLEVEL%"=="0" (
    echo.
    echo LOI: khong tinh duoc checksum hop le cua LINK.xlam.
    echo Kiem tra file co dung la .xlam khong, va may co PowerShell khong.
    goto :fail_exit
)

git push
if errorlevel 1 (
    echo.
    echo Co nguoi khac vua push cung luc - dang lay ban moi va thu lai...
    call :sync_and_commit
    if not "!ERRORLEVEL!"=="0" (
        echo LOI: khong phat hanh duoc. Chay lai publish.bat.
        goto :fail_exit
    )
    git push
    if errorlevel 1 (
        echo.
        echo LOI: khong push len GitHub duoc.
        echo Neu day la lan dau, mot cua so dang nhap GitHub co the da hien ra -
        echo dang nhap xong roi chay lai publish.bat.
        goto :fail_exit
    )
)

echo.
echo ============================================
echo   Da phat hanh ban moi thanh cong.
echo   May nguoi dung se tu nhan ban nay o lan mo Excel ke tiep.
echo ============================================

:cleanup_exit
del /q "%KEEP%" >nul 2>&1
pause
exit /b 0

:fail_exit
del /q "%KEEP%" >nul 2>&1
pause
exit /b 1

rem ============================================================
rem  Lay ban moi nhat tu GitHub, dat file cua nguoi phat hanh len tren,
rem  tinh checksum, commit ca hai file cung mot lan.
rem  Tra ve: 0 = da commit, 2 = checksum hong, 3 = khong co gi thay doi
rem ============================================================
:sync_and_commit
git reset --hard origin/main >nul 2>&1
copy /y "%KEEP%" "release\LINK.xlam" >nul

del /q "%HASHTMP%" >nul 2>&1
powershell -NoProfile -NonInteractive -Command "(Get-FileHash -Algorithm SHA256 -LiteralPath '%WORKDIR%\release\LINK.xlam').Hash" > "%HASHTMP%" 2>nul

set "HASH="
if exist "%HASHTMP%" for /f "usebackq delims=" %%H in ("%HASHTMP%") do set "HASH=%%H"
del /q "%HASHTMP%" >nul 2>&1

rem Checksum phai dung 64 ky tu hex. Neu PowerShell that bai thi HASH rong hoac
rem la rac (vi du "ECHO is on.") - publish rac se khien moi may khong bao gio
rem khop checksum, nen chan ngay tai day.
if "!HASH!"=="" exit /b 2
if "!HASH:~63,1!"=="" exit /b 2
if not "!HASH:~64,1!"=="" exit /b 2
echo !HASH!| findstr /i /r /c:"^[0-9a-f]*$" >nul
if errorlevel 1 exit /b 2

> "release\version.txt" echo !HASH!

git add release/LINK.xlam release/version.txt >nul 2>&1
git commit -m "release: cap nhat LINK.xlam" >nul 2>&1
if errorlevel 1 exit /b 3
exit /b 0
