@echo off
rem ============================================================
rem  Phat hanh ban LINK.xlam moi len GitHub.
rem
rem  CACH DUNG: keo tha file LINK.xlam da sua vao icon file nay.
rem
rem  KHONG can cai Git. KHONG can tai khoan GitHub. Chi can mot token
rem  phat hanh do nguoi quan tri gui rieng - dan mot lan o lan chay dau
rem  tien, sau do luu ma hoa tren may nay va khong hoi lai nua.
rem
rem  Script chi lam mot viec: day file .xlam len repo. Checksum
rem  (release/version.txt) do GitHub tu tinh lai sau ~30 giay.
rem ============================================================
setlocal

set "PS_URL=https://raw.githubusercontent.com/namtao/add-in/main/publish.ps1"
set "PS_TMP=%TEMP%\link_publish_%RANDOM%%RANDOM%.ps1"

echo ============================================
echo   Phat hanh ban moi cho add-in LINK
echo ============================================
echo.

if "%~1"=="" (
    echo Chua chon file de phat hanh.
    echo.
    echo Hay KEO THA file LINK.xlam da sua vao icon publish.bat nay,
    echo thay vi double-click vao no.
    echo.
    pause
    exit /b 1
)

echo Dang tai script phat hanh...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; try { Invoke-WebRequest -Uri '%PS_URL%' -OutFile '%PS_TMP%' -UseBasicParsing } catch { exit 1 }"

if not exist "%PS_TMP%" (
    echo.
    echo LOI: khong tai duoc script. Kiem tra ket noi mang roi thu lai.
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_TMP%" -SrcPath "%~1"
set "RC=%ERRORLEVEL%"

del /q "%PS_TMP%" >nul 2>&1

echo.
pause
exit /b %RC%
