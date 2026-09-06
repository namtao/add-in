<#
  Phat hanh ban LINK.xlam moi len GitHub qua Contents API.
  Duoc publish.bat tai ve va chay - nguoi dung khong chay truc tiep file nay.

  Khong can cai Git, khong can tai khoan GitHub. Chi can mot token phat hanh
  do nguoi quan tri gui rieng (dan mot lan o lan chay dau tien).

  Token duoc luu ma hoa bang DPAPI tai %USERPROFILE%\.link-addin-token, chi
  giai ma duoc boi dung tai khoan Windows do tren dung may do.
#>
param(
    [Parameter(Mandatory = $true)][string]$SrcPath
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Owner  = 'namtao'
$Repo   = 'add-in'
$Path   = 'release/LINK.xlam'
$Branch = 'main'
$MinBytes = 51200

$tokenFile = Join-Path $env:USERPROFILE '.link-addin-token'

function Get-PublishToken {
    if (Test-Path -LiteralPath $tokenFile) {
        try {
            $sec = Get-Content -LiteralPath $tokenFile | ConvertTo-SecureString
            $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
            return [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
        } catch {
            Write-Host "Token da luu bi hong, can dan lai." -ForegroundColor Yellow
            Remove-Item -LiteralPath $tokenFile -Force -ErrorAction SilentlyContinue
        }
    }

    Write-Host ""
    Write-Host "Lan dau chay - can token phat hanh." -ForegroundColor Cyan
    Write-Host "Token do nguoi quan tri gui rieng cho ban (bat dau bang github_pat_ hoac ghp_)."
    $sec = Read-Host -AsSecureString "Dan token vao day roi Enter"
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
    $plain = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
    if ([string]::IsNullOrWhiteSpace($plain)) { throw "Chua nhap token." }

    $sec | ConvertFrom-SecureString | Set-Content -LiteralPath $tokenFile -Encoding ASCII
    Write-Host "Da luu token (ma hoa). Lan sau khong phai nhap lai." -ForegroundColor Green
    return $plain
}

# Kiem tra file .xlam da chua module tu cap nhat chua.
#
# Vi sao can: neu ai do lo sua nham mot ban LINK.xlam cu (ban chua co
# modAutoUpdate) roi phat hanh, thi ban "khong biet tu cap nhat" do se duoc day
# xuong MOI may. Tu do tro di khong may nao nhan duoc ban moi nua, phai di cai
# lai tung may bang install.bat. Chan ngay tai day.
#
# Ten module nam dang text doc duoc trong xl/vbaProject.bin, o ca hai dang
# ASCII va UTF-16LE - do ca hai cho chac.
function Test-HasAutoUpdate {
    param([Parameter(Mandatory = $true)][string]$XlamPath)

    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

    $zip = [IO.Compression.ZipFile]::OpenRead($XlamPath)
    try {
        $entry = $zip.Entries | Where-Object { $_.FullName -eq 'xl/vbaProject.bin' }
        if (-not $entry) { return $false }   # khong co VBA project nao ca

        $ms = New-Object IO.MemoryStream
        $stream = $entry.Open()
        try { $stream.CopyTo($ms) } finally { $stream.Dispose() }
        $vba = $ms.ToArray()
        $ms.Dispose()
    } finally {
        $zip.Dispose()
    }

    # Latin1 anh xa 1 byte = 1 ky tu, nho vay dung duoc String.IndexOf de do byte
    $latin1 = [Text.Encoding]::GetEncoding('ISO-8859-1')
    $hay    = $latin1.GetString($vba)
    $needle = 'modAutoUpdate'
    $needleUtf16 = $latin1.GetString([Text.Encoding]::Unicode.GetBytes($needle))

    return ($hay.IndexOf($needle) -ge 0) -or ($hay.IndexOf($needleUtf16) -ge 0)
}

# --- Kiem tra file nguon ---
if (-not (Test-Path -LiteralPath $SrcPath)) { throw "Khong tim thay file: $SrcPath" }
$bytes = [IO.File]::ReadAllBytes($SrcPath)
if ($bytes.Length -lt $MinBytes) {
    throw "File chi co $($bytes.Length) bytes - qua nho, co ve khong phai LINK.xlam."
}
if ($bytes[0] -ne 0x50 -or $bytes[1] -ne 0x4B) {
    throw "File khong phai dinh dang .xlam hop le (thieu chu ky ZIP 'PK')."
}

if (-not (Test-HasAutoUpdate -XlamPath $SrcPath)) {
    throw @"
File nay CHUA co module tu cap nhat (modAutoUpdate).

Co ve ban dang sua nham mot ban LINK.xlam cu. Neu phat hanh ban nay, MOI may se
mat kha nang tu cap nhat va phai di cai lai tung may bang install.bat.

Hay tai ban moi nhat ve roi sua lai tren do:
  https://github.com/namtao/add-in/raw/main/release/LINK.xlam
"@
}

$token = Get-PublishToken
$headers = @{
    Authorization          = "Bearer $token"
    'User-Agent'           = 'link-addin-publish'
    Accept                 = 'application/vnd.github+json'
    'X-GitHub-Api-Version' = '2022-11-28'
}
$api = "https://api.github.com/repos/$Owner/$Repo/contents/$Path"

# --- Lay sha cua ban dang phat hanh (de GitHub biet ta ghi de len ban nao) ---
Write-Host "Dang lay thong tin ban hien tai..." -ForegroundColor Cyan
$sha = $null
try {
    $cur = Invoke-RestMethod -Uri "${api}?ref=$Branch" -Headers $headers -Method Get
    $sha = $cur.sha
} catch {
    if ($_.Exception.Response.StatusCode.value__ -eq 401) {
        Remove-Item -LiteralPath $tokenFile -Force -ErrorAction SilentlyContinue
        throw "Token khong hop le hoac da het han. Da xoa token cu - chay lai va dan token moi."
    }
    if ($_.Exception.Response.StatusCode.value__ -eq 404) {
        throw "Token khong co quyen ghi vao $Owner/$Repo, hoac repo khong ton tai."
    }
    throw
}

# --- Day file len ---
Write-Host "Dang tai len ($([math]::Round($bytes.Length/1KB)) KB)..." -ForegroundColor Cyan
$body = @{
    message = 'release: cap nhat LINK.xlam'
    content = [Convert]::ToBase64String($bytes)
    branch  = $Branch
}
if ($sha) { $body.sha = $sha }

try {
    $res = Invoke-RestMethod -Uri $api -Headers $headers -Method Put `
        -Body ($body | ConvertTo-Json -Compress) -ContentType 'application/json'
} catch {
    $code = $_.Exception.Response.StatusCode.value__
    if ($code -eq 409) { throw "Co nguoi khac vua phat hanh cung luc. Chay lai publish.bat." }
    if ($code -eq 403) { throw "Token khong du quyen (can Contents: Read and write)." }
    throw
}

if ($null -eq $res.commit.sha) {
    Write-Host ""
    Write-Host "Noi dung giong het ban dang phat hanh - khong co gi de cap nhat." -ForegroundColor Yellow
    exit 0
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "  Da phat hanh thanh cong."                   -ForegroundColor Green
Write-Host "  Cho ~30 giay de he thong tinh lai checksum." -ForegroundColor Green
Write-Host "  May nguoi dung se tu nhan ban nay o lan mo Excel ke tiep." -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
