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

# --- Kiem tra file nguon ---
if (-not (Test-Path -LiteralPath $SrcPath)) { throw "Khong tim thay file: $SrcPath" }
$bytes = [IO.File]::ReadAllBytes($SrcPath)
if ($bytes.Length -lt $MinBytes) {
    throw "File chi co $($bytes.Length) bytes - qua nho, co ve khong phai LINK.xlam."
}
if ($bytes[0] -ne 0x50 -or $bytes[1] -ne 0x4B) {
    throw "File khong phai dinh dang .xlam hop le (thieu chu ky ZIP 'PK')."
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
