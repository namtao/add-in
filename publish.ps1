<#
  Phat hanh ban LINK.xlam moi len GitHub qua Contents API.
  Duoc publish.bat tai ve va chay - nguoi dung khong chay truc tiep file nay.

  Khong can cai Git, khong can tai khoan GitHub. Chi can mot token phat hanh
  do nguoi quan tri gui rieng (dan mot lan o lan chay dau tien).

  CACH DUNG: keo tha MOT file .xlam HOAC .xlsm vao icon publish.bat. Ten file
  la gi cung duoc - script luon day len dung duong dan release/LINK.xlam.

  Script tu lo toan bo phan ky thuat: neu file thieu module tu cap nhat, thieu
  doan goi trong ThisWorkbook, quen bat IsAddin, hoac con o dang .xlsm chua
  chuyen sang add-in, no tu xu ly roi moi phat hanh. Nguoi phat hanh chi can
  sua noi dung roi keo tha - khong phai nho thao tac Excel nao.

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

$RawBase   = "https://raw.githubusercontent.com/$Owner/$Repo/$Branch/"
$ModuleUrl = $RawBase + 'src/modAutoUpdate.bas'
$AddinContentType = 'application/vnd.ms-excel.addin.macroEnabled.main+xml'

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

# =====================================================================
#  Kiem tra o muc goi .xlam - doc thang file zip, khong can Excel
# =====================================================================

function Get-ZipEntryText {
    param([string]$ZipPath, [string]$EntryName)

    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    $zip = [IO.Compression.ZipFile]::OpenRead($ZipPath)
    try {
        $entry = $zip.Entries | Where-Object { $_.FullName -eq $EntryName }
        if (-not $entry) { return $null }
        $reader = New-Object IO.StreamReader($entry.Open())
        try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
    } finally {
        $zip.Dispose()
    }
}

function Get-ZipEntryBytes {
    param([string]$ZipPath, [string]$EntryName)

    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    $zip = [IO.Compression.ZipFile]::OpenRead($ZipPath)
    try {
        $entry = $zip.Entries | Where-Object { $_.FullName -eq $EntryName }
        if (-not $entry) { return $null }
        $ms = New-Object IO.MemoryStream
        $stream = $entry.Open()
        try { $stream.CopyTo($ms) } finally { $stream.Dispose() }
        return $ms.ToArray()
    } finally {
        $zip.Dispose()
    }
}

# Kiem tra nhanh o muc goi. Khong doc duoc ma VBA o day - ma VBA nam nen ben
# trong vbaProject.bin - nen chi ket luan duoc 3 dieu, du de bat cac loi pho bien.
function Get-PackageState {
    param([string]$XlamPath)

    $types = Get-ZipEntryText -ZipPath $XlamPath -EntryName '[Content_Types].xml'
    $book  = Get-ZipEntryText -ZipPath $XlamPath -EntryName 'xl/workbook.xml'
    $vba   = Get-ZipEntryBytes -ZipPath $XlamPath -EntryName 'xl/vbaProject.bin'

    $veryHidden = $false
    if ($book) {
        $m = [regex]::Match($book, '<workbookView[^>]*>')
        $veryHidden = $m.Success -and $m.Value.Contains('visibility="veryHidden"')
    }

    $hasModule = $false
    if ($vba) {
        # Ten module nam dang text doc duoc trong vbaProject.bin o ca hai dang
        # ASCII va UTF-16LE. Rieng TEN module thi khong bi nen nen tim duoc;
        # NOI DUNG code thi bi nen, phai mo bang Excel moi doc duoc.
        $latin1 = [Text.Encoding]::GetEncoding('ISO-8859-1')
        $hay = $latin1.GetString($vba)
        $needle = 'modAutoUpdate'
        $needleUtf16 = $latin1.GetString([Text.Encoding]::Unicode.GetBytes($needle))
        $hasModule = ($hay.IndexOf($needle) -ge 0) -or ($hay.IndexOf($needleUtf16) -ge 0)
    }

    return [pscustomobject]@{
        # Luon dat $null ben TRAI: neu ve phai la mang, PowerShell se loc tung
        # phan tu va tra ve mang chu khong phai True/False.
        IsAddinType = ($null -ne $types) -and $types.Contains($AddinContentType)
        IsAddinFlag = $veryHidden
        HasModule   = $hasModule
        HasVba      = ($null -ne $vba)
    }
}

# =====================================================================
#  Kiem tra + tu vá o muc ma VBA - can Excel
# =====================================================================

# Doc/ghi VBProject qua COM doi Excel bat "Trust access to the VBA project
# object model". Mac dinh tat. Day la khoa trong HKCU nen khong can quyen admin.
function Enable-VbomAccess {
    $changed = $false
    $roots = Get-ChildItem 'HKCU:\Software\Microsoft\Office' -ErrorAction SilentlyContinue |
             Where-Object { $_.PSChildName -match '^\d+\.\d+$' }
    foreach ($r in $roots) {
        $sec = Join-Path $r.PSPath 'Excel\Security'
        if (Test-Path $sec) {
            Set-ItemProperty -Path $sec -Name 'AccessVBOM' -Value 1 -Type DWord
            $changed = $true
        }
    }
    return $changed
}

function Get-ThisWorkbookComponent {
    param($Workbook, $Project)

    $codeName = $Workbook.CodeName
    foreach ($c in $Project.VBComponents) {
        if ($c.Name -eq $codeName) { return $c }
    }
    foreach ($c in $Project.VBComponents) {
        if ($c.Name -eq 'ThisWorkbook') { return $c }
    }
    return $null
}

function Get-ComponentText {
    param($Component)
    $cm = $Component.CodeModule
    if ($cm.CountOfLines -lt 1) { return '' }
    return $cm.Lines(1, $cm.CountOfLines)
}

# Mo file bang Excel, soi ma VBA, vá nhung gi thieu, luu lai, roi soi lai lan
# nua de chac chan. Tra ve danh sach viec da vá, hoac nem loi neu van chua dat.
function Repair-Addin {
    param(
        [string]$SourceFile,
        [string]$ModuleBasPath,
        [string]$TargetXlam
    )

    $fixed = @()
    $xl = $null
    $wb = $null

    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        # msoAutomationSecurityForceDisable - chan macro tu chay khi mo file,
        # neu khong bo tu cap nhat cua chinh add-in se khoi dong ngay luc nay.
        $xl.AutomationSecurity = 3

        $wb = $xl.Workbooks.Open($SourceFile)

        $vbp = $null
        try {
            $vbp = $wb.VBProject
        } catch {
            throw 'VBOM_BLOCKED'
        }

        # --- 1. Module tu cap nhat ---
        $module = $null
        foreach ($c in $vbp.VBComponents) {
            if ($c.Name -eq 'modAutoUpdate') { $module = $c }
        }

        $needImport = $true
        if ($module) {
            $text = Get-ComponentText -Component $module
            if ($text.Contains('Sub CheckForUpdate') -and $text.Contains($RawBase)) {
                $needImport = $false
            }
        }

        if ($needImport) {
            if ($module) {
                $vbp.VBComponents.Remove($module)
                $fixed += 'thay module modAutoUpdate cu bang ban chuan tu repo'
            } else {
                $fixed += 'them module tu cap nhat modAutoUpdate'
            }
            $vbp.VBComponents.Import($ModuleBasPath) | Out-Null
        }

        # --- 2. Doan goi trong ThisWorkbook ---
        $twb = Get-ThisWorkbookComponent -Workbook $wb -Project $vbp
        if (-not $twb) { throw 'Khong tim thay module ThisWorkbook trong file.' }

        $cm = $twb.CodeModule
        $text = Get-ComponentText -Component $twb
        $hasOpen = $text -match '(?im)^\s*(Private\s+|Public\s+)?Sub\s+Workbook_Open\s*\('
        $hasCall = $text -match '(?i)ScheduleUpdateCheck'

        if (-not $hasOpen) {
            # Chen vao CUOI module, khong dung AddFromString: vi tri chen cua
            # AddFromString khong duoc bao dam, ma neu no dat Sub len truoc dong
            # "Option Explicit" co san thi VBA khong bien dich duoc.
            $snippet = "Private Sub Workbook_Open()`r`n    modAutoUpdate.ScheduleUpdateCheck`r`nEnd Sub"
            $cm.InsertLines($cm.CountOfLines + 1, $snippet)
            $fixed += 'them Workbook_Open goi modAutoUpdate.ScheduleUpdateCheck'
        } elseif (-not $hasCall) {
            # Da co Workbook_Open san - chi chen them dong goi vao trong, khong
            # tao Sub thu hai (VBA se bao loi trung ten).
            $bodyLine = $cm.ProcBodyLine('Workbook_Open', 0)
            $cm.InsertLines($bodyLine + 1, '    modAutoUpdate.ScheduleUpdateCheck')
            $fixed += 'chen loi goi ScheduleUpdateCheck vao Workbook_Open co san'
        }

        # --- 3. Co add-in ---
        if (-not $wb.IsAddin) {
            $wb.IsAddin = $true
            $fixed += 'bat IsAddin = True'
        }

        if ($SourceFile -ne $TargetXlam) {
            # 55 = xlOpenXMLAddIn. Phai dat IsAddin = True TRUOC (da lam o tren),
            # vi dinh dang file va thuoc tinh IsAddin la hai thu doc lap nhau -
            # luu dung duoi .xlam khong tu dong bat thuoc tinh do.
            $wb.SaveAs($TargetXlam, 55)
            $fixed += 'chuyen tu .xlsm sang dinh dang add-in .xlam'
        } elseif ($fixed.Count -gt 0) {
            $wb.Save()
        }

        # --- 4. Soi lai sau khi vá ---
        $twb = Get-ThisWorkbookComponent -Workbook $wb -Project $vbp
        $text = Get-ComponentText -Component $twb
        $okOpen = $text -match '(?im)^\s*(Private\s+|Public\s+)?Sub\s+Workbook_Open\s*\('
        $okCall = $text -match '(?i)ScheduleUpdateCheck'

        $okModule = $false
        foreach ($c in $vbp.VBComponents) {
            if ($c.Name -eq 'modAutoUpdate') { $okModule = $true }
        }

        if (-not ($okOpen -and $okCall -and $okModule -and $wb.IsAddin)) {
            throw 'Da thu vá nhung file van chua dat. Khong phat hanh.'
        }

        return [pscustomobject]@{ Fixed = $fixed; Path = $TargetXlam }
    } finally {
        if ($wb) { try { $wb.Close($false) } catch { } }
        if ($xl) { try { $xl.Quit() } catch { } }
        foreach ($o in @($wb, $xl)) {
            if ($o) { try { [Runtime.InteropServices.Marshal]::ReleaseComObject($o) | Out-Null } catch { } }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# =====================================================================
#  Kiem tra file nguon
# =====================================================================

if (-not (Test-Path -LiteralPath $SrcPath)) { throw "Khong tim thay file: $SrcPath" }

$srcItem = Get-Item -LiteralPath $SrcPath
$srcExt = $srcItem.Extension.ToLowerInvariant()
if ($srcExt -ne '.xlam' -and $srcExt -ne '.xlsm') {
    throw @"
File nay co duoi '$($srcItem.Extension)'. Chi nhan .xlam hoac .xlsm.

Neu file cua ban la .xlsx thi no khong chua macro, khong phai add-in LINK.

Ten file dat gi cung duoc - script luon day len dung duong dan release/LINK.xlam.
"@
}
$isXlsm = ($srcExt -eq '.xlsm')

$bytes = [IO.File]::ReadAllBytes($SrcPath)
if ($bytes.Length -lt $MinBytes) {
    throw "File chi co $($bytes.Length) bytes - qua nho, co ve khong phai LINK.xlam."
}
if ($bytes[0] -ne 0x50 -or $bytes[1] -ne 0x4B) {
    throw "File khong phai dinh dang .xlam hop le (thieu chu ky ZIP 'PK')."
}

# Lam viec tren mot BAN SAO trong thu muc tam - khong bao gio sua file goc cua
# nguoi dung.
$workDir = Join-Path $env:TEMP ("link_publish_" + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDir | Out-Null
# Ten ban sao phai DUY NHAT, khong duoc trung 'LINK.xlam': may nguoi phat hanh
# thuong da cai san add-in, Excel nap no tu XLSTART ngay khi khoi dong, va Excel
# tu choi mo hai file trung ten cung luc.
$stem = 'publish_' + [Guid]::NewGuid().ToString('N')
$workFile   = Join-Path $workDir ($stem + $srcExt)
$targetFile = Join-Path $workDir ($stem + '.xlam')
Copy-Item -LiteralPath $SrcPath -Destination $workFile

try {
    $state = Get-PackageState -XlamPath $workFile
    if (-not $state.HasVba) {
        throw "File khong chua ma VBA nao. Day khong phai add-in LINK."
    }

    # File .xlsm thi chac chan phai xu ly: doi dinh dang va bat IsAddin.
    $needRepair = $isXlsm -or -not ($state.IsAddinType -and $state.IsAddinFlag -and $state.HasModule)

    Write-Host "Dang kiem tra file..." -ForegroundColor Cyan

    $basPath = Join-Path $workDir 'modAutoUpdate.bas'
    Invoke-WebRequest -Uri $ModuleUrl -OutFile $basPath -UseBasicParsing

    $fixed = @()
    $finalFile = $workFile
    try {
        $r = Repair-Addin -SourceFile $workFile -ModuleBasPath $basPath -TargetXlam $targetFile
        $fixed = $r.Fixed
        $finalFile = $r.Path
    } catch {
        if ($_.Exception.Message -eq 'VBOM_BLOCKED') {
            Write-Host ""
            Write-Host "Excel dang chan script doc phan code VBA cua file." -ForegroundColor Yellow
            Write-Host "Can bat mot lan tuy chon 'Trust access to the VBA project object model'."
            Write-Host "Day la thiet lap cua rieng tai khoan Windows nay, khong can quyen admin."
            Write-Host ""
            $answer = Read-Host "Bat ngay bay gio? (C/k)"
            if ($answer -eq '' -or $answer -match '^[cCyY]') {
                if (Enable-VbomAccess) {
                    Write-Host "Da bat. Dang thu lai..." -ForegroundColor Green
                    # Excel doc thiet lap bao mat luc khoi dong tien trinh, nen
                    # phai de tien trinh cu thoat han roi mo tien trinh moi.
                    Start-Sleep -Seconds 3
                    $r = Repair-Addin -SourceFile $workFile -ModuleBasPath $basPath -TargetXlam $targetFile
                    $fixed = $r.Fixed
                    $finalFile = $r.Path
                } else {
                    throw "Khong tim thay thiet lap Excel tren may nay. Excel da duoc cai chua?"
                }
            } else {
                # Nguoi dung tu choi: khong doc duoc code VBA nua, chi con cac
                # kiem tra o muc goi. Cong chan tren GitHub van se soi lai day du.
                if ($isXlsm) {
                    throw @"
File .xlsm bat buoc phai qua Excel de chuyen sang dinh dang add-in .xlam,
nen khong the bo qua buoc nay.

Chay lai publish.bat va chon 'C' khi duoc hoi.
"@
                }
                if ($needRepair) {
                    throw @"
File nay chua dat va script khong duoc phep tu vá.

Thieu:
$(if (-not $state.IsAddinType) { "  - khong phai dinh dang add-in (dang la workbook)`r`n" })$(if (-not $state.IsAddinFlag) { "  - chua bat IsAddin = True`r`n" })$(if (-not $state.HasModule) { "  - thieu module tu cap nhat modAutoUpdate`r`n" })
Chay lai va chon 'C' de script tu xu ly, hoac sua tay roi phat hanh lai.
"@
                }
                Write-Host "Bo qua buoc kiem tra sau. Cong chan tren GitHub van se soi lai." -ForegroundColor Yellow
            }
        } else {
            throw
        }
    }

    if ($fixed.Count -gt 0) {
        Write-Host ""
        Write-Host "File con thieu vai thu - script da tu xu ly:" -ForegroundColor Yellow
        foreach ($f in $fixed) { Write-Host "  - $f" -ForegroundColor Yellow }
        Write-Host ""
        Write-Host "Nen cap nhat lai file goc cua ban theo nhung diem tren, de lan sau" -ForegroundColor Yellow
        Write-Host "khong phai vá nua." -ForegroundColor Yellow
        Write-Host ""
    } else {
        Write-Host "File dat yeu cau." -ForegroundColor Green
    }

    $bytes = [IO.File]::ReadAllBytes($finalFile)

    # --- Xac nhan truoc khi day len ---
    # Buoc nay la co tinh: phat hanh se toi MOI may da cai add-in, nen nguoi
    # phat hanh phai nhin thay minh sap day cai gi roi moi gat dau. Tra loi
    # "khong" o day cung la cach chay thu: file van duoc kiem tra va vá day du,
    # bao cao ra man hinh, nhung khong co gi roi khoi may.
    $localHash = (Get-FileHash -LiteralPath $finalFile -Algorithm SHA256).Hash.ToUpperInvariant()

    $liveHash = $null
    try {
        $nc = [Guid]::NewGuid().ToString('N')
        $liveHash = (Invoke-WebRequest -Uri ($RawBase + 'release/version.txt?nc=' + $nc) `
            -UseBasicParsing).Content.Trim().ToUpperInvariant()
    } catch {
        # Khong lay duoc ban dang chay thi bo qua phan doi chieu, van cho phat hanh.
    }

    Write-Host ""
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "  SAP PHAT HANH"                              -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host ("  File nguon : " + $srcItem.Name)
    Write-Host ("  Kich thuoc : " + [math]::Round($bytes.Length / 1KB) + " KB")
    Write-Host ("  Day len    : $Owner/$Repo  ->  $Path")
    if ($fixed.Count -gt 0) {
        Write-Host "  Da tu sua  :"
        foreach ($f in $fixed) { Write-Host "                 - $f" }
    } else {
        Write-Host "  Da tu sua  : khong phai sua gi"
    }

    if ($null -ne $liveHash -and $liveHash -eq $localHash) {
        Write-Host ""
        Write-Host "  LUU Y: noi dung GIONG HET ban dang chay." -ForegroundColor Yellow
        Write-Host "  Phat hanh cung se khong thay doi gi tren may nguoi dung." -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "  Moi may da cai add-in se tu nhan ban nay." -ForegroundColor Yellow
    Write-Host ""
    $confirm = Read-Host "Go 'c' roi Enter de phat hanh (bat ky phim nao khac = huy)"
    if ($confirm -notmatch '^\s*[cC]\s*$') {
        Write-Host ""
        Write-Host "Da huy - khong co gi duoc day len." -ForegroundColor Yellow
        Write-Host "File goc cua ban khong bi thay doi." -ForegroundColor Yellow
        exit 0
    }

    # --- Lay sha cua ban dang phat hanh (de GitHub biet ta ghi de len ban nao) ---
    $token = Get-PublishToken
    $headers = @{
        Authorization          = "Bearer $token"
        'User-Agent'           = 'link-addin-publish'
        Accept                 = 'application/vnd.github+json'
        'X-GitHub-Api-Version' = '2022-11-28'
    }
    $api = "https://api.github.com/repos/$Owner/$Repo/contents/$Path"

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
} finally {
    Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
}
