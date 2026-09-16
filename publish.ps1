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
function Get-ExcelSecurityKeys {
    $keys = @()
    $roots = Get-ChildItem 'HKCU:\Software\Microsoft\Office' -ErrorAction SilentlyContinue |
             Where-Object { $_.PSChildName -match '^\d+\.\d+$' }
    foreach ($r in $roots) {
        $excel = Join-Path $r.PSPath 'Excel'
        if (Test-Path $excel) { $keys += (Join-Path $excel 'Security') }
    }
    return $keys
}

function Test-VbomAccess {
    foreach ($sec in Get-ExcelSecurityKeys) {
        if (Test-Path $sec) {
            $v = (Get-ItemProperty -Path $sec -Name 'AccessVBOM' -ErrorAction SilentlyContinue).AccessVBOM
            if ($v -eq 1) { return $true }
        }
    }
    return $false
}

function Enable-VbomAccess {
    $changed = $false
    foreach ($sec in Get-ExcelSecurityKeys) {
        if (-not (Test-Path $sec)) {
            New-Item -Path $sec -Force -ErrorAction SilentlyContinue | Out-Null
        }
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

        # Khi Excel chan truy cap VBA project, PowerShell KHONG nem loi ma tra
        # ve $null. Phai kiem tra null, neu khong thi $vbp.VBComponents lang le
        # cho ra null va vo o cho khac, kho lan ra nguyen nhan.
        $vbp = $null
        try {
            $vbp = $wb.VBProject
        } catch {
            $vbp = $null
        }
        if ($null -eq $vbp -or $null -eq $vbp.VBComponents) {
            throw 'VBOM_BLOCKED'
        }

        # --- 1. Module tu cap nhat ---
        $module = $null
        foreach ($c in $vbp.VBComponents) {
            if ($c.Name -eq 'modAutoUpdate') { $module = $c }
        }

        # Dau hieu nhan biet ban module da sua loi Application.OnTime. Ban cu
        # goi OnTime bang ten macro khong co tien to ten workbook, nen Excel di
        # tim CheckForUpdate trong workbook dang active thay vi trong add-in va
        # lich hen bi bo qua trong im lang - do la ly do bo tu cap nhat khong
        # chay. Ban cu VAN co 'Sub CheckForUpdate' va van tro dung repo, nen neu
        # chi kiem tra hai thu do thi module hong se duoc giu lai va file phat
        # hanh ra tiep tuc chet lang.
        $qualifiedOnTime = '"''!modAutoUpdate.CheckForUpdate"'

        $needImport = $true
        if ($module) {
            $text = Get-ComponentText -Component $module
            if ($text.Contains('Sub CheckForUpdate') -and
                $text.Contains($RawBase) -and
                $text.Contains($qualifiedOnTime)) {
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
            if ($c.Name -eq 'modAutoUpdate') {
                $okModule = (Get-ComponentText -Component $c).Contains($qualifiedOnTime)
            }
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

    $basPath = Join-Path $workDir 'modAutoUpdate.bas'
    Invoke-WebRequest -Uri $ModuleUrl -OutFile $basPath -UseBasicParsing

    # Bat quyen doc VBA project neu chua bat. Lam im lang, khong hoi: day la
    # thiet lap cua rieng tai khoan Windows nay, khong can quyen admin, va
    # khong bat thi khong lam gi duoc.
    if (-not (Test-VbomAccess)) { Enable-VbomAccess | Out-Null }

    # Kiem tra va vá im lang. Khong in ra da sua gi - nguoi phat hanh khong can
    # biet, va co in cung khong lam gi khac di. Cong chan tren GitHub van soi lai.
    $finalFile = $workFile
    try {
        $finalFile = (Repair-Addin -SourceFile $workFile -ModuleBasPath $basPath -TargetXlam $targetFile).Path
    } catch {
        if ($_.Exception.Message -ne 'VBOM_BLOCKED') { throw }

        # Excel chi doc thiet lap bao mat luc khoi dong tien trinh. Neu vua bat
        # o tren ma van bi chan thi thuong la dang bam vao mot tien trinh Excel
        # cu - de no thoat han roi thu lai mot lan.
        Start-Sleep -Seconds 3
        try {
            $finalFile = (Repair-Addin -SourceFile $workFile -ModuleBasPath $basPath -TargetXlam $targetFile).Path
        } catch {
            if ($_.Exception.Message -ne 'VBOM_BLOCKED') { throw }
            throw @"
Excel dang chan script doc phan code VBA cua file.

Thu lan luot:
  1. Dong HET Excel dang mo (ke ca cua so an), roi chay lai publish.bat.
     Excel chi doc thiet lap bao mat luc khoi dong tien trinh.
  2. Kiem tra bang tay da tick chua:
     Excel -> File -> Options -> Trust Center -> Trust Center Settings
           -> Macro Settings -> "Trust access to the VBA project object model"
  3. Neu may do cong ty quan ly, thiet lap nay co the bi chinh sach khoa.
     Lien he IT, hoac phat hanh tu mot may khac.
"@
        }
    }

    $bytes = [IO.File]::ReadAllBytes($finalFile)

    # Chi hoi mot cau. Den day nghia la file da dat, khong con gi phai bao cao.
    $confirm = Read-Host "Phat hanh? (y/n)"
    if ($confirm -notmatch '^\s*[yY]\s*$') {
        Write-Host "Da huy." -ForegroundColor Yellow
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
