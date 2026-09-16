Attribute VB_Name = "modAutoUpdate"
'=====================================================================
' modAutoUpdate - Cap nhat add-in LINK.xlam tu GitHub
'---------------------------------------------------------------------
' Luong hoat dong:
'   ThisWorkbook.Workbook_Open
'     -> ScheduleUpdateCheck  (Application.OnTime tre ~3 giay, khong chan Excel)
'     -> CheckForUpdate
'          - tinh checksum SHA256 cua chinh file add-in dang chay (local)
'          - tai release/version.txt tren GitHub raw - la checksum cua ban
'            dang phat hanh (remote)
'          - neu khac  -> tai release/LINK.xlam moi ve file staging canh add-in
'          - kiem tra file tai: kich thuoc + chu ky "PK" cua ZIP + checksum
'            cua file tai ve PHAI trung remote (chan release bi lech, xem duoi)
'          - HOI NGUOI DUNG: OK = cap nhat ngay, Cancel = bo qua phien nay
'          - neu OK: luu cac workbook da tung luu, ghi 1 script .bat ra %TEMP%,
'            chay an bang cmd.exe, roi goi Application.Quit de dong Excel
'          - .bat doi toan bo EXCEL.EXE thoat -> copy de len add-in that
'            -> mo lai Excel -> tu xoa
'
' Khong dung so version thu cong: version.txt la checksum cua chinh file .xlam,
' publish.bat tu tinh - nguoi phat hanh khong can go so version nao, khong the
' quen bump. Sua noi dung file la du de kich hoat update.
'
' Vi sao ten macro trong Application.OnTime phai co tien to ten workbook:
'   OnTime phan giai ten thu tuc theo workbook DANG ACTIVE luc timer ban, chu
'   khong theo workbook da goi OnTime. Voi add-in, workbook active la file cua
'   nguoi dung - khong he co modAutoUpdate - nen lich hen bi bo qua trong im
'   lang va CheckForUpdate khong bao gio chay. Phai truyen dang
'   "'LINK.xlam'!modAutoUpdate.CheckForUpdate" thi Excel moi tim dung trong
'   add-in. Thieu tien to nay chinh la ly do ban cu khong tu cap nhat duoc.
'
' Vi sao phai kiem checksum file vua tai (guard chong lap vo han):
'   Neu version.txt tren server lech voi LINK.xlam that (push thieu, go nham),
'   ma cu swap bua thi sau khi swap xong checksum local van khac remote -> lan
'   mo Excel sau lai tai, lai swap... lap mai mai tren MOI may. Chi swap khi
'   file tai ve dung bang checksum server bao thi loi do khong the xay ra.
'
' Nguyen tac: MOI loi mang / HTTPS / file deu bi nuot lang.
' Add-in luon chay tiep binh thuong o phien ban hien tai. Nguoi dung chi thay
' hop thoai khi da co san mot ban moi tai xong va kiem tra dat.
'=====================================================================
Option Explicit

'--- Cau hinh (chinh khi doi repo / doi hanh vi) ---------------------
Private Const GH_BASE As String = "https://raw.githubusercontent.com/namtao/add-in/main/release/"
Private Const VERSION_FILE As String = "version.txt"
Private Const ADDIN_FILE As String = "LINK.xlam"
Private Const STAGING_NAME As String = "LINK.update.xlam"
Private Const HELPER_LOG As String = "link_addin_update.log"
Private Const MIN_VALID_BYTES As Long = 51200   ' 50 KB - chan file tai loi/rong
Private Const WAIT_ROUNDS As Long = 300         ' so vong .bat cho Excel thoat, ~2 giay/vong

Private mChecked As Boolean

'--- Goi tu ThisWorkbook.Workbook_Open -------------------------------
Public Sub ScheduleUpdateCheck()
    If mChecked Then Exit Sub

    ' Tien to "'LINK.xlam'!" la bat buoc - xem ghi chu dau file.
    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, 3), _
                       "'" & ThisWorkbook.Name & "'!modAutoUpdate.CheckForUpdate"
    If Err.Number <> 0 Then
        ' Khong hen gio duoc (hiem). Kiem tra thang, chap nhan cho vai giay.
        Err.Clear
        CheckForUpdate
    End If
    On Error GoTo 0
End Sub

'--- Chay tay de kiem tra: Alt+F8 -> go LINK_CheckUpdateNow ----------
Public Sub LINK_CheckUpdateNow()
    mChecked = False
    CheckForUpdate
End Sub

'--- Kiem tra + kich hoat cap nhat (chay 1 lan / phien) -------------
Public Sub CheckForUpdate()
    On Error GoTo Done
    If mChecked Then Exit Sub
    mChecked = True

    Dim staging As String
    staging = AddinFolder() & STAGING_NAME

    Dim remoteHash As String
    remoteHash = CleanHash(HttpGetText(GH_BASE & VERSION_FILE & "?nc=" & NoCache()))
    If Len(remoteHash) = 0 Then GoTo Done

    Dim localHash As String
    localHash = FileHash(ThisWorkbook.FullName)
    If Len(localHash) = 0 Then GoTo Done   ' khong tinh duoc checksum - thu lai lan sau

    If StrComp(remoteHash, localHash, vbTextCompare) = 0 Then
        ' Da la ban moi nhat - don rac neu lan truoc con sot staging
        SafeKill staging
        GoTo Done
    End If

    ' Noi dung khac -> tai ve staging
    If Not HttpDownloadFile(GH_BASE & ADDIN_FILE & "?nc=" & NoCache(), staging) Then GoTo Done
    If Not LooksLikeZip(staging) Then
        SafeKill staging
        GoTo Done
    End If

    ' Guard: file vua tai phai dung bang checksum server bao. Neu release tren
    ' server khong nhat quan thi bo qua han, khong swap (tranh lap vo han).
    If StrComp(FileHash(staging), remoteHash, vbTextCompare) <> 0 Then
        SafeKill staging
        GoTo Done
    End If

    ' Ban moi da san sang - tu day tro di moi lam phien nguoi dung.
    If Not UserAcceptsUpdate() Then
        ' Bo qua phien nay. Xoa staging cho sach; lan mo Excel sau se tai lai
        ' va hoi lai - ton vai tram KB nhung khong de lai file la canh add-in.
        SafeKill staging
        GoTo Done
    End If

    SaveOpenWorkbooks
    LaunchSwap staging, ThisWorkbook.FullName
    Application.Quit
Done:
    Exit Sub
End Sub

'--- Hoi nguoi dung -------------------------------------------------
Private Function UserAcceptsUpdate() As Boolean
    On Error Resume Next
    Dim msg As String, title As String

    msg = ChrW(272) & ChrW(227) & " c" & ChrW(243) & " b" & ChrW(7843) & _
          "n c" & ChrW(7853) & "p nh" & ChrW(7853) & "t m" & ChrW(7899) & _
          "i cho add-in LINK."
    msg = msg & vbCrLf & vbCrLf
    msg = msg & "B" & ChrW(7845) & "m OK: Excel s" & ChrW(7869) & " l" & _
          ChrW(432) & "u c" & ChrW(225) & "c file " & ChrW(273) & _
          "ang m" & ChrW(7903) & ", " & ChrW(273) & ChrW(243) & "ng l" & _
          ChrW(7841) & "i " & ChrW(273) & ChrW(7875) & " c" & ChrW(7853) & _
          "p nh" & ChrW(7853) & "t, r" & ChrW(7891) & "i t" & ChrW(7921) & _
          " m" & ChrW(7903) & " l" & ChrW(7841) & "i." & vbCrLf
    msg = msg & "B" & ChrW(7845) & "m Cancel: b" & ChrW(7887) & " qua l" & _
          ChrW(7847) & "n n" & ChrW(224) & "y, l" & ChrW(7847) & "n m" & _
          ChrW(7903) & " Excel sau s" & ChrW(7869) & " h" & ChrW(7887) & _
          "i l" & ChrW(7841) & "i." & vbCrLf

    title = "C" & ChrW(7853) & "p nh" & ChrW(7853) & "t add-in LINK"

    ' Dung VBA.MsgBox tuong minh: chuoi da la Unicode (ChrW) nen hien dung dau.
    ' vbMsgBoxSetForeground de hop thoai khong bi lap sau cua so khac luc Excel
    ' vua khoi dong xong.
    UserAcceptsUpdate = (VBA.MsgBox(msg, _
        vbOKCancel Or vbInformation Or vbMsgBoxSetForeground, title) = vbOK)
End Function

'--- Luu truoc khi dong Excel ---------------------------------------
' Chi luu workbook DA TUNG duoc luu (co duong dan). File chua tung luu van de
' Excel tu hoi cho luu luc Quit - khong tu quyet dinh giup nguoi dung, va cung
' khong am tham vut du lieu di.
Private Sub SaveOpenWorkbooks()
    On Error Resume Next
    Dim wb As Object
    For Each wb In Application.Workbooks
        If Not wb Is ThisWorkbook Then
            If Not wb.IsAddin Then
                If Not wb.Saved Then
                    If Len(wb.Path) > 0 Then wb.Save
                End If
            End If
        End If
    Next wb
    ' Add-in tu danh dau da luu de Excel khong hoi luu chinh no luc Quit.
    ThisWorkbook.Saved = True
End Sub

'--- HTTP -----------------------------------------------------------
Private Function HttpGetText(ByVal url As String) As String
    On Error GoTo fail
    Dim h As Object
    Set h = CreateObject("WinHttp.WinHttpRequest.5.1")
    h.SetTimeouts 3000, 3000, 5000, 5000   ' file 64 byte - khong doi lau
    h.Open "GET", url, False
    h.setRequestHeader "Cache-Control", "no-cache"
    h.send
    If h.Status = 200 Then HttpGetText = h.responseText
    Exit Function
fail:
    HttpGetText = vbNullString
End Function

Private Function HttpDownloadFile(ByVal url As String, ByVal path As String) As Boolean
    On Error GoTo fail
    Dim h As Object, st As Object
    Set h = CreateObject("WinHttp.WinHttpRequest.5.1")
    h.SetTimeouts 5000, 5000, 60000, 60000
    h.Open "GET", url, False
    h.setRequestHeader "Cache-Control", "no-cache"
    h.send
    If h.Status <> 200 Then GoTo fail

    Set st = CreateObject("ADODB.Stream")
    st.Type = 1                       ' adTypeBinary
    st.Open
    st.Write h.responseBody
    st.SaveToFile path, 2            ' adSaveCreateOverWrite
    st.Close
    HttpDownloadFile = (Len(Dir$(path)) > 0)
    Exit Function
fail:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
    HttpDownloadFile = False
End Function

'=== Checksum SHA256 =================================================
' Uu tien tinh NGAY TRONG TIEN TRINH Excel bang .NET (mscorlib da dang ky COM
' san tren moi may co .NET Framework, tuc la gan nhu moi Windows 7+). Cach nay
' chi ton vai mili giay.
'
' KHONG goi PowerShell lam cach chinh: WScript.Shell.Run(..., waitOnReturn:=True)
' chan luon luong chinh cua Excel, va PowerShell khoi dong nguoi tren may co
' antivirus mat 1-3 giay -> moi lan mo Excel deu bi don. PowerShell chi con la
' duong lui khi .NET khong dung duoc.
'=====================================================================
Private Function FileHash(ByVal path As String) As String
    Dim h As String
    h = HashViaDotNet(path)
    If Len(h) > 0 Then
        FileHash = h
    Else
        FileHash = HashViaPowerShell(path)
    End If
End Function

Private Function HashViaDotNet(ByVal path As String) As String
    On Error GoTo fail
    Dim f As Integer, n As Long, bytes() As Byte
    Dim sha As Object, digest As Variant
    Dim i As Long, s As String

    ' Doc bang lenh Open cua VBA, KHONG dung ADODB.Stream. ADODB.Stream mo file
    ' o che do doc quyen nen no bao loi 3002 "File could not be opened" khi file
    ' dang bi tien trinh khac giu - ma o day file can tinh checksum chinh la
    ' LINK.xlam dang duoc Excel mo. Do la ly do bo tu cap nhat khong tinh noi
    ' checksum cua chinh no va thoat im lang o moi phien.
    f = FreeFile
    Open path For Binary Access Read As #f
    n = LOF(f)
    If n = 0 Then
        Close #f
        Exit Function
    End If
    ReDim bytes(0 To n - 1)
    Get #f, 1, bytes
    Close #f

    Set sha = CreateObject("System.Security.Cryptography.SHA256Managed")
    digest = sha.ComputeHash_2((bytes))   ' ngoac kep de truyen ByVal

    For i = LBound(digest) To UBound(digest)
        s = s & Right$("0" & Hex$(digest(i)), 2)
    Next i
    HashViaDotNet = s                 ' Hex$ tra ve chu HOA, trung dinh dang PowerShell
    Exit Function
fail:
    On Error Resume Next
    Close #f
    HashViaDotNet = vbNullString
End Function

Private Function HashViaPowerShell(ByVal path As String) As String
    On Error GoTo fail
    Dim outFile As String
    outFile = Environ$("TEMP") & "\link_hash_" & Format$(Now, "yyyymmddhhnnss") & ".txt"
    SafeKill outFile

    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    Dim q As String: q = Chr$(34)
    Dim cmd As String
    ' Nhay don cua PowerShell cho duong dan -> khoi phai escape nhay kep long nhau
    cmd = "powershell.exe -NoProfile -NonInteractive -Command " & q & _
          "(Get-FileHash -Algorithm SHA256 -LiteralPath '" & path & "').Hash | " & _
          "Out-File -Encoding ascii '" & outFile & "'" & q

    sh.Run cmd, 0, True               ' 0 = an cua so, True = doi chay xong
    HashViaPowerShell = CleanHash(ReadTextFile(outFile))
    SafeKill outFile
    Exit Function
fail:
    HashViaPowerShell = vbNullString
End Function

Private Function ReadTextFile(ByVal path As String) As String
    On Error GoTo fail
    If Len(Dir$(path)) = 0 Then Exit Function
    Dim f As Integer, ln As String, buf As String
    f = FreeFile
    Open path For Input As #f
    Do While Not EOF(f)
        Line Input #f, ln
        buf = buf & ln
    Loop
    Close #f
    ReadTextFile = buf
    Exit Function
fail:
    On Error Resume Next
    Close #f
    ReadTextFile = vbNullString
End Function

'--- Kiem tra file tai la ZIP hop le (xlam = ZIP, magic "PK") -------
Private Function LooksLikeZip(ByVal path As String) As Boolean
    On Error GoTo fail
    If FileLen(path) < MIN_VALID_BYTES Then Exit Function
    Dim f As Integer, b1 As Byte, b2 As Byte
    f = FreeFile
    Open path For Binary Access Read As #f
    Get #f, 1, b1
    Get #f, 2, b2
    Close #f
    LooksLikeZip = (b1 = &H50 And b2 = &H4B)   ' "P" "K"
    Exit Function
fail:
    On Error Resume Next
    Close #f
    LooksLikeZip = False
End Function

'--- Ghi .bat helper + chay an ------------------------------------
' Script chay ngoai Excel nen no van song sau khi Excel thoat. No doi Excel
' thoat han roi moi de len file add-in dang bi khoa, sau do mo lai Excel.
' Vong cho co han WAIT_ROUNDS: neu nguoi dung bam Cancel o hop thoai luu file,
' Excel khong dong - script bo cuoc va don rac thay vi treo mai trong nen.
Private Sub LaunchSwap(ByVal staging As String, ByVal target As String)
    On Error Resume Next
    Dim q As String: q = Chr$(34)
    Dim bat As String
    bat = Environ$("TEMP") & "\link_addin_update_" & Format$(Now, "yyyymmddhhnnss") & ".bat"

    Dim logPath As String
    logPath = q & "%TEMP%\" & HELPER_LOG & q

    Dim f As Integer
    f = FreeFile
    Open bat For Output As #f
    Print #f, "@echo off"
    Print #f, "setlocal"
    Print #f, "set " & q & "SRC=%~1" & q
    Print #f, "set " & q & "DST=%~2" & q
    Print #f, "set " & q & "RUNNING=1" & q
    Print #f, "for /l %%i in (1,1," & CStr(WAIT_ROUNDS) & ") do ("
    Print #f, "  tasklist /fi " & q & "imagename eq excel.exe" & q & " 2>nul | find /i " & q & "excel.exe" & q & " >nul"
    Print #f, "  if errorlevel 1 ( set " & q & "RUNNING=0" & q & " & goto closed )"
    Print #f, "  ping -n 3 127.0.0.1 >nul"
    Print #f, ")"
    Print #f, ":closed"
    Print #f, "if " & q & "%RUNNING%" & q & "==" & q & "1" & q & " ("
    Print #f, "  >>" & logPath & " echo %date% %time% TIMEOUT " & q & "%DST%" & q
    Print #f, "  del /q " & q & "%SRC%" & q & " >nul 2>&1"
    Print #f, "  goto cleanup"
    Print #f, ")"
    Print #f, "set " & q & "OK=0" & q
    Print #f, "for /l %%i in (1,1,5) do ("
    Print #f, "  copy /y " & q & "%SRC%" & q & " " & q & "%DST%" & q & " >nul 2>&1"
    Print #f, "  if not errorlevel 1 ( set " & q & "OK=1" & q & " & goto done )"
    Print #f, "  ping -n 3 127.0.0.1 >nul"
    Print #f, ")"
    Print #f, ":done"
    Print #f, ">>" & logPath & " echo %date% %time% OK=%OK% " & q & "%DST%" & q
    Print #f, "if " & q & "%OK%" & q & "==" & q & "1" & q & " ("
    Print #f, "  del /q " & q & "%SRC%" & q & " >nul 2>&1"
    Print #f, "  start " & q & q & " excel.exe"
    Print #f, ")"
    Print #f, ":cleanup"
    Print #f, "(goto) 2>nul & del " & q & "%~f0" & q
    Close #f

    Dim cmd As String
    cmd = "cmd.exe /c " & q & q & bat & q & " " & q & staging & q & " " & q & target & q & q
    Shell cmd, vbHide
End Sub

'--- Tien ich -----------------------------------------------------
Private Function AddinFolder() As String
    Dim p As String
    p = ThisWorkbook.FullName
    AddinFolder = Left$(p, InStrRev(p, "\"))
End Function

' Tham so chong cache. LUU Y: raw.githubusercontent BO QUA query string khi
' cache (da do thuc te: x-cache HIT, max-age=300), nen tham so nay KHONG pha
' duoc cache cua GitHub - ban moi phat hanh co the mat toi 5 phut moi thay.
' Dieu do vo hai: trong luc do checksum khong khop nen add-in bo qua, lan mo
' Excel sau se nhan. Van giu tham so vi nhieu proxy noi bo cua cong ty co
' cache theo URL day du, va no khong ton gi.
Private Function NoCache() As String
    NoCache = Format$(Now, "yyyymmddhhnnss")
End Function

' Giu lai dung cac ky tu hex. KHONG dung Trim$: Trim$ trong VBA chi cat dau
' CACH, khong cat CR/LF. version.txt do GitHub Action ghi bang `echo "$HASH" >`
' nen luon ket thuc bang newline, va chuoi con dinh newline lam MOI phep StrComp
' sai - add-in tuong nhu co ban moi, tai ve, roi lai tuong file tai ve khong
' khop server, xoa staging va thoat im lang. Lam sach o day de khong con phu
' thuoc vao viec file tren server co newline hay khong.
Private Function CleanHash(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = UCase$(Mid$(s, i, 1))
        If (ch >= "0" And ch <= "9") Or (ch >= "A" And ch <= "F") Then out = out & ch
    Next i
    CleanHash = out
End Function

Private Sub SafeKill(ByVal path As String)
    On Error Resume Next
    If Len(Dir$(path)) > 0 Then Kill path
End Sub
