Attribute VB_Name = "modAutoUpdate"
'=====================================================================
' modAutoUpdate - Tu dong cap nhat add-in LINK.xlam tu GitHub
'---------------------------------------------------------------------
' Luong hoat dong:
'   ThisWorkbook.Workbook_Open
'     -> ScheduleUpdateCheck  (Application.OnTime tre ~2 giay, khong chan Excel)
'     -> CheckForUpdate
'          - tinh checksum SHA256 cua chinh file add-in dang chay (local)
'          - tai release/version.txt tren GitHub raw - la checksum cua ban
'            dang phat hanh (remote)
'          - neu khac  -> tai release/LINK.xlam moi ve file staging canh add-in
'          - kiem tra file tai: kich thuoc + chu ky "PK" cua ZIP + checksum
'            cua file tai ve PHAI trung remote (chan release bi lech, xem duoi)
'          - ghi 1 script .bat ra %TEMP%, chay an bang cmd.exe
'          - .bat doi toan bo EXCEL.EXE thoat -> copy de len add-in that -> tu xoa
'
' Khong dung so version thu cong: version.txt la checksum cua chinh file .xlam,
' publish.bat tu tinh - nguoi phat hanh khong can go so version nao, khong the
' quen bump. Sua noi dung file la du de kich hoat update.
'
' Vi sao phai kiem checksum file vua tai (guard chong lap vo han):
'   Neu version.txt tren server lech voi LINK.xlam that (push thieu, go nham),
'   ma cu swap bua thi sau khi swap xong checksum local van khac remote -> lan
'   mo Excel sau lai tai, lai swap... lap mai mai tren MOI may. Chi swap khi
'   file tai ve dung bang checksum server bao thi loi do khong the xay ra.
'
' Nguyen tac: MOI loi mang / HTTPS / file deu bi nuot lang.
' Add-in luon chay tiep binh thuong o phien ban hien tai.
'=====================================================================
Option Explicit

'--- Cau hinh (chinh khi doi repo / doi hanh vi) ---------------------
Private Const GH_BASE As String = "https://raw.githubusercontent.com/namtao/add-in/main/release/"
Private Const VERSION_FILE As String = "version.txt"
Private Const ADDIN_FILE As String = "LINK.xlam"
Private Const STAGING_NAME As String = "LINK.update.xlam"
Private Const HELPER_LOG As String = "link_addin_update.log"
Private Const MIN_VALID_BYTES As Long = 51200      ' 50 KB - chan file tai loi/rong
Private Const AUTO_REOPEN_EXCEL As Boolean = False ' True = tu mo lai Excel sau khi swap
Private Const NOTIFY_USER As Boolean = False       ' True = hien MsgBox bao co ban moi

Private mChecked As Boolean

'--- Goi tu ThisWorkbook.Workbook_Open -------------------------------
Public Sub ScheduleUpdateCheck()
    If mChecked Then Exit Sub
    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, 2), "modAutoUpdate.CheckForUpdate"
End Sub

'--- Kiem tra + kich hoat cap nhat (chay 1 lan / phien) -------------
Public Sub CheckForUpdate()
    On Error GoTo Done
    If mChecked Then Exit Sub
    mChecked = True

    Dim staging As String
    staging = AddinFolder() & STAGING_NAME

    Dim remoteHash As String
    remoteHash = Trim$(HttpGetText(GH_BASE & VERSION_FILE & "?nc=" & NoCache()))
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

    LaunchSwap staging, ThisWorkbook.FullName
    If NOTIFY_USER Then NotifyUpdated
Done:
    Exit Sub
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
    Dim st As Object, sha As Object
    Dim bytes As Variant, digest As Variant

    Set st = CreateObject("ADODB.Stream")
    st.Type = 1                       ' adTypeBinary
    st.Open
    st.LoadFromFile path
    bytes = st.Read
    st.Close

    Set sha = CreateObject("System.Security.Cryptography.SHA256Managed")
    digest = sha.ComputeHash_2((bytes))   ' ngoac kep de truyen ByVal

    Dim i As Long, s As String
    For i = LBound(digest) To UBound(digest)
        s = s & Right$("0" & Hex$(digest(i)), 2)
    Next i
    HashViaDotNet = s                 ' Hex$ tra ve chu HOA, trung dinh dang PowerShell
    Exit Function
fail:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
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
    HashViaPowerShell = Trim$(ReadTextFile(outFile))
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
Private Sub LaunchSwap(ByVal staging As String, ByVal target As String)
    On Error Resume Next
    Dim q As String: q = Chr$(34)
    Dim bat As String
    bat = Environ$("TEMP") & "\link_addin_update_" & Format$(Now, "yyyymmddhhnnss") & ".bat"

    Dim f As Integer
    f = FreeFile
    Open bat For Output As #f
    Print #f, "@echo off"
    Print #f, "setlocal"
    Print #f, "set " & q & "SRC=%~1" & q
    Print #f, "set " & q & "DST=%~2" & q
    Print #f, ":waitloop"
    Print #f, "tasklist /fi " & q & "imagename eq excel.exe" & q & " 2>nul | find /i " & q & "excel.exe" & q & " >nul"
    Print #f, "if not errorlevel 1 ("
    Print #f, "  ping -n 3 127.0.0.1 >nul"
    Print #f, "  goto waitloop"
    Print #f, ")"
    Print #f, "set " & q & "OK=0" & q
    Print #f, "for /l %%i in (1,1,5) do ("
    Print #f, "  copy /y " & q & "%SRC%" & q & " " & q & "%DST%" & q & " >nul 2>&1"
    Print #f, "  if not errorlevel 1 ( set " & q & "OK=1" & q & " & goto done )"
    Print #f, "  ping -n 3 127.0.0.1 >nul"
    Print #f, ")"
    Print #f, ":done"
    Print #f, ">>" & q & "%TEMP%\" & HELPER_LOG & q & " echo %date% %time% OK=%OK% " & q & "%DST%" & q
    Print #f, "if " & q & "%OK%" & q & "==" & q & "1" & q & " ("
    Print #f, "  del /q " & q & "%SRC%" & q & " >nul 2>&1"
    If AUTO_REOPEN_EXCEL Then Print #f, "  start " & q & q & " excel.exe"
    Print #f, ")"
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

Private Function NoCache() As String
    NoCache = Format$(Now, "yyyymmddhhnnss")
End Function

Private Sub SafeKill(ByVal path As String)
    On Error Resume Next
    If Len(Dir$(path)) > 0 Then Kill path
End Sub

'--- Chi dung khi NOTIFY_USER = True ------------------------------
Private Sub NotifyUpdated()
    On Error Resume Next
    Dim msg As String, title As String
    msg = ChrW(272) & ChrW(227) & " c" & ChrW(243) & " b" & ChrW(7843) & "n c" & ChrW(7853) & "p nh" & ChrW(7853) & "t m" & ChrW(7899) & "i cho add-in LINK." & vbCrLf
    If AUTO_REOPEN_EXCEL Then
        msg = msg & "Vui l" & ChrW(242) & "ng " & ChrW(273) & ChrW(243) & "ng to" & ChrW(224) & "n b" & ChrW(7897) & " c" & ChrW(7917) & "a s" & ChrW(7893) & " Excel " & ChrW(273) & ChrW(7875) & " h" & ChrW(7879) & " th" & ChrW(7889) & "ng t" & ChrW(7921) & " c" & ChrW(7853) & "p nh" & ChrW(7853) & "t, sau " & ChrW(273) & ChrW(243) & " Excel s" & ChrW(7869) & " t" & ChrW(7921) & " m" & ChrW(7903) & " l" & ChrW(7841) & "i."
    Else
        msg = msg & "B" & ChrW(7843) & "n m" & ChrW(7899) & "i s" & ChrW(7869) & " t" & ChrW(7921) & " " & ChrW(273) & ChrW(7897) & "ng " & ChrW(225) & "p d" & ChrW(7909) & "ng sau khi b" & ChrW(7841) & "n " & ChrW(273) & ChrW(243) & "ng to" & ChrW(224) & "n b" & ChrW(7897) & " Excel."
    End If
    title = "C" & ChrW(7853) & "p nh" & ChrW(7853) & "t add-in LINK"
    ' Dung VBA.MsgBox tuong minh: chuoi da la Unicode (ChrW) nen hien dung dau.
    VBA.MsgBox msg, vbInformation, title
End Sub
