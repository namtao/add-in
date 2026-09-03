Attribute VB_Name = "modAutoUpdate"
'=====================================================================
' modAutoUpdate - Tu dong cap nhat add-in LINK.xlam tu GitHub
'---------------------------------------------------------------------
' Luong hoat dong:
'   ThisWorkbook.Workbook_Open
'     -> ScheduleUpdateCheck  (Application.OnTime tre ~2 giay, khong chan Excel)
'     -> CheckForUpdate
'          - tai release/version.txt tren GitHub raw, so voi ADDIN_VERSION
'          - neu khac  -> tai release/LINK.xlam moi ve file staging canh add-in
'          - kiem tra file tai (kich thuoc + chu ky "PK" cua ZIP)
'          - ghi 1 script .bat ra %TEMP%, chay an bang cmd.exe
'          - .bat doi toan bo EXCEL.EXE thoat -> copy de len add-in that
'            -> (tuy chon) mo lai Excel -> tu xoa
'
' Nguyen tac: MOI loi mang / HTTPS / file deu bi nuot lang.
' Add-in luon chay tiep binh thuong o phien ban hien tai.
'=====================================================================
Option Explicit

'--- Cau hinh (chinh khi doi repo / doi hanh vi) ---------------------
Public Const ADDIN_VERSION As String = "1.0.0"          ' bump moi lan phat hanh
Private Const GH_BASE As String = "https://raw.githubusercontent.com/namtao/add-in/main/release/"
Private Const VERSION_FILE As String = "version.txt"
Private Const ADDIN_FILE As String = "LINK.xlam"
Private Const STAGING_NAME As String = "LINK.update.xlam"
Private Const HELPER_LOG As String = "link_addin_update.log"
Private Const MIN_VALID_BYTES As Long = 51200           ' 50 KB - chan file tai loi/rong
Private Const AUTO_REOPEN_EXCEL As Boolean = True        ' mo lai Excel sau khi cap nhat

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

    Dim remote As String
    remote = Trim$(HttpGetText(GH_BASE & VERSION_FILE & "?nc=" & NoCache()))
    If Len(remote) = 0 Then GoTo Done

    If StrComp(remote, ADDIN_VERSION, vbTextCompare) = 0 Then
        ' Da la ban moi nhat - don rac neu lan truoc con sot staging
        SafeKill staging
        GoTo Done
    End If

    ' Co ban khac -> tai ve staging
    If Not HttpDownloadFile(GH_BASE & ADDIN_FILE & "?nc=" & NoCache(), staging) Then GoTo Done
    If Not LooksLikeZip(staging) Then
        SafeKill staging
        GoTo Done
    End If

    LaunchSwap staging, ThisWorkbook.FullName
    NotifyUpdated remote
Done:
    Exit Sub
End Sub

'--- HTTP -----------------------------------------------------------
Private Function HttpGetText(ByVal url As String) As String
    On Error GoTo fail
    Dim h As Object
    Set h = CreateObject("WinHttp.WinHttpRequest.5.1")
    h.SetTimeouts 5000, 5000, 15000, 15000
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

Private Sub NotifyUpdated(ByVal newVer As String)
    On Error Resume Next
    Dim head As String, tail As String, l2 As String, title As String
    head = ChrW(272) & ChrW(227) & " t" & ChrW(7843) & "i b" & ChrW(7843) & "n c" & ChrW(7853) & "p nh" & ChrW(7853) & "t "
    tail = " cho add-in LINK."
    If AUTO_REOPEN_EXCEL Then
        l2 = "Vui l" & ChrW(242) & "ng " & ChrW(273) & ChrW(243) & "ng to" & ChrW(224) & "n b" & ChrW(7897) & " c" & ChrW(7917) & "a s" & ChrW(7893) & " Excel " & ChrW(273) & ChrW(7875) & " h" & ChrW(7879) & " th" & ChrW(7889) & "ng t" & ChrW(7921) & " c" & ChrW(7853) & "p nh" & ChrW(7853) & "t, sau " & ChrW(273) & ChrW(243) & " Excel s" & ChrW(7869) & " t" & ChrW(7921) & " m" & ChrW(7903) & " l" & ChrW(7841) & "i."
    Else
        l2 = "Vui l" & ChrW(242) & "ng " & ChrW(273) & ChrW(243) & "ng to" & ChrW(224) & "n b" & ChrW(7897) & " c" & ChrW(7917) & "a s" & ChrW(7893) & " Excel " & ChrW(273) & ChrW(7875) & " h" & ChrW(7879) & " th" & ChrW(7889) & "ng t" & ChrW(7921) & " c" & ChrW(7853) & "p nh" & ChrW(7853) & "t."
    End If
    title = "C" & ChrW(7853) & "p nh" & ChrW(7853) & "t add-in LINK"
    ' Dung VBA.MsgBox tuong minh: chuoi da la Unicode (ChrW) nen hien dung dau.
    VBA.MsgBox head & newVer & tail & vbCrLf & l2, vbInformation, title
End Sub
