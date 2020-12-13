Attribute VB_Name = "Module1"
Sub TEST001()
    With ThisWorkbook.Sheets("Sheet1")
        Set rangeTmp = .Cells(1, 1).CurrentRegion
        Debug.Print rangeTmp.Rows.Count
        Debug.Print rangeTmp.Columns.Count
        Call .Range(.Cells(2, 1), .Cells(rangeTmp.Rows.Count, rangeTmp.Columns.Count)).Sort( _
        key1:=.Cells(1, 3), _
        Order1:=xlAscending)
    End With
End Sub

Sub TEST002()
    ' JVLink関連の変数セット
    Dim retval As Long
    Dim readcount As Long
    Dim dlcount As Long
    Dim lastfiletimestamp As String
    Dim status As Long
    Dim buff As String
    Dim filename As String
    
    With ThisWorkbook.Sheets("Sheet1")
    '速報オッズ（３連単）の呼び出し
    UserForm1.JVLink1.JVClose
    UserForm1.JVLink1.JVInit ("EXCELSAMPLE")
        
    strdateTarg = 20190630
    placeTarg = "函館"
    racenumTarg = 1
    
    retval = UserForm1.JVLink1.JVOpen("RACE", strdateTarg - 1 & "000000", 1, readcount, dlcount, lastfiletimestamp)
    If (retval = -1) Then GoTo ERROR_PROCESS
    If (retval < -1) Then GoTo ERROR_PROCESS
    Do While status <> dlcount
        If Cancelflg = True Then Exit Do
        status = UserForm1.JVLink1.JVStatus
        'UserForm_Wait.Label1.Caption = dlcount & "ファイル中 " & status & " ファイルダウンロード完了"
        DoEvents
        Sleep (10)
    Loop
    
    Dim mO6Data(0) As JV_O6_ODDS_SANRENTAN
    retval = 1
    isGetData = False
    Do While retval <> 0
        retval = UserForm1.JVLink1.JVRead(buff, 110000, filename)
        If (retval < -1) Then GoTo LOOP_END1

        If Left(buff, 2) = "O6" Then
            Call SetData_O6(buff, mO6Data(0))
            ' 選んだ日付を超えたらループ抜ける
            If isGetData = True And Val(mO6Data(0).id.Year & mO6Data(0).id.MonthDay) > Val(strdateTarg) Then
                GoTo LOOP_END1
            End If
            If mO6Data(0).id.Year & mO6Data(0).id.MonthDay = strdateTarg Then
                strJyo = JyoCord(mO6Data(0).id.JyoCD)
                If IsEmpty(strJyo) Then GoTo LOOP_END1
                If strJyo = placeTarg And _
                    Val(mO6Data(0).id.racenum) = racenumTarg Then
                    isGetData = True
                    For i = 0 To 4895 'UBound(mO6Data(0).OddSanrentanInfo)
                        If Trim(mO6Data(0).OddsSanrentanInfo(i).Kumi) <> "" Then
                            .Cells(i + 1, 1) = mO6Data(0).OddsSanrentanInfo(i).Kumi
                            .Cells(i + 1, 2) = Val(Left(mO6Data(0).OddsSanrentanInfo(i).Kumi, 2))
                            .Cells(i + 1, 3) = Val(Mid(mO6Data(0).OddsSanrentanInfo(i).Kumi, 3, 2))
                            .Cells(i + 1, 4) = Val(Right(mO6Data(0).OddsSanrentanInfo(i).Kumi, 2))
                            .Cells(i + 1, 5) = Format(Val(mO6Data(0).OddsSanrentanInfo(i).Odds / 10), "0.0")
                        End If
                    Next i
                End If
            End If
        Else
            UserForm1.JVLink1.JVSkip
        End If
        DoEvents
    Loop
    
LOOP_END1:
ERROR_PROCESS:

        UserForm1.JVLink1.JVClose
    End With

End Sub

Function TestGetStockUmatanOdds(WSbase, datacsv As datacsv, strdateTarg As String, _
                        placeTarg As String, racenumTarg As Integer) As Long
    ' JVLink関連の変数セット
    Dim retval As Long
    Dim readcount As Long
    Dim dlcount As Long
    Dim lastfiletimestamp As String
    Dim status As Long
    Dim buff As String
    Dim filename As String
    Dim codeLength As Long
    
    Dim collUmatanOdds As New Collection
'    Set collUmatanOdds = New Collection
    Dim collUmatanOddsO1 As New Collection
'    Set collUmatanOddsO1 = New Collection
    Dim collUmatanOddsH1 As New Collection
'    Set collUmatanOddsH1 = New Collection
    Dim collOddsSanrentan As Collection
'    Set collOddsSanrentan = New Collection
    
    retval = UserForm1.JVLink1.JVOpen("RACE", strdateTarg - 1 & "000000", 1, readcount, dlcount, lastfiletimestamp)
    If (retval = -1) Then GoTo LOOP_END1
    If (retval < -1) Then GoTo ERROR_PROCESS
    Do While status <> dlcount
        If Cancelflg = True Then Exit Do
        status = UserForm1.JVLink1.JVStatus
        UserForm_Wait.Label1.Caption = dlcount & "ファイル中 " & status & " ファイルダウンロード完了"
        DoEvents
        Sleep (10)
    Loop
    
    ' オッズ（馬単）
    Dim mO4Data As JV_O4_ODDS_UMATAN
    ' オッズ（単複枠）
    Dim mO1Data As JV_O1_ODDS_TANFUKUWAKU
    ' オッズ（３連単）
    Dim mO6Data(0) As JV_O6_ODDS_SANRENTAN
    ' 票数1
    Dim mH1Data As JV_H1_HYOSU_ZENKAKE
    
    retval = 1
    isGetData = False
    codeLength = 30000
    cnt = 1
    With ThisWorkbook.Sheets("Sheet2")
    Do While retval <> 0
        retval = UserForm1.JVLink1.JVRead(buff, codeLength, filename)
        If (retval < -1) Then GoTo ERROR_PROCESS
        
        If Left(buff, 2) = "H1" Then
            Call SetData_H1(buff, mH1Data)
            If mH1Data.id.Year & mH1Data.id.MonthDay = strdateTarg Then
                strJyo = JyoCord(mH1Data.id.JyoCD)
                If IsEmpty(strJyo) Then GoTo LOOP_NEXT1
                If strJyo = placeTarg And _
                    Val(mH1Data.id.racenum) = racenumTarg Then
                    .Cells(cnt, 1) = buff
                    cnt = cnt + 1
                    codeLength = 1000
                End If
            End If
        ElseIf Left(buff, 2) = "O1" Then
            Call SetData_O1(buff, mO1Data)
            If mO1Data.id.Year & mO1Data.id.MonthDay = strdateTarg Then
                strJyo = JyoCord(mO1Data.id.JyoCD)
                If IsEmpty(strJyo) Then GoTo LOOP_NEXT1
                If strJyo = placeTarg And _
                    Val(mO1Data.id.racenum) = racenumTarg Then
                    .Cells(cnt, 1) = buff
                    cnt = cnt + 1
                    codeLength = 4100
                End If
            End If
        ElseIf Left(buff, 2) = "O4" Then
            Call SetData_O4(buff, mO4Data)
            If mO4Data.id.Year & mO4Data.id.MonthDay = strdateTarg Then
                strJyo = JyoCord(mO4Data.id.JyoCD)
                If IsEmpty(strJyo) Then GoTo LOOP_NEXT1
                If strJyo = placeTarg And _
                    Val(mO4Data.id.racenum) = racenumTarg Then
                    .Cells(cnt, 1) = buff
                    cnt = cnt + 1
                    codeLength = 84000
                End If
            End If
        ElseIf Left(buff, 2) = "O6" Then
            Call SetData_O6(buff, mO6Data(0))
            ' 選んだ日付を超えたらループ抜ける
'            If isGetData = True And Val(mO6Data(0).id.Year & mO6Data(0).id.MonthDay) > Val(strdateTarg) Then
'                GoTo LOOP_END2
'            End If
            If mO6Data(0).id.Year & mO6Data(0).id.MonthDay = strdateTarg Then
                strJyo = JyoCord(mO6Data(0).id.JyoCD)
                If IsEmpty(strJyo) Then GoTo LOOP_NEXT1
                If strJyo = placeTarg And _
                    Val(mO6Data(0).id.racenum) = racenumTarg Then
                    isGetData = True
                    .Cells(cnt, 1) = buff
                    cnt = cnt + 1
                End If
            End If
        Else
            UserForm1.JVLink1.JVSkip
        End If
        DoEvents
LOOP_NEXT1:
    Loop
LOOP_END2:
    End With
    Exit Function

LOOP_END1:
ERROR_PROCESS:
    Debug.Print "15:" & Err.Description
    MsgBox "エラー " & retval
End Function

Sub TEST003()
    Beep
End Sub

Sub TEST004()
    targdate = "20190425"
    Debug.Print Date
    Debug.Print targdate
    
    dateTarg = CDate(Format(targdate, "####/##/##"))
    If Date < dateTarg Then
        Debug.Print targdate
    End If
    Debug.Print DateDiff("d", dateTarg, Date)
End Sub

Sub TEST005()
Dim retval As Long
    Dim readcount As Long
    Dim dlcount As Long
    Dim lastfiletimestamp As String
    Dim status As Long
    Dim buff As String
    Dim buff2 As String
    Dim filename As String
    
    UserForm1.JVLink1.JVClose
    UserForm1.JVLink1.JVInit ("EXCELSAMPLE")
    strdate = Val(Format(Date, "yyyy")) - 4 & Format(Date, "mmdd")
    retval = UserForm1.JVLink1.JVOpen("YSCH", strdate & "000000", 4, _
    readcount, dlcount, lastfiletimestamp)
    
    hwnd1 = FindWindow("セットアップ", vbNullString)
    hwnd2 = FindWindowEx(hwnd1, 0, "TPanel", vbNullString)
    hwnd3 = FindWindowEx(hwnd2, 0, "TButton", "OK")
    ' OKをクリック
    ReturnSM1 = SendMessage(hwnd3, BM_CLICK, 0, 0)
End Sub
