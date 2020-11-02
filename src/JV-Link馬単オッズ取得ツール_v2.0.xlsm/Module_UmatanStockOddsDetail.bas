Attribute VB_Name = "Module_UmatanStockOddsDetail"
Option Explicit

Sub GetStockUmatanOdds(datacsv As datacsv, strdateTarg As String, _
                        placeTarg As String, racenumTarg As Integer, _
                        isCalcSanrentan As Boolean)
    ' JVLink関連の変数セット
    Dim retval As Long
    Dim readcount As Long
    Dim dlcount As Long
    Dim lastfiletimestamp As String
    Dim status As Long
    Dim buff As String
    Dim filename As String
    Dim codeLength As Long
    Dim Cancelflg As Boolean
    Dim isGetData As Boolean
    Dim opt As Integer
    Dim dateTarg As Date
    
    Dim strJyo As String
    Dim i As Long
    
    opt = 1
    dateTarg = CDate(Format(strdateTarg, "####/##/##"))
    dateTarg = DateAdd("d", -1, dateTarg)
    If DateDiff("d", dateTarg, Date) >= 365 Then
        opt = 4
    End If
    
    retval = UserForm1.JVLink1.JVOpen("RACE", Format(dateTarg, "yyyymmdd") & "000000", opt, readcount, dlcount, lastfiletimestamp)
    If (retval = -1) Then GoTo LOOP_END1
    If (retval < -1) Then GoTo ERROR_PROCESS
    Do While status <> dlcount
        If Cancelflg = True Then Exit Do
        status = UserForm1.JVLink1.JVStatus
        UserForm_Wait.Label1.Caption = dlcount & "ファイル中 " & status & " ファイルダウンロード完了"
        DoEvents
        Sleep (10)
    Loop
    
    retval = 1
    isGetData = False
    codeLength = 840000
    Dim collOutput As Collection
    Set collOutput = New Collection
    
    Do While retval <> 0
        retval = UserForm1.JVLink1.JVRead(buff, codeLength, filename)
        If (retval < -1) Then GoTo ERROR_PROCESS
'        If isGetData = False Then GoTo LOOP_END2
        If Left(buff, 2) = "H1" Then
            If isGetData And Mid(buff, 12, 8) > strdateTarg Then
                GoTo LOOP_END2
            End If
            If Mid(buff, 12, 8) = strdateTarg Then
                collOutput.Add (buff)
            End If
        ElseIf Left(buff, 2) = "O1" Then
            If Mid(buff, 12, 8) = strdateTarg Then
                collOutput.Add (buff)
            End If
        ElseIf Left(buff, 2) = "O4" Then
            If Mid(buff, 12, 8) = strdateTarg Then
                collOutput.Add (buff)
            End If
        ElseIf isCalcSanrentan And Left(buff, 2) = "O6" Then
            If Mid(buff, 12, 8) = strdateTarg Then
                collOutput.Add (buff)
                isGetData = True
            End If
        Else
            UserForm1.JVLink1.JVSkip
        End If
        DoEvents
LOOP_NEXT1:
    Loop
    
LOOP_END2:

    Dim strOutputFile As String
    Dim strOutput As Variant
    strOutputFile = ThisWorkbook.Path & "\buff.txt"
    Open strOutputFile For Output As #1
    For Each strOutput In collOutput
        Print #1, strOutput
    Next
    Close #1
    
    Call GetStockUmatanOddsText(datacsv, strdateTarg, placeTarg, racenumTarg, strOutputFile, isCalcSanrentan)
    
    Kill strOutputFile
    
    Exit Sub

LOOP_END1:
ERROR_PROCESS:
    MsgBox "エラー " & retval
    
End Sub


Sub GetStockUmatanOddsText(datacsv As datacsv, strdateTarg As String, _
                        placeTarg As String, racenumTarg As Integer, _
                        strOutputFile As String, isCalcSanrentan As Boolean)
    
    Dim strJyo As String
    Dim i As Long
    Dim collUmatanOdds As Collection
    Set collUmatanOdds = New Collection
    Dim collUmatanOddsO1 As Collection
    Set collUmatanOddsO1 = New Collection
    Dim collUmatanOddsH1 As Collection
    Set collUmatanOddsH1 = New Collection
    Dim collOddsSanrentan As Collection
    Set collOddsSanrentan = New Collection
    Dim OddsSanrentanInfo As Variant
    
    ' オッズ（馬単）
    Dim mO4Data As JV_O4_ODDS_UMATAN
    ' オッズ（単複枠）
    Dim mO1Data As JV_O1_ODDS_TANFUKUWAKU
    ' オッズ（３連単）
    Dim mO6Data As JV_O6_ODDS_SANRENTAN2
    ' 票数1
    Dim mH1Data As JV_H1_HYOSU_ZENKAKE
    
    Dim UmatanOdd As UmatanOdds
    Dim RaceUm As RaceUma
    Dim OddsSanrenta As OddsSanrentan
    Dim buff As String
    Dim strOutput As String
    Dim strOutputFile2 As String
    
    strOutputFile2 = ThisWorkbook.Path & "\OddsSanrenta.txt"
    Open strOutputFile For Input As #1
    Open strOutputFile2 For Append As #2
'    Do While Not EOF(1)
'        Line Input #1, buff
''        Debug.Print Trim(buff)
'    Loop
'    Close #1
    
    Do While Not EOF(1)
        Line Input #1, buff
        If Left(buff, 2) = "H1" Then
            Call SetData_H1(buff, mH1Data)
            strJyo = JyoCord(mH1Data.id.JyoCD)
            If strJyo = placeTarg And _
                Val(mH1Data.id.racenum) = racenumTarg Then
                For i = 0 To 305 'UBound(mO4Data.OddsUmatanInfo)
                    If Trim(mH1Data.HyoUmatan(i).Kumi) <> "" Then
                        Set UmatanOdd = New UmatanOdds
                        UmatanOdd.Kumi = mH1Data.HyoUmatan(i).Kumi
                        UmatanOdd.Hyou = Val(mH1Data.HyoUmatan(i).Hyo)
                        collUmatanOddsH1.Add UmatanOdd
                        Set UmatanOdd = Nothing
                    End If
                Next i
            End If
        ElseIf Left(buff, 2) = "O1" Then
            Call SetData_O1(buff, mO1Data)
            strJyo = JyoCord(mO1Data.id.JyoCD)
            If strJyo = placeTarg And _
                Val(mO1Data.id.racenum) = racenumTarg Then
                For i = 0 To 18  'UBound(mO1Data.OddsTansyoInfo)
                    If Trim(mO1Data.OddsTansyoInfo(i).Umaban) <> "" Then
                        Set RaceUm = New RaceUma
                        RaceUm.Umaban = Val(mO1Data.OddsTansyoInfo(i).Umaban)
                        RaceUm.Ninki = Val(mO1Data.OddsTansyoInfo(i).Ninki)
                        collUmatanOddsO1.Add RaceUm
                        Set RaceUm = Nothing
                    End If
                Next i
            End If
        ElseIf Left(buff, 2) = "O4" Then
            Call SetData_O4(buff, mO4Data)
            strJyo = JyoCord(mO4Data.id.JyoCD)
            If strJyo = placeTarg And _
                Val(mO4Data.id.racenum) = racenumTarg Then
                For i = 0 To 305 'UBound(mO4Data.OddsUmatanInfo)
                    If Trim(mO4Data.OddsUmatanInfo(i).Kumi) <> "" And _
                       Trim(mO4Data.OddsUmatanInfo(i).Odds <> "") And _
                        Val(mO4Data.OddsUmatanInfo(i).Odds) <> 0 And _
                        InStr(mO4Data.OddsUmatanInfo(i).Odds, "-") = 0 And _
                        InStr(mO4Data.OddsUmatanInfo(i).Odds, "*") = 0 Then
                        Set UmatanOdd = New UmatanOdds
                        UmatanOdd.Kumi = mO4Data.OddsUmatanInfo(i).Kumi
                        UmatanOdd.Umaban1 = Val(Left(mO4Data.OddsUmatanInfo(i).Kumi, 2))
                        UmatanOdd.Umaban2 = Val(Right(mO4Data.OddsUmatanInfo(i).Kumi, 2))
                        UmatanOdd.Odds = Format(Val(mO4Data.OddsUmatanInfo(i).Odds / 10), "0.0")
                        collUmatanOdds.Add UmatanOdd
                        Set UmatanOdd = Nothing
                    End If
                Next i
            End If
        ElseIf isCalcSanrentan And Left(buff, 2) = "O6" Then
            Call SetData_O6Z(buff, mO6Data)
            If mO6Data.id.Year & mO6Data.id.MonthDay = strdateTarg Then
                strJyo = JyoCord(mO6Data.id.JyoCD)
                If strJyo = placeTarg And _
                    Val(mO6Data.id.racenum) = racenumTarg Then
'                    Debug.Print (mO6Data.OddsSanrentanInfo.Count)
                    For Each OddsSanrentanInfo In mO6Data.OddsSanrentanInfo
                        If Trim(OddsSanrentanInfo.Kumi) <> "" And _
                           Trim(OddsSanrentanInfo.Odds <> "") And _
                            Val(OddsSanrentanInfo.Odds) <> 0 And _
                            InStr(OddsSanrentanInfo.Odds, "-") = 0 And _
                            InStr(OddsSanrentanInfo.Odds, "*") = 0 Then
                            Set OddsSanrenta = New OddsSanrentan
                            OddsSanrenta.Kumi = OddsSanrentanInfo.Kumi
                            OddsSanrenta.Umaban1 = Val(Left(OddsSanrentanInfo.Kumi, 2))
                            OddsSanrenta.Umaban2 = Val(Mid(OddsSanrentanInfo.Kumi, 3, 2))
                            OddsSanrenta.Umaban3 = Val(Right(OddsSanrentanInfo.Kumi, 2))
                            OddsSanrenta.OddsSanrentan = Format(Val(OddsSanrentanInfo.Odds / 10), "0.0")
                            collOddsSanrentan.Add OddsSanrenta
                            
'                            strOutput = OddsSanrenta.Umaban1 & " " &
'                                        OddsSanrenta.Umaban2 & " " & _
'                                        OddsSanrenta.Umaban3 & " " & _
'                                        OddsSanrenta.OddsSanrentan
'
'                            Print #2, strOutput
                            
                            Set OddsSanrenta = Nothing
                        End If
                    Next
                End If
            End If

        End If

    Loop
    
    Close #1
    Close #2

'    Debug.Print collOddsSanrentan.Count
    UserForm_Wait.Label1.Caption = strdateTarg & " " & placeTarg & " " & _
                                    StrConv(Format(racenumTarg, "00"), vbWide) & _
                                    vbCrLf & " 馬単オッズ計算中です。"
    
    Call CreateCompositeOdds(datacsv, collUmatanOddsO1, _
                            collUmatanOddsH1, collUmatanOdds, _
                            collOddsSanrentan, isCalcSanrentan)
    

End Sub


