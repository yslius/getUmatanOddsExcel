Attribute VB_Name = "Module_UmatanRealtimeOddsDetail"
Sub GetRealTimeUmatanOdds(datacsv As datacsv, strdateTarg As String, _
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
    
    Dim codeJyo As String
    Dim numRace As String
    Dim i As Long
    Dim j As Long
    Dim collUmatanOdds As Collection
    Set collUmatanOdds = New Collection
    Dim collUmatanOddsO1 As Collection
    Set collUmatanOddsO1 = New Collection
    Dim collUmatanOddsH1 As Collection
    Set collUmatanOddsH1 = New Collection
    Dim collOddsSanrentan As Collection
    Set collOddsSanrentan = New Collection
    
    Dim rowTarget As Long
    rowTarget = 2
    
    '場コードの特定
    codeJyo = JyogyakuCord(placeTarg)
    'レース番号
    numRace = Format(racenumTarg, "00")
 
    Dim numUma As Integer
    Dim cnt As Long
    
    '速報オッズ（馬単）の呼び出し
    UserForm1.JVLink1.JVClose
    UserForm1.JVLink1.JVInit ("EXCELSAMPLE")
    retval = UserForm1.JVLink1.JVRTOpen("0B34", strdateTarg & codeJyo & numRace)
    If (retval <= -1) Then GoTo ERROR_PROCESS
    retval = UserForm1.JVLink1.JVRead(buff, 4100, filename)
    If (retval < -1) Then GoTo ERROR_PROCESS
    ' オッズ（馬単）
    Dim mO4Data As JV_O4_ODDS_UMATAN
    Call SetData_O4(buff, mO4Data)
    
    Dim numHyou As Long
    Dim UmatanOdd As UmatanOdds
    numUma = Val(mO4Data.SyussoTosu)
    If numUma = 0 Then numUma = Val(mO4Data.TorokuTosu)
    
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


    '速報オッズ（単複枠）の呼び出し
    UserForm1.JVLink1.JVClose
    UserForm1.JVLink1.JVInit ("EXCELSAMPLE")
    retval = UserForm1.JVLink1.JVRTOpen("0B31", strdateTarg & codeJyo & numRace)
    If (retval <= -1) Then GoTo ERROR_PROCESS
    retval = UserForm1.JVLink1.JVRead(buff, 1000, filename)
    If (retval < -1) Then GoTo ERROR_PROCESS
    ' オッズ（単複枠）
    Dim mO1Data As JV_O1_ODDS_TANFUKUWAKU
    Call SetData_O1(buff, mO1Data)
    If numUma <> Val(mO1Data.SyussoTosu) Then

    End If
    
    ' 人気順を入れる
    Dim collUmatanOdd As UmatanOdds
    For i = 1 To Val(mO4Data.TorokuTosu)
        cnt = 1
        For Each collUmatanOdd In collUmatanOdds
            If Val(collUmatanOdd.Umaban1) = Val(mO1Data.OddsTansyoInfo(i - 1).Umaban) Then
                collUmatanOdds.item(cnt).Ninki1 = Val(mO1Data.OddsTansyoInfo(i - 1).Ninki)
            End If
            If Val(collUmatanOdd.Umaban2) = Val(mO1Data.OddsTansyoInfo(i - 1).Umaban) Then
                collUmatanOdds.item(cnt).Ninki2 = Val(mO1Data.OddsTansyoInfo(i - 1).Ninki)
            End If
            cnt = cnt + 1
        Next
    Next i


    '速報票数(全賭式)の呼び出し
    UserForm1.JVLink1.JVClose
    UserForm1.JVLink1.JVInit ("EXCELSAMPLE")
    retval = UserForm1.JVLink1.JVRTOpen("0B20", strdateTarg & codeJyo & numRace)
    If (retval = -1) Then GoTo SKIPHYOU
    If (retval <= -1) Then GoTo ERROR_PROCESS
    retval = UserForm1.JVLink1.JVRead(buff, 30000, filename)
    If (retval < -1) Then GoTo ERROR_PROCESS
    ' 票数1
    Dim mH1Data As JV_H1_HYOSU_ZENKAKE
    Call SetData_H1(buff, mH1Data)
    If numUma <> Val(mH1Data.SyussoTosu) Then

    End If
    For i = 0 To 305 'UBound(mH1Data.HyoUmatan)
        If Trim(mH1Data.HyoUmatan(i).Kumi) <> "" Then
            cnt = 1
            For Each collUmatanOdd In collUmatanOdds
                If collUmatanOdd.Kumi = mH1Data.HyoUmatan(i).Kumi Then
                    collUmatanOdds.item(cnt).Hyou = Val(mH1Data.HyoUmatan(i).Hyo)
                    Exit For
                End If
                cnt = cnt + 1
            Next
        End If
    Next i
SKIPHYOU:
    
    ' 馬単裏
    For i = 1 To collUmatanOdds.Count
        For Each collUmatanOdd In collUmatanOdds
            If collUmatanOdds.item(i).Umaban1 = collUmatanOdd.Umaban2 And _
                collUmatanOdds.item(i).Umaban2 = collUmatanOdd.Umaban1 Then
                collUmatanOdds.item(i).RevOdds = collUmatanOdd.Odds
                Exit For
            End If
        Next
    Next i
    
    ' 馬単合成
    Dim collOddsGousei As Collection
    Dim denom As Double
    For i = 1 To collUmatanOdds.Count
        Set collOddsGousei = New Collection
        If collUmatanOdds.item(i).Odds <> "" Then
            collOddsGousei.Add collUmatanOdds.item(i).Odds
        End If
        If collUmatanOdds.item(i).RevOdds <> "" Then
            collOddsGousei.Add collUmatanOdds.item(i).RevOdds
        End If
        denom = 0
        For j = 1 To collOddsGousei.Count
            denom = denom + 1 / Val(collOddsGousei.item(j))
        Next
        collUmatanOdds.item(i).SyntheticOdds1 = Format(1 / denom, "0.0")
    Next i
    
    If isCalcSanrentan Then
        ' 3連単1・2着軸総流し
        '３連単オッズの呼び出し
        UserForm1.JVLink1.JVClose
        UserForm1.JVLink1.JVInit ("EXCELSAMPLE")
        retval = UserForm1.JVLink1.JVRTOpen("0B36", strdateTarg & codeJyo & numRace)
        If (retval <= -1) Then GoTo ERROR_PROCESS
        retval = UserForm1.JVLink1.JVRead(buff, 110000, filename)
        If (retval < -1) Then GoTo ERROR_PROCESS
        ' ３連単オッズ
        Dim mO6Data As JV_O6_ODDS_SANRENTAN2
    '    ReDim mO6Data(0)
    '    Call SetData_O6(buff, mO6Data(0))
        Call SetData_O6Z(buff, mO6Data)
        
        Dim cOddsSanrentan As OddsSanrentan
        
        For Each OddsSanrentanInfo In mO6Data.OddsSanrentanInfo
            If Trim(OddsSanrentanInfo.Kumi) <> "" And _
               Trim(OddsSanrentanInfo.Odds <> "") And _
                Val(OddsSanrentanInfo.Odds) <> 0 And _
                InStr(OddsSanrentanInfo.Odds, "-") = 0 And _
                InStr(OddsSanrentanInfo.Odds, "*") = 0 Then
                Set cOddsSanrentan = New OddsSanrentan
                cOddsSanrentan.Kumi = OddsSanrentanInfo.Kumi
                cOddsSanrentan.Umaban1 = Val(Left(OddsSanrentanInfo.Kumi, 2))
                cOddsSanrentan.Umaban2 = Val(Mid(OddsSanrentanInfo.Kumi, 3, 2))
                cOddsSanrentan.Umaban3 = Val(Right(OddsSanrentanInfo.Kumi, 2))
                cOddsSanrentan.OddsSanrentan = Format(Val(OddsSanrentanInfo.Odds / 10), "0.0")
                collOddsSanrentan.Add cOddsSanrentan
                Set cOddsSanrentan = Nothing
            End If
        Next
        
        Debug.Print collOddsSanrentan.Count
        
        For i = 1 To collUmatanOdds.Count
            Set collOddsGousei = New Collection
            For j = 1 To collOddsSanrentan.Count
                If collUmatanOdds.item(i).Umaban1 = collOddsSanrentan.item(j).Umaban1 And _
                    collUmatanOdds.item(i).Umaban2 = collOddsSanrentan.item(j).Umaban2 Then
                    collOddsGousei.Add collOddsSanrentan.item(j).OddsSanrentan
                End If
            Next
            denom = 0
            For j = 1 To collOddsGousei.Count
                If Val(collOddsGousei.item(j)) <> 0 Then
                    denom = denom + 1 / Val(collOddsGousei.item(j))
                End If
            Next
            If denom > 0 Then collUmatanOdds.item(i).SyntheticOdds2 = Format(1 / denom, "0.0")
            DoEvents
        Next i
    End If
    
    
    Dim rowWrite As Long
    rowWrite = 2
    For Each collUmatanOdd In collUmatanOdds
        datacsv.setData(rowWrite, 1) = collUmatanOdd.Umaban1
        datacsv.setData(rowWrite, 2) = collUmatanOdd.Umaban2
        datacsv.setData(rowWrite, 3) = collUmatanOdd.Odds
        datacsv.setData(rowWrite, 4) = collUmatanOdd.Ninki1
        datacsv.setData(rowWrite, 5) = collUmatanOdd.Ninki2
        datacsv.setData(rowWrite, 6) = collUmatanOdd.Hyou
        datacsv.setData(rowWrite, 7) = collUmatanOdd.RevOdds
        datacsv.setData(rowWrite, 8) = collUmatanOdd.SyntheticOdds1
        If isCalcSanrentan Then
            datacsv.setData(rowWrite, 9) = collUmatanOdd.SyntheticOdds2
        End If
        rowWrite = rowWrite + 1
    Next
    
LOOP_END1:
    Exit Sub
    
ERROR_PROCESS:
    Debug.Print "13:" & Err.Description
    MsgBox "エラー " & retval
    
End Sub



