Attribute VB_Name = "Module_UmatanOddsCalc"
Sub CreateCompositeOdds(datacsv As datacsv, _
                        collUmatanOddsO1 As Collection, _
                        collUmatanOddsH1 As Collection, _
                        collUmatanOdds As Collection, _
                        collOddsSanrentan As Collection, _
                        isCalcSanrentan As Boolean)
    Dim strExpectTime As String
    If collOddsSanrentan.Count >= 4896 Then
        strExpectTime = "120"
    ElseIf collOddsSanrentan.Count >= 4080 Then
        strExpectTime = "60"
    ElseIf collOddsSanrentan.Count >= 3360 Then
        strExpectTime = "40"
    ElseIf collOddsSanrentan.Count >= 2730 Then
        strExpectTime = "30"
    ElseIf collOddsSanrentan.Count >= 2184 Then
        strExpectTime = "20"
    Else
        strExpectTime = "20"
    End If
    UserForm_Wait.Label2.Caption = "3連単の数:" & collOddsSanrentan.Count & _
                                    " 推定計算時間約:" & strExpectTime & "秒"
    DoEvents
    
    Dim collumatanoddO1 As RaceUma
    Set collumatanoddO1 = New RaceUma
    Dim collumatanoddH1 As UmatanOdds
    Set collumatanoddH1 = New UmatanOdds
    Dim collUmatanOdd As UmatanOdds
    Set collUmatanOdd = New UmatanOdds
    Dim collOddsGousei As Collection
    Dim cnt As Long
    Dim i As Long
    Dim j As Long
    Dim rowWrite As Long
    Dim denom As Double
    
    ' オッズ（単複枠）
    For Each collumatanoddO1 In collUmatanOddsO1
        cnt = 1
        For Each collUmatanOdd In collUmatanOdds
            If collUmatanOdd.Umaban1 = collumatanoddO1.Umaban Then
                collUmatanOdds.item(cnt).Ninki1 = collumatanoddO1.Ninki
            End If
            If collUmatanOdd.Umaban2 = collumatanoddO1.Umaban Then
                collUmatanOdds.item(cnt).Ninki2 = collumatanoddO1.Ninki
            End If
            cnt = cnt + 1
        Next
    Next
    
    ' 票数1
    For Each collumatanoddH1 In collUmatanOddsH1
        cnt = 1
        For Each collUmatanOdd In collUmatanOdds
            If collUmatanOdd.Kumi = collumatanoddH1.Kumi Then
                collUmatanOdds.item(cnt).Hyou = Val(collumatanoddH1.Hyou)
                Exit For
            End If
            cnt = cnt + 1
        Next
    Next
    
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
    For i = 1 To collUmatanOdds.Count
        Set collOddsGousei = New Collection
        collOddsGousei.Add collUmatanOdds.item(i).Odds
        collOddsGousei.Add collUmatanOdds.item(i).RevOdds
        denom = 0
        For j = 1 To collOddsGousei.Count
            denom = denom + 1 / Val(collOddsGousei.item(j))
        Next
        collUmatanOdds.item(i).SyntheticOdds1 = Format(1 / denom, "0.0")
    Next i
    
    ' ３連単オッズ
    If isCalcSanrentan Then
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
    
End Sub
