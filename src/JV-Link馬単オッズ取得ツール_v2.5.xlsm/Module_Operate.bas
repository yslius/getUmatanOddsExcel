Attribute VB_Name = "Module_Operate"
Sub GetDate()
    Dim retval As Long
    Dim readcount As Long
    Dim dlcount As Long
    Dim lastfiletimestamp As String
    Dim status As Long
    Dim buff As String
    Dim buff2 As String
    Dim filename As String
    
    Set WBbase = ThisWorkbook
    Set WSbase = WBbase.Sheets("開催日")
    WSbase.Cells.Clear
    
    UserForm1.Show vbModeless
    
    'JVLinkを初期化
    UserForm1.JVLink1.JVClose
    UserForm1.JVLink1.JVInit ("EXCELSAMPLE")

    ' 開催スケジュール
    Dim mYsData As JV_YS_SCHEDULE

    'セットアップデータの呼び出し
    strdate = Val(Format(Date, "yyyy")) - 4 & Format(Date, "mmdd")
    retval = UserForm1.JVLink1.JVOpen("YSCH", strdate & "000000", 4, readcount, dlcount, lastfiletimestamp)
    If (retval < -1) Then
        Debug.Print "17:" & Err.Description
        MsgBox ("JVOpenエラー。RC=" & retval)
        GoTo CommandButton1_END
    End If
    
    cnt = 1
    retval = 1
    While retval <> 0
        'キャンセルボタンチェック
        If Cancelflg = True Then GoTo CommandButton1_END
        
        'JVOpenで指定したデータを１レコードずつ取り込み
        retval = UserForm1.JVLink1.JVRead(buff, 40000, filename)
        ' JVReadエラー処理
        If (retval < -1) Then
            Debug.Print "18:" & Err.Description
            MsgBox ("JVReadエラー。RC=" & retval)
            GoTo CommandButton1_END
        End If
        
        If Left(buff, 2) = "YS" And filename <> "YSMW2020999920200406150936.jvd" Then
            Call SetData_YS(buff, mYsData)
            ' 本日を超えたらループ抜ける
'            If Val(mYsData.id.Year & mYsData.id.MonthDay) > Val(Format(DateAdd("d", 1, Date), "yyyymmdd")) Then
'                GoTo LOOP_END
'            End If
            
            If Val(mYsData.id.JyoCD) <= 10 Then
                rowT = findAlreadyDate(mYsData.id.Year & mYsData.id.MonthDay)
                tmpJyo = JyoCord(mYsData.id.JyoCD)
                
                If rowT = 0 Then
                    WSbase.Cells(cnt, 1) = mYsData.id.Year & mYsData.id.MonthDay
                    WSbase.Cells(cnt, 2) = tmpJyo
                    cnt = cnt + 1
                Else
                    i = 2
                    Do While WSbase.Cells(rowT, i) <> ""
                        If WSbase.Cells(rowT, i) = tmpJyo Then
                            Exit Do
                        End If
                        i = i + 1
                    Loop
                    WSbase.Cells(rowT, i) = tmpJyo
                End If
            End If
            UserForm1.Label1.Caption = buff
            DoEvents
        Else
            UserForm1.JVLink1.JVSkip
        End If
        
        DoEvents
        
    Wend

LOOP_END:
    ' 重複削除と並び替え
    endR = WSbase.Cells(WSbase.Rows.Count, 1).End(xlUp).row
    endC = WSbase.UsedRange.Columns.Count
    WSbase.Range(WSbase.Cells(1, 1), WSbase.Cells(endR, endC)).RemoveDuplicates (Array(1))
    WSbase.Range(WSbase.Cells(1, 1), WSbase.Cells(endR, endC)).Sort _
    key1:=WSbase.Cells(1, 1), Order1:=xlAscending
    For i = 1 To endR
        If Val(WSbase.Cells(i, 1)) > Val(Format(DateAdd("d", 7, Date), "yyyymmdd")) Then
'            WSbase.Cells(i, 1).Clear
            WSbase.Rows(i).Clear
        End If
    Next
    
'    Call UserForm_Initialize

CommandButton1_END:

    '一通り読み込みが終わった後はJVCloseを行う
    UserForm1.JVLink1.JVClose
    If Cancelflg = True Then
        MsgBox "キャンセルされました。"
    Else
        UserForm1.Label1.Caption = "読み込みが終了しました。"
    End If
    DLflg = False
    
    Unload UserForm1

End Sub

Function findAlreadyDate(strdateTarg As String) As Long
    Set WBbase = ThisWorkbook
    Set WSbase = WBbase.Sheets("開催日")
    endR = WSbase.UsedRange.Rows.Count
    For i = 1 To endR
        If strdateTarg = WSbase.Cells(i, 1) Then
            findAlreadyDate = i
            Exit Function
        End If
    Next i
    findAlreadyDate = 0
End Function

Function isFindDate(strdateTarg As String, placeTarg As String, racenumTarg As Integer) As Boolean
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

    ' 開催スケジュール
    Dim mYsData As JV_YS_SCHEDULE
    
    opt = 1
    dateTarg = CDate(Format(strdateTarg, "####/##/##"))
    dateTarg = DateAdd("d", -1, dateTarg)
    If DateDiff("d", dateTarg, Date) >= 365 Then
        opt = 4
    End If

    'セットアップデータの呼び出し
    retval = UserForm1.JVLink1.JVOpen("YSCH", Format(dateTarg, "yyyymmdd") & "000000", opt, readcount, dlcount, lastfiletimestamp)
    If (retval < -1) Then
        Debug.Print "19:" & Err.Description
        MsgBox ("JVOpenエラー。RC=" & retval)
        GoTo CommandButton1_END
    End If
    retval = 1
    While retval <> 0
        retval = UserForm1.JVLink1.JVRead(buff, 40000, filename)
        If (retval < -1) Then
            Debug.Print "1:" & Err.Description
            MsgBox ("JVReadエラー。RC=" & retval)
            GoTo CommandButton1_END
        End If
        If Left(buff, 2) = "YS" Then
            Call SetData_YS(buff, mYsData)
            ' 超えたらループ抜ける
'            If Val(mYsData.id.Year & mYsData.id.MonthDay) > Val(strdateTarg) Then
'                GoTo LOOP_END
'            End If
            
            If strdateTarg = mYsData.id.Year & mYsData.id.MonthDay Then
                strJyo = JyoCord(mYsData.id.JyoCD)
                If IsEmpty(strJyo) Then GoTo LOOP_END
                If strJyo = placeTarg And _
                    (racenumTarg >= 1 And racenumTarg <= 12) Then
                    isFindDate = True
                    GoTo LOOP_END
                End If
            End If
        Else
            UserForm1.JVLink1.JVSkip
        End If
        
        DoEvents
        
    Wend

LOOP_END:

CommandButton1_END:

    UserForm1.JVLink1.JVClose
    
End Function


Sub GetPlaceInfo(targdate)
    Dim retval As Long
    Dim readcount As Long
    Dim dlcount As Long
    Dim lastfiletimestamp As String
    Dim status As Long
    Dim buff As String
    Dim filename As String

    Cancelflg = False
    DLflg = False
    UserForm1.ListBox4.Clear
    
'    If Val(Format(Date, "yyyymmdd")) < Val(targdate) - 1 Then
'        targdate = Format(Date, "yyyymmdd")
'    End If

    'JVLinkを初期化
    UserForm1.JVLink1.JVClose
    UserForm1.JVLink1.JVInit ("EXCELSAMPLE")
    '蓄積系データのRACEをListBoxの日付以降について取り込み呼び出し
    retval = UserForm1.JVLink1.JVOpen("RACE", targdate - 1 & "000000", 1, readcount, dlcount, lastfiletimestamp)
    If (retval < -1) Then
        Debug.Print "2:" & Err.Description
        MsgBox ("JVOpenエラー。RC=" & retval)
        GoTo LOOP_END
    End If

    Dim mRaData As JV_RA_RACE
   
    status = 0
    DLflg = True
    UserForm1.CommandButton3.Caption = "キャンセル"
    Do While status <> dlcount
        'キャンセルボタンチェック
        If Cancelflg = True Then Exit Do
        status = UserForm1.JVLink1.JVStatus
        UserForm1.Label1.Caption = dlcount & "ファイル中 " & status & " ファイルダウンロード完了"
        DoEvents
        Sleep (10)
    Loop
    
    retval = 1
    isGetData = False
    While retval <> 0
        'JVOpenで指定したデータを１レコードずつ取り込み
        retval = UserForm1.JVLink1.JVRead(buff, 40000, filename)
        If (retval < -1) Then
            Debug.Print "3:" & Err.Description
            MsgBox ("JVReadエラー。RC=" & retval)
            GoTo LOOP_END
        End If

        If Left(buff, 2) = "RA" Then

            Call SetData_RA(buff, mRaData)
            
            ' 選んだ日付を超えたらループ抜ける
            If isGetData = True And Val(mRaData.id.Year & mRaData.id.MonthDay) > Val(targdate) Then
                GoTo LOOP_END
            End If
            
            If mRaData.id.Year & mRaData.id.MonthDay = targdate Then
                tmpJyo = JyoCord(mRaData.id.JyoCD)
                If Val(mRaData.id.JyoCD) <= 10 Then
                    isFind = False
                    For i = 0 To UserForm1.ListBox4.ListCount - 1
                        If UserForm1.ListBox4.List(i) = tmpJyo Then
                            isFind = True
                            Exit For
                        End If
                    Next i
                    If isFind = False Then
                        isGetData = True
                        UserForm1.ListBox4.AddItem tmpJyo
                    End If
                End If
            End If

        Else
            UserForm1.JVLink1.JVSkip
        End If
                
        UserForm1.Label1.Caption = buff
        DoEvents
    Wend

LOOP_END:

    UserForm1.JVLink1.JVClose
    
    If Cancelflg = True Then
        MsgBox "キャンセルされました。"
    Else
        UserForm1.Label1.Caption = "読み込みが終了しました。"
    End If
    UserForm1.CommandButton3.Caption = "Exit"
    DLflg = False

End Sub


Sub GetPlaceInfoZ(targdate)
    
    Set WBbase = ThisWorkbook
    Set WSbase = WBbase.Sheets("開催日")
    
    UserForm1.ListBox4.Clear
    endR = WSbase.UsedRange.Rows.Count
    
    For i = 1 To endR
        If Val(targdate) = WSbase.Cells(i, 1) Then
'            j = 2
'            Do While WSbase.Cells(i, j) <> ""
'                UserForm1.ListBox4.AddItem WSbase.Cells(i, j)
'                j = j + 1
'                DoEvents
'            Loop
            For j = 2 To 4
                If WSbase.Cells(i, j) <> "" Then
                    UserForm1.ListBox4.AddItem WSbase.Cells(i, j)
                End If
            Next j
            Exit Sub
        End If
    Next i
    
End Sub

Function GetRaceNumInfo(targdate, targJyo) As Collection
    Dim retval As Long
    Dim readcount As Long
    Dim dlcount As Long
    Dim lastfiletimestamp As String
    Dim status As Long
    Dim buff As String
    Dim filename As String
    Dim opt As Integer
    Dim dateTarg As Date
    Dim colRace As New Collection

    Cancelflg = False
    DLflg = False
    UserForm1.ListBox5.Clear
    
    opt = 1
    dateTarg = CDate(Format(targdate, "####/##/##"))
    ' 日付の変換
     If Date < dateTarg Then
'        targdate = Format(Date - 1, "yyyymmdd")
        dateTarg = DateAdd("d", -1, Date)
    ElseIf DateDiff("d", dateTarg, Date) >= 365 Then
        opt = 4
        dateTarg = DateAdd("d", -1, dateTarg)
    Else
        dateTarg = DateAdd("d", -1, dateTarg)
    End If
    
    'JVLinkを初期化
    UserForm1.JVLink1.JVClose
    UserForm1.JVLink1.JVInit ("EXCELSAMPLE")
    '蓄積系データのRACEをListBoxの日付以降について取り込み呼び出し
    retval = UserForm1.JVLink1.JVOpen("RACE", Format(dateTarg, "yyyymmdd") & "000000", opt, readcount, dlcount, lastfiletimestamp)
    If (retval < -1) Then
        Debug.Print "4:" & Err.Description
        MsgBox ("JVOpenエラー。RC=" & retval)
        GoTo CommandButton1_END
    End If

    Dim mRaData As JV_RA_RACE
   
    status = 0
    DLflg = True
    UserForm1.CommandButton3.Caption = "キャンセル"
    Do While status <> dlcount
        Debug.Print "discount=" & CStr(discount)
        'キャンセルボタンチェック
        If Cancelflg = True Then Exit Do
        status = UserForm1.JVLink1.JVStatus
        UserForm1.Label1.Caption = dlcount & "ファイル中 " & status & " ファイルダウンロード完了"
        DoEvents
        Sleep (10)
    Loop
    
    retval = 1
    isIndata = False
    While retval <> 0
        'キャンセルボタンチェック
        If Cancelflg = True Then GoTo CommandButton1_END
         
        'JVOpenで指定したデータを１レコードずつ取り込み
        retval = UserForm1.JVLink1.JVRead(buff, 40000, filename)
        If (retval < -1) Then
            Debug.Print "5:" & Err.Description
            MsgBox ("JVReadエラー。RC=" & retval)
            GoTo CommandButton1_END
        End If
        
        If Left(buff, 2) = "RA" Then
            Call SetData_RA(buff, mRaData)
            
            If isIndata = True And Val(mRaData.id.Year & mRaData.id.MonthDay) > Val(targdate) Then
                GoTo LOOP_END
            End If
            
            If mRaData.id.Year & mRaData.id.MonthDay = targdate And _
               Val(mRaData.head.DataKubun) <> 9 Then
                tmpJyo = JyoCord(mRaData.id.JyoCD)
                If tmpJyo = targJyo Then
                    isIndata = True
                    UserForm1.ListBox5.AddItem mRaData.id.racenum
                    colRace.Add mRaData.id.racenum
                End If
            End If
        Else
            UserForm1.JVLink1.JVSkip
        End If
        UserForm1.Label1.Caption = buff
        DoEvents
    Wend

CommandButton1_END:
LOOP_END:

    UserForm1.JVLink1.JVClose
    If Cancelflg = True Then
        MsgBox "キャンセルされました。"
    Else
        UserForm1.Label1.Caption = "読み込みが終了しました。"
    End If
    UserForm1.CommandButton3.Caption = "Exit"
    DLflg = False
    Set GetRaceNumInfo = colRace

End Function

Sub GetRaceUma(targdate, targJyo, racenum)
    Set WBbase = ThisWorkbook
    Set WSbase = WBbase.Sheets("レース")
    WSbase.Range(WSbase.Rows(2), WSbase.Rows(WSbase.Rows.Count)).ClearContents
    
'    targdate = 20190112
'    targJyo = "中山"
'    RaceNum = 1
    
    Dim retval As Long
    Dim readcount As Long
    Dim dlcount As Long
    Dim lastfiletimestamp As String
    Dim status As Long
    Dim buff As String
    Dim filename As String
    Dim opt As Integer
    Dim dateTarg As Date
    
    opt = 1
    dateTarg = CDate(Format(targdate, "####/##/##"))
    ' 日付の変換
     If Date < dateTarg Then
'        targdate = Format(Date - 1, "yyyymmdd")
        dateTarg = DateAdd("d", -1, Date)
    ElseIf DateDiff("d", dateTarg, Date) >= 365 Then
        opt = 4
        dateTarg = DateAdd("d", -1, dateTarg)
    Else
        dateTarg = DateAdd("d", -1, dateTarg)
    End If
    
    'JVLinkを初期化
    UserForm1.JVLink1.JVClose
    UserForm1.JVLink1.JVInit ("EXCELSAMPLE")
    
    '蓄積系データの呼び出し
    retval = UserForm1.JVLink1.JVOpen("RACE", targdate - 1 & "000000", opt, readcount, dlcount, lastfiletimestamp)
    'JVOpenエラー処理
    If (retval < -1) Then
        Debug.Print "6:" & Err.Description
        MsgBox ("JVOpenエラー " & retval)
        GoTo CommandButton1_END
    End If
    
    ' 馬毎レース情報
    Dim mSeData As JV_SE_RACE_UMA
    
    'ファイルのダウンロード
    status = 0
    DLflg = True
    Do While status <> dlcount
        'キャンセルボタンチェック
        If Cancelflg = True Then Exit Do
        status = UserForm1.JVLink1.JVStatus
        UserForm1.Label1.Caption = dlcount & "ファイル中 " & status & " ファイルダウンロード完了"
        DoEvents
        Sleep (120)
'        Debug.Print "status:" & status
    Loop
    
    Cancelflg = False
    retval = 1
    cnt = 1
    cntWhile = 1
    isIntoSE = False
    While retval <> 0
        'キャンセルボタンチェック
        If Cancelflg = True Then GoTo CommandButton1_END

        'JVOpenで指定したデータを１レコードずつ取り込み
        retval = UserForm1.JVLink1.JVRead(buff, 40000, filename)
        Debug.Print "retval:" & retval
        ' JVReadエラー処理
        If (retval < -1) Then
            Debug.Print "7:" & Err.Description
            MsgBox ("JVReadエラー。RC=" & retval)
            GoTo CommandButton1_END
        End If
        If Left(buff, 2) = "SE" Then
            'JVData構造体にSEのレコードをセットする
            Call SetData_SE(buff, mSeData)
            If isIntoSE = True And Val(mSeData.id.Year & mSeData.id.MonthDay) > Val(targdate) Then
                GoTo CommandButton1_END
            End If
            If targdate = mSeData.id.Year & mSeData.id.MonthDay And _
               targJyo = JyoCord(mSeData.id.JyoCD) And _
               racenum = Val(mSeData.id.racenum) Then
                isIntoSE = True
                WSbase.Cells(cnt, 1) = mSeData.id.Year & mSeData.id.MonthDay
                WSbase.Cells(cnt, 2) = JyoCord(mSeData.id.JyoCD)
                WSbase.Cells(cnt, 3) = mSeData.id.racenum
                WSbase.Cells(cnt, 4) = mSeData.Umaban
                WSbase.Cells(cnt, 5) = mSeData.Bamei
                WSbase.Cells(cnt, 6) = mSeData.KettoNum
                
                UserForm1.ListBox6.AddItem mSeData.Umaban & " " & Trim(mSeData.Bamei)
                cnt = cnt + 1
            End If
        Else
            UserForm1.JVLink1.JVSkip
        End If
        UserForm1.Label1.Caption = buff
        
        cntWhile = cntWhile + 1
'        Debug.Print "cntWhile:" & cntWhile
        DoEvents
    Wend
    
    
    UserForm1.JVLink1.JVClose
    
CommandButton1_END:
    UserForm1.JVLink1.JVClose
    
    UserForm1.CommandButton5.Enabled = True
End Sub

Function CreateCSVData(WS) As String
    Dim CSVData As String
    With WS
    
    Dim region As Range
    Set region = .Cells(1, 1).CurrentRegion ' データの範囲を自動取得
    
    Dim row As Range
    For i = 1 To region.Rows.Count  ' 行のループ
        Line = ""
        For j = 1 To region.Columns.Count ' 列のループ
            ' カンマ区切りで結合
            Dim item As Variant
            item = .Cells(i, j).Value
            If Line = "" Then
                Line = item
            Else
                If Not IsError(item) Then
                    Line = Line & "," & item
                Else
                    Line = Line & "," & .Cells(i, j).Text
                End If
            End If
        Next
        ' 行を結合
        If CSVData = "" Then
            CSVData = Line
        Else
            CSVData = CSVData & vbCrLf & Line
        End If
    Next
    
    CreateCSVData = CSVData
    
    End With
End Function


Function CreateCSVDataZ(WS) As String
    Dim CSVData As String
    With WS
    
    Dim region As Range
    Dim row As Range
    Set region = .Cells(1, 1).CurrentRegion
    
    For i = 3 To region.Rows.Count  ' 行のループ
        Line = ""
        For j = 1 To 12 ' 列のループ
            ' カンマ区切りで結合
            Dim item As Variant
            item = .Cells(i, j).Value
            If Line = "" Then
                Line = item
            Else
                Line = Line & "," & item
            End If
        Next
        ' 行を結合
        If CSVData = "" Then
            CSVData = Line
        Else
            CSVData = CSVData & vbCrLf & Line
        End If
    Next
    
    CreateCSVDataZ = CSVData
    
    End With
End Function


Sub OutputLog(datalog)
    fileLog = ThisWorkbook.Path & "\" & "log_" & Format(Date, "yyyymmdd") & ".txt"
    
    Open fileLog For Append As #1
        Print #1, datalog
    Close #1
End Sub


Function getrowT(WStarg, strShortJyo, racenum) As Integer
    With WStarg
        rowEnd = .Cells(.Rows.Count, 1).End(xlUp).row
        For i = 1 To rowEnd
            If InStr(.Cells(i, 3), strShortJyo) And _
               .Cells(i, 6) = racenum Then
               getrowT = i
               Exit Function
            End If
        Next i
    End With
End Function



