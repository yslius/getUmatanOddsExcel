Attribute VB_Name = "Module_UmatanOdds"
Option Explicit

Sub getUmatanOdds(strdateTarg As String, placeTarg As String, racenumTarg As Integer, _
                    isCalcSanrentan As Boolean)
    Dim pathTarg As String
    
    Call initArray
    Set WBbase = ThisWorkbook
    Set WSbase = WBbase.Sheets("TOP")
    Set WStarg = WBbase.Sheets("Sheet1")
    
    ' 3連単計算するか
    Debug.Print isCalcSanrentan
    
    Dim pathfileTarg As String
    pathTarg = WBbase.Sheets(1).OLEObjects("TextBox1").Object.Text
    If pathTarg = "" Then
        MsgBox "出力ファイルを保存するパスを選択してください。"
        Exit Sub
    End If
    
    ' 出力ファイルを作る
    Dim datacsv As datacsv
    Set datacsv = New datacsv
    datacsv.setIniData = arrAddHead
    ' シート1のデータ削除
    WStarg.Range(WStarg.Cells(5, 1), WStarg.Cells(Rows.Count, 9)).ClearContents
    
    pathfileTarg = pathTarg & "\馬単オッズ_" & strdateTarg & "_" & _
                    placeTarg & StrConv(Format(racenumTarg, "00"), vbWide) & ".csv"
    
    UserForm_Wait.Show vbModeless
    UserForm_Wait.Label1.Caption = strdateTarg & " " & placeTarg & " " & _
                                    StrConv(Format(racenumTarg, "00"), vbWide) & _
                                    vbCrLf & " 馬単オッズ取得中です。"
                       
    ' 速報開催情報(一括)の呼び出し
    'JVLinkを初期化
    UserForm1.JVLink1.JVClose
    UserForm1.JVLink1.JVInit ("EXCELSAMPLE")
    Dim retval As Long
    retval = UserForm1.JVLink1.JVRTOpen("0B14", strdateTarg)
    If (retval < -1) Then GoTo ERROR_PROCESS
    
    If retval = -1 Then
        UserForm1.JVLink1.JVClose
        If Not isFindDate(strdateTarg, placeTarg, racenumTarg) Then
            MsgBox "開催されていません。"
            Unload UserForm_Wait
            Exit Sub
        End If
        UserForm1.JVLink1.JVClose
        UserForm1.JVLink1.JVInit ("EXCELSAMPLE")
        Call GetStockUmatanOdds(datacsv, strdateTarg, placeTarg, racenumTarg, isCalcSanrentan)
    Else
        Call GetRealTimeUmatanOdds(datacsv, strdateTarg, placeTarg, racenumTarg, isCalcSanrentan)
    End If
    
END_PROCESS:
    UserForm1.JVLink1.JVClose
    Unload UserForm_Wait
    ' ファイル出力
    Dim rangeTmp As Range
    Dim CSVData As String
    
    datacsv.pasteData(ThisWorkbook.Sheets("Output")) = 1
    'ソート
    With ThisWorkbook.Sheets("Output")
    Set rangeTmp = .Cells(1, 1).CurrentRegion
'    Debug.Print rangeTmp.Rows.Count
'    Debug.Print rangeTmp.Columns.Count
    Call .Range(.Cells(2, 1), .Cells(rangeTmp.Rows.Count, rangeTmp.Columns.Count)).Sort( _
    key1:=.Cells(1, 3), _
    Order1:=xlAscending)
    Set rangeTmp = .Range(.Cells(2, 1), .Cells(rangeTmp.Rows.Count, rangeTmp.Columns.Count))
    End With
    
    ' ファイル出力
    With ThisWorkbook.Sheets("Output")
    CSVData = CreateCSVData(ThisWorkbook.Sheets("Output"))
    Open pathfileTarg For Output As #1
        Print #1, CSVData
    Close #1
    
    ' Sheet1に反映
    With ThisWorkbook.Sheets("Sheet1")
    rangeTmp.Copy .Cells(5, 1)
    End With
    
    ThisWorkbook.Sheets("Output").Cells.Clear
    End With
                        
    Beep
    MsgBox "終了しました。"
    Exit Sub
    
ERROR_PROCESS:
    Beep
    Unload UserForm_Wait
    UserForm1.JVLink1.JVClose
    Debug.Print "12:" & Err.Description
    MsgBox "エラー " & retval
    
End Sub

