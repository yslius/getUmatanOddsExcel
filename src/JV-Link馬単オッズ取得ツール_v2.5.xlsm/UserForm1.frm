VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "JV-Link Get Umatan Odds Tool"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cancelflg As Boolean
Dim DLflg As Boolean
Private isEvents As Boolean


Private Sub UserForm_Initialize()
   
    Set WBbase = ThisWorkbook
    Set WSbase = WBbase.Sheets("開催日")
    endR = WSbase.Cells(WSbase.Rows.Count, 1).End(xlUp).row
    
    ListBox1.Clear
    ListBox2.Clear
    ListBox3.Clear
    ListBox4.Clear
    ListBox5.Clear
    ListBox6.Clear
    
    ' 前回選んだ日付を選択する｡
    Call selectPrevious
    
End Sub

Sub selectPrevious()
    Set WBbase = ThisWorkbook
    Set WSbase = WBbase.Sheets("開催日")
    endR = WSbase.Cells(WSbase.Rows.Count, 1).End(xlUp).row
    isAlready = False
    rowEnd = WSbase.Cells(WSbase.Rows.Count, 1).End(xlUp).row
    For i = 1 To rowEnd
        If WSbase.Cells(i, 5) <> "" Then
            isAlready = True
            Exit For
        End If
    Next
    
    If isAlready Then
'        colT = 2
'        Do While WSbase.Cells(i, colT) <> ""
'            Me.ListBox4.AddItem WSbase.Cells(i, colT)
'            colT = colT + 1
'        Loop
        isEvents = False
        Call getDateYear(Left(WSbase.Cells(i, 1), 4))
        Call getDateMonth(Mid(WSbase.Cells(i, 1), 5, 2))
        Call getDateDate(Mid(WSbase.Cells(i, 1), 7, 2))
        Call getPlacePrevious(i)
        isEvents = True
    Else
        Call getDateYear("")
        isEvents = True
    End If
    
End Sub

Sub getPlacePrevious(i)
    Set WBbase = ThisWorkbook
    Set WSbase = WBbase.Sheets("開催日")
'    colT = 2
'    Do While WSbase.Cells(i, colT) <> ""
'        Me.ListBox4.AddItem WSbase.Cells(i, colT)
'        colT = colT + 1
'    Loop
    For j = 2 To 4
        If WSbase.Cells(i, j) <> "" Then
            Me.ListBox4.AddItem WSbase.Cells(i, j)
        End If
    Next j
    If WSbase.Cells(i, 6) <> "" Then
        Call getPlace(WSbase.Cells(i, 6))
    End If
    colT = 7
    Do While WSbase.Cells(i, colT) <> ""
        Me.ListBox5.AddItem WSbase.Cells(i, colT)
        colT = colT + 1
    Loop
    
End Sub

Sub getPlace(placeT)
    For i = 0 To Me.ListBox4.ListCount - 1
        If placeT = Me.ListBox4.List(i) Then
            Me.ListBox4.Selected(i) = True
            Exit For
        End If
    Next i
End Sub

Sub getDateYear(yearT)
    Set WBbase = ThisWorkbook
    Set WSbase = WBbase.Sheets("開催日")
    endR = WSbase.Cells(WSbase.Rows.Count, 1).End(xlUp).row
    For i = 1 To endR
        Lef = Left(WSbase.Cells(i, 1), 4)
        If tmp = "" Or tmp <> Lef Then
            Me.ListBox1.AddItem Lef
            tmp = Lef
        End If
    Next i
    
    If yearT <> "" Then
        For i = 0 To Me.ListBox1.ListCount - 1
            If yearT = Me.ListBox1.List(i) Then
                Me.ListBox1.Selected(i) = True
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub CommandButton4_Click()
    Call GetDate
End Sub

Private Sub ListBox1_Click()
    If Not isEvents Then
        Exit Sub
    End If
    Me.ListBox2.Locked = True
    Me.ListBox3.Locked = True
    Me.ListBox4.Locked = True
    Me.ListBox5.Locked = True
    Me.ListBox6.Locked = True
    Me.CommandButton5.Locked = True

    Me.ListBox2.Clear
    Me.ListBox3.Clear
    Me.ListBox4.Clear
    Me.ListBox5.Clear
    Me.ListBox6.Clear
    
    Call getDateMonth("")
    
    Me.ListBox2.Locked = False
    Me.ListBox3.Locked = False
    Me.ListBox4.Locked = False
    Me.ListBox5.Locked = False
    Me.ListBox6.Locked = False
    Me.CommandButton5.Locked = False
    
End Sub


Sub getDateMonth(monthT)
    Set WBbase = ThisWorkbook
    Set WSbase = WBbase.Sheets("開催日")
    endR = WSbase.Cells(WSbase.Rows.Count, 1).End(xlUp).row
    tmp = ""
    For i = 1 To endR
        Md1 = Mid(WSbase.Cells(i, 1), 1, 4)
        Md2 = Mid(WSbase.Cells(i, 1), 5, 2)
        If Me.ListBox1.List(Me.ListBox1.ListIndex) = Md1 And tmp <> Md2 Then
            Me.ListBox2.AddItem Md2
            tmp = Md2
        End If
    Next i
    
    If monthT <> "" Then
        For i = 0 To Me.ListBox2.ListCount - 1
            If monthT = Me.ListBox2.List(i) Then
                Me.ListBox2.Selected(i) = True
                Exit For
            End If
        Next i
    End If
    
End Sub


Private Sub ListBox2_Click()
    If Not isEvents Then
        Exit Sub
    End If
    Me.ListBox1.Locked = True
    Me.ListBox3.Locked = True
    Me.ListBox4.Locked = True
    Me.ListBox5.Locked = True
    Me.ListBox6.Locked = True
    Me.CommandButton5.Locked = True

    Me.ListBox3.Clear
    Me.ListBox4.Clear
    Me.ListBox5.Clear
    Me.ListBox6.Clear
    
    Call getDateDate("")
    
    Me.ListBox1.Locked = False
    Me.ListBox3.Locked = False
    Me.ListBox4.Locked = False
    Me.ListBox5.Locked = False
    Me.ListBox6.Locked = False
    Me.CommandButton5.Locked = False
    
End Sub

Sub getDateDate(DateT)
    Set WBbase = ThisWorkbook
    Set WSbase = WBbase.Sheets("開催日")
    endR = WSbase.Cells(WSbase.Rows.Count, 1).End(xlUp).row
    
    For i = 1 To endR
        Md1 = Mid(WSbase.Cells(i, 1), 1, 4)
        Md2 = Mid(WSbase.Cells(i, 1), 5, 2)
        Md3 = Mid(WSbase.Cells(i, 1), 7, 2)
'        If Me.ListBox1.Text = Md1 And Me.ListBox2.Text = Md2 Then
        If Me.ListBox1.List(Me.ListBox1.ListIndex) = Md1 And _
            Me.ListBox2.List(Me.ListBox2.ListIndex) = Md2 Then
            Me.ListBox3.AddItem Md3
        End If
    Next i
    
    If DateT <> "" Then
        For i = 0 To Me.ListBox3.ListCount - 1
            If DateT = Me.ListBox3.List(i) Then
                Me.ListBox3.Selected(i) = True
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub ListBox3_Click()
    If Not isEvents Then
        Exit Sub
    End If
    Me.ListBox1.Locked = True
    Me.ListBox2.Locked = True
    Me.ListBox4.Locked = True
    Me.ListBox5.Locked = True
    Me.ListBox6.Locked = True
    Me.CommandButton5.Locked = True

    Me.ListBox4.Clear
    Me.ListBox5.Clear
    Me.ListBox6.Clear
    
    targdate = ""
'    If Me.ListBox1.Text = "" Or _
'       Me.ListBox2.Text = "" Or _
'       Me.ListBox3.Text = "" Then
    If Me.ListBox1.List(Me.ListBox1.ListIndex) = "" Or _
       Me.ListBox2.List(Me.ListBox2.ListIndex) = "" Or _
       Me.ListBox3.List(Me.ListBox3.ListIndex) = "" Then
        MsgBox ("日付を選択してください。")
        Exit Sub
    Else
'        targdate = Me.ListBox1.Text & Me.ListBox2.Text & Me.ListBox3.Text
        targdate = (Me.ListBox1.List(Me.ListBox1.ListIndex) & _
                    Me.ListBox2.List(Me.ListBox2.ListIndex) & _
                    Me.ListBox3.List(Me.ListBox3.ListIndex))
        Call putPreviousDate(targdate)
    End If
    
    Call GetPlaceInfoZ(targdate)
    
    Me.ListBox1.Locked = False
    Me.ListBox2.Locked = False
    Me.ListBox4.Locked = False
    Me.ListBox5.Locked = False
    Me.ListBox6.Locked = False
    Me.CommandButton5.Locked = False
    
End Sub


Sub putPreviousDate(targdate)
    Set WBbase = ThisWorkbook
    Set WSbase = WBbase.Sheets("開催日")
    WSbase.Range(WSbase.Cells(1, 5), _
    WSbase.Cells(WSbase.Rows.Count, WSbase.Columns.Count)).ClearContents
    With WSbase
    rowEnd = .Cells(.Rows.Count, 1).End(xlUp).row
    For i = 1 To rowEnd
        If CStr(.Cells(i, 1)) = targdate Then
            .Cells(i, 5) = True
            Exit For
        End If
    Next i
    End With
End Sub


Private Sub ListBox4_Click()
    If Not isEvents Then
        Exit Sub
    End If
    Me.ListBox1.Locked = True
    Me.ListBox2.Locked = True
    Me.ListBox3.Locked = True
    Me.ListBox5.Locked = True
    Me.ListBox6.Locked = True
    Me.CommandButton5.Locked = True

    Me.ListBox5.Clear
    Me.ListBox6.Clear
    
'    If Me.ListBox1.Text = "" Or _
'       Me.ListBox2.Text = "" Or _
'       Me.ListBox3.Text = "" Or _
'       Me.ListBox4.Text = "" Then
    If Me.ListBox1.List(Me.ListBox1.ListIndex) = "" Or _
       Me.ListBox2.List(Me.ListBox2.ListIndex) = "" Or _
       Me.ListBox2.List(Me.ListBox3.ListIndex) = "" Or _
       Me.ListBox3.List(Me.ListBox4.ListIndex) = "" Then
        MsgBox ("日付、場所を選択してください。")
        Exit Sub
    Else
'        targdate = (Me.ListBox1.Text & Me.ListBox2.Text & Me.ListBox3.Text)
        targdate = (Me.ListBox1.List(Me.ListBox1.ListIndex) & _
                    Me.ListBox2.List(Me.ListBox2.ListIndex) & _
                    Me.ListBox3.List(Me.ListBox3.ListIndex))
'        targJyo = Me.ListBox4.Text
        targJyo = Me.ListBox4.List(Me.ListBox4.ListIndex)
    End If
    Set colRace = GetRaceNumInfo(targdate, targJyo)
    Call putPreviousRace(targdate, targJyo, colRace)
    
    Me.ListBox1.Locked = False
    Me.ListBox2.Locked = False
    Me.ListBox3.Locked = False
    Me.ListBox5.Locked = False
    Me.ListBox6.Locked = False
    Me.CommandButton5.Locked = False
    
End Sub

Sub putPreviousRace(targdate, targJyo, colRace)
    Set WBbase = ThisWorkbook
    Set WSbase = WBbase.Sheets("開催日")
    WSbase.Range(WSbase.Cells(1, 6), _
    WSbase.Cells(WSbase.Rows.Count, WSbase.Columns.Count)).ClearContents
    With WSbase
    rowEnd = .Cells(.Rows.Count, 1).End(xlUp).row
    colT = 7
    For i = 1 To rowEnd
        If CStr(.Cells(i, 1)) = targdate Then
            .Cells(i, 6) = targJyo
            For Each ele In colRace
                .Cells(i, colT) = ele
                colT = colT + 1
            Next
            
        End If
    Next i
    End With
End Sub

Private Sub ListBox5_Click()
    If Not isEvents Then
        Exit Sub
    End If
'    Me.Repaint
'    DoEvents
'    Debug.Print "ListBox5_Click"
    Me.ListBox1.Locked = True
    Me.ListBox2.Locked = True
    Me.ListBox3.Locked = True
    Me.ListBox4.Locked = True
    Me.ListBox6.Locked = True
    Me.CommandButton5.Locked = True

    Me.ListBox6.Clear
    
'    If Me.ListBox1.Text = "" Or _
'       Me.ListBox2.Text = "" Or _
'       Me.ListBox3.Text = "" Or _
'       Me.ListBox4.Text = "" Or _
'       Me.ListBox5.Text = "" Then
    If Me.ListBox1.List(Me.ListBox1.ListIndex) = "" Or _
       Me.ListBox2.List(Me.ListBox2.ListIndex) = "" Or _
       Me.ListBox3.List(Me.ListBox3.ListIndex) = "" Or _
       Me.ListBox4.List(Me.ListBox4.ListIndex) = "" Or _
       Me.ListBox5.List(Me.ListBox5.ListIndex) = "" Then
        MsgBox ("日付、場所、レース番号を選択してください。")
        Exit Sub
    Else
'        targdate = (Me.ListBox1.Text & Me.ListBox2.Text & Me.ListBox3.Text)
        targdate = (Me.ListBox1.List(Me.ListBox1.ListIndex) & _
                    Me.ListBox2.List(Me.ListBox2.ListIndex) & _
                    Me.ListBox3.List(Me.ListBox3.ListIndex))
'        targJyo = Me.ListBox4.Text
        targJyo = Me.ListBox4.List(Me.ListBox4.ListIndex)
'        racenum = Val(Me.ListBox5.Text)
        racenum = Val(Me.ListBox5.List(Me.ListBox5.ListIndex))
    End If
    
    Call GetRaceUma(targdate, targJyo, racenum)
    
'    Debug.Print "ListBox5_Click return"
    Me.ListBox1.Locked = False
    Me.ListBox2.Locked = False
    Me.ListBox3.Locked = False
    Me.ListBox4.Locked = False
    Me.ListBox6.Locked = False
    Me.CommandButton5.Locked = False
    
End Sub

Private Sub CommandButton2_Click()
    retval = JVLink1.JVSetUIProperties()
    'JVSetUIPropatiesエラー処理
    If (retval < -1) Then
        Debug.Print "16:" & Err.Description
        MsgBox ("エラーのためJV-Linkの設定に失敗しました。")
    End If
End Sub

Private Sub CommandButton3_Click()
    If DLflg = True Then
        Cancelflg = True
    Else
        Unload Me
    End If
End Sub

Private Sub CommandButton5_Click()

    Me.ListBox1.Locked = True
    Me.ListBox2.Locked = True
    Me.ListBox3.Locked = True
    Me.ListBox4.Locked = True
    Me.ListBox5.Locked = True
    Me.ListBox6.Locked = True
    Me.CommandButton5.Locked = True

'    If Me.ListBox1.Text = "" Or _
'       Me.ListBox2.Text = "" Or _
'       Me.ListBox3.Text = "" Or _
'       Me.ListBox4.Text = "" Or _
'       Me.ListBox5.Text = "" Then
    If Me.ListBox1.List(Me.ListBox1.ListIndex) = "" Or _
       Me.ListBox2.List(Me.ListBox2.ListIndex) = "" Or _
       Me.ListBox3.List(Me.ListBox3.ListIndex) = "" Or _
       Me.ListBox4.List(Me.ListBox4.ListIndex) = "" Or _
       Me.ListBox5.List(Me.ListBox5.ListIndex) = "" Then
        MsgBox ("日付、場所、レース番号を選択してください。")
        Exit Sub
    Else
'        targdate = (Me.ListBox1.Text & Me.ListBox2.Text & Me.ListBox3.Text)
        targdate = (Me.ListBox1.List(Me.ListBox1.ListIndex) & _
                    Me.ListBox2.List(Me.ListBox2.ListIndex) & _
                    Me.ListBox3.List(Me.ListBox3.ListIndex))
'        targJyo = Me.ListBox4.Text
        targJyo = Me.ListBox4.List(Me.ListBox4.ListIndex)
'        racenum = Val(Me.ListBox5.Text)
        racenum = Val(Me.ListBox5.List(Me.ListBox5.ListIndex))
    End If
    
    ' 3連単計算するか
    Dim isCalcSanrentan As Boolean
    isCalcSanrentan = Me.CheckBox1.Value
    Call getUmatanOdds(CStr(targdate), CStr(targJyo), CInt(racenum), isCalcSanrentan)
    
    Me.ListBox1.Locked = False
    Me.ListBox2.Locked = False
    Me.ListBox3.Locked = False
    Me.ListBox4.Locked = False
    Me.ListBox5.Locked = False
    Me.ListBox6.Locked = False
    Me.CommandButton5.Locked = False
    
    Unload Me
    UserForm1.Show vbModeless
    
End Sub
