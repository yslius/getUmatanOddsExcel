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
    
    ' 年
    For i = 1 To endR
        Lef = Left(WSbase.Cells(i, 1), 4)
        If tmp = "" Or tmp <> Lef Then
            Me.ListBox1.AddItem Lef
            tmp = Lef
        End If
    Next i
    
    If Me.ListBox1.ListCount = 1 Then
        Exit Sub
    End If
    
End Sub

Private Sub CommandButton4_Click()
    Call GetDate
End Sub

Private Sub ListBox1_Click()

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

    Set WBbase = ThisWorkbook
    Set WSbase = WBbase.Sheets("開催日")
    endR = WSbase.Cells(WSbase.Rows.Count, 1).End(xlUp).row
    tmp = ""
    For i = 1 To endR
        Md1 = Mid(WSbase.Cells(i, 1), 1, 4)
        Md2 = Mid(WSbase.Cells(i, 1), 5, 2)
        If Me.ListBox1.Text = Md1 And tmp <> Md2 Then
            Me.ListBox2.AddItem Md2
            tmp = Md2
        End If
    Next i
    
    Me.ListBox2.Locked = False
    Me.ListBox3.Locked = False
    Me.ListBox4.Locked = False
    Me.ListBox5.Locked = False
    Me.ListBox6.Locked = False
    Me.CommandButton5.Locked = False
    
End Sub

Private Sub ListBox2_Click()

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
    
    Set WBbase = ThisWorkbook
    Set WSbase = WBbase.Sheets("開催日")
    endR = WSbase.Cells(WSbase.Rows.Count, 1).End(xlUp).row
    
    For i = 1 To endR
        Md1 = Mid(WSbase.Cells(i, 1), 1, 4)
        Md2 = Mid(WSbase.Cells(i, 1), 5, 2)
        Md3 = Mid(WSbase.Cells(i, 1), 7, 2)
        If Me.ListBox1.Text = Md1 And Me.ListBox2.Text = Md2 Then
            Me.ListBox3.AddItem Md3
        End If
    Next i

    Me.ListBox1.Locked = False
    Me.ListBox3.Locked = False
    Me.ListBox4.Locked = False
    Me.ListBox5.Locked = False
    Me.ListBox6.Locked = False
    Me.CommandButton5.Locked = False
    
End Sub

Private Sub ListBox3_Click()

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
    If Me.ListBox1.Text = "" Or _
       Me.ListBox2.Text = "" Or _
       Me.ListBox3.Text = "" Then
        MsgBox ("日付を選択してください。")
        Exit Sub
    Else
        targdate = Me.ListBox1.Text & Me.ListBox2.Text & Me.ListBox3.Text
    End If
    
    Call GetPlaceInfoZ(targdate)
    
    Me.ListBox1.Locked = False
    Me.ListBox2.Locked = False
    Me.ListBox4.Locked = False
    Me.ListBox5.Locked = False
    Me.ListBox6.Locked = False
    Me.CommandButton5.Locked = False
    
End Sub

Private Sub ListBox4_Click()

    Me.ListBox1.Locked = True
    Me.ListBox2.Locked = True
    Me.ListBox3.Locked = True
    Me.ListBox5.Locked = True
    Me.ListBox6.Locked = True
    Me.CommandButton5.Locked = True

    Me.ListBox5.Clear
    Me.ListBox6.Clear
    
    If Me.ListBox1.Text = "" Or _
       Me.ListBox2.Text = "" Or _
       Me.ListBox3.Text = "" Or _
       Me.ListBox4.Text = "" Then
        MsgBox ("日付、場所を選択してください。")
        Exit Sub
    Else
        targdate = (Me.ListBox1.Text & Me.ListBox2.Text & Me.ListBox3.Text)
        targJyo = Me.ListBox4.Text
    End If
    Call GetRaceNumInfo(targdate, targJyo)
    
    Me.ListBox1.Locked = False
    Me.ListBox2.Locked = False
    Me.ListBox3.Locked = False
    Me.ListBox5.Locked = False
    Me.ListBox6.Locked = False
    Me.CommandButton5.Locked = False
    
End Sub

Private Sub ListBox5_Click()

    Me.ListBox1.Locked = True
    Me.ListBox2.Locked = True
    Me.ListBox3.Locked = True
    Me.ListBox4.Locked = True
    Me.ListBox6.Locked = True
    Me.CommandButton5.Locked = True

    Me.ListBox6.Clear
    
    If Me.ListBox1.Text = "" Or _
       Me.ListBox2.Text = "" Or _
       Me.ListBox3.Text = "" Or _
       Me.ListBox4.Text = "" Or _
       Me.ListBox5.Text = "" Then
        MsgBox ("日付、場所、レース番号を選択してください。")
        Exit Sub
    Else
        targdate = (Me.ListBox1.Text & Me.ListBox2.Text & Me.ListBox3.Text)
        targJyo = Me.ListBox4.Text
        racenum = Val(Me.ListBox5.Text)
    End If
    
    Call GetRaceUma(targdate, targJyo, racenum)
    
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

    If Me.ListBox1.Text = "" Or _
       Me.ListBox2.Text = "" Or _
       Me.ListBox3.Text = "" Or _
       Me.ListBox4.Text = "" Or _
       Me.ListBox5.Text = "" Then
        MsgBox ("日付、場所、レース番号を選択してください。")
        Exit Sub
    Else
        targdate = Me.ListBox1.Text & Me.ListBox2.Text & Me.ListBox3.Text
        targJyo = Me.ListBox4.Text
        racenum = Val(Me.ListBox5.Text)
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
    
End Sub
