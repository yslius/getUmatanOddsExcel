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
    Set WSbase = WBbase.Sheets("�J�Ó�")
    WSbase.Cells.Clear
    
    UserForm1.Show vbModeless
    
    'JVLink��������
    UserForm1.JVLink1.JVClose
    UserForm1.JVLink1.JVInit ("EXCELSAMPLE")

    ' �J�ÃX�P�W���[��
    Dim mYsData As JV_YS_SCHEDULE

    '�Z�b�g�A�b�v�f�[�^�̌Ăяo��
    strdate = Val(Format(Date, "yyyy")) - 4 & Format(Date, "mmdd")
    retval = UserForm1.JVLink1.JVOpen("YSCH", strdate & "000000", 4, readcount, dlcount, lastfiletimestamp)
    If (retval < -1) Then
        Debug.Print "17:" & Err.Description
        MsgBox ("JVOpen�G���[�BRC=" & retval)
        GoTo CommandButton1_END
    End If
    
    cnt = 1
    retval = 1
    While retval <> 0
        '�L�����Z���{�^���`�F�b�N
        If Cancelflg = True Then GoTo CommandButton1_END
        
        'JVOpen�Ŏw�肵���f�[�^���P���R�[�h����荞��
        retval = UserForm1.JVLink1.JVRead(buff, 40000, filename)
        ' JVRead�G���[����
        If (retval < -1) Then
            Debug.Print "18:" & Err.Description
            MsgBox ("JVRead�G���[�BRC=" & retval)
            GoTo CommandButton1_END
        End If
        
        If Left(buff, 2) = "YS" And filename <> "YSMW2020999920200406150936.jvd" Then
            Call SetData_YS(buff, mYsData)
            ' �{���𒴂����烋�[�v������
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
    ' �d���폜�ƕ��ёւ�
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

    '��ʂ�ǂݍ��݂��I��������JVClose���s��
    UserForm1.JVLink1.JVClose
    If Cancelflg = True Then
        MsgBox "�L�����Z������܂����B"
    Else
        UserForm1.Label1.Caption = "�ǂݍ��݂��I�����܂����B"
    End If
    DLflg = False
    
    Unload UserForm1

End Sub

Function findAlreadyDate(strdateTarg As String) As Long
    Set WBbase = ThisWorkbook
    Set WSbase = WBbase.Sheets("�J�Ó�")
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

    ' �J�ÃX�P�W���[��
    Dim mYsData As JV_YS_SCHEDULE
    
    opt = 1
    dateTarg = CDate(Format(strdateTarg, "####/##/##"))
    dateTarg = DateAdd("d", -1, dateTarg)
    If DateDiff("d", dateTarg, Date) >= 365 Then
        opt = 4
    End If

    '�Z�b�g�A�b�v�f�[�^�̌Ăяo��
    retval = UserForm1.JVLink1.JVOpen("YSCH", Format(dateTarg, "yyyymmdd") & "000000", opt, readcount, dlcount, lastfiletimestamp)
    If (retval < -1) Then
        Debug.Print "19:" & Err.Description
        MsgBox ("JVOpen�G���[�BRC=" & retval)
        GoTo CommandButton1_END
    End If
    retval = 1
    While retval <> 0
        retval = UserForm1.JVLink1.JVRead(buff, 40000, filename)
        If (retval < -1) Then
            Debug.Print "1:" & Err.Description
            MsgBox ("JVRead�G���[�BRC=" & retval)
            GoTo CommandButton1_END
        End If
        If Left(buff, 2) = "YS" Then
            Call SetData_YS(buff, mYsData)
            ' �������烋�[�v������
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

    'JVLink��������
    UserForm1.JVLink1.JVClose
    UserForm1.JVLink1.JVInit ("EXCELSAMPLE")
    '�~�όn�f�[�^��RACE��ListBox�̓��t�ȍ~�ɂ��Ď�荞�݌Ăяo��
    retval = UserForm1.JVLink1.JVOpen("RACE", targdate - 1 & "000000", 1, readcount, dlcount, lastfiletimestamp)
    If (retval < -1) Then
        Debug.Print "2:" & Err.Description
        MsgBox ("JVOpen�G���[�BRC=" & retval)
        GoTo LOOP_END
    End If

    Dim mRaData As JV_RA_RACE
   
    status = 0
    DLflg = True
    UserForm1.CommandButton3.Caption = "�L�����Z��"
    Do While status <> dlcount
        '�L�����Z���{�^���`�F�b�N
        If Cancelflg = True Then Exit Do
        status = UserForm1.JVLink1.JVStatus
        UserForm1.Label1.Caption = dlcount & "�t�@�C���� " & status & " �t�@�C���_�E�����[�h����"
        DoEvents
        Sleep (10)
    Loop
    
    retval = 1
    isGetData = False
    While retval <> 0
        'JVOpen�Ŏw�肵���f�[�^���P���R�[�h����荞��
        retval = UserForm1.JVLink1.JVRead(buff, 40000, filename)
        If (retval < -1) Then
            Debug.Print "3:" & Err.Description
            MsgBox ("JVRead�G���[�BRC=" & retval)
            GoTo LOOP_END
        End If

        If Left(buff, 2) = "RA" Then

            Call SetData_RA(buff, mRaData)
            
            ' �I�񂾓��t�𒴂����烋�[�v������
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
        MsgBox "�L�����Z������܂����B"
    Else
        UserForm1.Label1.Caption = "�ǂݍ��݂��I�����܂����B"
    End If
    UserForm1.CommandButton3.Caption = "Exit"
    DLflg = False

End Sub


Sub GetPlaceInfoZ(targdate)
    
    Set WBbase = ThisWorkbook
    Set WSbase = WBbase.Sheets("�J�Ó�")
    
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
    ' ���t�̕ϊ�
     If Date < dateTarg Then
'        targdate = Format(Date - 1, "yyyymmdd")
        dateTarg = DateAdd("d", -1, Date)
    ElseIf DateDiff("d", dateTarg, Date) >= 365 Then
        opt = 4
        dateTarg = DateAdd("d", -1, dateTarg)
    Else
        dateTarg = DateAdd("d", -1, dateTarg)
    End If
    
    'JVLink��������
    UserForm1.JVLink1.JVClose
    UserForm1.JVLink1.JVInit ("EXCELSAMPLE")
    '�~�όn�f�[�^��RACE��ListBox�̓��t�ȍ~�ɂ��Ď�荞�݌Ăяo��
    retval = UserForm1.JVLink1.JVOpen("RACE", Format(dateTarg, "yyyymmdd") & "000000", opt, readcount, dlcount, lastfiletimestamp)
    If (retval < -1) Then
        Debug.Print "4:" & Err.Description
        MsgBox ("JVOpen�G���[�BRC=" & retval)
        GoTo CommandButton1_END
    End If

    Dim mRaData As JV_RA_RACE
   
    status = 0
    DLflg = True
    UserForm1.CommandButton3.Caption = "�L�����Z��"
    Do While status <> dlcount
        Debug.Print "discount=" & CStr(discount)
        '�L�����Z���{�^���`�F�b�N
        If Cancelflg = True Then Exit Do
        status = UserForm1.JVLink1.JVStatus
        UserForm1.Label1.Caption = dlcount & "�t�@�C���� " & status & " �t�@�C���_�E�����[�h����"
        DoEvents
        Sleep (10)
    Loop
    
    retval = 1
    isIndata = False
    While retval <> 0
        '�L�����Z���{�^���`�F�b�N
        If Cancelflg = True Then GoTo CommandButton1_END
         
        'JVOpen�Ŏw�肵���f�[�^���P���R�[�h����荞��
        retval = UserForm1.JVLink1.JVRead(buff, 40000, filename)
        If (retval < -1) Then
            Debug.Print "5:" & Err.Description
            MsgBox ("JVRead�G���[�BRC=" & retval)
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
        MsgBox "�L�����Z������܂����B"
    Else
        UserForm1.Label1.Caption = "�ǂݍ��݂��I�����܂����B"
    End If
    UserForm1.CommandButton3.Caption = "Exit"
    DLflg = False
    Set GetRaceNumInfo = colRace

End Function

Sub GetRaceUma(targdate, targJyo, racenum)
    Set WBbase = ThisWorkbook
    Set WSbase = WBbase.Sheets("���[�X")
    WSbase.Range(WSbase.Rows(2), WSbase.Rows(WSbase.Rows.Count)).ClearContents
    
'    targdate = 20190112
'    targJyo = "���R"
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
    ' ���t�̕ϊ�
     If Date < dateTarg Then
'        targdate = Format(Date - 1, "yyyymmdd")
        dateTarg = DateAdd("d", -1, Date)
    ElseIf DateDiff("d", dateTarg, Date) >= 365 Then
        opt = 4
        dateTarg = DateAdd("d", -1, dateTarg)
    Else
        dateTarg = DateAdd("d", -1, dateTarg)
    End If
    
    'JVLink��������
    UserForm1.JVLink1.JVClose
    UserForm1.JVLink1.JVInit ("EXCELSAMPLE")
    
    '�~�όn�f�[�^�̌Ăяo��
    retval = UserForm1.JVLink1.JVOpen("RACE", targdate - 1 & "000000", opt, readcount, dlcount, lastfiletimestamp)
    'JVOpen�G���[����
    If (retval < -1) Then
        Debug.Print "6:" & Err.Description
        MsgBox ("JVOpen�G���[ " & retval)
        GoTo CommandButton1_END
    End If
    
    ' �n�����[�X���
    Dim mSeData As JV_SE_RACE_UMA
    
    '�t�@�C���̃_�E�����[�h
    status = 0
    DLflg = True
    Do While status <> dlcount
        '�L�����Z���{�^���`�F�b�N
        If Cancelflg = True Then Exit Do
        status = UserForm1.JVLink1.JVStatus
        UserForm1.Label1.Caption = dlcount & "�t�@�C���� " & status & " �t�@�C���_�E�����[�h����"
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
        '�L�����Z���{�^���`�F�b�N
        If Cancelflg = True Then GoTo CommandButton1_END

        'JVOpen�Ŏw�肵���f�[�^���P���R�[�h����荞��
        retval = UserForm1.JVLink1.JVRead(buff, 40000, filename)
        Debug.Print "retval:" & retval
        ' JVRead�G���[����
        If (retval < -1) Then
            Debug.Print "7:" & Err.Description
            MsgBox ("JVRead�G���[�BRC=" & retval)
            GoTo CommandButton1_END
        End If
        If Left(buff, 2) = "SE" Then
            'JVData�\���̂�SE�̃��R�[�h���Z�b�g����
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
    Set region = .Cells(1, 1).CurrentRegion ' �f�[�^�͈̔͂������擾
    
    Dim row As Range
    For i = 1 To region.Rows.Count  ' �s�̃��[�v
        Line = ""
        For j = 1 To region.Columns.Count ' ��̃��[�v
            ' �J���}��؂�Ō���
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
        ' �s������
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
    
    For i = 3 To region.Rows.Count  ' �s�̃��[�v
        Line = ""
        For j = 1 To 12 ' ��̃��[�v
            ' �J���}��؂�Ō���
            Dim item As Variant
            item = .Cells(i, j).Value
            If Line = "" Then
                Line = item
            Else
                Line = Line & "," & item
            End If
        Next
        ' �s������
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



