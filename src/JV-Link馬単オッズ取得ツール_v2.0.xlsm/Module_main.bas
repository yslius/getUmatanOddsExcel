Attribute VB_Name = "Module_main"
Sub GetTyoukyouData(strdate, targJyo, racenum)
    ' �ۑ��t�H���_�[�̎w��
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "�ۑ��t�H���_�[�̎w��"
        .InitialFileName = ThisWorkbook.Path & "\"
        If .Show = 0 Then
            MsgBox "�L�����Z���{�^�����N���b�N���܂����B"
            Exit Sub
        End If
        pathUserSave = .SelectedItems(1)
    End With
    
startTime = Timer

    Set WBbase = ThisWorkbook
    Set WSbase = WBbase.Sheets("Template")
    WSbase.Rows(1).ClearContents
    WSbase.Range(WSbase.Cells(3, 1), WSbase.Cells(WSbase.Rows.Count, 12)).ClearContents
    
    Dim retval As Long
    Dim readcount As Long
    Dim dlcount As Long
    Dim lastfiletimestamp As String
    Dim status As Long
    Dim buff As String
    Dim filename As String
    
    Dim targdate As Date
    targdate = Mid(strdate, 1, 4) & "/" & Mid(strdate, 5, 2) & "/" & Mid(strdate, 7, 2)
    targdate = DateAdd("d", -181, targdate)
    strStartdate = Format(targdate, "yyyymmdd")
    
    'JVLink��������
    UserForm1.JVLink1.JVClose
    UserForm1.JVLink1.JVInit ("EXCELSAMPLE")
    
    '�~�όn�f�[�^�̌Ăяo��
    retval = UserForm1.JVLink1.JVOpen("SLOP", strStartdate & "000000", 1, readcount, dlcount, lastfiletimestamp)
    'JVOpen�G���[����
    If (retval < -1) Then
        MsgBox ("JVOpen�G���[ " & retval)
        GoTo CommandButton1_END
    End If
    
    '�t�@�C���̃_�E�����[�h
    Dim mHcData As JV_HC_HANRO
    status = 0
    DLflg = True
    Do While status <> dlcount
        '�L�����Z���{�^���`�F�b�N
        If Cancelflg = True Then Exit Do
        status = UserForm1.JVLink1.JVStatus
        UserForm1.Label1.Caption _
         = dlcount & "�t�@�C���� " & status & " �t�@�C���_�E�����[�h����"
        DoEvents
        Sleep (120)
    Loop
    
    Cancelflg = False
    retval = 1
    cnt = 3
    While retval <> 0
        '�L�����Z���{�^���`�F�b�N
        If Cancelflg = True Then GoTo CommandButton1_END

        'JVOpen�Ŏw�肵���f�[�^���P���R�[�h����荞��
        retval = UserForm1.JVLink1.JVRead(buff, 40000, filename)
        ' JVRead�G���[����
        If (retval < -1) Then
            MsgBox ("JVRead�G���[�BRC=" & retval)
            GoTo CommandButton1_END
        End If
        If Left(buff, 2) = "HC" Then
            'JVData�\���̂�RA�̃��R�[�h���Z�b�g����
            Call SetData_HC(buff, mHcData)
            
            If Val(mHcData.ChokyoDate.Year & _
                mHcData.ChokyoDate.Month & _
                mHcData.ChokyoDate.Day) > Val(strdate) Then
                GoTo LOOP_END
            End If
            
            If Val(mHcData.LapTime1) = 0 Then
                GoTo LOOP_NEXT
            End If
            isFind = False
            For i = 1 To 16
                If WBbase.Sheets("���[�X").Cells(i, 6) = mHcData.KettoNum Then
                    isFind = True
                    strUma = WBbase.Sheets("���[�X").Cells(i, 5)
                    Exit For
                End If
            Next i
            If isFind = True Then
                If Int(mHcData.TresenKubun) = 0 Then
                    WSbase.Cells(cnt, 1) = "���Y"
                Else
                    WSbase.Cells(cnt, 1) = "�I��"
                End If
                WSbase.Cells(cnt, 2) = mHcData.ChokyoDate.Year & _
                                mHcData.ChokyoDate.Month & _
                                mHcData.ChokyoDate.Day
                WSbase.Cells(cnt, 3) = strUma  '�@mHcData.KettoNum
                WSbase.Cells(cnt, 4) = WBbase.Sheets("���[�X").Cells(i, 4)
                WSbase.Cells(cnt, 5) = Val(mHcData.HaronTime4) / 10
                WSbase.Cells(cnt, 6) = Val(mHcData.HaronTime3) / 10
                WSbase.Cells(cnt, 7) = Val(mHcData.HaronTime2) / 10
                WSbase.Cells(cnt, 8) = Val(mHcData.LapTime1) / 10
                WSbase.Cells(cnt, 9) = Val(mHcData.LapTime4) / 10
                WSbase.Cells(cnt, 10) = Val(mHcData.LapTime3) / 10
                WSbase.Cells(cnt, 11) = Val(mHcData.LapTime2) / 10
                WSbase.Cells(cnt, 12) = Val(mHcData.LapTime1) / 10
                cnt = cnt + 1
'                If cnt > 10 Then
'                    GoTo LOOP_END
'                End If
                UserForm1.Label1.Caption = buff
                DoEvents
            End If
        Else
            UserForm1.JVLink1.JVSkip
        End If
LOOP_NEXT:
    Wend

LOOP_END:
    UserForm1.JVLink1.JVClose
    
    ' �t�@�C�����̓���
    WSbase.Cells(1, 1) = "TrainData_" & strdate & "_" & targJyo & "_" & Format(racenum, "00")
    
    '�����_�̕\��
    endR = WSbase.Cells(WSbase.Rows.Count, 1).End(xlUp).row
    WSbase.Range(WSbase.Cells(3, 5), WSbase.Cells(endR, 12)).NumberFormatLocal = "0.0"
    
    ' �����\�[�g
    WSbase.Range(WSbase.Cells(3, 1), WSbase.Cells(endR, 12)).Sort _
    key1:=WSbase.Cells(3, 12), Order1:=xlAscending
    
    ' CSV�f�[�^(Shift-JIS)�̕ۑ�
    Dim FilePath As String
    strSaveName = "TrainData_" & strdate & "_" & targJyo & "_" & Format(racenum, "00") & ".csv"
    csvFileSJ = pathUserSave & "\" & strSaveName
    CSVData = CreateCSVData(WSbase)
    Open csvFileSJ For Output As #1
        Print #1, CSVData
    Close #1
    
CommandButton1_END:
    UserForm1.JVLink1.JVClose
    Unload UserForm1

endTime = Timer
Debug.Print "�������ԁF" & endTime - startTime

    MsgBox "����ɏI�����܂����B"
    
End Sub


Sub GetUmatanOddsData()
    ' �ۑ��t�H���_�[�̎w��
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "�ۑ��t�H���_�[�̎w��"
        .InitialFileName = ThisWorkbook.Path & "\"
        If .Show = 0 Then
            MsgBox "�L�����Z���{�^�����N���b�N���܂����B"
            Exit Sub
        End If
        pathUserSave = .SelectedItems(1)
    End With
    
startTime = Timer

    Set WBbase = ThisWorkbook
    Set WSbase = WBbase.Sheets("Template")
    WSbase.Rows(1).ClearContents
    WSbase.Range(WSbase.Cells(3, 1), WSbase.Cells(WSbase.Rows.Count, 12)).ClearContents
    
    Dim retval As Long
    Dim readcount As Long
    Dim dlcount As Long
    Dim lastfiletimestamp As String
    Dim status As Long
    Dim buff As String
    Dim filename As String
    
    Dim targdate As Date
    targdate = Mid(strdate, 1, 4) & "/" & Mid(strdate, 5, 2) & "/" & Mid(strdate, 7, 2)
    targdate = DateAdd("d", -181, targdate)
    strStartdate = Format(targdate, "yyyymmdd")
    
    'JVLink��������
    UserForm1.JVLink1.JVClose
    UserForm1.JVLink1.JVInit ("EXCELSAMPLE")
    
    '�~�όn�f�[�^�̌Ăяo��
    retval = UserForm1.JVLink1.JVOpen("SLOP", strStartdate & "000000", 1, readcount, dlcount, lastfiletimestamp)
    'JVOpen�G���[����
    If (retval < -1) Then
        MsgBox ("JVOpen�G���[ " & retval)
        GoTo CommandButton1_END
    End If
    
    '�t�@�C���̃_�E�����[�h
    Dim mHcData As JV_HC_HANRO
    status = 0
    DLflg = True
    Do While status <> dlcount
        '�L�����Z���{�^���`�F�b�N
        If Cancelflg = True Then Exit Do
        status = UserForm1.JVLink1.JVStatus
        UserForm1.Label1.Caption _
         = dlcount & "�t�@�C���� " & status & " �t�@�C���_�E�����[�h����"
        DoEvents
        Sleep (120)
    Loop
    
    Cancelflg = False
    retval = 1
    cnt = 3
    While retval <> 0
        '�L�����Z���{�^���`�F�b�N
        If Cancelflg = True Then GoTo CommandButton1_END

        'JVOpen�Ŏw�肵���f�[�^���P���R�[�h����荞��
        retval = UserForm1.JVLink1.JVRead(buff, 40000, filename)
        ' JVRead�G���[����
        If (retval < -1) Then
            MsgBox ("JVRead�G���[�BRC=" & retval)
            GoTo CommandButton1_END
        End If
        If Left(buff, 2) = "HC" Then
            'JVData�\���̂�RA�̃��R�[�h���Z�b�g����
            Call SetData_HC(buff, mHcData)
            
            If Val(mHcData.ChokyoDate.Year & _
                mHcData.ChokyoDate.Month & _
                mHcData.ChokyoDate.Day) > Val(strdate) Then
                GoTo LOOP_END
            End If
            
            If Val(mHcData.LapTime1) = 0 Then
                GoTo LOOP_NEXT
            End If
            isFind = False
            For i = 1 To 16
                If WBbase.Sheets("���[�X").Cells(i, 6) = mHcData.KettoNum Then
                    isFind = True
                    strUma = WBbase.Sheets("���[�X").Cells(i, 5)
                    Exit For
                End If
            Next i
            If isFind = True Then
                If Int(mHcData.TresenKubun) = 0 Then
                    WSbase.Cells(cnt, 1) = "���Y"
                Else
                    WSbase.Cells(cnt, 1) = "�I��"
                End If
                WSbase.Cells(cnt, 2) = mHcData.ChokyoDate.Year & _
                                mHcData.ChokyoDate.Month & _
                                mHcData.ChokyoDate.Day
                WSbase.Cells(cnt, 3) = strUma  '�@mHcData.KettoNum
                WSbase.Cells(cnt, 4) = WBbase.Sheets("���[�X").Cells(i, 4)
                WSbase.Cells(cnt, 5) = Val(mHcData.HaronTime4) / 10
                WSbase.Cells(cnt, 6) = Val(mHcData.HaronTime3) / 10
                WSbase.Cells(cnt, 7) = Val(mHcData.HaronTime2) / 10
                WSbase.Cells(cnt, 8) = Val(mHcData.LapTime1) / 10
                WSbase.Cells(cnt, 9) = Val(mHcData.LapTime4) / 10
                WSbase.Cells(cnt, 10) = Val(mHcData.LapTime3) / 10
                WSbase.Cells(cnt, 11) = Val(mHcData.LapTime2) / 10
                WSbase.Cells(cnt, 12) = Val(mHcData.LapTime1) / 10
                cnt = cnt + 1
'                If cnt > 10 Then
'                    GoTo LOOP_END
'                End If
                UserForm1.Label1.Caption = buff
                DoEvents
            End If
        Else
            UserForm1.JVLink1.JVSkip
        End If
LOOP_NEXT:
    Wend

LOOP_END:
    UserForm1.JVLink1.JVClose
    
    ' �t�@�C�����̓���
    WSbase.Cells(1, 1) = "TrainData_" & strdate & "_" & targJyo & "_" & Format(racenum, "00")
    
    '�����_�̕\��
    endR = WSbase.Cells(WSbase.Rows.Count, 1).End(xlUp).row
    WSbase.Range(WSbase.Cells(3, 5), WSbase.Cells(endR, 12)).NumberFormatLocal = "0.0"
    
    ' �����\�[�g
    WSbase.Range(WSbase.Cells(3, 1), WSbase.Cells(endR, 12)).Sort _
    key1:=WSbase.Cells(3, 12), Order1:=xlAscending
    
    ' CSV�f�[�^(Shift-JIS)�̕ۑ�
    Dim FilePath As String
    strSaveName = "TrainData_" & strdate & "_" & targJyo & "_" & Format(racenum, "00") & ".csv"
    csvFileSJ = pathUserSave & "\" & strSaveName
    CSVData = CreateCSVData(WSbase)
    Open csvFileSJ For Output As #1
        Print #1, CSVData
    Close #1
    
CommandButton1_END:
    UserForm1.JVLink1.JVClose
    Unload UserForm1

endTime = Timer
Debug.Print "�������ԁF" & endTime - startTime

    MsgBox "����ɏI�����܂����B"
    
End Sub


