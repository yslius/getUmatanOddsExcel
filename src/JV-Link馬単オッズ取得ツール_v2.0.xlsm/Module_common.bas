Attribute VB_Name = "Module_common"
'API宣言
'ウィンドウハンドルを取得する関数
Public Declare PtrSafe Function FindWindow Lib "USER32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "USER32.dll" Alias "FindWindowExA" _
  (ByVal hwndParent As Long, ByVal hwndChildAfter As Long, _
  ByVal lpszClass As String, ByVal lpszWindow As String) As Long

'ウィンドウに関する情報を返す関数
Public Declare PtrSafe Function GetWindowLong Lib "USER32" Alias "GetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'ウィンドウの属性を変更する関数
Public Declare PtrSafe Function SetWindowLong Lib "USER32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'メニューバーを描画する関数
Public Declare PtrSafe Function DrawMenuBar Lib "USER32" (ByVal hWnd As Long) As Long

Declare PtrSafe Function SetForegroundWindow Lib "USER32.dll" (ByVal hWnd As Long) As Long

Declare PtrSafe Function GetInputState Lib "USER32" () As Long

Public Declare Function SendMessage Lib "USER32.dll" Alias "SendMessageA" _
  (ByVal hWnd As Long, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer

Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'スタイルを取得する定数
Public Const GWL_STYLE As Long = -16
'ウィンドウスタイル
Public Const WS_SYSMENU As Long = &H80000

Public WBbase As Workbook
Public WSbase As Worksheet
Public WBtarg As Workbook
Public WStarg As Worksheet

Public cntRaceUma As Integer
Public isERROREND As Boolean

Public arrAddHead() As Variant
Public arrAddTrain(6) As Variant


Sub initArray()
    arrAddHead = Array("目1", "目2", "馬単オッズ", _
                   "人気1", "人気2", "馬単票数", _
                   "馬単裏", "馬単合成", "3連単1・2着軸総流し")
    
'    With ThisWorkbook.Sheets("Template")
'        arrAddTrain(0) = .Cells(2, 152)
'        arrAddTrain(1) = .Cells(2, 153)
'        arrAddTrain(2) = .Cells(2, 154)
'        arrAddTrain(3) = .Cells(2, 155)
'        arrAddTrain(4) = .Cells(2, 156)
'        arrAddTrain(5) = .Cells(2, 157)
'        arrAddTrain(6) = .Cells(2, 158)
'    End With
End Sub


Sub writeAddHeadData(datacsv As datacsv)
    Dim rowTarget As Long
    rowTarget = 2
    Do While rowTarget < datacsv.getDataMaxRow()
'        For i = 0 To UBound(arrAddHead)
'            datacsv.setData(rowTarget - 1, 16 + i) = arrAddHead(i)
'        Next i
        For i = 0 To UBound(arrAddTrain)
            datacsv.setData(rowTarget + 1, 31 + i) = arrAddTrain(i)
        Next i
        rowTarget = rowTarget + datacsv.getData(rowTarget, 4) + 3
    Loop
End Sub


Function createPayData(strPay As String, strKumi As String, Optional flag = False) As String
    If Val(strPay) = 0 Then
        createPayData = ""
        Exit Function
    End If
    
    Dim tmpstrKumi As String
    If flag Then
        tmpstrKumi = Format(Left(strKumi, 1), "00") & "・" & Format(Right(strKumi, 1), "00")
    ElseIf Len(strKumi) >= 6 Then
        tmpstrKumi = Left(strKumi, 2) & "・" & _
                     Mid(strKumi, 3, 2) & "・" & _
                     Right(strKumi, 2)
    ElseIf Len(strKumi) >= 4 Then
        tmpstrKumi = Left(strKumi, 2) & "・" & _
                     Right(strKumi, 2)
    Else
        tmpstrKumi = strKumi
    End If
    
    createPayData = Val(strPay) & "(" & tmpstrKumi & ")"
    
End Function


Sub AppDrawCalStart()
     Application.ScreenUpdating = True
     Application.Calculation = xlCalculationAutomatic
End Sub


Sub AppDrawCalStop()
     Application.ScreenUpdating = False
     Application.Calculation = xlCalculationManual
End Sub
