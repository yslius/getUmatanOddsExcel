VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataCSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public dataCsvAll As String

Public Property Let setIniData(InputData() As Variant)
    Me.dataCsvAll = Join(InputData, ",") & vbCrLf
End Property

Public Property Let setData(indRow As Long, indCol As Long, InputData As String)
    Dim arrdataCsv As Variant
    arrdataCsv = Split(Me.dataCsvAll, vbCrLf)
    If indRow > UBound(arrdataCsv) + 1 Then
        ReDim Preserve arrdataCsv(indRow - 1)
    End If
    Dim arrdatadataCsv As Variant
    arrdatadataCsv = Split(arrdataCsv(indRow - 1), ",")
    If indCol > UBound(arrdatadataCsv) Then
        ReDim Preserve arrdatadataCsv(indCol - 1)
    End If
    arrdatadataCsv(indCol - 1) = InputData
    arrdataCsv(indRow - 1) = Join(arrdatadataCsv, ",")
    Me.dataCsvAll = Join(arrdataCsv, vbCrLf)
End Property


Public Property Get getData(indRow As Long, indCol As Long)
    arrdataCsv = Split(Me.dataCsvAll, vbCrLf)
    arrdatadataCsv = Split(arrdataCsv(indRow - 1), ",")
    getData = arrdatadataCsv(indCol - 1)
End Property


Public Property Let addDataLine(indRow As Long, lenArr As Long, InputData As String)
    Dim arrdataCsv As Variant
    arrdataCsv = Split(Me.dataCsvAll, vbCrLf)
    Dim arrdatadataCsv As Variant
    arrdatadataCsv = Split(arrdataCsv(indRow - 1), ",")
    If UBound(arrdatadataCsv) <> lenArr Then
        ReDim Preserve arrdatadataCsv(lenArr - 1)
    End If
    arrdataCsv(indRow - 1) = Join(arrdatadataCsv, ",")
    arrdataCsv(indRow - 1) = arrdataCsv(indRow - 1) & "," & InputData
    Me.dataCsvAll = Join(arrdataCsv, vbCrLf)
End Property


Public Function getDataLine(indRow As Long) As String
    arrdataCsv = Split(Me.dataCsvAll, vbCrLf)
    getDataLine = arrdataCsv(indRow - 1)
End Function


Public Sub setDataSub(indRow, indCol, InputData)
    Dim arrdataCsv As Variant
    arrdataCsv = Split(Me.dataCsvAll, vbCrLf)
    Dim arrdatadataCsv As Variant
    arrdatadataCsv = Split(arrdataCsv(indRow - 1), ",")
    If indCol > UBound(arrdatadataCsv) Then
        ReDim Preserve arrdatadataCsv(indCol - 1)
    End If
    arrdatadataCsv(indCol - 1) = InputData
    arrdataCsv(indRow - 1) = Join(arrdatadataCsv, ",")
    Me.dataCsvAll = Join(arrdataCsv, vbCrLf)
End Sub


Public Sub addDataLineSub(indRow As Long, InputData As String)
    Dim arrdataCsv As Variant
    arrdataCsv = Split(Me.dataCsvAll, vbCrLf)
    arrdataCsv(indRow - 1) = arrdataCsv(indRow - 1) & "," & InputData
    Me.dataCsvAll = Join(arrdataCsv, vbCrLf)
End Sub


Public Function getDataMaxRow() As Long
    arrdataCsv = Split(Me.dataCsvAll, vbCrLf)
    getDataMaxRow = UBound(arrdataCsv) - 1
End Function


Public Function getDataRow(strShortJyo, racenum) As Long
    arrdataCsv = Split(Me.dataCsvAll, vbCrLf)
    For i = 0 To UBound(arrdataCsv) - 1
        arrdatadataCsv = Split(arrdataCsv(i), ",")
        If InStr(arrdatadataCsv(2), strShortJyo) And _
           Val(arrdatadataCsv(5)) = racenum Then
           getDataRow = i + 1
           Exit Function
        End If
    Next i
End Function


Public Property Let pasteData(WS As Worksheet, indStartRow As Long)
    Dim arrdataCsv As Variant
    arrdataCsv = Split(Me.dataCsvAll, vbCrLf)
    Dim arrdatadataCsv As Variant
    Dim rowInsert As Long
    rowInsert = 1
    For i = 0 To UBound(arrdataCsv)
        If i >= indStartRow - 1 Then
            arrdatadataCsv = Split(arrdataCsv(i), ",")
            For j = 0 To UBound(arrdatadataCsv)
                WS.Cells(rowInsert, j + 1) = arrdatadataCsv(j)
            Next j
            rowInsert = rowInsert + 1
        End If
    Next i
End Property
