VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UmatanOdds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Kumi As String '
Public Umaban1 As String ' 目1
Public Umaban2 As String ' 目2
Public Odds As String ' 馬単オッズ
Public Ninki1 As String ' 人気1
Public Ninki2 As String ' 人気2
Public Hyou As String ' 馬単票数
Public RevOdds As String ' 馬単裏
Public SyntheticOdds1 As String ' 馬単合成
Public SyntheticOdds2 As String ' 3連単1・2着軸総流し

Public Property Get Self() As UmatanOdds
    Set Self = Me
End Property
