VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Wait 
   Caption         =   "情報取得中です..."
   ClientHeight    =   1485
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4890
   OleObjectBlob   =   "UserForm_Wait.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm_Wait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
  Dim hWnd As Long
  Dim lngWstyle As Long
 
  'ユーザーフォームのハンドル
  hWnd = FindWindow(vbNullString, Me.Caption)
  lngWstyle = GetWindowLong(hWnd, GWL_STYLE)
 
  '閉じるボタンの消去
  SetWindowLong hWnd, GWL_STYLE, lngWstyle And (Not WS_SYSMENU)
  'メニュー再描画
  DrawMenuBar hWnd
End Sub

Private Sub CommandButton1_Click()
    FLAG_STOP = True
End Sub

