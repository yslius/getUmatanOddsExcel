VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Wait 
   Caption         =   "���擾���ł�..."
   ClientHeight    =   1485
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4890
   OleObjectBlob   =   "UserForm_Wait.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm_Wait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
  Dim hWnd As Long
  Dim lngWstyle As Long
 
  '���[�U�[�t�H�[���̃n���h��
  hWnd = FindWindow(vbNullString, Me.Caption)
  lngWstyle = GetWindowLong(hWnd, GWL_STYLE)
 
  '����{�^���̏���
  SetWindowLong hWnd, GWL_STYLE, lngWstyle And (Not WS_SYSMENU)
  '���j���[�ĕ`��
  DrawMenuBar hWnd
End Sub

Private Sub CommandButton1_Click()
    FLAG_STOP = True
End Sub

