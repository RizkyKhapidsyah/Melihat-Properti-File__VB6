VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Melihat Properti File"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Propersties"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   2040
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub ShowProps(FileName As String, OwnerhWnd _
As Long)
Dim SEI As SHELLEXECUTEINFO
Dim r As Long
  With SEI
    .cbSize = Len(SEI)
    .fMask = SEE_MASK_NOCLOSEPROCESS Or _
    SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
    .hwnd = OwnerhWnd
    .lpVerb = "properties"
    .lpFile = FileName
    .lpParameters = vbNullChar
    .lpDirectory = vbNullChar
    .nShow = 0
    .hInstApp = 0
    .lpIDList = 0
  End With
  r = ShellExecuteEX(SEI)
End Sub

Private Sub Command1_Click()
  'Ganti 'FileContoh.txt' dengan nama file yang Anda
  'ingin lihat kotak dialog property-nya, dan letakkan
  'file tersebut satu direktori dengan source program ini (App.Path)
  Call ShowProps(App.Path + "\FileContoh.txt", Me.hwnd)
End Sub


