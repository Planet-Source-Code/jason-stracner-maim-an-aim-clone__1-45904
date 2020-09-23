VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00000000&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "MAIM Help"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   9285
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHelp 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmHelp.frx":0000
      Top             =   60
      Width           =   9195
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
  If Me.WindowState <> FormWindowStateConstants.vbMinimized Then
    txtHelp.Width = Me.ScaleWidth - 90
    txtHelp.Height = Me.ScaleHeight - 75
  End If
End Sub


Private Sub txtHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If txtHelp.SelText <> "" Then
    On Error Resume Next
    VB.Clipboard.SetText txtHelp.SelText
    Err.Clear
  End If
End Sub
