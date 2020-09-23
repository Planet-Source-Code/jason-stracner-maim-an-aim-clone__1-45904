VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmChat 
   BackColor       =   &H00000000&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Chat Messages"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4620
   ForeColor       =   &H00FFFFFF&
   ScaleHeight     =   3540
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer txtCheckForMore 
      Interval        =   30
      Left            =   3960
      Top             =   3060
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close Window"
      Default         =   -1  'True
      Height          =   435
      Left            =   60
      TabIndex        =   1
      Top             =   3060
      Width           =   4515
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2955
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4515
      ExtentX         =   7964
      ExtentY         =   5212
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
  Unload Me
  End
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> FormWindowStateConstants.vbMinimized Then
    On Error Resume Next
    Me.WebBrowser1.Width = Me.ScaleWidth - 105
    Me.WebBrowser1.Height = Me.ScaleHeight - 585
    Me.cmdClose.Width = Me.WebBrowser1.Width
    Me.cmdClose.Top = Me.ScaleHeight - 480
  End If
End Sub

Private Sub txtCheckForMore_Timer()
  Call addTextToBrowser
End Sub

