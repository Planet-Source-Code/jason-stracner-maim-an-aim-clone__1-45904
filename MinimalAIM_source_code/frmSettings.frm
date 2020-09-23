VERSION 5.00
Begin VB.Form frmSettings 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Settings"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkLogIMs 
      BackColor       =   &H00000000&
      Caption         =   "Log IMs"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   17
      Top             =   1980
      Width           =   3795
   End
   Begin VB.TextBox txtNotification 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   5040
      Width           =   4695
   End
   Begin VB.CheckBox chkAutoSwichOnIM 
      BackColor       =   &H00000000&
      Caption         =   "Auto switch to and display user's IM screen when a new IM arrives."
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   720
      TabIndex        =   14
      Top             =   1440
      Width           =   3795
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Auto logon"
      ForeColor       =   &H00FFFFFF&
      Height          =   2115
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   4695
      Begin VB.TextBox txtAutoServerPort 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2400
         TabIndex        =   12
         Text            =   "80"
         Top             =   1620
         Width           =   2115
      End
      Begin VB.TextBox txtAutoServer 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2400
         TabIndex        =   10
         Text            =   "toc.oscar.aol.com"
         Top             =   1080
         Width           =   2115
      End
      Begin VB.TextBox txtAutoPassword 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   180
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1620
         Width           =   2115
      End
      Begin VB.TextBox txtAutoUsername 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Top             =   1080
         Width           =   2115
      End
      Begin VB.CheckBox chkAutoLogon 
         BackColor       =   &H00000000&
         Caption         =   "Auto logon with the following setting."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   180
         TabIndex        =   5
         Top             =   420
         Width           =   2835
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "Server port:"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2400
         TabIndex        =   13
         Top             =   1380
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "Server:"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2400
         TabIndex        =   11
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Password:"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   9
         Top             =   1380
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Username:"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   7
         Top             =   840
         Width           =   1935
      End
   End
   Begin VB.CheckBox chkEnablePlugins 
      BackColor       =   &H00000000&
      Caption         =   "Enable plugins"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   480
      Width           =   2895
   End
   Begin VB.CheckBox chkMinimizeToTray 
      BackColor       =   &H00000000&
      Caption         =   "Minimize to system tray"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   2955
   End
   Begin VB.CommandButton cmdSettings 
      BackColor       =   &H00808080&
      Caption         =   "Save Settings"
      Height          =   375
      Left            =   3120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5940
      UseMaskColor    =   -1  'True
      Width           =   1875
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Notify me when any of these users come online (use commas to seperate names):"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   16
      Top             =   4620
      Width           =   4515
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Plugins must be in a subfolder called 'plugins'."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1020
      TabIndex        =   3
      Top             =   840
      Width           =   3315
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSettings_Click()
  Dim iEachNotificationUser As Long
  Dim sCheckForUsers2Delete As String
  
  Call WriteINIString("settings", "enable plugins", Me.chkEnablePlugins.Value, App.Path & "\maim_settings.ini")
  Call WriteINIString("settings", "use tray", Me.chkMinimizeToTray, App.Path & "\maim_settings.ini")
  Call WriteINIString("settings", "auto display IMs", Me.chkAutoSwichOnIM, App.Path & "\maim_settings.ini")
  Call WriteINIString("settings", "auto login", Me.chkAutoLogon.Value, App.Path & "\maim_settings.ini")
  Call WriteINIString("settings", "auto login username", Me.txtAutoUsername.Text, App.Path & "\maim_settings.ini")
  Call WriteINIString("settings", "auto login password", functionEncryptText(Me.txtAutoPassword.Text), App.Path & "\maim_settings.ini")
  Call WriteINIString("settings", "auto login server", Me.txtAutoServer.Text, App.Path & "\maim_settings.ini")
  Call WriteINIString("settings", "auto login server port", Me.txtAutoServerPort.Text, App.Path & "\maim_settings.ini")
  Call WriteINIString("settings", "log ims", Me.chkLogIMs, App.Path & "\maim_settings.ini")
  
  arrsNotificationList() = Split(Me.txtNotification.Text, ",")
  'fix some possible problems in the input
  For iEachNotificationUser = LBound(arrsNotificationList) To UBound(arrsNotificationList)
    'no spaces
    arrsNotificationList(iEachNotificationUser) = Replace(arrsNotificationList(iEachNotificationUser), " ", "")
    'lowercase
    arrsNotificationList(iEachNotificationUser) = LCase(arrsNotificationList(iEachNotificationUser))
    'no newlines
    arrsNotificationList(iEachNotificationUser) = Replace(arrsNotificationList(iEachNotificationUser), vbNewLine, "")
    'write it
    Call WriteINIString("settings", _
          "notification list buddy " & iEachNotificationUser, _
          arrsNotificationList(iEachNotificationUser), _
          App.Path & "\maim_settings.ini")
  Next iEachNotificationUser
  
  'We may have intended to remove someone from the end of the list.
  'We should make sure that there are no entrys beyond those that
  'should be there.
  Do
    sCheckForUsers2Delete = GetINIString("settings", _
          "notification list buddy " & iEachNotificationUser, _
          App.Path & "\maim_settings.ini", "")
    If sCheckForUsers2Delete <> "" Then
      'this shouldn't be there
      Call WriteINIString("settings", _
            "notification list buddy " & iEachNotificationUser, _
            "", _
            App.Path & "\maim_settings.ini")
    Else
      Exit Do
    End If
    iEachNotificationUser = iEachNotificationUser + 1
  Loop Until sCheckForUsers2Delete = ""
  
  'This makes sure that the array has been
  'setup so that you can do a ubound on it
  'without getting the dreaded "Subscript out of bounds" error.
  'I wish I knew a better way to do this- yuck.
  On Error Resume Next
  iEachNotificationUser = UBound(arrsNotificationList())
  If Err Then
    ReDim arrsNotificationList(0)
    Err.Clear
  End If
  On Error GoTo 0

  If Me.chkLogIMs.Value = vbChecked Then
    fLogIMs = True
  Else
    fLogIMs = False
  End If
  If Me.chkEnablePlugins.Value = vbChecked Then
    fEnablePlugins = True
  Else
    fEnablePlugins = False
  End If
  If Me.chkMinimizeToTray.Value = vbChecked Then
    fUseTray = True
  Else
    fUseTray = False
  End If
  If Me.chkAutoSwichOnIM.Value = vbChecked Then
    fAutoIMDisplayMode = True
  Else
    fAutoIMDisplayMode = False
  End If
  Unload Me
End Sub

Private Sub Form_Load()
  Dim iEachNotificationUser As Long
  Dim sCurrentNotificationListBuddy As String
  fAutoLogin = GetINIString("settings", "auto login", App.Path & "\maim_settings.ini", "0") = "1"
  sAutoLoginUsername = GetINIString("settings", "auto login username", App.Path & "\maim_settings.ini", "")
  sAutoLoginPassword = GetINIString("settings", "auto login password", App.Path & "\maim_settings.ini", "")
  sAutoLoginPassword = functionDecryptText(sAutoLoginPassword)
  sAutoLoginServer = GetINIString("settings", "auto login server", App.Path & "\maim_settings.ini", "")
  sAutoLoginServerPort = GetINIString("settings", "auto login server port", App.Path & "\maim_settings.ini", "")
  fEnablePlugins = GetINIString("settings", "enable plugins", App.Path & "\maim_settings.ini", "0") = "1"
  fUseTray = GetINIString("settings", "use tray", App.Path & "\maim_settings.ini", "0") = "1"
  fAutoIMDisplayMode = GetINIString("settings", "auto display IMs", App.Path & "\maim_settings.ini", "0") = "1"
  fLogIMs = GetINIString("settings", "log ims", App.Path & "\maim_settings.ini", "0") = "1"
  
  iEachNotificationUser = 0
  Do
    sCurrentNotificationListBuddy = GetINIString("settings", "notification list buddy " & iEachNotificationUser, App.Path & "\maim_settings.ini", "")
    If sCurrentNotificationListBuddy = "" Then
      Exit Do
    Else
      ReDim Preserve arrsNotificationList(iEachNotificationUser)
      arrsNotificationList(iEachNotificationUser) = sCurrentNotificationListBuddy
      iEachNotificationUser = iEachNotificationUser + 1
    End If
  Loop

  'This makes sure that the array has been
  'setup so that you can do a ubound on it
  'without getting the dreaded "Subscript out of bounds" error.
  'I wish I knew a better way to do this- yuck.
  On Error Resume Next
  iEachNotificationUser = UBound(arrsNotificationList())
  If Err Then
    ReDim arrsNotificationList(0)
    Err.Clear
  End If
  On Error GoTo 0

  If fLogIMs Then
    Me.chkLogIMs.Value = vbChecked
  End If
  Me.txtNotification.Text = Join(arrsNotificationList, ", ")
  
  If fAutoIMDisplayMode Then
    Me.chkAutoSwichOnIM.Value = vbChecked
  End If
  If fEnablePlugins Then
    Me.chkEnablePlugins.Value = vbChecked
  End If
  If fUseTray Then
    Me.chkMinimizeToTray.Value = vbChecked
  End If
  If fAutoLogin Then
    Me.chkAutoLogon.Value = vbChecked
  End If
  If sAutoLoginServer <> "" Then
    Me.txtAutoServer.Text = MMain.sAutoLoginServer
  End If
  If sAutoLoginServerPort <> "" Then
    Me.txtAutoServerPort.Text = MMain.sAutoLoginServerPort
  End If
  If sAutoLoginUsername <> "" Then
    Me.txtAutoUsername = MMain.sAutoLoginUsername
  End If
  If sAutoLoginPassword <> "" Then
    Me.txtAutoPassword = MMain.sAutoLoginPassword
  End If
End Sub
