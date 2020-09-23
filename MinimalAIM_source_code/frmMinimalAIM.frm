VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMinimalAIM 
   BackColor       =   &H00000000&
   Caption         =   "MAIM"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9285
   Icon            =   "frmMinimalAIM.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtDataViews 
      Height          =   1875
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Visible         =   0   'False
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   3307
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmMinimalAIM.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdDataView 
      BackColor       =   &H00808080&
      Caption         =   "&V"
      Height          =   375
      Left            =   8940
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Click here to change to a different view of the data."
      Top             =   2760
      UseMaskColor    =   -1  'True
      Width           =   315
   End
   Begin VB.Timer tmrAutoConnect 
      Interval        =   500
      Left            =   8760
      Top             =   540
   End
   Begin MSWinsockLib.Winsock wskAIM 
      Left            =   8760
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSettings 
      BackColor       =   &H00808080&
      Caption         =   "&S"
      Height          =   375
      Left            =   8580
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Click here to adjust your settings."
      Top             =   2760
      UseMaskColor    =   -1  'True
      Width           =   315
   End
   Begin VB.ComboBox cboInput 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   2760
      Width           =   5775
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H00808080&
      Caption         =   "&?"
      Height          =   375
      Left            =   8220
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Click here for help with commands."
      Top             =   2760
      UseMaskColor    =   -1  'True
      Width           =   315
   End
   Begin VB.TextBox txtInputPassword 
      BackColor       =   &H00000000&
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
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   60
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   8115
   End
   Begin RichTextLib.RichTextBox txtMessages 
      Height          =   2295
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   4048
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmMinimalAIM.frx":038A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Pop up"
      Visible         =   0   'False
      Begin VB.Menu mnuViewAllData 
         Caption         =   "All data"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewWhosOnline 
         Caption         =   "Who's online"
      End
      Begin VB.Menu mnuViewOnlyChatMessages 
         Caption         =   "Chat in "
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewChatUsers 
         Caption         =   "Chat users in "
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuIMsFromUser 
         Caption         =   "IMs w/"
         Index           =   0
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMinimalAIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'% Integer
'& Long
'! Single
'# Double
'$ String
'@ Currency
Dim iLastSelStartInInput          As Long
Dim iLastSelLengthInInput         As Long
Dim sServer                       As String
Dim sUsername                     As String
Dim sPassword                     As String
Dim sPort                         As String
Dim iLoginStep                    As Long
Dim fLoggedin                     As Boolean
Dim iCurrentMessageNumber         As Long
Public iPreMinimizedState         As Long
Dim sDataForWhosOnline            As String

Dim sDataForMessagesOnlyViews()   As String

Dim sDataForChatMessagesViews()   As String
Dim sDataForChatUsersViews()      As String

Dim sChatRoomNamesLookup()        As String
Dim sChatRoomIDsLookup()          As String

Dim sNameOfCurrentDataView        As String
Dim fLogonInvisable               As Boolean
Public sLongCommandToSend_from_frmLongCommands As String

Private Sub cboInput_KeyUp(KeyCode As Integer, Shift As Integer)
  Dim sUsernameForIMMode As String
  Dim sChatIDForChatMode As String
  
  If KeyCode = vbKeyL And Shift = 3 And wskAIM.State = sckConnected Then
    'crtl+shift+L
    Dim sLongCommandToSend As String
    
    Load frmLongCommands
    frmLongCommands.txtCommand.Text = cboInput.Text
    frmLongCommands.txtCommand.SelStart = Len(frmLongCommands.txtCommand.Text)
    frmLongCommands.txtCommand.SelLength = 0
    frmLongCommands.Show vbModal
    sLongCommandToSend = sLongCommandToSend_from_frmLongCommands
    sLongCommandToSend_from_frmLongCommands = ""
    If Len(sLongCommandToSend) > 0 Then
      If InStr(sNameOfCurrentDataView, "IMs w/") Then
        'if we are on the screen that just shows messages from a single user
        'then we can omit the send_im, username and quotes from the command.
        sUsernameForIMMode = Replace(sNameOfCurrentDataView, "IMs w/", "")
        Call sendTocCommand(2, "toc_send_im """ & sUsernameForIMMode & """ """ & NormalizeForToc(sLongCommandToSend) & """" & Chr(0))
      ElseIf InStr(sNameOfCurrentDataView, "Chat in ") Then
        'if we are on the screen that just shows messages for a chatroom
        'then we can omit the send_chat, chatroom id and quotes from the command.
        sChatIDForChatMode = Replace(sNameOfCurrentDataView, "Chat in ", "")
        sChatIDForChatMode = getChatNumberFromName(sChatIDForChatMode)
        Call sendTocCommand(2, "toc_send_chat " & sChatIDForChatMode & " """ & NormalizeForToc(sLongCommandToSend) & """" & Chr(0))
      Else
        Call sendTocCommand(2, "toc_" & sLongCommandToSend & Chr(0))
      End If
    End If
  End If
End Sub

Private Sub cmdDataView_Click()
  PopupMenu mnuPopUp, , Me.cmdDataView.Left, cmdDataView.Top + cmdDataView.Height
End Sub

Private Sub cmdHelp_Click()
  frmHelp.Show
End Sub

Private Sub cmdSettings_Click()
  frmSettings.Show vbModal
End Sub

Private Sub Form_Load()
  ReDim sDataForChatMessagesViews(0)
  ReDim sDataForChatUsersViews(0)
  ReDim sDataForMessagesOnlyViews(0)
  
  If GetINIString("settings", "top of window position", App.Path & "\maim_settings.ini", "") = "MAX" Then
    Me.WindowState = FormWindowStateConstants.vbMaximized
  Else
    Me.Top = GetINIString("settings", "top of window position", App.Path & "\maim_settings.ini", Me.Top)
    Me.Left = GetINIString("settings", "left edge of window position", App.Path & "\maim_settings.ini", Me.Left)
    Me.Width = GetINIString("settings", "window width", App.Path & "\maim_settings.ini", Me.Width)
    Me.Height = GetINIString("settings", "window height", App.Path & "\maim_settings.ini", Me.Height)
  End If
  
  Call doDataDisplay("Welcome to {{-SetColor=" & vbGreen & "-}}MAIM{{-SetColor=" & vbYellow & "-}}.", vbYellow)
  Call doDataDisplay("build: " & App.Major & "." & App.Minor & "." & App.Revision, vbYellow)
  If Not fAutoLogin Then
    Call doDataDisplay("A minimal text based TOC client written by Jason Stracner.", vbYellow)
    Call doDataDisplay("MAIM lets you to see every piece of data that is sent and", vbYellow)
    Call doDataDisplay("recieved.", vbYellow)
    Call doDataDisplay(" ", vbYellow)
    Call doDataDisplay("What server would you like to connect to (Just press <enter> for toc.oscar.aol.com)?", vbYellow)
  Else
    Me.Visible = False
    Load frmTray
    Call doDataDisplay("Auto connecting to " & _
              MMain.sAutoLoginServer & " on port " & _
              MMain.sAutoLoginServerPort & " as " & _
              MMain.sAutoLoginUsername & ".", vbYellow)
    sUsername = MMain.sAutoLoginUsername
    sPassword = MMain.sAutoLoginPassword
    sServer = MMain.sAutoLoginServer
    sPort = MMain.sAutoLoginServerPort
    fLoggedin = True
    Call doDataDisplay("Attepting to connect.", vbYellow)
    doLogin
  End If
  
  cboInput.SelStart = 0
  cboInput.SelLength = Len(cboInput.Text)
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> FormWindowStateConstants.vbMinimized Then
    iPreMinimizedState = Me.WindowState
    On Error Resume Next
    txtMessages.Height = Me.ScaleHeight - 510
    txtMessages.Width = Me.ScaleWidth - 90
    txtDataViews.Height = txtMessages.Height
    txtDataViews.Width = txtMessages.Width
    cboInput.Top = Me.ScaleHeight - 405
    cboInput.Width = Me.ScaleWidth - 1170
    txtInputPassword.Top = cboInput.Top
    txtInputPassword.Width = cboInput.Width
    cmdHelp.Top = cboInput.Top
    cmdHelp.Left = Me.ScaleWidth - 1065
    cmdSettings.Top = cmdHelp.Top
    cmdSettings.Left = Me.ScaleWidth - 705
    cmdDataView.Top = cmdHelp.Top
    cmdDataView.Left = Me.ScaleWidth - 345
    txtMessages.SelStart = Len(txtMessages.Text)
    Err.Clear
  Else 'is min.
    If fUseTray Then
      Me.Visible = False
      Load frmTray
    End If
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call oAutomation.Quit
  Set oAutomation = Nothing
  If Me.WindowState = FormWindowStateConstants.vbNormal Then
    Call WriteINIString("settings", "top of window position", Me.Top, App.Path & "\maim_settings.ini")
    Call WriteINIString("settings", "left edge of window position", Me.Left, App.Path & "\maim_settings.ini")
    Call WriteINIString("settings", "window width", Me.Width, App.Path & "\maim_settings.ini")
    Call WriteINIString("settings", "window height", Me.Height, App.Path & "\maim_settings.ini")
  ElseIf Me.WindowState = FormWindowStateConstants.vbMaximized Then
    Call WriteINIString("settings", "top of window position", "MAX", App.Path & "\maim_settings.ini")
  End If
  End
End Sub

Private Sub cboInput_GotFocus()
  cboInput.SelStart = 0
  cboInput.SelLength = Len(cboInput.Text)
End Sub

Private Sub cboInput_KeyPress(KeyAscii As Integer)
  Dim sUsernameForIMMode As String
  Dim sChatIDForChatMode As String
  
  If KeyAscii = 13 Then
    If fLoggedin = False Then
      doLoginSteps
    Else
      If wskAIM.State = sckConnected Then
        If cboInput.Text = "signoff" Then
          Me.wskAIM.Close
          iLoginStep = 0
          fLoggedin = False
          cboInput.Text = ""
          Call doDataDisplay(" ", vbYellow)
          Call doDataDisplay("You are no longer connected.  Please reconnect.", vbRed)
          Call doDataDisplay("What server would you like to connect to (Just press <enter> for toc.oscar.aol.com)?", vbYellow)
          Call updateWhosOnlineData("", "", "", "", "", "", True) 'clear who's online screen data
          Exit Sub
        ElseIf cboInput.Text = "cls" Then
          'clear the screen
          txtMessages.Text = ""
          cboInput.Text = ""
          Exit Sub
        End If
        
        If InStr(sNameOfCurrentDataView, "IMs w/") Then
          'if we are on the screen that just shows messages from a single user
          'then we can omit the send_im, username and quotes from the command.
          sUsernameForIMMode = Replace(sNameOfCurrentDataView, "IMs w/", "")
          Call sendTocCommand(2, "toc_send_im """ & sUsernameForIMMode & """ """ & NormalizeForToc(cboInput.Text) & """" & Chr(0))
        ElseIf InStr(sNameOfCurrentDataView, "Chat in ") Then
          'if we are on the screen that just shows messages for a chatroom
          'then we can omit the send_chat, chatroom id and quotes from the command.
          sChatIDForChatMode = Replace(sNameOfCurrentDataView, "Chat in ", "")
          sChatIDForChatMode = getChatNumberFromName(sChatIDForChatMode)
          Call sendTocCommand(2, "toc_send_chat " & sChatIDForChatMode & " """ & NormalizeForToc(cboInput.Text) & """" & Chr(0))
        Else
          Call sendTocCommand(2, "toc_" & cboInput.Text & Chr(0))
        End If
        
      Else
        'Login again.  We must have been disconnected.
        Call doDataDisplay(" ", vbYellow)
        Call doDataDisplay("You are no longer connected.  Please reconnect.", vbRed)
        Call doDataDisplay("What server would you like to connect to (Just press <enter> for toc.oscar.aol.com)?", vbYellow)
        cboInput.Text = ""
        iLoginStep = 0
        fLoggedin = False
      End If
    End If
    If cboInput.Text <> "" Then
      cboInput.AddItem cboInput.Text, 0
    End If
    cboInput.Text = ""
    KeyAscii = 0
  End If
End Sub

Sub doLoginSteps()
  If iLoginStep = 0 Then
    If cboInput.Text = "" Or cboInput.Text = "Type here." Then
      sServer = "toc.oscar.aol.com"
    Else
      sServer = LCase(cboInput.Text)
    End If
    Call doDataDisplay(sServer, vbYellow)
    cboInput.Text = ""
    'Next step
    Call doDataDisplay("What port on that server would you like to connect to (Just press <enter> for 80)?", vbYellow)
    iLoginStep = iLoginStep + 1
  ElseIf iLoginStep = 1 Then
    If cboInput.Text = "" Then
      sPort = "80"
    Else
      sPort = cboInput.Text
    End If
    Call doDataDisplay(sPort, vbYellow)
    cboInput.Text = ""
    If Not IsNumeric(sPort) Then
      Call doDataDisplay("Error: The port number has to be a number.  Try again:", vbYellow)
      Exit Sub
    End If
    'Next step
    Call doDataDisplay("Enter your username (note: add the text 'hideme' after your name to start invisabe - like 'myusername hideme'):", vbYellow)
    iLoginStep = iLoginStep + 1
  ElseIf iLoginStep = 2 Then
    sUsername = cboInput.Text
    sUsername = LCase(sUsername)
    Call doDataDisplay(sUsername, vbYellow)
    'Next step
    Call doDataDisplay("Enter your password:", vbYellow)
    cboInput.Text = ""
    txtInputPassword.Text = ""
    txtInputPassword.Visible = True
    cboInput.Visible = False
    txtInputPassword.SetFocus
    iLoginStep = iLoginStep + 1
  ElseIf iLoginStep = 3 Then
    sPassword = txtInputPassword.Text
    sPassword = encryptPasswordForTOCServer(sPassword)
    txtMessages = txtMessages & " *************"
    cboInput.Text = ""
    txtInputPassword.Text = ""
    txtInputPassword.Visible = False
    cboInput.Visible = True
    cboInput.SetFocus
    fLoggedin = True
    Call doLogin
  End If
End Sub

Sub doLogin()
  Randomize
  iCurrentMessageNumber = Int((65535 * Rnd) + 1)
  wskAIM.Close
  wskAIM.RemoteHost = sServer
  wskAIM.RemotePort = sPort
  wskAIM.Connect
End Sub

Private Sub cboInput_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Shift = 2 Then
    'The ctrl key was down so past in the data.
    On Error Resume Next
    
    cboInput.SelStart = iLastSelStartInInput
    cboInput.SelLength = iLastSelLengthInInput
    cboInput.SelText = VB.Clipboard.GetText
    Err.Clear
  End If
End Sub

Private Sub cboInput_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  iLastSelLengthInInput = cboInput.SelLength
  iLastSelStartInInput = cboInput.SelStart
End Sub

Private Sub mnuIMsFromUser_Click(Index As Integer)
  Dim iEachIMMenu As Long
  sNameOfCurrentDataView = mnuIMsFromUser(Index).Caption
  Me.txtMessages.Visible = False
  Me.txtDataViews.Visible = True
  Me.mnuViewAllData.Checked = False
  Me.mnuViewWhosOnline.Checked = False
  
  For iEachIMMenu = mnuIMsFromUser.LBound To mnuIMsFromUser.UBound
    Me.mnuIMsFromUser(iEachIMMenu).Checked = False
  Next iEachIMMenu
  Me.mnuIMsFromUser(Index).Checked = True
  For iEachIMMenu = mnuViewOnlyChatMessages.LBound To mnuViewOnlyChatMessages.UBound
    Me.mnuViewOnlyChatMessages(iEachIMMenu).Checked = False
  Next iEachIMMenu
  For iEachIMMenu = mnuViewChatUsers.LBound To mnuViewChatUsers.UBound
    Me.mnuViewChatUsers(iEachIMMenu).Checked = False
  Next iEachIMMenu

  txtDataViews.Text = ""
  Call doDataDisplay(sDataForMessagesOnlyViews(Index), vbWhite, , txtDataViews, True)
  If Me.Visible Then cboInput.SetFocus
End Sub

Private Sub mnuViewAllData_Click()
  Dim iEachIMMenu As Long
  
  sNameOfCurrentDataView = "" 'all
  Me.txtMessages.Visible = True
  Me.txtDataViews.Visible = False
  
  Me.mnuViewAllData.Checked = True
  Me.mnuViewWhosOnline.Checked = False
  
  For iEachIMMenu = mnuIMsFromUser.LBound To mnuIMsFromUser.UBound
    Me.mnuIMsFromUser(iEachIMMenu).Checked = False
  Next iEachIMMenu
  For iEachIMMenu = mnuViewOnlyChatMessages.LBound To mnuViewOnlyChatMessages.UBound
    Me.mnuViewOnlyChatMessages(iEachIMMenu).Checked = False
  Next iEachIMMenu
  For iEachIMMenu = mnuViewChatUsers.LBound To mnuViewChatUsers.UBound
    Me.mnuViewChatUsers(iEachIMMenu).Checked = False
  Next iEachIMMenu
  
  If Me.Visible Then cboInput.SetFocus
End Sub

Private Sub mnuViewChatUsers_Click(Index As Integer)
  Dim iEachIMMenu As Long
  
  sNameOfCurrentDataView = mnuViewChatUsers(Index).Caption
  Me.txtMessages.Visible = False
  Me.txtDataViews.Visible = True
  
  Me.mnuViewAllData.Checked = False
  Me.mnuViewWhosOnline.Checked = False
  
  For iEachIMMenu = mnuIMsFromUser.LBound To mnuIMsFromUser.UBound
    Me.mnuIMsFromUser(iEachIMMenu).Checked = False
  Next iEachIMMenu
  For iEachIMMenu = mnuViewOnlyChatMessages.LBound To mnuViewOnlyChatMessages.UBound
    Me.mnuViewOnlyChatMessages(iEachIMMenu).Checked = False
  Next iEachIMMenu
  For iEachIMMenu = mnuViewChatUsers.LBound To mnuViewChatUsers.UBound
    Me.mnuViewChatUsers(iEachIMMenu).Checked = False
  Next iEachIMMenu
  Me.mnuViewChatUsers(Index).Checked = True
  
  txtDataViews.Text = ""
  Call doDataDisplay(sDataForChatUsersViews(Index), vbCyan, , txtDataViews, True)
  If Me.Visible Then cboInput.SetFocus
End Sub

Private Sub mnuViewOnlyChatMessages_Click(Index As Integer)
  Dim iEachIMMenu As Long
  
  sNameOfCurrentDataView = mnuViewOnlyChatMessages(Index).Caption
  Me.txtMessages.Visible = False
  Me.txtDataViews.Visible = True
  
  Me.mnuViewAllData.Checked = False
  Me.mnuViewWhosOnline.Checked = False

  For iEachIMMenu = mnuIMsFromUser.LBound To mnuIMsFromUser.UBound
    Me.mnuIMsFromUser(iEachIMMenu).Checked = False
  Next iEachIMMenu
  For iEachIMMenu = mnuViewOnlyChatMessages.LBound To mnuViewOnlyChatMessages.UBound
    Me.mnuViewOnlyChatMessages(iEachIMMenu).Checked = False
  Next iEachIMMenu
  For iEachIMMenu = mnuViewChatUsers.LBound To mnuViewChatUsers.UBound
    Me.mnuViewChatUsers(iEachIMMenu).Checked = False
  Next iEachIMMenu
  Me.mnuViewOnlyChatMessages(Index).Checked = True

  txtDataViews.Text = ""
  Call doDataDisplay(sDataForChatMessagesViews(Index), vbGreen, , txtDataViews, True)
  If Me.Visible Then cboInput.SetFocus
End Sub

Private Sub mnuViewWhosOnline_Click()
  Dim iEachIMMenu As Long
  
  sNameOfCurrentDataView = "who's online"
  Me.txtMessages.Visible = False
  Me.txtDataViews.Visible = True
  
  Me.mnuViewAllData.Checked = False
  Me.mnuViewWhosOnline.Checked = True

  For iEachIMMenu = mnuIMsFromUser.LBound To mnuIMsFromUser.UBound
    Me.mnuIMsFromUser(iEachIMMenu).Checked = False
  Next iEachIMMenu
  For iEachIMMenu = mnuViewOnlyChatMessages.LBound To mnuViewOnlyChatMessages.UBound
    Me.mnuViewOnlyChatMessages(iEachIMMenu).Checked = False
  Next iEachIMMenu
  For iEachIMMenu = mnuViewChatUsers.LBound To mnuViewChatUsers.UBound
    Me.mnuViewChatUsers(iEachIMMenu).Checked = False
  Next iEachIMMenu

  txtDataViews.Text = ""
  Call doDataDisplay(sDataForWhosOnline, vbCyan, , txtDataViews, True)
  If Me.Visible Then cboInput.SetFocus
End Sub

Private Sub tmrAutoConnect_Timer()
  If fAutoLogin Then
    If wskAIM.State = sckError Then
      Call doLogin
    End If
  End If
End Sub

Private Sub txtDataViews_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If txtDataViews.SelText <> "" Then
    On Error Resume Next
    VB.Clipboard.SetText txtDataViews.SelText
    Err.Clear
  End If
End Sub

Private Sub txtInputPassword_KeyPress(KeyAscii As Integer)
  cboInput_KeyPress KeyAscii
End Sub

Private Sub txtMessages_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If txtMessages.SelText <> "" Then
    On Error Resume Next
    VB.Clipboard.SetText txtMessages.SelText
    Err.Clear
  End If
End Sub

Private Sub wskAIM_Connect()
  'the FLAPON is our first message sent to the aim toc server after a connection is made.
  'from here we will a flap response containing the flap version.
  If wskAIM.State = sckConnected Then
    wskAIM.SendData "FLAPON" & vbNewLine & vbNewLine
    Call doDataDisplay(">>>: FLAPON" & vbNewLine & vbNewLine, vbYellow, True)
  End If
End Sub

Private Sub wskAIM_DataArrival(ByVal bytesTotal As Long)
  'this procedure is where all the data is handled. it is important for us to pay attention
  'to the flap headers since more than one command may be sent per packet. the payload in
  'the flap header is very important here. it allows us to know how much data is in that
  'command.
  Dim sDataSent As String, iCurrentPosition As Long, iDataLength As Long
  Dim iFrameType As Long, iSeqA As Long, iSeqB As Long
  Dim iPayLo As Long, iPayHi As Long, iPayload As Long
  Dim sCommandPortion As String
  wskAIM.GetData sDataSent$, vbString
  Debug.Print sDataSent$
  
  iDataLength& = Len(sDataSent$)
  iCurrentPosition& = 1
  Do While iCurrentPosition& < iDataLength&
    iFrameType& = Asc(Mid(sDataSent$, iCurrentPosition& + 1))
    iSeqA& = Asc(Mid(sDataSent$, iCurrentPosition& + 2))
    iSeqB& = Asc(Mid(sDataSent$, iCurrentPosition& + 3))
    iPayLo& = Asc(Mid(sDataSent$, iCurrentPosition& + 4))
    iPayHi& = Asc(Mid(sDataSent$, iCurrentPosition& + 5))
    iPayload& = MakeLong(iPayHi&, iPayLo&)
    sCommandPortion$ = Mid(sDataSent$, iCurrentPosition& + 6, iPayload&)
    Call doDataDisplay("<<<: *{\" & getZeroPadToThree(iFrameType&) & _
                "," & getZeroPadToThree(iSeqA&) & _
                "," & getZeroPadToThree(iSeqB&) & _
                "," & getZeroPadToThree(iPayLo&) & _
                "," & getZeroPadToThree(iPayHi&) & _
                "}" & sCommandPortion$, vbWhite, True)
    Call interpretIncommingData(iFrameType&, sCommandPortion$)
    'here we have seperated a command from the incoming data. we will call the handlproc
    'procedure for each command found in the incoming data.
    iCurrentPosition& = iCurrentPosition& + iPayload& + 6
  Loop
End Sub

Public Sub doDataDisplay( _
              ByVal sDataToDisplay As String, _
              ByVal iColor As Long, _
              Optional ByVal fSendTimeStamp As Boolean = False, _
              Optional txtTextbox2Update As RichTextBox, _
              Optional fDoNotEscapeThis As Boolean = False)
  Dim iEachLetter As Long
  Dim sMessageText As String
  
  If txtTextbox2Update Is Nothing Then
    Set txtTextbox2Update = Me.txtMessages
  End If
  
  If Not fDoNotEscapeThis Then
    For iEachLetter = 0 To 31
      'If iEachLetter <> 10 And iEachLetter <> 13 Then ' This was a bad idea so I took it out.
      sDataToDisplay = Replace(sDataToDisplay, Chr(iEachLetter), "{\" & getZeroPadToThree(iEachLetter) & "}")
      'End If
    Next iEachLetter
    
    For iEachLetter = 127 To 144
      sDataToDisplay = Replace(sDataToDisplay, Chr(iEachLetter), "{\" & getZeroPadToThree(iEachLetter) & "}")
    Next iEachLetter
    
    For iEachLetter = 147 To 255
      sDataToDisplay = Replace(sDataToDisplay, Chr(iEachLetter), "{\" & getZeroPadToThree(iEachLetter) & "}")
    Next iEachLetter
  End If
  'this is not really needed now that I think about it.
''  'Replace lone cr's and lf's.
''  sDataToDisplay = Replace(sDataToDisplay, vbNewLine, "{vbNewLine}")
''  sDataToDisplay = Replace(sDataToDisplay, vbLf, "{vbNewLine}")
''  sDataToDisplay = Replace(sDataToDisplay, vbCr, "{vbNewLine}")
''  sDataToDisplay = Replace(sDataToDisplay, "{vbNewLine}", vbNewLine)
  
  'Make sure the password doesn't ever show up on the screen.
  sDataToDisplay = Replace(sDataToDisplay, sPassword, "<password removed>")
  
  If fSendTimeStamp Then
    'Insert a time stamp:
    sDataToDisplay = getTimeStamp & " " & sDataToDisplay
  End If
  
  'Ignore keep alive messages.
  'Check the data to see if it is a 'keep alive' message
  If InStr(sDataToDisplay, ",000,004}{\000}{\000}{\000}2") Then
    Exit Sub
  End If
  
'  I used this before I switched to a rich textbox.
'  'trim off the text at the top to make room for more text.
'  If Len(txtTextbox2Update.Text & vbNewLine & sDataToDisplay$) >= 65534 Then
'    txtTextbox2Update.Text = Mid(txtTextbox2Update.Text, Len(vbNewLine & sDataToDisplay$))
'    txtTextbox2Update.SelStart = Len(txtTextbox2Update.Text)
'  End If
  
  Dim iEachLine As Long
  Dim arrsLinesOfDataToDisplay() As String
  Dim arrsColorZonesOfLinesDataToDisplay() As String
  Dim sColorOfZone As String
  Dim sTextInColorZone As String
  Dim iEachZone As Long
  
  arrsLinesOfDataToDisplay() = Split(sDataToDisplay$, vbNewLine)
  For iEachLine = LBound(arrsLinesOfDataToDisplay) To UBound(arrsLinesOfDataToDisplay)
    'Like: {{-SetColor=16777215-}} for white {{SetColor=0}} for black.
    arrsColorZonesOfLinesDataToDisplay() = Split(arrsLinesOfDataToDisplay(iEachLine), "{{-SetColor")
    'set the inital color
    txtTextbox2Update.SelStart = Len(txtTextbox2Update.Text)
    txtTextbox2Update.SelLength = 0
    txtTextbox2Update.SelColor = iColor 'default color
    For iEachZone = LBound(arrsColorZonesOfLinesDataToDisplay) To UBound(arrsColorZonesOfLinesDataToDisplay)
      'Is there a color zone here?
      If Strings.Left(arrsColorZonesOfLinesDataToDisplay(iEachZone), 1) = "=" _
              And InStr(arrsColorZonesOfLinesDataToDisplay(iEachZone), "-}}") <= 10 Then
        txtTextbox2Update.SelStart = Len(txtTextbox2Update.Text)
        txtTextbox2Update.SelLength = 0
        'now we can be reasonably sure that this was meant to be a zone of a different color.
        sTextInColorZone = arrsColorZonesOfLinesDataToDisplay(iEachZone)
        'Get the color of the zone.
        sColorOfZone = Mid(sTextInColorZone, 2, InStr(sTextInColorZone, "-}}") - 2)
        If IsNumeric(sColorOfZone) Then
          'set the color
          txtTextbox2Update.SelColor = CLng(sColorOfZone)
        Else
          MsgBox "Error: couldn't change colors.", vbCritical, "Error"
        End If
        'take out the color command thing.
        sTextInColorZone = Mid(sTextInColorZone, InStr(sTextInColorZone, "-}}") + 3)
        'Add the text to the display
        txtTextbox2Update.SelText = sTextInColorZone
        sTextInColorZone = ""
      Else
        'this is not a color zone so we will just make it the default color passed in.
        txtTextbox2Update.SelText = arrsColorZonesOfLinesDataToDisplay(iEachZone)
      End If
    Next iEachZone
    'add a newline
    txtTextbox2Update.SelStart = Len(txtTextbox2Update.Text)
    txtTextbox2Update.SelLength = 0
    txtTextbox2Update.SelText = vbNewLine
  Next iEachLine
  txtTextbox2Update.SelStart = Len(txtTextbox2Update.Text)
End Sub

Public Function filterOutHTML(ByVal sInput As String) As String
  If InStr(sInput, "<") = 0 And InStr(sInput, ">") = 0 Then 'shortcut
    filterOutHTML = sInput
    Exit Function
  End If
  sInput$ = Replace(sInput$, "<HTML>", "", , , vbTextCompare)
  sInput$ = Replace(sInput$, "</HTML>", "", , , vbTextCompare)
  sInput$ = Replace(sInput$, "<SUP>", "", , , vbTextCompare)
  sInput$ = Replace(sInput$, "</SUP>", "", , , vbTextCompare)
  sInput$ = Replace(sInput$, "<HR>", "", , , vbTextCompare)
  sInput$ = Replace(sInput$, "<H1>", "", , , vbTextCompare)
  sInput$ = Replace(sInput$, "<H2>", "", , , vbTextCompare)
  sInput$ = Replace(sInput$, "<H3>", "", , , vbTextCompare)
  sInput$ = Replace(sInput$, "<PRE>", "", , , vbTextCompare)
  sInput$ = Replace(sInput$, "</PRE>", "", , , vbTextCompare)
  sInput$ = Replace(sInput$, "<PRE=", "", , , vbTextCompare) '?
  sInput$ = Replace(sInput$, "<B>", "", , , vbTextCompare)
  sInput$ = Replace(sInput$, "</B>", "", , , vbTextCompare)
  sInput$ = Replace(sInput$, "<U>", "", , , vbTextCompare)
  sInput$ = Replace(sInput$, "</U>", "", , , vbTextCompare)
  sInput$ = Replace(sInput$, "<I>", "", , , vbTextCompare)
  sInput$ = Replace(sInput$, "</I>", "", , , vbTextCompare)
  sInput$ = Replace(sInput$, "<FONT>", "", , , vbTextCompare)
  sInput$ = Replace(sInput$, "</FONT>", "", , , vbTextCompare)
  sInput$ = Replace(sInput$, "<BODY>", "", , , vbTextCompare)
  sInput$ = Replace(sInput$, "</BODY>", "", , , vbTextCompare)
  sInput$ = Replace(sInput$, "<BR>", "", , , vbTextCompare)
  sInput$ = Replace(sInput$, "</A>", "", , , vbTextCompare)
  
  sInput$ = Replace(sInput$, "&amp;", "&", , , vbTextCompare)
  sInput$ = Replace(sInput$, "&lt;", "<", , , vbTextCompare)
  sInput$ = Replace(sInput$, "&nbsp;", " ", , , vbTextCompare)
  sInput$ = Replace(sInput$, "&quot;", """", , , vbTextCompare) 'a quote
  sInput$ = Replace(sInput$, "&gt;", ">", , , vbTextCompare)
  
  If InStr(sInput, "<") = 0 And InStr(sInput, ">") = 0 Then 'shortcut
    filterOutHTML = sInput
    Exit Function
  End If
  
  sInput$ = removeMarkupTag(sInput$, "A")
  sInput$ = removeMarkupTag(sInput$, "Body")
  sInput$ = removeMarkupTag(sInput$, "FONT")
  
  filterOutHTML$ = sInput$
End Function

Private Function removeMarkupTag(ByVal sInput As String, ByVal sTag As String) As String
  Dim iStartingPoint As Long
  Dim iLength As Long
  Dim iEndPont As Long
  Dim sLeftPart As String
  Dim sRightPart As String
  Dim sOutput As String
  
  sOutput = sInput
  sTag = LCase(sTag)
  sOutput = Replace(sOutput, "<" & sTag, "<" & LCase(sTag), , , vbTextCompare) 'ensure lowercase
  iStartingPoint& = InStr(sOutput$, "<" & sTag) 'find opening tag
  iLength& = Len(sOutput$)
  Do While iStartingPoint& <> 0
    iEndPont& = InStr(iStartingPoint&, sOutput$, ">") 'fist occurance of >
    If iEndPont& <> 0 Then
      'Found one
      sLeftPart$ = Left(sOutput$, iStartingPoint& - 1)
      sRightPart$ = Right(sOutput$, iLength& - iEndPont&)
      sOutput$ = sLeftPart$ & sRightPart$
      iLength& = Len(sOutput$)
    End If
    iStartingPoint& = InStr(sOutput$, "<" & sTag)
    DoEvents
  Loop
  removeMarkupTag = sOutput 'return
End Function

Private Sub wskAIM_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  Dim sErrorToShow As String
  Dim iEachView As Long
  Dim sUser As String
  Dim sChatRoom As String
  
  If Number <> 11004 Then
    sErrorToShow = "EXPLANATION: Winsock error " & Number & " " & Description
    If Right(Me.txtMessages.Text, Len(sErrorToShow & vbNewLine)) <> sErrorToShow & vbNewLine Then
      'don't repeat the same error over and over
      Call doDataDisplay(sErrorToShow, vbRed, True)
      Call updateWhosOnlineData("", "", "", "", "", "", True) 'clear who's online screen data
      
      'Add text to im messages too.
      For iEachView = mnuViewOnlyChatMessages.LBound To mnuViewOnlyChatMessages.UBound
        If mnuViewOnlyChatMessages.UBound = 0 Then Exit For 'no items
        sChatRoom = mnuViewOnlyChatMessages(iEachView).Caption
        sChatRoom = Replace(sChatRoom, "Chat in ", "")
        sChatRoom = getChatNumberFromName(sChatRoom)
        If sChatRoom <> "" Then
          addChatMessageToChatDataView sChatRoom, "{{-SetColor=" & vbRed & "-}}" & " Disconnected from server."
          updateChatUsersData sChatRoom, "", "", True
        End If
      Next iEachView
      
      For iEachView = mnuIMsFromUser.LBound To mnuIMsFromUser.UBound
        If mnuIMsFromUser.UBound = 0 Then Exit For 'no items
        sUser = Me.mnuIMsFromUser(iEachView).Caption
        sUser = Replace(sUser, "IMs w/", "")
        If sUser <> "" Then
          Call addIMToIMOnlyDataViews(sUser, "{{-SetColor=" & vbRed & "-}}" & " Disconnected from server.")
          Call logIMs(sUser, "Disconnected from server.")
        End If
      Next iEachView
    End If
  End If


End Sub

Private Sub interpretIncommingData(ByVal iFrameType As Long, ByVal sRecivedData As String)
  Dim arrCommand() As String
  Dim arrCommandArguments() As String
  Dim sServerAuthorizerHost As String
  Dim sServerAuthorizerHostPort As String
  Dim sUserClass As String
  Dim sUserClassData As String
  Dim sLogonLogoffAnnouncement As String
  
  Select Case iFrameType&
    Case 1 'a frame type of "1" indicates this message is part of the signon sequence
      If sRecivedData$ = Chr(0) & Chr(0) & Chr(0) & Chr(1) Then
        'set fLogonInvisable variable
        If InStr(sUsername, " hideme") Then
          fLogonInvisable = True
          sUsername = Replace(sUsername, " hideme", "")
        Else
          fLogonInvisable = False
        End If
        sUsername = Replace(sUsername, " ", "")
        Call sendTocCommand(1, Chr(0) & Chr(0) & Chr(0) & Chr(1) & Chr(0) & Chr(1) & Chr(0) & Chr(CLng(CStr(Len(sUsername)))) & sUsername)
        Call doDataDisplay("EXPLANATION: Sent username.", 14715847)
        'here we send our flap version, tvl tag, normalized screen name length, and
        'our normalized screen name
        
        'This assumes that the authorizer host is the same as the login server name except "toc."
        'should be replaced with "login.".  toc.oscar.aol.com -> login.oscar.aol.com
        sServerAuthorizerHost = Mid(sServer, InStr(sServer, "toc.") + 4)
        sServerAuthorizerHost = "login." & sServerAuthorizerHost
        Call sendTocCommand(2, "toc_signon " & sServerAuthorizerHost & " " & "0234  " & sUsername & " " & sPassword & " english " & Chr(34) & "AOLInstantMessenger" & Chr(34) & Chr(0))
        Call doDataDisplay("EXPLANATION: toc_signon sent: " & _
                  "(authorizer host = " & sServerAuthorizerHost & ") " & _
                  "(authorizer port = 0234) " & _
                  "(user name = " & sUsername & ") " & _
                  "(roasted password = " & sPassword & ") " & _
                  "(language = english) " & _
                  "(version = AOLInstantMessenger).", 14715847)
      End If
    Case 2 'a frame type of "2" indicates normal data
      arrCommand$ = Split(sRecivedData$, ":", 2)
      Select Case UCase(arrCommand$(0))
        Case "GOTO_URL"
          arrCommandArguments$() = Split(arrCommand$(1), ":", 3)
          
          '<Window Name>:<Url>
          Call doDataDisplay("EXPLANATION: GOTO_URL (suggested new window name = " & _
                    arrCommandArguments$(0) & ") (url = " & _
                    arrCommandArguments$(1) & ").", ColorConstants.vbCyan)
          If fEnablePlugins And FileExists(App.Path & "\plugins\GOTO_URL.exe") Then
            oAutomation.PluginCommandData = sRecivedData
            Shell App.Path & "\plugins\GOTO_URL.exe"
          End If
        Case "RVOUS_PROPOSE"
          Call doDataDisplay("EXPLANATION: got rendezvous proposal.", ColorConstants.vbCyan)
          If fEnablePlugins And FileExists(App.Path & "\plugins\RVOUS_PROPOSE.exe") Then
            oAutomation.PluginCommandData = sRecivedData
            Shell App.Path & "\plugins\RVOUS_PROPOSE.exe"
          End If
        Case "CONFIG"
          Dim sNamesToSend As String
          Dim iNumberOfNames As Integer
          Dim iEachLetter As Long
          Dim sBuddyData As String
          Dim sConfigData As String
          Dim sConfigDataForNewMode As String
          
          sConfigData = arrCommand$(1)
          
          If InStr(sConfigData, "m 1") = 1 Then
            'Permit All
            Call doDataDisplay("EXPLANATION: You are in permit all mode.", 14715847)
          ElseIf InStr(sConfigData, "m 2") = 1 Then
            'Deny All
            Call doDataDisplay("EXPLANATION: You are in deny all mode.", 14715847)
          ElseIf InStr(sConfigData, "m 3") = 1 Then
            'Permit Some
            Call doDataDisplay("EXPLANATION: You are in permit some mode.  " & _
                      "User names that have a 'p' character in front " & _
                      "of their names will be allowed to see and send you messages.", 14715847)
          ElseIf InStr(sConfigData, "m 4") = 1 Then
            'Deny Some
            Call doDataDisplay("EXPLANATION: You are in deny some mode.  " & _
                      "User names that have a 'd' character in front " & _
                      "of their names will not be allowed to see or send you messages.", 14715847)
          End If
          
          If fLogonInvisable Then
            'this makes us invisable
            Call sendTocCommand(2, "toc_add_permit" & Chr(0))
            Call doDataDisplay("EXPLANATION: This will hide us from other users.", 14715847)
          Else
            'this ensures that we are visable
            Call sendTocCommand(2, "toc_add_deny" & Chr(0))
            Call doDataDisplay("EXPLANATION: This will ensure that we are visable to other users.", 14715847)
          End If
          
          Dim sLinesOfConfigData() As String
          Dim iEachLine As Long
          Dim sCurrentLineOfConfigData As String
          Dim sCurrentName As String
          
          sLinesOfConfigData() = Split(sConfigData, vbLf)
          iNumberOfNames = 0
          sNamesToSend = ""
          For iEachLine = LBound(sLinesOfConfigData()) To UBound(sLinesOfConfigData())
            sCurrentLineOfConfigData = sLinesOfConfigData(iEachLine) 'simplify
            If Left(sCurrentLineOfConfigData, 1) = "b" Then 'the b indicates that it is a buddy
              sCurrentName = sCurrentLineOfConfigData
              sCurrentName = Replace(sCurrentName, "b ", "", , 1) 'take off the b
              sCurrentName = Replace(sCurrentName, " ", "") 'remove spaces
              sCurrentName = Chr(34) & sCurrentName & Chr(34) 'add quotes
              sCurrentName = " " & sCurrentName 'add a space to the front
              sNamesToSend = sNamesToSend & sCurrentName 'add the names to the list
              iNumberOfNames = iNumberOfNames + 1
            End If
            If iNumberOfNames >= 10 Or iEachLine = UBound(sLinesOfConfigData()) Then
              'if this is the tenth name or if this is the last name then...
              If sNamesToSend <> "" Then
                Call sendTocCommand(2, "toc_add_buddy" & sNamesToSend & Chr(0))
                sNamesToSend = ""
                iNumberOfNames = 0
              End If
            End If
          Next iEachLine
          Call sendTocCommand(2, "toc_set_caps ""Chat, other stuff""" & Chr$(0))
          Call doDataDisplay("EXPLANATION: Capabilities are set to Chat only.", 14715847)

          Call sendTocCommand(2, "toc_init_done" & Chr(0))
          Dim sSetPersonalInformation As String
          
          sSetPersonalInformation = ""
          sSetPersonalInformation = "<HTML><BODY>I am using my own TOC protocol client called MAIM (written by: Jason Stracner).  He he he!</BODY></HTML>"
          Call sendTocCommand(2, "toc_set_info " & Chr(34) & escapeSpecialCharactersForTransmission(sSetPersonalInformation) & Chr(34) & Chr(0))
          Call doDataDisplay("EXPLANATION: Locate user information set to " & sSetPersonalInformation & ".", 14715847)
          If fEnablePlugins And FileExists(App.Path & "\plugins\CONFIG.exe") Then
            oAutomation.PluginCommandData = sRecivedData
            Shell App.Path & "\plugins\CONFIG.exe"
          End If
        Case "CHAT_IN"
          'incoming chat room text
          '<Chat Room Id>:<Source User>:<Whisper? T/F>:<Message>
          'argument 1: chat room id
          'argument 2: sender's screen name
          'argument 3: whisper t/f (not handled here)
          'argument 4: message
          arrCommandArguments$() = Split(arrCommand$(1), ":", 4)
          Call doDataDisplay("EXPLANATION: Got a chat message. (chatroom id=" & _
                    arrCommandArguments$(0) & ") (sender={{-SetColor=" & vbCyan & "-}}" & _
                    arrCommandArguments$(1) & "{{-SetColor=" & vbMagenta & "-}}) (whisper t/f=" & _
                    arrCommandArguments$(2) & ") (message={{-SetColor=" & vbCyan & "-}}" & _
                    filterOutHTML(arrCommandArguments$(3)) & "{{-SetColor=" & vbMagenta & "-}})", vbMagenta)
                    
          Call addChatMessageToChatDataView(arrCommandArguments$(0), _
                    padStringWithSpaces(arrCommandArguments$(1), 17) & ": " & _
                    filterOutHTML(arrCommandArguments$(3)))
          If fEnablePlugins And FileExists(App.Path & "\plugins\CHAT_IN.exe") Then
            oAutomation.PluginCommandData = sRecivedData
            Shell App.Path & "\plugins\CHAT_IN.exe"
          End If
        Case "CHAT_INVITE"
          'incoming invitation to a chat room
          'argument 1: chat room name
          'argument 2: chat room id
          'argument 3: invite sender
          'argument 4: invitation message
          arrCommandArguments$() = Split(arrCommand$(1), ":", 4)
          Call doDataDisplay("EXPLANATION: Got a chat invitation. (chatroom=" & _
                    arrCommandArguments$(0) & ") (chatroom id=" & _
                    arrCommandArguments$(1) & ") (sender={{-SetColor=" & vbCyan & "-}}" & _
                    arrCommandArguments$(2) & "{{-SetColor=" & 14715847 & "-}}) (message={{-SetColor=" & vbCyan & "-}}" & _
                    arrCommandArguments$(3) & "{{-SetColor=" & 14715847 & "-}})", 14715847)
          If fEnablePlugins And FileExists(App.Path & "\plugins\CHAT_INVITE.exe") Then
            oAutomation.PluginCommandData = sRecivedData
            Shell App.Path & "\plugins\CHAT_INVITE.exe"
          End If
        Case "CHAT_LEFT"
          Call doDataDisplay("EXPLANATION: Successfuly left the chatroom. (chatroom id=" & _
                    arrCommand$(1) & ")", 14715847)
          Call addChatMessageToChatDataView(arrCommand$(1), "EXPLANATION: Successfuly left the chatroom.")
          Call updateChatUsersData(arrCommand$(1), "", "", True)
          'sDataForChatMessagesView = ""
          'sDataForChatUsersView = ""
          'If sNameOfCurrentDataView = "chat users" Or sNameOfCurrentDataView = "chat messages" Then
          '  Me.txtDataViews.Text = ""
          'End If
        Case "CHAT_JOIN"
          'indicates that our attempt to join a chat room was successful
          'argument 1: chat room id
          'argument 2: chat room name
          arrCommandArguments$() = Split(arrCommand$(1), ":", 2)
          Call doDataDisplay("EXPLANATION: Chatroom join success. (chatroom id=" & _
                    arrCommandArguments$(0) & ") (chatroom=" & _
                    arrCommandArguments$(1) & ")", 14715847)
          Call addChatRoomNameLookup(arrCommandArguments$(1), arrCommandArguments$(0))
          Call addChatMessageToChatDataView(arrCommandArguments$(1), "{{-SetColor=" & vbYellow & "-}} Successfuly joined the chatroom.")
          'clear the user's data for the room
          
          If fEnablePlugins And FileExists(App.Path & "\plugins\CHAT_JOIN.exe") Then
            oAutomation.PluginCommandData = sRecivedData
            Shell App.Path & "\plugins\CHAT_JOIN.exe"
          End If
        Case "CHAT_UPDATE_BUDDY"
          'indicates that a user has either joined or parted a chat room
          'argument 1: chat room id
          'argument 2: joined t/f
          'argument 3: list of users joining or parting the room
          arrCommandArguments$() = Split(arrCommand$(1), ":", 3)
          Call doDataDisplay("EXPLANATION: Chatroom user update. (chatroom id=" & _
                    arrCommandArguments$(0) & ") (joined t/f=" & _
                    arrCommandArguments$(1) & ") (users=" & _
                    arrCommandArguments$(2) & ")", 14715847)
          Call updateChatUsersData(arrCommandArguments$(0), arrCommandArguments$(2), arrCommandArguments$(1))
          If fEnablePlugins And FileExists(App.Path & "\plugins\CHAT_UPDATE_BUDDY.exe") Then
            oAutomation.PluginCommandData = sRecivedData
            Shell App.Path & "\plugins\CHAT_UPDATE_BUDDY.exe"
          End If
        Case "ERROR"
          'indicates there was an error
          'argument 1: error id number
          'argument 2: varies depending on the error
          arrCommandArguments$() = Split(arrCommand$(1), ":", 2)
          Select Case arrCommandArguments$(0)
            Case "901"
              Call doDataDisplay("EXPLANATION: General Error: 901. " & arrCommandArguments$(1) & " is not currently available.", ColorConstants.vbRed)
            Case "902"
              Call doDataDisplay("EXPLANATION: General Error: 902. Warning of " & arrCommandArguments$(1) & " is not currently available.", ColorConstants.vbRed)
            Case "903"
              Call doDataDisplay("EXPLANATION: General Error: 903. A message has been dropped, you are exceeding the server speed limit.", ColorConstants.vbRed)
            Case "950"
              Call doDataDisplay("EXPLANATION: Chat Error: 950. Chat in " & arrCommandArguments$(1) & " is unavailable.", ColorConstants.vbRed)
            Case "960"
              Call doDataDisplay("EXPLANATION: IM/Info Error: 960. You are sending messages too fast to " & arrCommandArguments$(1) & ".", ColorConstants.vbRed)
            Case "961"
              Call doDataDisplay("EXPLANATION: IM/Info Error: 961. You missed a a message from " & arrCommandArguments$(1) & " because it was too big.", ColorConstants.vbRed)
            Case "962"
              Call doDataDisplay("EXPLANATION: IM/Info Error: 962. You missed a a message from " & arrCommandArguments$(1) & " because it was sent too fast.", ColorConstants.vbRed)
            Case "970"
              Call doDataDisplay("EXPLANATION: Directory Error: 970. Failure.", ColorConstants.vbRed)
            Case "971"
              Call doDataDisplay("EXPLANATION: Directory Error: 971. Too many matches.", ColorConstants.vbRed)
            Case "972"
              Call doDataDisplay("EXPLANATION: Directory Error: 972. Need more qualifiers.", ColorConstants.vbRed)
            Case "973"
              Call doDataDisplay("EXPLANATION: Directory Error: 973. Dir service temporarily unavailable.", ColorConstants.vbRed)
            Case "974"
              Call doDataDisplay("EXPLANATION: Directory Error: 974. Email lookup restricted.", ColorConstants.vbRed)
            Case "975"
              Call doDataDisplay("EXPLANATION: Directory Error: 975. Keyword Ignored.", ColorConstants.vbRed)
            Case "976"
              Call doDataDisplay("EXPLANATION: Directory Error: 976. No Keywords.", ColorConstants.vbRed)
            Case "977"
              Call doDataDisplay("EXPLANATION: Directory Error: 977. Language not supported.", ColorConstants.vbRed)
            Case "978"
               Call doDataDisplay("EXPLANATION: Directory Error: 978. Country not supported.", ColorConstants.vbRed)
            Case "979"
               Call doDataDisplay("EXPLANATION: Directory Error: 979. Failure unknown " & arrCommandArguments$(1) & ".", ColorConstants.vbRed)
            Case "980"
               Call doDataDisplay("EXPLANATION: Authorization Error: 980. Incorrect nickname or password.", ColorConstants.vbRed)
            Case "981"
               Call doDataDisplay("EXPLANATION: Authorization Error: 981. The service is temporarily unavailable.", ColorConstants.vbRed)
            Case "982"
               Call doDataDisplay("EXPLANATION: Authorization Error: 982. Your warning level is currently too high to sign on.", ColorConstants.vbRed)
            Case "983"
               Call doDataDisplay("EXPLANATION: Authorization Error: 983. You have been connecting and disconnecting too frequently." & vbNewLine & "Wait 10 minutes and try again. If you continue to try, you will need to wait even longer.", ColorConstants.vbRed)
            Case "989"
               Call doDataDisplay("EXPLANATION: Authorization Error: 989. An unknown signon error has occurred " & arrCommandArguments$(1) & ".", ColorConstants.vbRed)
          End Select
          If fEnablePlugins And FileExists(App.Path & "\plugins\ERROR.exe") Then
            oAutomation.PluginCommandData = sRecivedData
            Shell App.Path & "\plugins\ERROR.exe"
          End If
        Case "EVILED"
          'EVILED:<new evil>:<name of eviler, blank if anonymous>
          'The user was just eviled.
          arrCommandArguments$() = Split(arrCommand$(1), ":", 2)
          Call doDataDisplay("EXPLANATION: You were just warned. (new warning value=" & _
                    arrCommandArguments$(0) & "%) (warner=" & _
                    arrCommandArguments$(1) & ")", ColorConstants.vbRed)
          If fEnablePlugins And FileExists(App.Path & "\plugins\EVILED.exe") Then
            oAutomation.PluginCommandData = sRecivedData
            Shell App.Path & "\plugins\EVILED.exe"
          End If
        Case "IM_IN"
          'indicates an incoming instant message
          'argument 1: sender's screen name
          'argument 2: auto resonse t/f (not handled here)
          'argument 3: message
          arrCommandArguments$() = Split(arrCommand$(1), ":", 3)
          Dim iEachMenuItem As Long
          
          If fAutoIMDisplayMode Then
            For iEachMenuItem = mnuIMsFromUser.LBound To mnuIMsFromUser.UBound
              If "IMs w/" & Replace(LCase(arrCommandArguments$(0)), " ", "") = mnuIMsFromUser(iEachMenuItem).Caption Then
                'switch to im mode
                Call mnuIMsFromUser_Click(CInt(iEachMenuItem))
              End If
            Next iEachMenuItem
            If Me.Visible = False Or Me.WindowState = FormWindowStateConstants.vbMinimized Then
              'come out of hiding in the tray
              Me.WindowState = Me.iPreMinimizedState
              Me.Visible = True
              Me.Show
              frmTray.DeleteIconFromTray
              Unload frmTray
            End If
          End If
          
          Call doDataDisplay("         IM: Incoming instant message. (sender={{-SetColor=" & vbCyan & "-}}" & _
                    arrCommandArguments$(0) & "{{-SetColor=" & vbGreen & "-}}) (auto resonse t/f=" & _
                    arrCommandArguments$(1) & ") (message={{-SetColor=" & vbCyan & "-}}" & _
                    filterOutHTML(arrCommandArguments$(2)) & "{{-SetColor=" & vbGreen & "-}})", vbGreen)
          
          Call addIMToIMOnlyDataViews(arrCommandArguments$(0), _
                    "{{-SetColor=" & vbGreen & "-}}" & arrCommandArguments$(0) & _
                    ": " & filterOutHTML(arrCommandArguments$(2)))
          'log it if needed
          If fLogIMs Then Call logIMs(arrCommandArguments$(0), filterOutHTML(arrCommandArguments$(2)))
          If fEnablePlugins And FileExists(App.Path & "\plugins\IM_IN.exe") Then
            oAutomation.PluginCommandData = sRecivedData
            Shell App.Path & "\plugins\IM_IN.exe"
          End If
        Case "NICK"
          'this sends us the format our screen name should be used for display
          'argument 1: formatted screen name
          Call doDataDisplay("EXPLANATION: The correct formatting of this nickname should be " & arrCommand$(1), 14715847)
          If fEnablePlugins And FileExists(App.Path & "\plugins\NICK.exe") Then
            oAutomation.PluginCommandData = sRecivedData
            Shell App.Path & "\plugins\NICK.exe"
          End If
        Case "SIGN_ON"
          'the sign on message is sent letting us know is is ok to send our configuriations.
          Call doDataDisplay("EXPLANATION: Sign on was successful.  (Client version supported=" & arrCommand$(1) & ")", 14715847)
          If fEnablePlugins And FileExists(App.Path & "\plugins\SIGN_ON.exe") Then
            oAutomation.PluginCommandData = sRecivedData
            Shell App.Path & "\plugins\SIGN_ON.exe"
          End If
        Case "UPDATE_BUDDY"
          'indicates an update in the online status of one of our buddies
          'argument 1: buddies screen name
          'argument 2: online status t/f
          'argument 3: evil amount (not handled here)
          'argument 4: sign on time (not handled here)
          'argument 5: idle time (not handled here)
          'argument 6: user class (not handled here)
          '
          'UPDATE_BUDDY:<Buddy User>:<Online? T/F>:<Evil Amount>:<Signon Time>:<IdleTime>:<UC>
          '   This one command handles arrival/depart/updates.  Evil Amount is
          '   a percentage, Signon Time is UNIX epoc, idle time is in minutes, UC (User Class)
          '   is a two/three character string.
          '   uc [0]:
          '   ' '  - Ignore
          '   'A'  - On AOL
          '   uc [1]
          '   ' '  - Ignore
          '   'A'  - Oscar Admin
          '   'U'  - Oscar Unconfirmed
          '   'O'  - Oscar Normal
          '   uc [2]
          '   '\0' - Ignore
          '   ' '  - Ignore
          '   'U'  - The user has set their unavailable flag.
          arrCommandArguments$() = Split(arrCommand$(1), ":", 6)
          
          sUserClass = arrCommandArguments$(5)
          sUserClassData = ""

          If Mid(sUserClass, 1, 1) = "A" Then
            sUserClassData = sUserClassData & "User is an AOL subscriber"
          End If
          
          If Mid(sUserClass, 2, 1) <> " " Then
            If sUserClassData <> "" Then
              sUserClassData = sUserClassData & ", "
            End If
            If Mid(sUserClass, 2, 1) = "A" Then
              sUserClassData = sUserClassData & "User is an Oscar Admin."
            ElseIf Mid(sUserClass, 2, 1) = "U" Then
              sUserClassData = sUserClassData & "Internet unconfirmed."
              'sUserClassData = sUserClassData & "User is Oscar unconfirmed."
            ElseIf Mid(sUserClass, 2, 1) = "O" Then
              sUserClassData = sUserClassData & "Internet confirmed."
              'sUserClassData = sUserClassData & "User is Oscar Normal."
            End If
          End If
          
          If Len(sUserClass) >= 3 Then
            If Mid(sUserClass, 3, 1) = "U" Then
              If sUserClassData <> "" Then
                sUserClassData = sUserClassData & ", "
              End If
              sUserClassData = sUserClassData & "User is currently unavailable."
            End If
          End If
          
          If arrCommandArguments$(1) = "T" Then
            sLogonLogoffAnnouncement = ""
            sLogonLogoffAnnouncement = sLogonLogoffAnnouncement & "{{-SetColor=16751001-}}"
            sLogonLogoffAnnouncement = sLogonLogoffAnnouncement & arrCommandArguments$(0)
            sLogonLogoffAnnouncement = sLogonLogoffAnnouncement & "{{-SetColor=" & vbMagenta & "-}}"
            sLogonLogoffAnnouncement = sLogonLogoffAnnouncement & " is now "
            sLogonLogoffAnnouncement = sLogonLogoffAnnouncement & "{{-SetColor=16751001-}}"
            sLogonLogoffAnnouncement = sLogonLogoffAnnouncement & "ONLINE."
            sLogonLogoffAnnouncement = sLogonLogoffAnnouncement & "{{-SetColor=" & vbMagenta & "-}}"
          Else
            sLogonLogoffAnnouncement = ""
            sLogonLogoffAnnouncement = sLogonLogoffAnnouncement & "{{-SetColor=16751001-}}"
            sLogonLogoffAnnouncement = sLogonLogoffAnnouncement & arrCommandArguments$(0)
            sLogonLogoffAnnouncement = sLogonLogoffAnnouncement & "{{-SetColor=" & vbMagenta & "-}}"
            sLogonLogoffAnnouncement = sLogonLogoffAnnouncement & " is now "
            sLogonLogoffAnnouncement = sLogonLogoffAnnouncement & "{{-SetColor=16751001-}}"
            sLogonLogoffAnnouncement = sLogonLogoffAnnouncement & "OFFLINE."
            sLogonLogoffAnnouncement = sLogonLogoffAnnouncement & "{{-SetColor=" & vbMagenta & "-}}"
          End If
          Call doDataDisplay("EXPLANATION: " & sLogonLogoffAnnouncement & _
                    " Buddy status update. (warning level=" & _
                    arrCommandArguments$(2) & "%) (sign on time=" & _
                    getDateFromUnixEpoc(arrCommandArguments$(3)) & ") (idle time=" & _
                    arrCommandArguments$(4) & ") (user class info=" & _
                    sUserClassData & ")", vbMagenta)
          Call updateWhosOnlineData( _
                    arrCommandArguments$(0), _
                    arrCommandArguments$(1), _
                    arrCommandArguments$(2) & "%", _
                    getDateFromUnixEpoc(arrCommandArguments$(3)), _
                    arrCommandArguments$(4), _
                    sUserClassData)
          
          'This sets the text that goes to the im only data views.
          If arrCommandArguments$(1) = "T" Then
            Call addIMToIMOnlyDataViews(arrCommandArguments$(0), _
                      sLogonLogoffAnnouncement & _
                      " Buddy status update. (warning level=" & _
                      arrCommandArguments$(2) & "%) (sign on time=" & _
                      getDateFromUnixEpoc(arrCommandArguments$(3)) & ") (idle time=" & _
                      arrCommandArguments$(4) & ") (user class info=" & _
                      sUserClassData & ")")
            'This was creating too many files.
            'I might make this an option sometime later.
            ''log it if needed
            'If fLogIMs Then Call logIMs(arrCommandArguments$(0), _
                      sLogonLogoffAnnouncement & _
                      " Buddy status update. (warning level=" & _
                      arrCommandArguments$(2) & "%) (sign on time=" & _
                      getDateFromUnixEpoc(arrCommandArguments$(3)) & ") (idle time=" & _
                      arrCommandArguments$(4) & ") (user class info=" & _
                      sUserClassData & ")")
          Else
            Call addIMToIMOnlyDataViews(arrCommandArguments$(0), _
                      sLogonLogoffAnnouncement)
            'This was creating too many files.
            'I might make this an option sometime later.
'            If fLogIMs Then Call logIMs(arrCommandArguments$(0), _
'                      sLogonLogoffAnnouncement)
          End If
          
          If fEnablePlugins And FileExists(App.Path & "\plugins\UPDATE_BUDDY.exe") Then
            oAutomation.PluginCommandData = sRecivedData
            Shell App.Path & "\plugins\UPDATE_BUDDY.exe"
          End If
      End Select
    Case 4
      Call doDataDisplay("EXPLANATION: You are now signed off from the server." & sRecivedData, ColorConstants.vbRed)
        'manufacture a signoff
        cboInput.Text = "signoff"
        Call cboInput_KeyPress(13)
    Case 5
      'We are going to ignore these now.  The server sends a lot of these and they
      'are not very important to this program.
      'Call doDataDisplay("EXPLANATION: Keep alive signal was recived from the server.")
    Case Else
      Call doDataDisplay("EXPLANATION: Invalid Frame " & iFrameType& & " = " & sRecivedData, ColorConstants.vbRed)
  End Select
End Sub

Public Sub sendTocCommand(ByVal iFrameInformation As Long, ByVal sDataToSend As String)
  'this procedure sends data to the aim server
  Dim iSeqHi As Long, iSeqLo As Long, sDataToSendOut As String
  Dim iLen As Long, iLenHi As Long, iLenLo As Long
  Dim iEachLetter As Long
  'the flap header is built here. see the protocol documentation for an explanation on this.
  
  If Len(sDataToSend) > 2048 Then Exit Sub ' Disconnects if bigger than 2k (client->server)
  
  iCurrentMessageNumber& = iCurrentMessageNumber& + 1
  If iCurrentMessageNumber& > 65535 Then
    iCurrentMessageNumber& = 0
  End If
  
  sDataToSendOut$ = sDataToSend$
  'replace escapes
  sDataToSendOut$ = Replace(sDataToSendOut$, "{\010}", Chr(10))
  If InStr(sDataToSendOut$, "{\") Then 'shortcut
    For iEachLetter = 0 To 255
      sDataToSendOut$ = Replace(sDataToSendOut$, _
                  "{\" & getZeroPadToThree(iEachLetter) & "}", _
                  Chr(iEachLetter))
      If InStr(sDataToSendOut$, "{\") Then 'shortcut
        Exit For 'got them all
      End If
    Next iEachLetter
  End If
  
  iSeqHi& = Hi(iCurrentMessageNumber&)
  iSeqLo& = Lo(iCurrentMessageNumber&)
  iLen& = Len(sDataToSendOut$)
  iLenHi& = Hi(iLen&)
  iLenLo& = Lo(iLen&)
  sDataToSendOut$ = "*" & Chr(iFrameInformation&) & Chr(iSeqLo&) & Chr(iSeqHi&) & Chr(iLenLo&) & Chr(iLenHi&) & sDataToSendOut$
  If wskAIM.State = sckConnected Then
    Call doDataDisplay(">>>: *{\" & _
                getZeroPadToThree(iFrameInformation&) & _
                "," & getZeroPadToThree(iSeqLo&) & _
                "," & getZeroPadToThree(iSeqHi&) & _
                "," & getZeroPadToThree(iLenLo&) & _
                "," & getZeroPadToThree(iLenHi&) & _
                "}" & sDataToSend$, ColorConstants.vbYellow, True)
    wskAIM.SendData sDataToSendOut$


    If InStr(sDataToSend$, "toc_send_im ") = 1 Then
      Call addIMToIMOnlyDataViews(Split(sDataToSend$, " ")(1), _
                "{{-SetColor=" & vbYellow & "-}}" & _
                Replace(Replace(sDataToSend$, "toc_", "", , 1), Chr(0), "") _
                )
      'log it if needed
      If fLogIMs Then Call logIMs( _
                Split(sDataToSend$, " ")(1), _
                Replace(Replace(sDataToSend$, "toc_", "", , 1), Chr(0), "") _
                )
    ElseIf InStr(sDataToSend$, "toc_send_chat ") = 1 Then
      Call addChatMessageToChatDataView(Split(sDataToSend$, " ")(1), _
                "{{-SetColor=" & vbYellow & "-}}" & _
                Replace(Replace(sDataToSend$, "toc_", "", , 1), Chr(0), "") _
                )
    End If

    Debug.Print ">>> " & sDataToSendOut$
  End If
End Sub

Public Function escapeSpecialCharactersForTransmission(ByVal sInput As String) As String
  'most strings sent to the aim toc server need to be normalized. this procedure formats
  'the strings as necessary.
  sInput$ = Replace(sInput$, "\", "\\")
  sInput$ = Replace(sInput$, "$", "$")
  sInput$ = Replace(sInput$, Chr(34), "\" & Chr(34))
  sInput$ = Replace(sInput$, "(", "\(")
  sInput$ = Replace(sInput$, ")", "\)")
  sInput$ = Replace(sInput$, "[", "\[")
  sInput$ = Replace(sInput$, "]", "\]")
  sInput$ = Replace(sInput$, "{", "\{")
  sInput$ = Replace(sInput$, "}", "\}")
  escapeSpecialCharactersForTransmission$ = sInput$
End Function

'the following three procedures are used to help create the sequence numbers for the
'flap headers being sent to the aim toc server.
Public Function MakeLong(ByVal iHi As Long, ByVal iLo As Long) As Long
  MakeLong& = iLo& * 256 + iHi&
End Function

Public Function Lo(ByVal iVal As Long) As Long
  Lo& = Fix(iVal& / 256)
End Function

Public Function Hi(ByVal iVal As Long) As Long
  Hi& = iVal& Mod 256
End Function

Function getTimeStamp() As String
  Dim sReturn As String
  Dim sFraction As String
  
  sReturn = ""
  sReturn = sReturn & Year(Now) & "."
  sReturn = sReturn & getZeroPadToTwo(Month(Now)) & "."
  sReturn = sReturn & getZeroPadToTwo(Day(Now)) & " "
  sReturn = sReturn & getZeroPadToTwo(Hour(Now)) & ":"
  sReturn = sReturn & getZeroPadToTwo(Minute(Now)) & ":"
  sReturn = sReturn & getZeroPadToTwo(Second(Now)) & "."
  'add the fractional part
  sFraction = Format(Timer, "0.00")
  sFraction = Split(sFraction, ".")(1)
  sReturn = sReturn & sFraction
  getTimeStamp = sReturn
End Function

Function getDateFromUnixEpoc(ByVal sEpoc As String) As String
  Dim iSeconds As Long
  Dim dtNewDate As Date
  
  iSeconds = CLng(sEpoc)
  dtNewDate = DateAdd("s", iSeconds - DateDiff("s", CDate("1/1/1970"), Now), Now)
  'Adjust to this timezone.
  'todo: Need to add support for different timezones and daylight savings time.
  dtNewDate = DateAdd("h", -5, dtNewDate)
  getDateFromUnixEpoc = Format(CStr(dtNewDate), "yyyy.mm.dd hh:mm:ss AM/PM")
End Function

Function getZeroPadToTwo(ByVal sInput As Integer) As String
  Dim sReturn As String
  
  sReturn = CStr(sInput)
  If Len(sReturn) = 0 Then
    sReturn = "00"
  ElseIf Len(sReturn) = 1 Then
    sReturn = "0" & sReturn
  End If
  getZeroPadToTwo = sReturn
End Function

Function getZeroPadToThree(ByVal sInput As Variant) As String
  Dim sReturn As String
  
  sReturn = CStr(sInput)
  If Len(sReturn) = 0 Then
    sReturn = "000"
  ElseIf Len(sReturn) = 1 Then
    sReturn = "00" & CStr(sReturn)
  ElseIf Len(sReturn) = 2 Then
    sReturn = "0" & CStr(sReturn)
  End If
  getZeroPadToThree = sReturn
End Function

Sub updateWhosOnlineData(ByVal sName As String, _
                  ByVal sOnlineTF As String, _
                  ByVal sWarningPercent As String, _
                  ByVal sSignOnTime As String, _
                  ByVal sIdleTime As String, _
                  ByVal sUserClassInfo As String, _
                  Optional fClearWhosOnlineData As Boolean = False)
  Dim iStartOfRemove As Long
  Dim iEndOfRemove As Long
  Dim iLengthOfRemove As Long
  Dim sDataToRemove As String
  Dim iEachNotificationName As Long
  Dim frmNotificaion As frmHelp
  Dim iUboundOfNotificationsForms As Long
  'length of each column: 17 , 6 , 25 , 5 , (no limit)
  
  sName = Replace(sName, " ", "") 'despace the name.
  sName = LCase(sName)
  If fClearWhosOnlineData Then
    'remove all data and clear the screen.
    sDataForWhosOnline = ""
    If sNameOfCurrentDataView = "who's online" Then
      txtDataViews.Text = ""
      Call doDataDisplay(sDataForWhosOnline, vbCyan, , txtDataViews, True)
    End If
    Exit Sub
  End If
  
  'do notifications.
  If InStr(sDataForWhosOnline, sName & " ") = 0 And sOnlineTF = "T" Then
    'this user was not allready in the list so they must have just come online
    For iEachNotificationName = LBound(arrsNotificationList) To UBound(arrsNotificationList)
      If Replace(LCase(arrsNotificationList(iEachNotificationName)), " ", "") = sName Then
        Dim frmNotificaions As New frmHelp
        frmNotificaions.Caption = "MAIM Notificaion: " & sName & " is now online."
        frmNotificaions.txtHelp.ToolTipText = "Click the x to close."
        frmNotificaions.txtHelp = "The user " & sName & " is now online (" & getTimeStamp & ")."
        frmNotificaions.Height = frmNotificaions.Height / 3
        frmNotificaions.Visible = True
        frmNotificaions.Show
        Exit For
      End If
    Next iEachNotificationName
  End If
  
  'add header info if needed
  If sDataForWhosOnline = "" Then
    sDataForWhosOnline = sDataForWhosOnline & padStringWithSpaces("User", 17) & " "
    sDataForWhosOnline = sDataForWhosOnline & padStringWithSpaces("Warn%", 6)
    sDataForWhosOnline = sDataForWhosOnline & padStringWithSpaces("Signon time", 25)
    sDataForWhosOnline = sDataForWhosOnline & padStringWithSpaces("Idle time", 10)
    sDataForWhosOnline = sDataForWhosOnline & "User class info" & vbNewLine
  End If
  
  'Remove info for this user (from username to vbnewline)
  iStartOfRemove = InStr(sDataForWhosOnline, sName & " ")
  If iStartOfRemove Then
    'user is in the list
    iEndOfRemove = InStr(iStartOfRemove, sDataForWhosOnline, vbNewLine)
    iLengthOfRemove = iEndOfRemove - iStartOfRemove
    sDataToRemove = Mid(sDataForWhosOnline, iStartOfRemove, iLengthOfRemove)
    'remove user's data
    sDataForWhosOnline = Replace(sDataForWhosOnline, sDataToRemove & vbNewLine, "")
  End If
  
  If sOnlineTF = "T" Then
    'add the user info for this person if they are online
    sDataForWhosOnline = sDataForWhosOnline & padStringWithSpaces(sName, 17) & " "
    sDataForWhosOnline = sDataForWhosOnline & padStringWithSpaces(sWarningPercent, 6)
    sDataForWhosOnline = sDataForWhosOnline & padStringWithSpaces(sSignOnTime, 25)
    sDataForWhosOnline = sDataForWhosOnline & padStringWithSpaces(sIdleTime, 10)
    sDataForWhosOnline = sDataForWhosOnline & sUserClassInfo & vbNewLine
  End If
  
  'if we are displaying this data screen then we need to update the screen
  If sNameOfCurrentDataView = "who's online" Then
    txtDataViews.Text = ""
    Call doDataDisplay(sDataForWhosOnline, vbCyan, , txtDataViews, True)
  End If
End Sub

Sub updateChatUsersData(ByVal sChatRoomID As String, _
                  ByVal sNames As String, _
                  ByVal sOnlineTF As String, _
                  Optional fClearChatroomData As Boolean = False)
  Dim iStartOfRemove          As Long
  Dim iEndOfRemove            As Long
  Dim iLengthOfRemove         As Long
  Dim sDataToRemove           As String
  Dim arrUserNames()          As String
  Dim iEachUser               As Long
  Dim iStartOnName            As Long
  Dim iEachMenuItem           As Long
  Dim iIndexOfMenu            As Long
  Dim sChatRoomName           As String
  
  sChatRoomName = getChatNameFromNumber(sChatRoomID)
  
  'see if we have a menu item for that user
  For iEachMenuItem = mnuViewChatUsers.LBound To mnuViewChatUsers.UBound
    If "Chat users in " & sChatRoomName = mnuViewChatUsers(iEachMenuItem).Caption Then
      iIndexOfMenu = iEachMenuItem
    End If
  Next iEachMenuItem
  
  'clear it if needed
  If fClearChatroomData Then
    sDataForChatUsersViews(iIndexOfMenu) = ""
    'if we are displaying this data screen then we need to update the screen
    If sNameOfCurrentDataView = "Chat users in " & sChatRoomName Then
      txtDataViews.Text = ""
      Call doDataDisplay("", vbCyan, , txtDataViews, True)
    End If
    Exit Sub
  End If
  
  If iIndexOfMenu = 0 Then
    iIndexOfMenu = mnuViewChatUsers.UBound + 1
    Load mnuViewChatUsers(iIndexOfMenu)
    mnuViewChatUsers(iIndexOfMenu).Visible = True
    mnuViewChatUsers(iIndexOfMenu).Caption = "Chat users in " & sChatRoomName
    ReDim Preserve sDataForChatUsersViews(iIndexOfMenu)
  End If
  
  'add header info if needed
  If sDataForChatUsersViews(iIndexOfMenu) = "" Then
    sDataForChatUsersViews(iIndexOfMenu) = sDataForChatUsersViews(iIndexOfMenu) & "-Chat Users-" & vbNewLine
  End If
  sNames = Replace(sNames, " ", "") 'despace the names.

  arrUserNames() = Split(sNames, ":")
  For iEachUser = LBound(arrUserNames()) To UBound(arrUserNames())
    sDataForChatUsersViews(iIndexOfMenu) = Replace(sDataForChatUsersViews(iIndexOfMenu), arrUserNames(iEachUser) & " " & vbNewLine, "")
    If sOnlineTF = "T" Then
      'add the user info for this person if they are online
      sDataForChatUsersViews(iIndexOfMenu) = sDataForChatUsersViews(iIndexOfMenu) & arrUserNames(iEachUser) & " " & vbNewLine
    End If
  Next iEachUser
  'if we are displaying this data screen then we need to update the screen
  If sNameOfCurrentDataView = "Chat users in " & sChatRoomName Then
    txtDataViews.Text = ""
    Call doDataDisplay(sDataForChatUsersViews(iIndexOfMenu), vbCyan, , txtDataViews, True)
  End If
End Sub

Function padStringWithSpaces(ByVal sInput As String, ByVal iResultingLength As Long) As String
  Dim sResult As String
  Dim iNumberOfSpacesToAdd As Long
  
  sResult = sInput
  'force to length if longer that the max
  If Len(sResult) > iResultingLength Then
    sResult = Left(sResult, iResultingLength)
  Else
    iNumberOfSpacesToAdd = iResultingLength - Len(sResult)
    sResult = sResult & Space(iNumberOfSpacesToAdd)
  End If
  padStringWithSpaces = sResult
End Function

Sub addIMToIMOnlyDataViews(ByVal sFrom As String, ByVal sIMData As String)
  Dim iEachMenuItem As Long
  Dim iIndexOfMenu As Long
  
  sFrom = LCase(sFrom) 'lowercase
  sFrom = Replace(sFrom, """", "") 'no quotes
  sFrom = Replace(sFrom, " ", "") 'no spaces
  
  'see if we have a menu item for that user
  For iEachMenuItem = mnuIMsFromUser.LBound To mnuIMsFromUser.UBound
    If "IMs w/" & sFrom = mnuIMsFromUser(iEachMenuItem).Caption Then
      iIndexOfMenu = iEachMenuItem
    End If
  Next iEachMenuItem
  
  If iIndexOfMenu = 0 Then
    iIndexOfMenu = mnuIMsFromUser.UBound + 1
    Load mnuIMsFromUser(iIndexOfMenu)
    mnuIMsFromUser(iIndexOfMenu).Visible = True
    mnuIMsFromUser(iIndexOfMenu).Caption = "IMs w/" & sFrom
    ReDim Preserve sDataForMessagesOnlyViews(iIndexOfMenu)
  End If
  
  'add the data to the array and - update textbox if we are in that view mode.
  If sDataForMessagesOnlyViews(iIndexOfMenu) = "" Then
    sDataForMessagesOnlyViews(iIndexOfMenu) = "Note: On this screen you can enter just the " & _
            "text you want to send to the other user.  The program will add send_im, the " & _
            "user's name and the quotes.  It will also escape special characters so you can " & _
            "use things like quotes in you input." & vbNewLine & vbNewLine & vbNewLine
    sDataForMessagesOnlyViews(iIndexOfMenu) = sDataForMessagesOnlyViews(iIndexOfMenu) & _
            getTimeStamp() & " " & sIMData
  Else
    sDataForMessagesOnlyViews(iIndexOfMenu) = sDataForMessagesOnlyViews(iIndexOfMenu) & vbNewLine & getTimeStamp() & " " & sIMData
  End If
  If sNameOfCurrentDataView = "IMs w/" & sFrom Then
    Call doDataDisplay(getTimeStamp() & " " & sIMData, vbWhite, , txtDataViews, True)
  End If
End Sub

'if fLogIMs then logIMs("","")
Sub logIMs(ByVal sBuddyName As String, ByVal sIMData As String)
  Dim iFreeFile As Long
  Dim sFile As String
  
  sBuddyName = LCase(sBuddyName) 'lowercase
  sBuddyName = Replace(sBuddyName, """", "") 'no quotes
  sBuddyName = Replace(sBuddyName, " ", "") 'no spaces
  
  sFile = App.Path & "\logs\"
  sFile = sFile & sBuddyName
  sFile = sFile & "__" & Year(Now)
  sFile = sFile & "-" & getZeroPadToTwo(Month(Now))
  sFile = sFile & "-" & getZeroPadToTwo(Day(Now))
  sFile = sFile & ".log"
  iFreeFile = FreeFile
  If Not FileExists(App.Path & "\logs") Then
    MkDir App.Path & "\logs"
  End If
  Open sFile For Append As #iFreeFile
  Print #iFreeFile, getTimeStamp() & " " & sIMData
  Close #iFreeFile
End Sub


Sub addChatMessageToChatDataView(ByVal sChatRoomID As String, ByVal sMessage As String)
  Dim iEachMenuItem As Long
  Dim iIndexOfMenu As Long
  Dim sChatRoomName As String
  
  sChatRoomName = getChatNameFromNumber(sChatRoomID)
  
  'see if we have a menu item for that user
  For iEachMenuItem = mnuViewOnlyChatMessages.LBound To mnuViewOnlyChatMessages.UBound
    If "Chat in " & sChatRoomName = mnuViewOnlyChatMessages(iEachMenuItem).Caption Then
      iIndexOfMenu = iEachMenuItem
    End If
  Next iEachMenuItem
  
  If iIndexOfMenu = 0 Then
    iIndexOfMenu = mnuViewOnlyChatMessages.UBound + 1
    Load mnuViewOnlyChatMessages(iIndexOfMenu)
    mnuViewOnlyChatMessages(iIndexOfMenu).Visible = True
    mnuViewOnlyChatMessages(iIndexOfMenu).Caption = "Chat in " & sChatRoomName
    ReDim Preserve sDataForChatMessagesViews(iIndexOfMenu)
  End If
  
  'add the data to the array and - update textbox if we are in that view mode.
  If sDataForChatMessagesViews(iIndexOfMenu) = "" Then
    sDataForChatMessagesViews(iIndexOfMenu) = "Note: On this screen you can enter just the " & _
            "text you want to send to the chat room.  The program will add send_chat, the " & _
            "chatroom id and the quotes.  It will also escape special characters so you can " & _
            "use things like quotes in you input." & vbNewLine & vbNewLine & vbNewLine
    sDataForChatMessagesViews(iIndexOfMenu) = sDataForChatMessagesViews(iIndexOfMenu) & _
            getTimeStamp() & sMessage
  Else
    sDataForChatMessagesViews(iIndexOfMenu) = sDataForChatMessagesViews(iIndexOfMenu) & _
            vbNewLine & getTimeStamp() & " " & sMessage
  End If
  If sNameOfCurrentDataView = "Chat in " & sChatRoomName Then
    Call doDataDisplay(getTimeStamp() & " " & sMessage, vbGreen, , txtDataViews, True)
  End If
End Sub

Sub addChatRoomNameLookup(sChatRoomName As String, sChatRoomID As String)
  Dim iNewIndex As Long
  
  On Error Resume Next
  iNewIndex = UBound(sChatRoomNamesLookup)
  If Err.Number = 9 Then
    'If this is the first time then there will be
    'an error.  Yuck.
    iNewIndex = 0
    Err.Clear
  End If
  On Error GoTo 0 'Resume Next = off
  
  ReDim Preserve sChatRoomNamesLookup(iNewIndex)
  ReDim Preserve sChatRoomIDsLookup(iNewIndex)
  
  sChatRoomNamesLookup(iNewIndex) = sChatRoomName
  sChatRoomIDsLookup(iNewIndex) = sChatRoomID
End Sub

Function getChatNameFromNumber(sChatRoomIDNumber As String) As String
  Dim iEachIndex As Long
  Dim sChatRoomName As String
  
  For iEachIndex = LBound(sChatRoomIDsLookup) To UBound(sChatRoomIDsLookup)
    If sChatRoomIDsLookup(iEachIndex) = sChatRoomIDNumber Then
      sChatRoomName = sChatRoomNamesLookup(iEachIndex)
    End If
  Next iEachIndex
  
  'if none found use the number.
  If sChatRoomName = "" Then
    sChatRoomName = sChatRoomIDNumber
  End If
  getChatNameFromNumber = sChatRoomName
End Function

Function getChatNumberFromName(sChatRoomName As String) As String
  Dim iEachIndex As Long
  Dim sChatRoomNumber As String
  
  For iEachIndex = LBound(sChatRoomNamesLookup) To UBound(sChatRoomNamesLookup)
    If LCase(sChatRoomNamesLookup(iEachIndex)) = LCase(sChatRoomName) Then
      sChatRoomNumber = sChatRoomIDsLookup(iEachIndex)
    End If
  Next iEachIndex
  
  getChatNumberFromName = sChatRoomNumber
End Function

Public Function NormalizeForToc(ByVal strIn As String) As String
  'most strings sent to the aim toc server need to be normalized. this procedure formats
  'the strings as necessary.
  strIn$ = Replace(strIn$, "\", "\\")
  strIn$ = Replace(strIn$, "$", "$")
  strIn$ = Replace(strIn$, Chr(34), "\" & Chr(34))
  strIn$ = Replace(strIn$, "(", "\(")
  strIn$ = Replace(strIn$, ")", "\)")
  strIn$ = Replace(strIn$, "[", "\[")
  strIn$ = Replace(strIn$, "]", "\]")
  strIn$ = Replace(strIn$, "{", "\{")
  strIn$ = Replace(strIn$, "}", "\}")
  NormalizeForToc$ = strIn$
End Function
