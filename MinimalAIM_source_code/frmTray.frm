VERSION 5.00
Begin VB.Form frmTray 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
'
''Messages for the system try icon
Private Const NIM_ADD = &H0             'Adds an icon to the system try
Private Const NIM_MODIFY = &H1          'Changes the icon, tooltip text or notification message for an icon in the system try
Private Const NIM_DELETE = &H2          'Deletes an icon from the system try
'
''Flags
Private Const NIF_MESSAGE = &H1         'hIcon is valid
Private Const NIF_ICON = &H2            'uCallbackMessage is valid
Private Const NIF_TIP = &H4             'szTip is valid
'
Private Const WM_MOUSEMOVE = &H200      'MouseMove windows message identifier
'
'                                        'Messages sent to the form's MouseMove event
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Type NOTIFYICONDATA
    cbSize              As Long
    hwnd                As Long         'Handle of window that receives notification messages
    uID                 As Long         'Application-defined identifier of the taskbar icon
    uFlags              As Long         'Flags indicating which structure members contain valid data
    uCallbackMessage    As Long         'Application defined callback message
    hIcon               As Long         'Handle of taskbar icon
    szTip               As String * 64  'Tooltip text to display for icon
End Type
'
Dim mtIconData          As NOTIFYICONDATA
'
Private Declare Function Shell_NotifyIcon Lib "Shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Long

Private Sub Form_Load()
  'Put the icon into the system tray.
  AddIconToTray
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'This line in AddIconToTray causes callback messages to be
  'sent to this event: .uCallbackMessage = WM_MOUSEMOVE
  '
  'The actual callback message is contained in the X parameter.
  'Note: when using this technique, X is a Windows message not a coordinate.
  '
  'Slick aye?
  Static bBusy As Boolean
  
  If bBusy = False Then           'Do one thing at a time
    bBusy = True
    Select Case CLng(X)
      Case WM_LBUTTONDBLCLK   'Double-click left mouse button
        'PopupMenu mnuMainTrayMenu, 2
        'mnuShow_Click
        
        'Strange things happen here.  Need to find a fix for this.
        'It doesn't seem to work exactly right with a double click.
        
      Case 7710 'WM_LBUTTONUP       'Left mouse button released
        mnuShow_Click
      Case WM_RBUTTONDBLCLK   'Double-click right mouse button
      Case WM_RBUTTONDOWN     'Right mouse button pressed
      Case WM_RBUTTONUP       'Right mouse button released: display popup menu
        'PopupMenu mnuMainTrayMenu, 2
    End Select
    bBusy = False
  End If
End Sub

Private Sub AddIconToTray() 'Adds an icon to the taskbar notification area
    With mtIconData
        .cbSize = Len(mtIconData)
        .hwnd = Me.hwnd                                     'Use the form to receive callback messages.
        'This is the magic that make the
        'mouse move event fire for our
        'Hotkeys.
        .uCallbackMessage = WM_MOUSEMOVE                    'Tell icon to send MouseMove messages.
        'Just a identifier.
        .uID = 1&
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .hIcon = frmMinimalAIM.Icon  'imgTrayIcon_CALC.Picture                     'Initial icon
        .szTip = frmMinimalAIM.Caption & Chr$(0)                'Initial tooltip for icon
        If Shell_NotifyIcon(NIM_ADD, mtIconData) = 0 Then   'Create icon in tray
            MsgBox "Unable to add icon to system tray!"
            Call mnuShow_Click
        End If
    End With
End Sub

Public Sub DeleteIconFromTray()
  If Shell_NotifyIcon(NIM_DELETE, mtIconData) = 0 Then
    'MsgBox "Unable to delete icon from system tray!"
  End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
  DeleteIconFromTray
End Sub

Private Sub mnuShow_Click()
  frmMinimalAIM.WindowState = frmMinimalAIM.iPreMinimizedState
  frmMinimalAIM.Visible = True
  frmMinimalAIM.Show
  DeleteIconFromTray
  Unload Me
End Sub

