VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAutomation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type GUIDs
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

'Declares needed to register object in the ROT
Private Const ACTIVEOBJECT_STRONG = 0
Private Const ACTIVEOBJECT_WEAK = 1
Private Declare Function CLSIDFromProgID Lib "ole32.dll" (ByVal ProgID As Long, rclsid As GUIDs) As Long
Private Declare Function CoDisconnectObject Lib "ole32.dll" (ByVal pUnk As IUnknown, pvReserved As Long) As Long
Private Declare Function RegisterActiveObject Lib "oleaut32.dll" (ByVal pUnk As IUnknown, rclsid As GUIDs, ByVal dwFlags As Long, pdwRegister As Long) As Long
Private Declare Function RevokeActiveObject Lib "oleaut32.dll" (ByVal dwRegister As Long, ByVal pvReserved As Long) As Long
Dim OLEInstance As Long
Private m_FromUsername As String
Private m_Parameters As String

'Public Event SaidHello()

Private Sub Class_Initialize()
'The magic happens here
'This code is responsible for creating the entry for "vbAIM.CAutomation" in the rot
    Dim mGUID As GUIDs
    Dim lp As Long
    OLEInstance = 0
    lp = CLSIDFromProgID(StrPtr("MinimalAIM.CAutomation"), mGUID)
    If lp = 0 Then
        lp = RegisterActiveObject(Me, mGUID, ACTIVEOBJECT_WEAK, OLEInstance)
    End If
End Sub

Friend Sub Quit()
    'Once we are done with the main program, lets clean up the rot
    'by removing the entry for our ActiveX Server
    
    If OLEInstance <> 0 Then
        RevokeActiveObject OLEInstance, 0
    End If
    CoDisconnectObject Me, 0
End Sub

Private Sub Class_Terminate()
    RevokeActiveObject OLEInstance, 0
End Sub

Property Get PluginCommandData() As String
  PluginCommandData = m_Parameters
  m_Parameters = "" 'Set it to blank after the plugin reads it.
End Property

Property Let PluginCommandData(sParameters As String)
  m_Parameters = sParameters
End Property

Public Function SendMessage(sToUsername As String, sMessage As String) As Boolean
    'Check to see if any plugins want to send out a message.
  If frmMinimalAIM.wskAIM.State = sckConnected Then
       
    If sMessage <> "" And sToUsername <> "" Then
      Call frmMinimalAIM.sendTocCommand(2, "toc_send_im " & sToUsername & " " & Chr(34) & Normalize("<HTML>" & sMessage & "</HTML>") & Chr(34) & Chr(0))
      SendMessage = True
    Else
      SendMessage = False
    End If
  Else
    SendMessage = False
  End If
End Function

Private Function Normalize(ByVal strIn As String) As String
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
  Normalize$ = strIn$
End Function
