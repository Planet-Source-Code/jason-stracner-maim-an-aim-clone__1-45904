Attribute VB_Name = "MMain"
'aim: jasonstracner
Option Explicit
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public oAutomation As CAutomation
Public fEnablePlugins As Boolean
Public fUseTray As Boolean
Public fAutoIMDisplayMode As Boolean
Public fAutoLogin As Boolean
Public sAutoLoginUsername As String
Public sAutoLoginPassword As String
Public sAutoLoginServer As String
Public sAutoLoginServerPort As String
Public arrsNotificationList() As String
Public fLogIMs As Boolean

Sub Main()
  Dim iEachNotificationUser As Long
  Dim sCurrentNotificationListBuddy As String
  
  fAutoLogin = GetINIString("settings", "auto login", App.Path & "\maim_settings.ini", "0") = "1"
  sAutoLoginUsername = GetINIString("settings", "auto login username", App.Path & "\maim_settings.ini", "")
  sAutoLoginPassword = GetINIString("settings", "auto login password", App.Path & "\maim_settings.ini", "")
  sAutoLoginPassword = functionDecryptText(sAutoLoginPassword)
  sAutoLoginPassword = encryptPasswordForTOCServer(sAutoLoginPassword)
  sAutoLoginServer = GetINIString("settings", "auto login server", App.Path & "\maim_settings.ini", "")
  sAutoLoginServerPort = GetINIString("settings", "auto login server port", App.Path & "\maim_settings.ini", "")
  fEnablePlugins = GetINIString("settings", "enable plugins", App.Path & "\maim_settings.ini", "0") = "1"
  fUseTray = GetINIString("settings", "use tray", App.Path & "\maim_settings.ini", "0") = "1"
  fAutoIMDisplayMode = GetINIString("settings", "auto display IMs", App.Path & "\maim_settings.ini", "1") = "1"
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
  
  If Not fAutoLogin Then
    frmMinimalAIM.Show
  Else
    Load frmMinimalAIM
  End If
  Set oAutomation = New CAutomation
End Sub

Public Function WriteINIString(ByVal strSection As String, ByVal strKeyName As String, ByVal strValue As String, ByVal strFile As String) As Long
  Dim lngStatus As Long
  Dim sNetworkUser As String
  
  sNetworkUser = GetNetworkUsername()
  lngStatus& = WritePrivateProfileString(strSection & " for " & sNetworkUser, strKeyName, strValue, strFile)
  WriteINIString& = (lngStatus& <> 0)
End Function

Public Function GetINIString(ByVal strSection As String, ByVal strKeyName As String, ByVal strFile As String, Optional ByVal strDefault As String = "") As String
  Dim strBuffer As String * 256, lngSize As Long
  Dim sNetworkUser As String
  
  sNetworkUser = GetNetworkUsername()
  lngSize& = GetPrivateProfileString(strSection$ & " for " & sNetworkUser, strKeyName$, strDefault$, strBuffer$, 256, strFile$)
  GetINIString$ = Left$(strBuffer$, lngSize&)
End Function

Public Function GetNetworkUsername() As String
    Dim lpUserID As String
    Dim nBuffer As Long
    Dim Ret As Long
    
    lpUserID = String(25, 0)
    nBuffer = 25
    Ret = GetUserName(lpUserID, nBuffer)

    If Ret Then
      lpUserID = Replace(lpUserID$, Chr(0), "")
      GetNetworkUsername$ = lpUserID$
    End If
End Function

Public Function FileExists(ByVal strFileName As String) As Boolean
  Dim intLen As Integer
  On Error Resume Next
  
  If strFileName$ <> "" Then
    intLen% = Len(Dir$(strFileName$))
    If intLen = 0 Then
      intLen% = Len(Dir$(strFileName$, vbDirectory))
    End If
    FileExists = (Not Err And intLen% > 0)
  Else
    FileExists = False
  End If
End Function

Public Function functionSimpleEncryption(ByVal sInput As String, Optional fDecrypt As Boolean = False) As String
  Dim iEachInputLetter As Long
  Dim fUpOrDown As Boolean
  Dim sEncryptedText As String
  Const sPassword As String = "akashdklf alsdlslsl doio dkk"
  Dim iEachPasswordLetter As Long
  Dim iNewLetterNumber As Long
  
  fUpOrDown = fDecrypt
  iEachPasswordLetter = 1
  For iEachInputLetter = 1 To Len(sInput)
    If fUpOrDown Then
      iNewLetterNumber = Asc(Mid(sInput, iEachInputLetter, 1)) + Asc(Mid(sPassword, iEachPasswordLetter, 1))
    Else
      iNewLetterNumber = Asc(Mid(sInput, iEachInputLetter, 1)) - Asc(Mid(sPassword, iEachPasswordLetter, 1))
    End If
    If iNewLetterNumber > 255 Then
      iNewLetterNumber = iNewLetterNumber - 255
    ElseIf iNewLetterNumber < 1 Then
      iNewLetterNumber = iNewLetterNumber + 255
    End If
    sEncryptedText = sEncryptedText & Chr(iNewLetterNumber)
    fUpOrDown = Not fUpOrDown
    iEachPasswordLetter = iEachPasswordLetter + 1
    If iEachPasswordLetter > Len(sPassword) Then
      iEachPasswordLetter = 1
    End If
  Next iEachInputLetter
  functionSimpleEncryption = sEncryptedText
End Function

Public Function functionDecryptText(ByVal sInput As String) As String
    Dim Password As String
    Dim sStringToDecode As String
    Dim Decrypted As String
    Dim sPreText As String
    Dim sPostText As String
    Dim X As Long
    Dim sNumberToDecode As String
    
    If InStr(sInput, "<encoded>") = 0 Then
      functionDecryptText = sInput
      Exit Function
    End If
    If InStr(sInput, "</encoded>") = 0 Then
      functionDecryptText = sInput
      Exit Function
    End If
    
    sPreText = Mid(sInput, 1, InStr(sInput, "<encoded>") - 1)
    
    sPostText = Mid(sInput, InStr(sInput, "</encoded>") + 11)
    
    sStringToDecode = Mid(sInput, InStr(sInput, "<encoded>") + Len("<encoded>"))  'take off first part
    sStringToDecode = Mid(sStringToDecode, 1, InStr(sStringToDecode, "</encoded>") - 1) 'take off last part
    
    'Convert from numbers.
    sNumberToDecode = ""
    For X = 1 To Len(sStringToDecode) Step 3
      sNumberToDecode = sNumberToDecode & Chr(Mid(sStringToDecode, X, 3))
    Next X
    Decrypted = sNumberToDecode
    Decrypted = functionSimpleEncryption(Decrypted, True)
    functionDecryptText = sPreText & Decrypted & sPostText
End Function

Public Function functionEncryptText(ByVal sInput As String) As String
    Dim Encrypted As String
    Dim X As Long
    Dim sNumberEncoded As String
    
    Encrypted = functionSimpleEncryption(sInput)
    'Convert to numbers.
    sNumberEncoded = ""
    For X = 1 To Len(Encrypted)
      sNumberEncoded = sNumberEncoded & func0Pad2ThreeChars(Asc(Mid(Encrypted, X, 1)))
    Next X
    functionEncryptText = "<encoded>" & sNumberEncoded & "</encoded>"
End Function

Function func0Pad2ThreeChars(ByVal sInput As String) As String
  If Len(sInput) = 0 Then
    func0Pad2ThreeChars = "000" & sInput
  ElseIf Len(sInput) = 1 Then
    func0Pad2ThreeChars = "00" & sInput
  ElseIf Len(sInput) = 2 Then
    func0Pad2ThreeChars = "0" & sInput
  ElseIf Len(sInput) >= 3 Then
    func0Pad2ThreeChars = sInput
  End If
End Function

Public Function encryptPasswordForTOCServer(ByVal sPassword As String) As String
  'this is a simple xor encryption used to encrypt the aim password. the roasting string
  'is "Tic/Toc"
  Dim arrEncodingTable() As Variant, sEncryptedPassword As String
  Dim iEachPasswordLetter As Long, sHexValue As String
  
  arrEncodingTable = Array("84", "105", "99", "47", "84", "111", "99")
  sEncryptedPassword$ = "0x"
  For iEachPasswordLetter& = 0 To Len(sPassword$) - 1
    sHexValue$ = Hex(Asc(Mid(sPassword$, iEachPasswordLetter& + 1, 1)) Xor CLng(arrEncodingTable((iEachPasswordLetter& Mod 7))))
    If CLng("&H" & sHexValue$) < 16 Then sEncryptedPassword$ = sEncryptedPassword$ & "0"
    sEncryptedPassword$ = sEncryptedPassword$ & sHexValue$
  Next
  encryptPasswordForTOCServer$ = LCase(sEncryptedPassword$)
End Function
