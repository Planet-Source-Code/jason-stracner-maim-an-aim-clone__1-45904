Attribute VB_Name = "modMain"
Option Explicit

Sub Main()
  Dim sCommandData As String
  Dim sMessage As String
  Dim sFrom As String
  Dim oMAIMAutomationServer As Object
  
  'Create the automation object.
  On Error Resume Next
  Set oMAIMAutomationServer = GetObject(, "MinimalAIM.CAutomation")
  If Err Then
    If Err.Description = "ActiveX component can't create object" Then
      MsgBox "Error: " & App.EXEName & ".exe couldn't find a running copy of MAIM to automate."
    Else
      MsgBox "Error: " & Err.Description & " (" & Err.Number & ")"
    End If
    Exit Sub
  End If
  On Error GoTo 0 'Resume Next = off
  
  'This can only be read once.  It is erased after that.
  sCommandData = oMAIMAutomationServer.PluginCommandData
  
  If sCommandData <> "" Then
   'Do the work.
   sFrom = Split(sCommandData, ":", 4)(1)
   sMessage = Split(sCommandData, ":", 4)(3)
   frmIM.Caption = frmIM.Caption & "   -( from " & sFrom & ")"
   frmIM.WebBrowser1.Navigate ("about:blank")
   Do While frmIM.WebBrowser1.LocationURL <> "about:blank"
     DoEvents
   Loop
   frmIM.WebBrowser1.Document.BODY.Innerhtml = sMessage
   frmIM.Show
   frmIM.WindowState = vbNormal
   
   'Send the response.
   'Call oMAIMAutomationServer.SendMessage(sFromName, sResponse)
  End If
End Sub
