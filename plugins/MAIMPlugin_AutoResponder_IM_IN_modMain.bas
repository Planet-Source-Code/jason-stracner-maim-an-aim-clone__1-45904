Attribute VB_Name = "modMain"
Option Explicit

Sub Main()
  Dim sCommandData As String
  Dim sMessage As String
  Dim sFrom As String
  Dim oMAIMAutomationServer As Object
  Dim iRandomResponse As Long
  Dim sColResponses As New Collection

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
  'On Error GoTo 0 'Resume Next = off
  
  'This can only be read once.  It is erased after that.
  sCommandData = oMAIMAutomationServer.PluginCommandData
  
  If sCommandData <> "" Then
    'Do the work.
    sFrom = Split(sCommandData, ":", 4)(1)
    
    'Send the response.
    
    'these are repeated to be more frequent
    sColResponses.Add "that is interesting. please continue."
    sColResponses.Add "please go on."
    sColResponses.Add "tell me more about that."
    sColResponses.Add "does talking about this bother you?"
    sColResponses.Add "please go on."
    sColResponses.Add "i'm not sure i understand you fully."
    sColResponses.Add "what is the connection, do you suppose?"
    sColResponses.Add "tell me more about that."
    sColResponses.Add "i'm not sure i understand you fully."
    sColResponses.Add "does talking about this bother you?"
    sColResponses.Add "i'm not sure i understand you fully."
    sColResponses.Add "oh, i please explain?"
    sColResponses.Add "please tell me some more about this."
    sColResponses.Add "that is interesting. please continue."
    sColResponses.Add "please go on."
    sColResponses.Add "tell me more about that."
    sColResponses.Add "does talking about this bother you?"
    sColResponses.Add "please go on."
    sColResponses.Add "i'm not sure i understand you fully."
    sColResponses.Add "what is the connection, do you suppose?"
    sColResponses.Add "tell me more about that."
    sColResponses.Add "i'm not sure i understand you fully."
    sColResponses.Add "does talking about this bother you?"
    sColResponses.Add "i'm not sure i understand you fully."
    sColResponses.Add "oh, i please explain?"
    sColResponses.Add "please tell me some more about this."
    
    'here is the list
    sColResponses.Add "that is interesting. please continue."
    sColResponses.Add "please go on."
    sColResponses.Add "tell me more about that."
    sColResponses.Add "does talking about this bother you?"
    sColResponses.Add "please go on."
    sColResponses.Add "i'm not sure i understand you fully."
    sColResponses.Add "what is the connection, do you suppose?"
    sColResponses.Add "tell me more about that."
    sColResponses.Add "i'm not sure i understand you fully."
    sColResponses.Add "does talking about this bother you?"
    sColResponses.Add "i'm not sure i understand you fully."
    sColResponses.Add "oh, i please explain?"
    sColResponses.Add "please tell me some more about this."
    sColResponses.Add "You're such a chazmo"
    sColResponses.Add "i hate chazmos"
    sColResponses.Add "Shut your man pleaser"
    sColResponses.Add "Shut your hole"
    sColResponses.Add "I'm pissed off talk about something else."
    sColResponses.Add "don't be such a chazmo"
    sColResponses.Add "what the hell is that supposed to mean"
    sColResponses.Add "i hate you"
    sColResponses.Add "you know exactly what I mean"
    sColResponses.Add "don't pretend you don't know what i'm talking about"
    sColResponses.Add "i'll do whatever the hell I want"
    sColResponses.Add "You think you invented that word"
    sColResponses.Add "do you always sound this stupid"
    sColResponses.Add "what on earth are you talking about, you chazmo"
    sColResponses.Add "i like women"
    sColResponses.Add "yeah do that"
    sColResponses.Add "yeah"
    sColResponses.Add "like im going to tell you"
    sColResponses.Add "did that chazmo tell you to say that"
    sColResponses.Add "heh i cant believe you fell for that one"
    sColResponses.Add "don't you dare tell anyone what i said earlier"
    sColResponses.Add "shut up"
    sColResponses.Add "your no fun anymore"
    sColResponses.Add "i know what your getting at"
    sColResponses.Add "hey you started it"
    sColResponses.Add "im sorry i dont speak stupid could you repeat that in english"
    sColResponses.Add "i dont speak moron could you repeat that in english"
    sColResponses.Add "don't you ever listen?"
    sColResponses.Add "I have shoes that are smarter than you"
    sColResponses.Add "no way"
    sColResponses.Add "say that again i missed it"
    sColResponses.Add "you have to be smoking crack to talk like that"
    sColResponses.Add "uh huh"
    sColResponses.Add "give me your phone number"
    sColResponses.Add "i know what youre up to"
    sColResponses.Add "are you always this stupid"
    sColResponses.Add "i hate my life"
    sColResponses.Add "if stupid were money, youd be rich"
    sColResponses.Add "maybe if you had a brain we could communicate better"
    sColResponses.Add "i bet youd be smarter with your shirt off"
    sColResponses.Add "do you have a sister?"
    sColResponses.Add "are all your friends chazmos like you"
    sColResponses.Add "sorry, i wasnt paying attention"
    sColResponses.Add "darn it, I closed my aim and lost what you said"
    sColResponses.Add "you talking to me"
    sColResponses.Add "too bad your moronic friends aren't here."
    sColResponses.Add "do you have a website"
    sColResponses.Add "do you have any friends you IM"
    sColResponses.Add "your sister is a better kisser than your mom"
    sColResponses.Add "poop"
    sColResponses.Add "ive got pics of you naked"
    sColResponses.Add "i hate people"
    sColResponses.Add "can you even tie your shoes?"
    sColResponses.Add "twenty bucks says you cant even spell your name"
    sColResponses.Add "are you really that stupid?"
    sColResponses.Add "you probably think that means something to me"
    sColResponses.Add "try saying that in english, you moron"
    sColResponses.Add "that's bull and you know it"
    sColResponses.Add "your the reason they still make shoes with velcro"
    sColResponses.Add "who ever told you that is a chazmo"
    
    Call Randomize
    iRandomResponse = Int((sColResponses.Count - 1 + 1) * Rnd + 1)
    sMessage = sColResponses(iRandomResponse)
    'pause a little bit before sending a message back
    doAPause Int((6 - 1 + 1) * Rnd + 1)
    Call oMAIMAutomationServer.SendMessage(sFrom, sMessage)
  End If
End Sub

Sub doAPause(rNumberOfSeconds As Single)
  Dim rStartTime As Single
  Dim rEndTime As Single
  
  rStartTime = Timer
  rEndTime = rStartTime + rNumberOfSeconds
  Do Until Timer > rEndTime
    DoEvents: DoEvents: DoEvents
    If rStartTime > Timer Then
      'we have gone past midnight
      Exit Do
    End If
  Loop
End Sub

