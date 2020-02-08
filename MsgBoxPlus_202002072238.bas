Attribute VB_Name = "MsgBoxPlus_202002072238"
Option Explicit

Public Enum mPlus
    doNothing = 0
    alwaysBreak = 1
    msg = 2
    bell = 4
End Enum

Sub MsgBoxPlus_Caller()
    'F5 or click Run to move through stops
    MsgBoxPlus 7, "<msg>", "<debugMsg>"
    Stop
    MsgBoxPlus doNothing, "<msg>", "<debugMsg>"
    Stop
    MsgBoxPlus bell, "<msg>", "<debugMsg>"
    Stop
    MsgBoxPlus alwaysBreak, "<msg>", "<debugMsg>"
    Stop
    Debug.Print MsgBoxPlus(msg, "<msg>", "<debugMsg>")
    Stop
End Sub


Public Function MsgBoxPlus(ByVal whoaController As mPlus, Optional ByVal msg As String, Optional debugMsg As String) As Long
    Dim mbpResponse As Long, debugMsgAlreadyExposed, x As String, y As String, beepOn As Boolean
    If debugMsg <> "" Then Debug.Print debugMsg
    DoEvents
    If whoaController > 7 Then
       'max value = 7, all on
       whoaController = 7
    End If
    If whoaController >= 4 Then
        'beep on = 4's bit on
        whoaController = whoaController Mod 4
        beepOn = True
    End If
    If whoaController >= 2 Then
        'msgbox on = 2's bit on
        whoaController = whoaController Mod 2
        If beepOn Then Beep
        mbpResponse = MsgBox(IIf(msg & debugMsg = "", "<no msg>", msg & IIf(debugMsg = "", "", "check immediate window")), IIf(whoaController > 0, vbOKOnly, vbOKCancel), IIf(whoaController > 0, "", "OKAY=Continue, Cancel=Break"))
    End If
    If whoaController = 1 Then
        If beepOn Then Beep
        Stop
    ElseIf mbpResponse = vbCancel Then
        y = "the secret code"
        If Now < #1/1/2020 7:00:00 PM# Then y = "Y" 'set to future date/time to reduce typing secret code
        x = InputBox("Type " & IIf(y = "", "<ZLS>", y) & " to BREAK into VBE, type CLOSE to terminate." & vbLf & "(If not sure, type CLOSE, no data will be lost)", "")
        If UCase(y) = UCase(x) Then
            Stop 'use Set Next Statement (Ctrl+F9) to skip over End if desired
            End 'terminates all procedures
        End If
    Else
        If beepOn Then Beep
    End If
    MsgBoxPlus = mbpResponse
End Function
