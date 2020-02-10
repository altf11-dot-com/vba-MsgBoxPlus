Attribute VB_Name = "MsgBoxPlus_202002072238"
Option Explicit

Public Enum mPlus
    doNothing = 0
    alwaysBreak = 1
    msg = 2
    bell = 4
End Enum

Sub MsgBoxPlus_Caller()
    Dim result As Long
    'F5 or click Run to move through stops
    result = MsgBoxPlus(bell + msg, "Hello world!") 'all functions but no message, will break into VBE
    Debug.Print "returned = " & IIf(result = vbCancel, "vbCancel ", "vbOK")
'    Stop
'    MsgBoxPlus 7 'all functions but no message, will break into VBE
'    Stop
'    MsgBoxPlus 7, "<msg>", "<debugMsg>"
'    Stop
'    MsgBoxPlus doNothing, "<msg>", "<debugMsg>"
'    Stop
'    MsgBoxPlus bell, "<msg>", "<debugMsg>"
'    Stop
'    MsgBoxPlus alwaysBreak, "<msg>", "<debugMsg>"
'    Stop
'    Debug.Print MsgBoxPlus(msg, "<msg>", "<debugMsg>")
'    Stop
End Sub


Public Function MsgBoxPlus(ByVal whoaController As mPlus, Optional ByVal msg As String, Optional debugMsg As String) As Long
    'returns long value of first msgbox keypress vbCancel = 2, vbOK = 1
    'sum mPlus enums for multiple functions
    'bell + msg will display a message with a sound alert
    '
    Dim inputBoxResult As String, secretCode As String, beepOn As Boolean, loopCount As Long
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
        beepOn = False 'don't want to beep twice
        MsgBoxPlus = MsgBox(IIf(msg & debugMsg = "", "<no msg>", msg & IIf(debugMsg = "", "", vbLf & "check immediate window")), IIf(whoaController > 0, vbOKOnly, vbOKCancel), IIf(whoaController > 0, "", "OKAY=Continue, Cancel=Break"))
    End If
    If whoaController = 1 Then
        If beepOn Then Beep
        beepOn = False 'don't want to beep twice
        Stop
    ElseIf MsgBoxPlus = vbCancel Then
        secretCode = "the secret code"
        If Now < #1/1/2020 7:00:00 PM# Then secretCode = "Y" 'set to future date/time to reduce typing secret code
        Do
            inputBoxResult = InputBox( _
                "Type " & IIf(secretCode = "", "<ZLS>", secretCode) & " to BREAK into VBE," & vbLf & _
                "type END to terminate processing," & vbLf & _
                "CLICK OK or Cancel to continue ..." & vbLf & _
                "(If not sure, click OKAY and continue.", "AltF11.com jeff@jeffbrown.us", "OK")
            If UCase(inputBoxResult) = UCase(secretCode) Then
                Stop
            ElseIf UCase(Left(inputBoxResult, 3)) = "END" Then
                MsgBox "Procedure will TERMINATE."
                End
            Else
                MsgBox "Procedure will continue."
                Exit Do
            End If
        Loop
    Else
        If beepOn Then Beep
    End If
End Function
