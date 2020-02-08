Attribute VB_Name = "MsgBoxPlus_202002072238"
Option Explicit


Public Enum mPlus
    off = 0
    Break = 1
    mplusmsg = 2
    bell = 4
End Enum


Public Function MsgBoxPlus(ByVal whoaController As mPlus, Optional ByVal msg As String, Optional debugMsg As String)
    Dim mbpResponse As Integer, debugMsgAlreadyExposed, x As String, y As String
    If debugMsg <> "" Then Debug.Print debugMsg
    DoEvents
    If whoaController > 7 Then
       'max value = 7, all on
       whoaController = 7
    End If
    If whoaController >= 4 Then
        'beep on = 4's bit on
        whoaController = whoaController Mod 4
        Beep
    End If
    If whoaController >= 2 Then
        'msgbox on = 2's bit on
        whoaController = whoaController Mod 2
        mbpResponse = MsgBox(IIf(msg & debugMsg = "", "<no msg>", msg & IIf(debugMsg = "", "", "check immediate window")), IIf(whoaController > 0, vbOKOnly, vbOKCancel), IIf(whoaController > 0, "", "OKAY to ignore, Cancel to ") & "BREAK")
    End If
    If whoaController = 1 Then
        Stop
    ElseIf mbpResponse = vbCancel Then
        y = "the secret code"
        If Now < #1/1/2020# Then y = "Y"
        x = InputBox("Type " & IIf(y = "", "<ZLS>", y) & " to BREAK into VBE, type CLOSE to terminate." & vbLf & "(If not sure, type CLOSE, no data will be lost)", "")
        If UCase(y) = UCase(x) Then
            Stop
        End If
    End If
End Function
