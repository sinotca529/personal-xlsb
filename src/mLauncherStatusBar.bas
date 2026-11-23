Attribute VB_Name = "mLauncherStatusBar"
Option Explicit
Option Private Module
' ステータスバー上でコマンドを入力・実行する Vim 風のコマンドモード。
' ExitCmdMode 以降、":" キーで EnterCmdMode に入る。

Private Buf As String

Private Sub ShowBuf()
    Application.StatusBar = ":" & Buf & "|"
End Sub

'
Public Sub EnterCmdMode()
    EnableOnKeyHandler
    Buf = ""
    ShowBuf
End Sub

'
Public Sub EnableOnKeyHandler()
    SetOnKeyHandler "CmdHandleKey"
End Sub

Public Sub DisableOnKeyHandler()
    SetOnKeyHandler ""
End Sub

'
Public Sub CmdHandleKey(key As String)
    Select Case key
        Case "~"
            ExecuteCmd
        Case "{ESC}"
            ExitCmdMode
        Case "{BS}"
            If Len(Buf) = 0 Then
                ExitCmdMode
                Exit Sub
            End If
            Buf = Left(Buf, Len(Buf) - 1)
            ShowBuf
        Case Else
            Buf = Buf & key
            ShowBuf
    End Select
End Sub

'
Private Sub SetOnKeyHandler(handlerName As String)
    Dim i As Integer
    Dim handler As String
    Dim ch As String
    Dim keyStr As Variant

    For i = Asc(" ") To Asc("~")
        ch = Chr(i)
        Select Case True
            Case Asc("A") <= i And i <= Asc("Z"):  keyStr = "+" & LCase(ch)
            Case InStr("+^%*(){}[]", Chr$(i)) > 0: keyStr = "{" & ch & "}"
            Case Else:                             keyStr = ch
        End Select
        
        handler = IIf(handlerName <> "", "'" & handlerName & " """ & ch & """'", "")
        Application.OnKey keyStr, handler
    Next i
    
    For Each keyStr In Array("{BS}", "~", "{ESC}")
        Application.OnKey keyStr, IIf(handlerName <> "", "'" & handlerName & " """ & keyStr & """'", "")
    Next
End Sub

Private Sub ExecuteCmd()
    On Error GoTo Cleanup
    DisableOnKeyHandler
    mLauncher.TryRun Buf
Cleanup:
    ExitCmdMode
End Sub

Private Sub ExitCmdMode()
    DisableOnKeyHandler
    Application.OnKey ":", "EnterCmdMode"
    Application.StatusBar = ""
End Sub
