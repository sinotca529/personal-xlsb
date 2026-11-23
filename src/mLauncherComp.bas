Attribute VB_Name = "mLauncherComp"
Option Explicit
Option Private Module

' 補完候補の表。build.py が生成する mLauncherCompData.CompletionText() を読む。
' 書式 (1行1項目):
'     [CONTEXT]      以降の項目の文脈 (直前までに入力済みのトークン列)。空なら先頭トークン
'     TOKEN : DESC   候補とその説明
'     ; ...          コメント

Private CP_Map As Object   ' Scripting.Dictionary

'-------------------------
' 公開API
'-------------------------
Public Sub CP_LoadCompletions()
    Dim lines As Variant
    Dim i As Long
    Dim line As String, ctx As String
    Dim tok As String, desc As String

    ' 表は埋め込みで変化しないため、一度読めば十分
    If Not CP_Map Is Nothing Then Exit Sub

    Set CP_Map = CreateObject("Scripting.Dictionary")
    ctx = ""

    lines = Split(mLauncherCompData.CompletionText(), vbLf)
    For i = LBound(lines) To UBound(lines)
        line = Trim$(lines(i))
        If line = "" Then GoTo NextLine

        Select Case True
            Case line Like ";*"
                ' コメント

            Case line Like "[[]*[]]"
                ctx = CollapseSpaces(Mid$(line, 2, Len(line) - 2))
                If Not CP_Map.Exists(ctx) Then CP_Map.Add ctx, New Collection

            Case Else
                ParseLineToTokenDesc line, tok, desc
                If tok <> "" Then
                    If Not CP_Map.Exists(ctx) Then CP_Map.Add ctx, New Collection
                    CP_Map(ctx).Add MakePair(tok, desc)
                End If
        End Select
NextLine:
    Next i
End Sub

Private Sub CP_GetContextAndPrefix(ByVal s As String, ByRef ctx As String, ByRef prefix As String)
    Dim p As Long

    If s = "" Then
        ctx = ""
        prefix = ""
        Exit Sub
    End If

    If s Like "* " Then
        ctx = CollapseSpaces(Trim$(s))
        prefix = ""
        Exit Sub
    End If

    s = CollapseSpaces(LTrim$(s))
    p = InStrRev(s, " ")

    If p = 0 Then
        ctx = ""
        prefix = s
        Exit Sub
    End If

    ctx = Left$(s, p - 1)
    prefix = Mid$(s, p + 1)
End Sub

Public Function CP_ListCandidates(ByVal query As String) As Collection
    Dim c As Collection
    Dim i As Long
    Dim pair As Variant

    Dim ctx As String
    Dim prefix As String

    CP_GetContextAndPrefix query, ctx, prefix

    Set CP_ListCandidates = New Collection
    CP_LoadCompletions

    If Not CP_Map.Exists(ctx) Then
        Exit Function
    End If

    Set c = CP_Map(ctx)
    For i = 1 To c.Count
        pair = c(i)
        If prefix = "" Or LCase$(CStr(pair(0))) Like LCase$(prefix) & "*" Then
            CP_ListCandidates.Add pair
        End If
    Next i
End Function

Public Function CP_ReplaceLastToken(ByVal s As String, ByVal newTok As String) As String
    Dim i As Long

    i = InStrRev(s, " ")
    If i = 0 Then
        CP_ReplaceLastToken = newTok
    Else
        If Right$(s, 1) = " " Then
            CP_ReplaceLastToken = s & newTok
        Else
            CP_ReplaceLastToken = Left$(s, i) & newTok
        End If
    End If
End Function

'-------------------------
' 内部ユーティリティ
'-------------------------
Private Sub ParseLineToTokenDesc(ByVal line As String, ByRef tok As String, ByRef desc As String)
    Dim p As Long
    p = InStr(line, ":")
    If p > 0 Then
        tok = Trim$(Left$(line, p - 1))
        desc = Trim$(Mid$(line, p + 1))
    Else
        tok = Trim$(line)
        desc = ""
    End If
End Sub

Private Function MakePair(ByVal tok As String, ByVal desc As String) As Variant
    Dim a(1) As Variant
    a(0) = tok
    a(1) = desc
    MakePair = a
End Function

Private Function CollapseSpaces(ByVal s As String) As String
    Dim i As Long, ch As String * 1, inSpace As Boolean, r As String
    s = Trim$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = " " Then
            If Not inSpace Then r = r & " ": inSpace = True
        Else
            r = r & ch
            inSpace = False
        End If
    Next
    CollapseSpaces = r
End Function
