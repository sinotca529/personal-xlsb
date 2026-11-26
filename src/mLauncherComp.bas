Attribute VB_Name = "mLauncherComp"
Option Explicit
Option Private Module

' 補完候補の表。build.py が生成する mLauncherCompData.CompletionText() を読む。
' 書式 (1行1項目):
'     [CONTEXT]      以降の項目の文脈 (直前までに入力済みのトークン列)。空なら先頭トークン
'     TOKEN : DESC   候補とその説明
'     @Module.Func   候補を作る関数 (入力のたびに変わる候補。シート名など)
'     ; ...          コメント
'
' 候補を作る関数は、入力中のトークンを受け取り、Array(候補, 説明) の配列を表示順に返す。
' 絞り込みと並べ替えは関数の責任で、ここでは返ってきたものをそのまま使う。

Private CP_Map As Object        ' Scripting.Dictionary: 文脈 -> Collection (候補と説明の組)
Private CP_Providers As Object  ' Scripting.Dictionary: 文脈 -> "Module.Func"

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
    Set CP_Providers = CreateObject("Scripting.Dictionary")
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

            Case line Like "@*"
                If Not CP_Map.Exists(ctx) Then CP_Map.Add ctx, New Collection
                CP_Providers(ctx) = Mid$(line, 2)

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
    Dim p As Long

    CP_GetContextAndPrefix query, ctx, prefix

    Set CP_ListCandidates = New Collection
    CP_LoadCompletions

    ' 文脈が表にない場合 (例: "ln 2 da" の "ln 2") は、末尾のトークンを落として
    ' 登録済みの文脈 ("ln") を探す。順不同に引数を並べるコマンドでも補完が効くように。
    ' コマンド名だけの文脈まで遡っても無ければ諦める (トップレベルには戻らない)
    Do While Not CP_Map.Exists(ctx)
        p = InStrRev(ctx, " ")
        If p = 0 Then Exit Function
        ctx = Left$(ctx, p - 1)
    Loop

    If CP_Providers.Exists(ctx) Then
        AddProvidedCandidates CP_ListCandidates, CP_Providers(ctx), prefix
        Exit Function
    End If

    ' 前方一致は Like だと [ # * ? が特殊文字になるため、先頭の文字列で比較する
    Set c = CP_Map(ctx)
    For i = 1 To c.Count
        pair = c(i)
        If StrComp(Left$(CStr(pair(0)), Len(prefix)), prefix, vbTextCompare) = 0 Then
            CP_ListCandidates.Add pair
        End If
    Next i
End Function

' s の最後のトークンを newTok にした文字列。newTok に空白や " があれば "..." で囲む
' (ランチャーが1つの引数として受け取れるように)
Public Function CP_ReplaceLastToken(ByVal s As String, ByVal newTok As String) As String
    Dim i As Long

    newTok = QuoteToken(newTok)
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

' 候補を作る関数 (Module.Func) を呼び、返ってきた候補を result に足す。
' 入力途中の開き引用符 ("設計 など) は外して渡す
Private Sub AddProvidedCandidates(ByVal result As Collection, ByVal func As String, ByVal prefix As String)
    Dim pairs As Variant
    Dim i As Long

    If Left$(prefix, 1) = """" Then prefix = Mid$(prefix, 2)

    On Error GoTo FAILED
    pairs = Application.Run("'" & ThisWorkbook.Name & "'!" & func, prefix)
    If Not IsArray(pairs) Then Exit Sub
    For i = LBound(pairs) To UBound(pairs)
        result.Add pairs(i)
    Next i
    Exit Sub

FAILED:
    ' 候補を出せなくても入力は続けられるので、何も足さずに戻る
End Sub

' 空白や " を含むトークンを、ランチャーの引数の書き方 ("..."、\" と \\) にする
Private Function QuoteToken(ByVal tok As String) As String
    If InStr(tok, " ") = 0 And InStr(tok, """") = 0 Then
        QuoteToken = tok
        Exit Function
    End If
    tok = Replace(tok, "\", "\\")
    tok = Replace(tok, """", "\""")
    QuoteToken = """" & tok & """"
End Function
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
