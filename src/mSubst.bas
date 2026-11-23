Attribute VB_Name = "mSubst"
Option Explicit
'===============================================================================
' NAME
'     s - 選択範囲のセルと図形内のテキストを正規表現で置換する (Vim の :s)
'
' SYNOPSIS
'     s/PATTERN/REPLACEMENT/[FLAGS]
'
' DESCRIPTION
'     ランチャーでは s の直後に区切り文字を続けて入力する。区切り文字以降は空白を
'     含めて丸ごと引数になる。区切り文字は / 以外の記号でもよい (例: s#a/b#c#)。
'     区切り文字そのものは \/ のようにエスケープする。末尾の区切り文字は省略できる。
'
'     対象:
'         セルを選択中   選択範囲内の文字列セルと、左上が選択範囲内にある図形
'         図形を選択中   選択中の図形 (グループ内の図形を含む)
'     シート全体を置換したいときは Ctrl+A で全体を選択してから実行する (:%s 相当)。
'
'     FLAGS:
'         g   すべての一致を置換する (省略時は各セル・各図形で最初の一致のみ)
'         i   大文字・小文字を区別しない
'
'     PATTERN は VBScript の正規表現 (後読み (?<=...) などは使えない)。
'     ^ / $ は各行の先頭 / 末尾に一致する (MultiLine)。
'
'     REPLACEMENT では次が使える。
'         \1 ～ \9 / $1 ～ $9   グループの参照
'         $&                    一致した文字列全体
'         \n  \t                改行 / タブ
'         \\  $$                \ / $ そのもの
'
'     例: s/株式会社/(株)/g
'         s/(\d+)円/\1 円/g
'         s/^\s+//          行頭の空白を削除
'
' NOTES
'     - 数式セル・数値セルは対象外 (文字列の定数セルのみ)。
'     - セル内の一部の書式 (文字ごとの色など) は、置換したセルでは失われる。
'       図形は一致した部分だけを差し替えるので、ほかの部分の書式は残る。
'     - 文字列セルは置換後も文字列のまま保つ。置換結果が数式・数値・日付として
'       解釈されそうな場合 (=abc, 0123, 1/2 など) は、先頭に ' を付けて入れる。
'     - マクロによる変更のため、Ctrl+Z で元に戻せない。
'
' COMPLETIONS
'     s               正規表現で置換 (s/PAT/REP/FLAGS)
'===============================================================================

Private Type SubstResult
    Replacements As Long
    CellCount As Long
    ShapeCount As Long
End Type

Public Sub s(Optional ByVal arg As String)
    Dim booster As cBooster: Set booster = New cBooster
    On Error GoTo FINALLY

    Dim pattern As String, replacement As String, flags As String
    Dim re As Object
    Dim result As SubstResult
    Dim errMsg As String

    If Not ParseSubstArg(arg, pattern, replacement, flags, errMsg) Then
        Application.StatusBar = "[s] " & errMsg
        Exit Sub
    End If
    If Not BuildRegExp(pattern, flags, re, errMsg) Then
        Application.StatusBar = "[s] " & errMsg
        Exit Sub
    End If
    replacement = ToVbsReplacement(replacement)

    If TypeName(Selection) = "Range" Then
        ReplaceInCells Selection, re, replacement, result
        ReplaceInShapesWithin Selection, re, replacement, result
    Else
        Dim sr As ShapeRange
        On Error Resume Next: Set sr = Selection.ShapeRange: On Error GoTo FINALLY
        If sr Is Nothing Then
            Application.StatusBar = "[s] Select cells or shapes."
            Exit Sub
        End If
        Dim i As Long
        For i = 1 To sr.Count
            ReplaceInShape sr(i), re, replacement, result
        Next i
    End If

    Application.StatusBar = "[s] " & result.Replacements & " replacement(s) in " & _
                            result.CellCount & " cell(s), " & result.ShapeCount & " shape(s)."
    Exit Sub

FINALLY:
    Application.StatusBar = "[s] E" & Err.Number & ": " & Err.Description
End Sub

'========================
' 引数の解析
'========================

' "/PAT/REP/FLAGS" を分解する。区切り文字は先頭の1文字。\区切り文字 は区切り文字そのもの。
' それ以外の \x はそのまま残す (正規表現・置換文字列側で解釈する)
Private Function ParseSubstArg(ByVal arg As String, ByRef outPattern As String, ByRef outReplacement As String, _
                               ByRef outFlags As String, ByRef errMsg As String) As Boolean
    Dim delim As String
    Dim parts(0 To 2) As String
    Dim n As Long, i As Long
    Dim ch As String

    If Len(arg) = 0 Then errMsg = "Usage: s/PATTERN/REPLACEMENT/[FLAGS]": Exit Function
    delim = Left$(arg, 1)
    If delim Like "[A-Za-z0-9\ """"]" Or delim = "\" Then
        errMsg = "Invalid delimiter: " & delim: Exit Function
    End If

    i = 2
    Do While i <= Len(arg)
        ch = Mid$(arg, i, 1)
        If ch = "\" And i < Len(arg) Then
            If Mid$(arg, i + 1, 1) = delim And n < 2 Then
                parts(n) = parts(n) & delim
            Else
                parts(n) = parts(n) & ch & Mid$(arg, i + 1, 1)
            End If
            i = i + 2
        ElseIf ch = delim And n < 2 Then
            n = n + 1
            i = i + 1
        Else
            parts(n) = parts(n) & ch
            i = i + 1
        End If
    Loop

    If Len(parts(0)) = 0 Then errMsg = "Empty pattern.": Exit Function
    For i = 1 To Len(parts(2))
        If InStr("gi", Mid$(parts(2), i, 1)) = 0 Then
            errMsg = "Unknown flag: " & Mid$(parts(2), i, 1): Exit Function
        End If
    Next i

    outPattern = parts(0)
    outReplacement = parts(1)
    outFlags = parts(2)
    ParseSubstArg = True
End Function

Private Function BuildRegExp(ByVal pattern As String, ByVal flags As String, ByRef outRe As Object, _
                             ByRef errMsg As String) As Boolean
    Set outRe = CreateObject("VBScript.RegExp")
    outRe.pattern = pattern
    outRe.Global = (InStr(flags, "g") > 0)
    outRe.IgnoreCase = (InStr(flags, "i") > 0)
    outRe.MultiLine = True

    ' 不正なパターンは Test で例外になる
    On Error Resume Next
    outRe.Test ""
    If Err.Number <> 0 Then errMsg = "Invalid pattern: " & pattern: Exit Function
    On Error GoTo 0
    BuildRegExp = True
End Function

' Vim 風の置換文字列を VBScript の書式 ($1, $$ など) に直す。
' \1～\9 → $1～$9、\n → 改行、\t → タブ、\\ → \、\$ → $$。$ で始まる VBScript の書式はそのまま
Private Function ToVbsReplacement(ByVal rep As String) As String
    Dim i As Long
    Dim ch As String, nx As String
    Dim out As String

    i = 1
    Do While i <= Len(rep)
        ch = Mid$(rep, i, 1)
        If ch = "\" And i < Len(rep) Then
            nx = Mid$(rep, i + 1, 1)
            Select Case True
                Case nx Like "[1-9]": out = out & "$" & nx
                Case nx = "n": out = out & vbLf
                Case nx = "t": out = out & vbTab
                Case nx = "$": out = out & "$$"
                Case Else: out = out & nx ' \\ → \、\x → x
            End Select
            i = i + 2
        Else
            out = out & ch
            i = i + 1
        End If
    Loop
    ToVbsReplacement = out
End Function

'========================
' セル
'========================
Private Sub ReplaceInCells(ByVal sel As Range, ByVal re As Object, ByVal rep As String, ByRef result As SubstResult)
    Dim target As Range
    Dim cell As Range
    Dim text As String, replaced As String

    Set target = Application.Intersect(sel, sel.Worksheet.UsedRange)
    If target Is Nothing Then Exit Sub

    ' 文字列の定数セルだけに絞る。SpecialCells は単一セルに対して呼ぶとシート全体を
    ' 対象にしてしまうので、単一セルは直接判定する
    If target.Cells.CountLarge = 1 Then
        If target.HasFormula Or VarType(target.Value2) <> vbString Then Exit Sub
    Else
        On Error Resume Next
        Set target = target.SpecialCells(xlCellTypeConstants, xlTextValues)
        If Err.Number <> 0 Then Exit Sub ' 該当セルなし
        On Error GoTo 0
    End If

    For Each cell In target.Cells
        text = cell.Value2
        If re.Test(text) Then
            result.Replacements = result.Replacements + re.Execute(text).Count
            replaced = re.Replace(text, rep)
            If replaced <> text Then
                cell.Value2 = AsTextInput(cell, replaced)
                result.CellCount = result.CellCount + 1
            End If
        End If
    Next cell
End Sub

' 文字列セルに書き戻す値。数式・数値・日付として解釈されそうな文字列は先頭に ' を
' 付け、文字列のまま入れる (表示形式が文字列 @ のセルは変換されないので付けない)
Private Function AsTextInput(ByVal cell As Range, ByVal text As String) As String
    AsTextInput = text
    If cell.NumberFormat = "@" Or Len(text) = 0 Then Exit Function
    If InStr("=+-@", Left$(text, 1)) > 0 Or IsNumeric(text) Or IsDate(text) Then
        AsTextInput = "'" & text
    End If
End Function

'========================
' 図形
'========================

' 左上のセルが sel の中にある図形を置換する
Private Sub ReplaceInShapesWithin(ByVal sel As Range, ByVal re As Object, ByVal rep As String, ByRef result As SubstResult)
    Dim shp As Shape
    Dim topLeft As Range

    For Each shp In sel.Worksheet.Shapes
        Set topLeft = Nothing
        On Error Resume Next
        Set topLeft = shp.TopLeftCell
        On Error GoTo 0
        If Not topLeft Is Nothing Then
            If Not Application.Intersect(sel, topLeft) Is Nothing Then
                ReplaceInShape shp, re, rep, result
            End If
        End If
    Next shp
End Sub

' 図形 (グループなら中の図形すべて) のテキストを置換する。
' 一致した部分だけを後ろから差し替え、ほかの部分の書式を保つ
Private Sub ReplaceInShape(ByVal shp As Shape, ByVal re As Object, ByVal rep As String, ByRef result As SubstResult)
    Dim child As Shape
    Dim tr As Object
    Dim matches As Object
    Dim i As Long
    Dim m As Object

    If shp.Type = msoGroup Then
        For Each child In shp.GroupItems
            ReplaceInShape child, re, rep, result
        Next child
        Exit Sub
    End If

    ' テキストを持たない図形 (画像・グラフなど) は例外になるので飛ばす
    On Error Resume Next
    If Not shp.TextFrame2.HasText Then Exit Sub
    Set tr = shp.TextFrame2.TextRange
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0

    Set matches = re.Execute(tr.text)
    If matches.Count = 0 Then Exit Sub

    For i = matches.Count - 1 To 0 Step -1
        Set m = matches(i)
        tr.Characters(m.FirstIndex + 1, m.Length).text = ExpandReplacement(rep, m)
    Next i
    result.Replacements = result.Replacements + matches.Count
    result.ShapeCount = result.ShapeCount + 1
End Sub

' VBScript 書式の置換文字列を、一致 m について展開する ($1～$9, $&, $$)
Private Function ExpandReplacement(ByVal rep As String, ByVal m As Object) As String
    Dim i As Long
    Dim ch As String, nx As String
    Dim out As String
    Dim k As Long

    i = 1
    Do While i <= Len(rep)
        ch = Mid$(rep, i, 1)
        If ch = "$" And i < Len(rep) Then
            nx = Mid$(rep, i + 1, 1)
            Select Case True
                Case nx = "$"
                    out = out & "$": i = i + 2
                Case nx = "&"
                    out = out & m.Value: i = i + 2
                Case nx Like "[1-9]"
                    k = CLng(nx)
                    If k <= m.SubMatches.Count Then
                        If Not IsEmpty(m.SubMatches(k - 1)) Then out = out & m.SubMatches(k - 1)
                    Else
                        out = out & ch & nx
                    End If
                    i = i + 2
                Case Else
                    out = out & ch: i = i + 1
            End Select
        Else
            out = out & ch
            i = i + 1
        End If
    Loop
    ExpandReplacement = out
End Function
