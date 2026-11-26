Attribute VB_Name = "mHtml"
Option Explicit
'===============================================================================
' NAME
'     html - 選択範囲を罫線・書式つきの HTML の表としてクリップボードにコピーする
'
' SYNOPSIS
'     html
'
' DESCRIPTION
'     選択範囲を HTML の <table> に変換し、HTML 形式とテキスト形式 (HTML のソース)
'     の両方でクリップボードにコピーする。Confluence・Word・Outlook などに貼ると
'     表として、テキストエディタに貼ると HTML のソースとして貼り付けられる。
'
'     - セルの値は画面の表示どおりの文字列を使う (md と同じ)。改行は <br> にする。
'     - 罫線は上下左右それぞれの線の種類・太さ・色を CSS の border にする。
'     - 塗り色・文字色・太字・斜体・下線・取り消し線・横位置・縦位置・折り返しの
'       有無を style 属性にする。条件付き書式による見た目も反映する。
'     - セル内の一部だけに付いた文字の書式 (文字色・太字・斜体・下線・取り消し線・
'       上付き・下付き) は <span> にする。対象は数式でない文字列のセルだけ。
'     - 結合セルは colspan / rowspan にする。
'     - 列幅・行の高さを pt 単位で指定する。
'     - 非表示の行・列は出力しない。
'
' NOTES
'     - 列全体などを選択した場合は、使用範囲 (UsedRange) と重なる部分だけを対象にする。
'     - 選択範囲からはみ出した結合セルは、選択範囲の中だけを結合したものとして扱う。
'     - 斜線の罫線・パターンやグラデーションの塗り・フォントの種類と大きさは
'       反映しない。
'     - セル内の一部だけの書式は1文字ずつ調べるので、長い文章のセルでは時間がかかる。
'       読み取りに失敗したセルは、書式なしの文字列にする。
'     - 線の種類は、一点鎖線などを破線 (dashed) にまとめる。
'     - 複数領域の選択 (Ctrl+クリックなど) には対応しない。
'
' COMPLETIONS
'     html            選択範囲を罫線・書式つきの HTML の表としてコピー
'===============================================================================

Public Sub html()
    Dim booster As cBooster: Set booster = New cBooster
    On Error GoTo FINALLY

    Dim rng As Range
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Areas.Count > 1 Then
        Application.StatusBar = "[html] Multiple areas are not supported."
        Exit Sub
    End If
    Set rng = Application.Intersect(Selection, Selection.Worksheet.UsedRange)
    If rng Is Nothing Then
        Application.StatusBar = "[html] The selection is empty."
        Exit Sub
    End If

    Dim nRows As Long, nCols As Long
    Dim r As Long, c As Long
    Dim rowShown() As Boolean, colShown() As Boolean
    Dim covered() As Boolean ' 出力済みの結合セルに含まれる
    nRows = rng.Rows.Count
    nCols = rng.Columns.Count
    ReDim rowShown(1 To nRows)
    ReDim colShown(1 To nCols)
    ReDim covered(1 To nRows, 1 To nCols)
    For r = 1 To nRows
        rowShown(r) = Not rng.Rows(r).EntireRow.Hidden
    Next r
    For c = 1 To nCols
        colShown(c) = Not rng.Columns(c).EntireColumn.Hidden
    Next c

    Dim lines As Collection: Set lines = New Collection
    Dim tr As String
    Dim nShownRows As Long, nShownCols As Long

    lines.Add "<table style=""border-collapse:collapse"">"
    lines.Add "<colgroup>"
    For c = 1 To nCols
        If colShown(c) Then
            lines.Add "<col style=""width:" & Pt(rng.Columns(c).Width) & """>"
            nShownCols = nShownCols + 1
        End If
    Next c
    lines.Add "</colgroup>"

    For r = 1 To nRows
        If rowShown(r) Then
            nShownRows = nShownRows + 1
            tr = "<tr style=""height:" & Pt(rng.Rows(r).Height) & """>"
            For c = 1 To nCols
                If colShown(c) And Not covered(r, c) Then
                    tr = tr & CellHtml(rng, r, c, rowShown, colShown, covered)
                End If
            Next c
            lines.Add tr & "</tr>"
        End If
    Next r
    lines.Add "</table>"

    If nShownRows = 0 Or nShownCols = 0 Then
        Application.StatusBar = "[html] All rows or columns in the selection are hidden."
        Exit Sub
    End If

    Dim src As String
    src = JoinLines(lines)
    If mClipboard.SetClipboardHtml(src, src) Then
        Application.StatusBar = "[html] Copied a " & nShownRows & "x" & nShownCols & " table."
    Else
        Application.StatusBar = "[html] Failed to write to the clipboard."
    End If
    Exit Sub

FINALLY:
    Application.StatusBar = "[html] E" & Err.Number & ": " & Err.Description
End Sub

' rng 内の (r, c) から始まる <td>。結合セルなら範囲内のセルを covered にする
Private Function CellHtml(ByVal rng As Range, ByVal r As Long, ByVal c As Long, _
                          ByRef rowShown() As Boolean, ByRef colShown() As Boolean, _
                          ByRef covered() As Boolean) As String
    Dim cell As Range, area As Range
    Dim r0 As Long, c0 As Long, r1 As Long, c1 As Long
    Dim i As Long, j As Long
    Dim rowSpan As Long, colSpan As Long
    Dim attrs As String

    Set cell = rng.Cells(r, c)
    If cell.MergeCells Then
        Set area = Application.Intersect(cell.MergeArea, rng)
    Else
        Set area = cell
    End If

    ' 選択範囲内での結合範囲 (非表示の行・列は数えない)
    r0 = area.Row - rng.Row + 1
    c0 = area.Column - rng.Column + 1
    r1 = r0 + area.Rows.Count - 1
    c1 = c0 + area.Columns.Count - 1
    For i = r0 To r1
        If rowShown(i) Then rowSpan = rowSpan + 1
        For j = c0 To c1
            covered(i, j) = True
        Next j
    Next i
    For j = c0 To c1
        If colShown(j) Then colSpan = colSpan + 1
    Next j

    If rowSpan > 1 Then attrs = attrs & " rowspan=""" & rowSpan & """"
    If colSpan > 1 Then attrs = attrs & " colspan=""" & colSpan & """"
    attrs = attrs & " style=""" & CellStyle(area) & """"

    ' 値と書式は結合セルの左上 (選択範囲外のこともある) が持つ
    Dim owner As Range
    Dim content As String
    Set owner = area.Cells(1, 1).MergeArea.Cells(1, 1)
    content = RichTextHtml(owner)
    If Len(content) = 0 Then content = EscapeHtml(mMarkdown.CellText(owner))
    CellHtml = "<td" & attrs & ">" & content & "</td>"
End Function

Private Function CellStyle(ByVal area As Range) As String
    Dim cell As Range
    Dim df As DisplayFormat
    Dim s As String

    Set cell = area.Cells(1, 1).MergeArea.Cells(1, 1)
    Set df = cell.DisplayFormat

    ' 罫線は範囲の外周にあるセルから取る
    s = s & BorderCss("border-top", area.Cells(1, 1).DisplayFormat.Borders(xlEdgeTop))
    s = s & BorderCss("border-bottom", area.Cells(area.Rows.Count, 1).DisplayFormat.Borders(xlEdgeBottom))
    s = s & BorderCss("border-left", area.Cells(1, 1).DisplayFormat.Borders(xlEdgeLeft))
    s = s & BorderCss("border-right", area.Cells(1, area.Columns.Count).DisplayFormat.Borders(xlEdgeRight))

    If df.Interior.Pattern <> xlPatternNone Then
        s = s & "background:" & mColorUtil.ColorLongToHex(df.Interior.Color) & ";"
    End If

    ' セル内で混在する項目は Null になり、FontCss が出さない (RichTextHtml の <span> が持つ)
    s = s & FontCss(df.Font, False)

    s = s & HAlignCss(cell)
    s = s & VAlignCss(cell.VerticalAlignment)
    If Not cell.WrapText Then s = s & "white-space:nowrap;"
    s = s & "padding:1pt 3pt"

    CellStyle = s
End Function

' 文字の書式の CSS。Null (セル内で混在) の項目は出さない。
' 上付き・下付きは <span> にだけ付ける (<td> の vertical-align を上書きしないため)
Private Function FontCss(ByVal f As Font, ByVal withScript As Boolean) As String
    Dim s As String, deco As String

    If Not IsNull(f.Color) Then
        If f.Color <> 0 Then s = s & "color:" & mColorUtil.ColorLongToHex(f.Color) & ";"
    End If
    If IsTrue(f.Bold) Then s = s & "font-weight:bold;"
    If IsTrue(f.Italic) Then s = s & "font-style:italic;"
    If Not IsNull(f.Underline) Then
        If f.Underline <> xlUnderlineStyleNone Then deco = deco & " underline"
    End If
    If IsTrue(f.Strikethrough) Then deco = deco & " line-through"
    If Len(deco) > 0 Then s = s & "text-decoration:" & Mid$(deco, 2) & ";"

    If withScript Then
        If IsTrue(f.Superscript) Then
            s = s & "vertical-align:super;font-size:smaller;"
        ElseIf IsTrue(f.Subscript) Then
            s = s & "vertical-align:sub;font-size:smaller;"
        End If
    End If

    FontCss = s
End Function

Private Function IsTrue(ByVal v As Variant) As Boolean
    If Not IsNull(v) Then IsTrue = CBool(v)
End Function

' セル内の一部だけに付いた書式を <span> にした中身。書式が混在しないセル、
' 数式や数値のセル、読み取りに失敗したセルは空文字列 (呼び出し側で書式なしにする)
Private Function RichTextHtml(ByVal cell As Range) As String
    Dim f As Font
    Dim text As String
    Dim n As Long, i As Long, runStart As Long
    Dim css As String, runCss As String
    Dim out As String

    If cell.HasFormula Then Exit Function
    If VarType(cell.Value2) <> vbString Then Exit Function
    Set f = cell.Font
    If Not (IsNull(f.Color) Or IsNull(f.Bold) Or IsNull(f.Italic) Or IsNull(f.Underline) Or _
            IsNull(f.Strikethrough) Or IsNull(f.Superscript) Or IsNull(f.Subscript)) Then Exit Function

    On Error GoTo FALLBACK
    text = cell.Value2
    n = Len(text)
    If n = 0 Then Exit Function

    ' 同じ書式が続く文字をまとめて1つの <span> にする
    runStart = 1
    runCss = FontCss(cell.Characters(1, 1).Font, True)
    For i = 2 To n + 1
        If i <= n Then css = FontCss(cell.Characters(i, 1).Font, True)
        If i > n Or css <> runCss Then
            out = out & SpanHtml(Mid$(text, runStart, i - runStart), runCss)
            runStart = i
            runCss = css
        End If
    Next i
    RichTextHtml = out
    Exit Function

FALLBACK:
    RichTextHtml = ""
End Function

Private Function SpanHtml(ByVal text As String, ByVal css As String) As String
    If Len(css) = 0 Then
        SpanHtml = EscapeHtml(text)
    Else
        SpanHtml = "<span style=""" & css & """>" & EscapeHtml(text) & "</span>"
    End If
End Function

' 罫線1本ぶんの CSS (例: border-top:1px solid #000000;)。線がなければ空文字列
Private Function BorderCss(ByVal prop As String, ByVal b As Border) As String
    Dim ls As Variant
    Dim lineStyle As String, width As String

    ls = b.LineStyle
    If IsNull(ls) Then Exit Function
    If ls = xlLineStyleNone Then Exit Function

    Select Case ls
        Case xlDouble: lineStyle = "double"
        Case xlDot: lineStyle = "dotted"
        Case xlDash, xlDashDot, xlDashDotDot, xlSlantDashDot: lineStyle = "dashed"
        Case Else: lineStyle = "solid"
    End Select

    If lineStyle = "double" Then
        width = "3px" ' 二重線は2本の線と隙間が収まる太さが要る
    Else
        Select Case b.Weight
            Case xlMedium: width = "2px"
            Case xlThick: width = "3px"
            Case Else: width = "1px" ' xlThin, xlHairline
        End Select
    End If

    BorderCss = prop & ":" & width & " " & lineStyle & " " & mColorUtil.ColorLongToHex(b.Color) & ";"
End Function

' 横位置。HTML の既定 (左寄せ) と同じときは空文字列
Private Function HAlignCss(ByVal cell As Range) As String
    Dim align As String
    Select Case cell.HorizontalAlignment
        Case xlCenter, xlCenterAcrossSelection: align = "center"
        Case xlRight: align = "right"
        Case xlJustify, xlDistributed: align = "justify"
        Case xlGeneral
            ' 標準: 数値は右寄せ、論理値とエラー値は中央揃え
            Select Case VarType(cell.Value2)
                Case vbDouble, vbCurrency, vbDecimal, vbInteger, vbLong, vbSingle: align = "right"
                Case vbBoolean, vbError: align = "center"
            End Select
    End Select
    If Len(align) > 0 Then HAlignCss = "text-align:" & align & ";"
End Function

' 縦位置。HTML の既定は中央、Excel の既定は下なので常に指定する
Private Function VAlignCss(ByVal v As Long) As String
    Select Case v
        Case xlTop: VAlignCss = "vertical-align:top;"
        Case xlBottom: VAlignCss = "vertical-align:bottom;"
        Case Else: VAlignCss = "vertical-align:middle;" ' xlCenter, xlJustify, xlDistributed
    End Select
End Function

Private Function EscapeHtml(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, vbCrLf, "<br>")
    s = Replace(s, vbLf, "<br>")
    s = Replace(s, vbCr, "<br>")
    EscapeHtml = s
End Function

' 地域設定によらず小数点を . にした pt 値 (例: 13.5pt)
Private Function Pt(ByVal points As Double) As String
    Pt = Trim$(Str$(Round(points, 2))) & "pt"
End Function

Private Function JoinLines(ByVal lines As Collection) As String
    Dim parts() As String
    Dim i As Long
    ReDim parts(1 To lines.Count)
    For i = 1 To lines.Count
        parts(i) = lines(i)
    Next i
    JoinLines = Join(parts, vbCrLf) & vbCrLf
End Function
