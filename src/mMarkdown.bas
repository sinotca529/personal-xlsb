Attribute VB_Name = "mMarkdown"
Option Explicit
'===============================================================================
' NAME
'     md - 選択範囲を Markdown の表としてクリップボードにコピーする
'
' SYNOPSIS
'     md
'
' DESCRIPTION
'     選択範囲の1行目を見出し行として、Markdown の表 (| A | B |) に変換し、
'     クリップボードにコピーする。
'
'     - セルの値は画面の表示どおりの文字列 (日付・桁区切りなどの書式を反映) を使う。
'     - セル内の | は \| に、改行は <br> に置き換える。
'     - 見出しを除く空でないセルがすべて数値の列は、右寄せ (---:) にする。
'     - テキストのままでも読みやすいよう、全角を2文字ぶんとして列幅を揃える。
'
' NOTES
'     - 列全体などを選択した場合は、使用範囲 (UsedRange) と重なる部分だけを対象にする。
'     - 列幅が足りず ### と表示されているセルは、書式を反映しない値を使う。
'     - 結合セルは左上のセルにだけ値が入り、残りは空欄になる。
'     - 複数領域の選択 (Ctrl+クリックなど) には対応しない。
'
' COMPLETIONS
'     md              選択範囲を Markdown の表としてコピー
'===============================================================================

Public Sub md()
    On Error GoTo FINALLY

    Dim rng As Range
    Dim cells_() As String
    Dim widths() As Long
    Dim rightAlign() As Boolean
    Dim nRows As Long, nCols As Long
    Dim r As Long, c As Long
    Dim lines() As String

    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Areas.Count > 1 Then
        Application.StatusBar = "[md] Multiple areas are not supported."
        Exit Sub
    End If
    Set rng = Application.Intersect(Selection, Selection.Worksheet.UsedRange)
    If rng Is Nothing Then
        Application.StatusBar = "[md] The selection is empty."
        Exit Sub
    End If

    nRows = rng.Rows.Count
    nCols = rng.Columns.Count
    ReDim cells_(1 To nRows, 1 To nCols)
    ReDim widths(1 To nCols)
    ReDim rightAlign(1 To nCols)

    ' セルの文字列・列幅・右寄せの判定
    For c = 1 To nCols
        rightAlign(c) = (nRows > 1)
        For r = 1 To nRows
            cells_(r, c) = EscapeCell(CellText(rng.Cells(r, c)))
            If DisplayWidth(cells_(r, c)) > widths(c) Then widths(c) = DisplayWidth(cells_(r, c))
            If r > 1 And Len(cells_(r, c)) > 0 Then
                If Not IsNumericValue(rng.Cells(r, c).Value2) Then rightAlign(c) = False
            End If
        Next r
        If widths(c) < 3 Then widths(c) = 3 ' 区切り行の --- が収まる幅
    Next c

    ' 見出し行・区切り行・データ行
    ReDim lines(0 To nRows)
    lines(0) = FormatRow(cells_, 1, widths, rightAlign)
    lines(1) = SeparatorRow(widths, rightAlign)
    For r = 2 To nRows
        lines(r) = FormatRow(cells_, r, widths, rightAlign)
    Next r

    If mClipboard.SetClipboardText(Join(lines, vbCrLf) & vbCrLf) Then
        Application.StatusBar = "[md] Copied a " & nRows & "x" & nCols & " table."
    Else
        Application.StatusBar = "[md] Failed to write to the clipboard."
    End If
    Exit Sub

FINALLY:
    Application.StatusBar = "[md] E" & Err.Number & ": " & Err.Description
End Sub

' 表示どおりの文字列。列幅不足で ### になっている場合は値をそのまま使う
Private Function CellText(ByVal cell As Range) As String
    Dim t As String
    t = cell.text
    If Len(t) > 0 And Replace(t, "#", "") = "" And Not IsError(cell.Value2) Then
        If VarType(cell.Value2) <> vbString Then t = CStr(cell.Value2)
    End If
    CellText = t
End Function

Private Function EscapeCell(ByVal s As String) As String
    s = Replace(s, "|", "\|")
    s = Replace(s, vbCrLf, "<br>")
    s = Replace(s, vbLf, "<br>")
    s = Replace(s, vbCr, "<br>")
    EscapeCell = s
End Function

Private Function IsNumericValue(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbDouble, vbCurrency, vbDecimal, vbInteger, vbLong, vbSingle
            IsNumericValue = True
    End Select
End Function

' 等幅フォントでの表示幅 (全角 2、半角 1) の近似
Private Function DisplayWidth(ByVal s As String) As Long
    DisplayWidth = LenB(StrConv(s, vbFromUnicode))
End Function

Private Function Pad(ByVal s As String, ByVal colWidth As Long, ByVal toRight As Boolean) As String
    Dim padding As String
    padding = Space$(colWidth - DisplayWidth(s))
    Pad = IIf(toRight, padding & s, s & padding)
End Function

Private Function FormatRow(ByRef cells_() As String, ByVal r As Long, _
                           ByRef widths() As Long, ByRef rightAlign() As Boolean) As String
    Dim c As Long
    Dim parts() As String
    ReDim parts(1 To UBound(widths))
    For c = 1 To UBound(widths)
        parts(c) = Pad(cells_(r, c), widths(c), rightAlign(c) And r > 1)
    Next c
    FormatRow = "| " & Join(parts, " | ") & " |"
End Function

Private Function SeparatorRow(ByRef widths() As Long, ByRef rightAlign() As Boolean) As String
    Dim c As Long
    Dim parts() As String
    ReDim parts(1 To UBound(widths))
    For c = 1 To UBound(widths)
        If rightAlign(c) Then
            parts(c) = String$(widths(c) - 1, "-") & ":"
        Else
            parts(c) = String$(widths(c), "-")
        End If
    Next c
    SeparatorRow = "| " & Join(parts, " | ") & " |"
End Function
