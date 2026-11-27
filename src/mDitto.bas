Attribute VB_Name = "mDitto"
Option Explicit
'===============================================================================
' NAME
'     ditto - すぐ上のセルと同じ値のセルの文字色をグレーにする
'
' SYNOPSIS
'     ditto
'
' DESCRIPTION
'     選択範囲の列ごとに、各セルをすぐ上のセルと比べ、表示どおりの文字列が同じなら
'     文字色をグレー (Tailwind の gray400) にする。セル結合を避けて同じ値を並べた表で、
'     繰り返しの部分を目立たなくするためのもの。選択範囲の1行目は比べない。
'
'     もう一度実行すると、上と違う値になったのに文字色が gray400 のままのセルは、
'     文字色を自動に戻す (ほかの色を付けたセルには触らない)。値を変えたあとは
'     実行し直す。
'
' NOTES
'     - 列ごとに独立して比べる (左の列の値は見ない)。
'     - 値ではなく画面の表示どおりの文字列で比べる (書式の違う同じ日付なども同じとみなす)。
'     - 空のセルと結合セルは対象外。空のセル・結合セルのすぐ下のセルはグレーにしない。
'     - 非表示の行は飛ばし、すぐ上の表示されている行と比べる。
'     - 列全体などを選択した場合は、使用範囲 (UsedRange) と重なる部分だけを対象にする。
'     - 複数領域の選択 (Ctrl+クリックなど) には対応しない。
'
' COMPLETIONS
'     ditto           すぐ上と同じ値のセルの文字色をグレーに
'===============================================================================

Public Sub ditto()
    Dim booster As cBooster: Set booster = New cBooster
    On Error GoTo FINALLY

    Dim rng As Range
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Areas.Count > 1 Then
        Application.StatusBar = "[ditto] Multiple areas are not supported."
        Exit Sub
    End If
    Set rng = Application.Intersect(Selection, Selection.Worksheet.UsedRange)
    If rng Is Nothing Then
        Application.StatusBar = "[ditto] The selection is empty."
        Exit Sub
    End If

    Dim grayed As Long, restored As Long
    ApplyDitto rng, grayed, restored
    Application.StatusBar = "[ditto] Grayed " & grayed & " cell(s), restored " & restored & " cell(s)."
    Exit Sub

FINALLY:
    Application.StatusBar = "[ditto] E" & Err.Number & ": " & Err.Description
End Sub

' rng に ditto の色付けをする。グレーにしたセルと自動に戻したセルの数を返す
Public Sub ApplyDitto(ByVal rng As Range, Optional ByRef grayed As Long, Optional ByRef restored As Long)
    Dim gray As Long
    Dim c As Long, r As Long
    Dim cell As Range
    Dim text As String, above As String
    Dim hasAbove As Boolean

    If Not mColorUtil.TryParseHexColor(mPalette.LookupHex("gray", "400"), gray) Then Exit Sub
    grayed = 0
    restored = 0

    For c = 1 To rng.Columns.Count
        hasAbove = False
        For r = 1 To rng.Rows.Count
            Set cell = rng.Cells(r, c)
            If Not cell.EntireRow.Hidden Then
                If cell.MergeCells Then
                    hasAbove = False
                Else
                    text = mMarkdown.CellText(cell)
                    If hasAbove And Len(text) > 0 And text = above Then
                        If FontColorIs(cell, gray) = False Then cell.Font.Color = gray
                        grayed = grayed + 1
                    ElseIf FontColorIs(cell, gray) Then
                        cell.Font.ColorIndex = xlColorIndexAutomatic
                        restored = restored + 1
                    End If
                    above = text
                    hasAbove = (Len(text) > 0)
                End If
            End If
        Next r
    Next c
End Sub

' セル全体の文字色が clr か (セル内で色が混在していれば False)
Private Function FontColorIs(ByVal cell As Range, ByVal clr As Long) As Boolean
    Dim v As Variant
    v = cell.Font.Color
    If IsNull(v) Then Exit Function
    FontColorIs = (CLng(v) = clr)
End Function
