Attribute VB_Name = "mBorder"
Option Explicit
'===============================================================================
' NAME
'     db - 選択範囲の罫線を消す
'
' SYNOPSIS
'     db SIDES
'
' DESCRIPTION
'     選択範囲の各セルについて、指定した辺の罫線を個別に消す。
'     「選択範囲の上端」のような範囲全体としての外周ではなく、選択した
'     全セルそれぞれの指定辺を対象とする (例: p を指定すると、
'     選択範囲内の各セルの上辺の罫線をすべて消す)。
'
' ARGUMENTS
'     SIDES は "oprl" の部分文字列 (順不同・重複可)。Excel リボンの罫線メニュー
'     ([Alt] H B の後のアクセスキー) に合わせた対応:
'         o - 下辺 (Bottom Border)
'         p - 上辺 (Top Border)
'         r - 右辺 (Right Border)
'         l - 左辺 (Left Border)
'
'     例: db o     選択範囲内の各セルの下辺を消す
'         db pl    選択範囲内の各セルの上辺・左辺を消す
'         db oprl  選択範囲内の各セルの全4辺を消す
'
' NOTES
'     - oprl 以外の文字は無視する。有効な文字が1つもない場合は何もしない。
'     - 対象は Range のみ (図形は対象外)。
'===============================================================================

Public Sub db(Optional a As String)
    Dim booster As cBooster: Set booster = New cBooster
    On Error GoTo FINALLY

    Dim rng As Range
    Dim edges() As Long
    Dim edgeCount As Long
    Dim cell As Range
    Dim i As Long

    If TypeName(Selection) <> "Range" Then Exit Sub
    Set rng = Selection

    edgeCount = ParseSides(LCase$(Trim$(a)), edges)
    If edgeCount = 0 Then Exit Sub

    For Each cell In rng.Cells
        For i = 0 To edgeCount - 1
            cell.Borders(edges(i)).LineStyle = xlNone
        Next i
    Next cell
    Exit Sub

FINALLY:
    Application.StatusBar = "[db] E" & Err.Number & ": " & Err.Description
End Sub

' sides ("oprl" の部分文字列、順不同・重複可) を Borders インデックス定数の配列へ変換する
Private Function ParseSides(ByVal sides As String, ByRef outEdges() As Long) As Long
    Dim hasO As Boolean, hasP As Boolean, hasR As Boolean, hasL As Boolean
    Dim i As Long, ch As String, n As Long

    For i = 1 To Len(sides)
        ch = Mid$(sides, i, 1)
        Select Case ch
            Case "o": hasO = True
            Case "p": hasP = True
            Case "r": hasR = True
            Case "l": hasL = True
        End Select
    Next i

    ReDim outEdges(0 To 3)
    n = 0
    If hasP Then outEdges(n) = xlEdgeTop: n = n + 1
    If hasO Then outEdges(n) = xlEdgeBottom: n = n + 1
    If hasL Then outEdges(n) = xlEdgeLeft: n = n + 1
    If hasR Then outEdges(n) = xlEdgeRight: n = n + 1

    If n > 0 Then ReDim Preserve outEdges(0 To n - 1)
    ParseSides = n
End Function
