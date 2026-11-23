Attribute VB_Name = "mMisc"
Option Explicit

Sub mg(Optional SizePt As Long = 20): MakeGrid (SizePt): End Sub
Sub mt(Optional ByVal headerRows As Long = 1): MakeTable (headerRows): End Sub


' -------------------------------------------------------------------

' Excel 方眼紙を作成
Public Sub MakeGrid(Optional SizePt As Long = 20)
    Dim booster As cBooster: Set booster = New cBooster
    On Error GoTo FINALLY

    Columns("A").ColumnWidth = 1
    Columns("A:XFD").ColumnWidth = SizePt / Columns("A").Width
    Rows("1:1048576").RowHeight = Columns("A").Width
    Exit Sub

FINALLY:
    Application.StatusBar = "[mg] E" & Err.Number & ": " & Err.Description
End Sub

Public Sub MakeTable(Optional ByVal headerRows As Long = 1)
    Dim booster As cBooster: Set booster = New cBooster
    On Error GoTo FINALLY

    Dim rng As Range
    Dim ws As Worksheet
    Dim i As Long
    Dim colOdd As Range, colEven As Range
    Dim rowOdd As Range, rowEven As Range
    Dim headBand As Range
    Dim firstRow As Long
    Dim firstCol As Long
    Dim firstColVals As Variant
    Dim headerVals As Variant
    Dim dataRows As Long
    Dim colCount As Long
    
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set rng = Selection
    If rng.Areas.Count > 1 Then Exit Sub
    If headerRows < 1 Or headerRows >= rng.Rows.Count Then Exit Sub
    
    Set ws = rng.Worksheet
    
    ' ---- 初期化 ----
    rng.Borders.LineStyle = xlNone
    firstRow = rng.Row
    firstCol = rng.Column
    
    ' Collect rows
    dataRows = rng.Rows.Count - headerRows
    If dataRows > 1 Then
        firstColVals = rng.Offset(headerRows, 0).Resize(dataRows, 1).Value2
        For i = 1 To dataRows
            If firstColVals(i, 1) <> "" Then
                If (i Mod 2) = 0 Then
                    Set rowEven = UnionSafe(rowEven, rng.Rows(i + headerRows))
                Else
                    Set rowOdd = UnionSafe(rowOdd, rng.Rows(i + headerRows))
                End If
            End If
        Next i
    End If

    ' ==== 列の収集 (ヘッダ1行目、最左列除外、偶奇に分割) ====
    colCount = rng.Columns.Count
    If colCount > 1 Then
        headerVals = rng.Rows(1).Value2
        For i = 2 To colCount
            If headerVals(1, i) <> "" Then
                If (i Mod 2) = 0 Then
                    Set colEven = UnionSafe(colEven, rng.Columns(i))
                Else
                    Set colOdd = UnionSafe(colOdd, rng.Columns(i))
                End If
            End If
        Next i
    End If

    ' ==== ヘッダ下端帯 ====
    Set headBand = rng.Rows(headerRows)

    ' ==== 罫線適用 ====
    ' 行区切り (破線:上エッジ)? 偶奇グループごとに一括
    ApplyBorder rowOdd, xlEdgeTop, xlDash
    ApplyBorder rowEven, xlEdgeTop, xlDash

    ' 列区切り (実線:左エッジ)? 偶奇グループごとに一括
    ApplyBorder colOdd, xlEdgeLeft, xlContinuous
    ApplyBorder colEven, xlEdgeLeft, xlContinuous

    ' ヘッダ下端 (実線:下エッジ)
    With headBand.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    ' 外枠 (実線)
    rng.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, ColorIndex:=xlAutomatic
    Exit Sub

FINALLY:
    Application.StatusBar = "[mt] E" & Err.Number & ": " & Err.Description
End Sub


' ---- 共通ユーティリティ ----
Private Function UnionSafe(base As Range, add As Range) As Range
    If base Is Nothing Then
        Set UnionSafe = add
    Else
        Set UnionSafe = Union(base, add)
    End If
End Function

Private Sub ApplyBorder(target As Range, edge As XlBordersIndex, style As XlLineStyle)
    If target Is Nothing Then Exit Sub
    With target.Borders(edge)
        .LineStyle = style
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End Sub

