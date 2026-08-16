Attribute VB_Name = "mUtil"
Option Explicit

Public Sub tb(ByVal headerRows As Long)
    Dim booster As cBooster: Set booster = New cBooster
    On Error GoTo FINALLY

    Dim rng As Range
    Dim ws As Worksheet
    
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set rng = Selection
    If rng.Areas.Count > 1 Then
        Set rng = rng.Areas(1).Worksheet.Range(rng.Address)
    End If
    If headerRows < 1 Or headerRows > rng.Rows.Count Then Exit Sub
    
    Set ws = rng.Worksheet
    
    
    ' ---- 既存罫線クリア ----
    rng.Borders.LineStyle = xlNone
    
    Dim firstRow As Long, firstCol As Long
    firstRow = rng.Row
    firstCol = rng.Column
    
    Dim colBands As Range ' 縦線 (列区切り:左エッジに実線)
    Dim rowBands As Range ' 行区切り (上エッジに破線)
    Dim headBand As Range ' ヘッダ下端 (実線)
    
    Set colBands = rng.Columns(1)
    Set rowBands = rng.Rows(headerRows)
    
    ' ---- 列区切りの検出 (ヘッダ1行目) ----
    Dim j As Long
    For j = 2 To rng.Columns.Count
        If IsNonEmpty(ws.Cells(firstRow, firstCol + j - 1).Value) Then
            Set colBands = Union(colBands, rng.Columns(j))
        End If
    Next j
    
    ' ---- 行区切りの検出 (データ領域) ----
    Dim i As Long
    If rng.Rows.Count > headerRows Then
        For i = headerRows + 1 To rng.Rows.Count
            If IsNonEmpty(ws.Cells(firstRow + i - 1, firstCol).Value) Then
                Set rowBands = Union(rowBands, rng.Rows(i))
            End If
        Next i
    End If
    
    ' ---- ヘッダ下端帯 ----
    Set headBand = rng.Rows(headerRows)
    
    ' ---- 罫線適用 ----
    ' 1) 行区切り (破線:上エッジ)
    With rowBands.Borders(xlEdgeTop)
        .LineStyle = xlDash
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    ' 2) 列区切り (実線:左エッジ)
    With colBands.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    ' 3) ヘッダ下端 (実線:下エッジ)
    With headBand.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    ' 4) 外枠 (実線: BorderAround 使用)
    rng.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, ColorIndex:=xlAutomatic
    Exit Sub

FINALLY:
    Application.StatusBar = "[tb] E" & Err.Number & ": " & Err.Description
End Sub

Private Function IsNonEmpty(ByVal v As Variant) As Boolean
    ' 数式で "" を返す場合は空扱い。空白のみの文字列も空扱い。
    If IsError(v) Then
        IsNonEmpty = True
    Else
        IsNonEmpty = (Len(Trim$(CStr(v))) > 0)
    End If
End Function

