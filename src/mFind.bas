Attribute VB_Name = "mFind"
Option Explicit
'===============================================================================
' NAME
'     find, fn, fp, fq - セルと図形内のテキストをブック横断で検索し、順に移動する
'
' SYNOPSIS
'     find TEXT
'     find
'     fn
'     fp
'     fq
'
' DESCRIPTION
'     find TEXT は、アクティブなブックの全シート (非表示シートを除く) について、
'     セルの値と図形内のテキスト (グループ内の図形も含む) から TEXT を探し、
'     ヒットした箇所の一覧を作って1件目へ移動する。空白を含む TEXT はランチャーで
'     "..." で囲む。
'
'     find 後は検索モードになり、次のキーで移動できる (Vim の n / N 相当)。
'         n       次のヒットへ (末尾の次は先頭へ戻る)
'         N       前のヒットへ (Shift+n)
'         Esc     検索モードを終了し、n / N / Esc を元に戻す
'     ランチャーから fn (次へ) / fp (前へ) / fq (終了) でも同じ操作ができる。
'     引数なしの find は、前回の検索結果のまま検索モードに戻る。
'
'     移動先は、セルならそのセルを選択し、図形ならヒットした文字を選択する。
'     一覧はシート順に、シート内では画面上の位置 (セルの行→列、図形は左上の
'     セル位置) の順に並ぶ。図形内で複数回ヒットした場合は、それぞれを1件とする。
'
' NOTES
'     - 大文字・小文字 (と全角・半角などの違い) は区別しない (vbTextCompare)。
'     - セルは表示文字列ではなく値 (Value2) を検索する。日付は内部の数値になる。
'     - 非表示の行・列にあるセル、非表示の図形は対象外。数式セルは計算結果を検索する。
'     - 検索モードに入るとき IME をオフ (半角英数) にする。日本語で検索した直後でも
'       そのまま n / N で移動できる。
'     - 検索モード中は n / N のキーを奪うため、セルに n から直接入力し始めることは
'       できない (F2 などで編集モードに入れば入力できる)。
'     - 一覧は find を実行した時点のもの。その後に編集した場合は find をやり直す。
'
' COMPLETIONS
'     find            セルと図形内のテキストを検索
'     fn              次の検索結果へ
'     fp              前の検索結果へ
'     fq              検索モードを終了
'===============================================================================

' ヒット1件 = Array(種別, 対象, 行, 列, 開始位置, 長さ)
'   種別 "cell": 対象は Range。開始位置・長さは未使用
'   種別 "shape": 対象は Shape。開始位置・長さはテキスト内のヒット位置
Private Const H_KIND As Long = 0
Private Const H_TARGET As Long = 1
Private Const H_ROW As Long = 2
Private Const H_COL As Long = 3
Private Const H_START As Long = 4
Private Const H_LEN As Long = 5

Private gHits As Collection
Private gIndex As Long   ' 現在位置 (1 始まり、0 は未移動)
Private gQuery As String

'========================
' エントリポイント
'========================
Public Sub find(Optional ByVal query As String)
    On Error GoTo FINALLY

    If Len(query) = 0 Then
        If gHits Is Nothing Then
            Application.StatusBar = "[find] Please specify the text to search."
        ElseIf gHits.Count = 0 Then
            Application.StatusBar = "[find] No matches: " & gQuery
        Else
            BeginSession
            ShowHit
        End If
        Exit Sub
    End If

    Dim booster As cBooster: Set booster = New cBooster
    gQuery = query
    Set gHits = CollectHits(ActiveWorkbook, query)
    Set booster = Nothing ' 画面更新を戻してから移動する

    If gHits.Count = 0 Then
        EndSession
        Application.StatusBar = "[find] No matches: " & query
        Exit Sub
    End If

    gIndex = 1
    BeginSession
    ShowHit
    Exit Sub

FINALLY:
    Application.StatusBar = "[find] E" & Err.Number & ": " & Err.Description
End Sub

Public Sub fn(): MoveBy 1: End Sub
Public Sub fp(): MoveBy -1: End Sub

Public Sub fq()
    EndSession
    Application.CutCopyMode = False ' Esc 本来の役割 (コピー範囲の解除) を代わりに行う
    Application.StatusBar = False
End Sub

'========================
' 検索モード (キー割り当て)
'========================
Private Sub BeginSession()
    Dim book As String
    book = "'" & ThisWorkbook.Name & "'!"
    Application.OnKey "n", book & "fn"
    Application.OnKey "+n", book & "fp"
    Application.OnKey "{ESC}", book & "fq"

    ' 日本語で検索した直後でも n / N が効くよう、IME をオフ (半角英数) にする
    mIme.ImeOffLater
End Sub

Private Sub EndSession()
    Application.OnKey "n"
    Application.OnKey "+n"
    Application.OnKey "{ESC}"
End Sub

'========================
' 移動
'========================
Private Sub MoveBy(ByVal delta As Long)
    If gHits Is Nothing Then
        Application.StatusBar = "[find] No search results. Run find TEXT first."
        Exit Sub
    End If
    If gHits.Count = 0 Then Exit Sub

    gIndex = ((gIndex - 1 + delta) Mod gHits.Count + gHits.Count) Mod gHits.Count + 1
    ShowHit
End Sub

Private Sub ShowHit()
    Dim hit As Variant
    Dim ws As Worksheet
    Dim shp As Shape
    Dim where As String

    On Error GoTo STALE
    hit = gHits(gIndex)

    If hit(H_KIND) = "cell" Then
        Dim cell As Range
        Set cell = hit(H_TARGET)
        Set ws = cell.Worksheet
        ws.Parent.Activate
        ws.Activate
        Application.Goto cell, False
        where = ws.Name & "!" & cell.Address(False, False)
    Else
        Set shp = hit(H_TARGET)
        Set ws = ShapeSheet(shp)
        ws.Parent.Activate
        ws.Activate
        If Application.Intersect(ActiveWindow.VisibleRange, shp.TopLeftCell) Is Nothing Then
            Application.Goto shp.TopLeftCell, True
        End If
        On Error Resume Next
        shp.TextFrame2.TextRange.Characters(hit(H_START), hit(H_LEN)).Select
        If Err.Number <> 0 Then Err.Clear: shp.Select
        On Error GoTo STALE
        where = ws.Name & "!" & shp.Name
    End If

    Application.StatusBar = "[find] " & gIndex & "/" & gHits.Count & "  " & where & _
                            "   (n: 次 / N: 前 / Esc: 終了)"
    Exit Sub

STALE:
    Application.StatusBar = "[find] The result is no longer valid. Run find again. (E" & Err.Number & ")"
End Sub

' 図形 (グループ内の図形を含む) が載っているシート
Private Function ShapeSheet(ByVal shp As Shape) As Worksheet
    Dim p As Object
    Set p = shp.Parent
    Do While Not TypeOf p Is Worksheet
        Set p = p.Parent
    Loop
    Set ShapeSheet = p
End Function

'========================
' 検索
'========================
Private Function CollectHits(ByVal wb As Workbook, ByVal query As String) As Collection
    Dim result As New Collection
    Dim ws As Worksheet
    Dim cellHits As Collection, shapeHits As Collection

    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            Set cellHits = CollectCellHits(ws, query)
            Set shapeHits = CollectShapeHits(ws, query)
            MergeByPosition cellHits, shapeHits, result
        End If
    Next ws

    Set CollectHits = result
End Function

' セルのヒットを行→列の順に集める
Private Function CollectCellHits(ByVal ws As Worksheet, ByVal query As String) As Collection
    Dim result As New Collection
    Dim used As Range
    Dim v As Variant
    Dim r As Long, c As Long
    Dim cell As Range

    Set CollectCellHits = result
    Set used = ws.UsedRange
    v = used.Value2
    If Not IsArray(v) Then
        Dim one(1 To 1, 1 To 1) As Variant
        one(1, 1) = v
        v = one
    End If

    For r = 1 To UBound(v, 1)
        For c = 1 To UBound(v, 2)
            If Not IsError(v(r, c)) Then
                If InStr(1, CStr(v(r, c)), query, vbTextCompare) > 0 Then
                    Set cell = used.Cells(r, c)
                    If Not (cell.EntireRow.Hidden Or cell.EntireColumn.Hidden) Then
                        result.Add Array("cell", cell, cell.Row, cell.Column, 0, 0)
                    End If
                End If
            End If
        Next c
    Next r
End Function

' 図形内テキストのヒットを、左上セルの行→列→テキスト内の位置の順に集める
Private Function CollectShapeHits(ByVal ws As Worksheet, ByVal query As String) As Collection
    Dim found As New Collection
    Dim shp As Shape

    For Each shp In ws.Shapes
        AddShapeHits shp, query, found
    Next shp

    Set CollectShapeHits = SortHits(found)
End Function

Private Sub AddShapeHits(ByVal shp As Shape, ByVal query As String, ByVal found As Collection)
    Dim child As Shape
    Dim text As String
    Dim pos As Long
    Dim topRow As Long, leftCol As Long

    If Not shp.Visible Then Exit Sub

    If shp.Type = msoGroup Then
        For Each child In shp.GroupItems
            AddShapeHits child, query, found
        Next child
        Exit Sub
    End If

    ' テキストを持たない図形 (画像・グラフ・フォームコントロールなど) は例外になるので飛ばす
    On Error Resume Next
    If shp.TextFrame2.HasText Then text = shp.TextFrame2.TextRange.text
    If Err.Number <> 0 Then Exit Sub
    topRow = shp.TopLeftCell.Row
    leftCol = shp.TopLeftCell.Column
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0

    pos = InStr(1, text, query, vbTextCompare)
    Do While pos > 0
        found.Add Array("shape", shp, topRow, leftCol, pos, Len(query))
        pos = InStr(pos + Len(query), text, query, vbTextCompare)
    Loop
End Sub

' (行, 列, 開始位置) の昇順に並べた新しい Collection を返す (図形は少数なので挿入ソート)
Private Function SortHits(ByVal hits As Collection) As Collection
    Dim result As New Collection
    Dim i As Long, j As Long
    Dim inserted As Boolean

    For i = 1 To hits.Count
        inserted = False
        For j = 1 To result.Count
            If ComesBefore(hits(i), result(j)) Then
                result.Add hits(i), Before:=j
                inserted = True
                Exit For
            End If
        Next j
        If Not inserted Then result.Add hits(i)
    Next i

    Set SortHits = result
End Function

Private Function ComesBefore(ByVal a As Variant, ByVal b As Variant) As Boolean
    If a(H_ROW) <> b(H_ROW) Then
        ComesBefore = a(H_ROW) < b(H_ROW)
    ElseIf a(H_COL) <> b(H_COL) Then
        ComesBefore = a(H_COL) < b(H_COL)
    Else
        ComesBefore = a(H_START) < b(H_START)
    End If
End Function

' 位置順に並んだ cellHits と shapeHits を、位置順を保って result の末尾に追加する。
' 同じ位置ではセルを先にする。
Private Sub MergeByPosition(ByVal cellHits As Collection, ByVal shapeHits As Collection, ByVal result As Collection)
    Dim i As Long, j As Long
    i = 1: j = 1
    Do While i <= cellHits.Count Or j <= shapeHits.Count
        If j > shapeHits.Count Then
            result.Add cellHits(i): i = i + 1
        ElseIf i > cellHits.Count Then
            result.Add shapeHits(j): j = j + 1
        ElseIf ComesBefore(shapeHits(j), cellHits(i)) Then
            result.Add shapeHits(j): j = j + 1
        Else
            result.Add cellHits(i): i = i + 1
        End If
    Loop
End Sub
