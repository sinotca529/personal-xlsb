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
'     引数なしの find は、前回の結果 (find または lint) のまま検索モードに戻る。
'     移動の仕組みは mHitList.bas にあり、lint (mLint.bas) と共通。
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

Private gQuery As String

'========================
' エントリポイント
'========================
Public Sub find(Optional ByVal query As String)
    On Error GoTo FINALLY

    If Len(query) = 0 Then
        If Not mHitList.ResumeSession() Then
            Application.StatusBar = "[find] Please specify the text to search."
        End If
        Exit Sub
    End If

    Dim booster As cBooster: Set booster = New cBooster
    Dim hits As Collection
    gQuery = query
    Set hits = CollectHits(ActiveWorkbook, query)
    Set booster = Nothing ' 画面更新を戻してから移動する

    If hits.Count = 0 Then
        mHitList.EndSession
        Application.StatusBar = "[find] No matches: " & query
        Exit Sub
    End If
    mHitList.Begin "find", hits
    Exit Sub

FINALLY:
    Application.StatusBar = "[find] E" & Err.Number & ": " & Err.Description
End Sub

' n / N / Esc に割り当てられる。find / lint 共通
Public Sub fn(): mHitList.MoveBy 1: End Sub
Public Sub fp(): mHitList.MoveBy -1: End Sub

Public Sub fq()
    mHitList.EndSession
    Application.CutCopyMode = False ' Esc 本来の役割 (コピー範囲の解除) を代わりに行う
    Application.StatusBar = False
End Sub

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
            mHitList.MergeByPosition cellHits, shapeHits, result
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
                        result.Add mHitList.CellHit(cell)
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

    Set CollectShapeHits = mHitList.SortHits(found)
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
        found.Add mHitList.ShapeHit(shp, topRow, leftCol, pos, Len(query))
        pos = InStr(pos + Len(query), text, query, vbTextCompare)
    Loop
End Sub
